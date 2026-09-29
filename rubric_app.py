import streamlit as st
import openpyxl
from io import BytesIO, StringIO
import csv
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="Scholarship Rubric")

WEIGHTS_STORAGE_KEY = "scholarship-rubric-weights-v1"
EVALUATIONS_STORAGE_KEY = "scholarship-rubric-evaluations-v1"

# This trusted, app-owned component keeps preferences on the visitor's device.
# It contains no user-provided content and sends only the saved weight values to
# Streamlit so the widgets can be initialized after a browser reload.
WEIGHT_STORAGE_COMPONENT = st.components.v2.component(
    "rubric_weight_storage",
    js="""
    export default function({ data, setStateValue }) {
        const storageKey = data.storageKey;
        try {
            if (data.clear) {
                window.localStorage.removeItem(storageKey);
                setStateValue("savedWeights", {});
                return;
            }

            if (!data.hydrated) {
                const savedWeights = JSON.parse(
                    window.localStorage.getItem(storageKey) || "{}"
                );
                setStateValue("savedWeights", savedWeights);
                return;
            }

            window.localStorage.setItem(storageKey, JSON.stringify(data.weights));
        } catch (error) {
            // Storage can be unavailable in private browsing or restrictive
            // browser settings. The rubric still works for the current session.
            setStateValue("savedWeights", {});
        }
    }
    """,
)

# Evaluations are also kept only in the visitor's browser. A record is replaced
# when its student key matches an existing record, allowing corrections without
# duplicate rows in the downloaded file.
EVALUATION_STORAGE_COMPONENT = st.components.v2.component(
    "rubric_evaluation_storage",
    js="""
    export default function({ data, setStateValue }) {
        const storageKey = data.storageKey;
        try {
            if (data.clear) {
                window.localStorage.removeItem(storageKey);
                setStateValue("records", []);
                return;
            }

            if (data.record) {
                const records = JSON.parse(
                    window.localStorage.getItem(storageKey) || "[]"
                );
                const safeRecords = Array.isArray(records) ? records : [];
                const existingIndex = safeRecords.findIndex(
                    (item) => item.studentKey === data.record.studentKey
                );
                if (existingIndex === -1) {
                    safeRecords.push(data.record);
                } else {
                    safeRecords[existingIndex] = data.record;
                }
                window.localStorage.setItem(storageKey, JSON.stringify(safeRecords));
                setStateValue("records", safeRecords);
                return;
            }

            if (!data.hydrated) {
                const records = JSON.parse(
                    window.localStorage.getItem(storageKey) || "[]"
                );
                setStateValue("records", Array.isArray(records) ? records : []);
            }
        } catch (error) {
            // The rubric remains usable if localStorage is blocked by browser
            // privacy settings; its results simply cannot persist after reload.
            setStateValue("records", []);
        }
    }
    """,
)


def reset_scores():
    for key in list(st.session_state.keys()):
        if key.startswith("score_"):
            del st.session_state[key]


def forget_saved_weights():
    for key in list(st.session_state.keys()):
        if key.startswith("weight_"):
            del st.session_state[key]
    st.session_state["clear_saved_weights"] = True
    st.session_state["weights_restored"] = False


def mark_evaluation_saved():
    if st.session_state.pop("pending_evaluation", None) is not None:
        st.session_state["evaluation_save_notice"] = True


def clear_saved_evaluations():
    st.session_state["clear_saved_evaluations"] = True
    st.session_state["evaluations_restored"] = False


col_title, col_forget_weights, col_reset = st.columns([5, 1, 1])
col_title.title("Scholarship Application Rubric")
col_title.caption("Weight preferences and saved evaluations stay in this browser.")
col_forget_weights.button(
    "Forget weights",
    type="secondary",
    use_container_width=True,
    on_click=forget_saved_weights,
)
col_reset.button(
    "Reset scores",
    type="secondary",
    use_container_width=True,
    on_click=reset_scores,
)
student_name = st.text_input(
    "Student name or ID",
    key="student_name",
    placeholder="Enter a unique student name or identifier",
)

RUBRIC_PATH = Path(__file__).parent / "rubric_template.xlsx"


@st.cache_data
def load_rubric():
    wb = openpyxl.load_workbook(RUBRIC_PATH, data_only=True)
    ws = wb["Sheet1"]
    rubric = []
    current_section = None
    for row in ws.iter_rows(values_only=True):
        if row[0] is not None:
            current_section = row[0]
        elif row[1] is not None and current_section:
            descriptors = {4: row[2], 3: row[3], 2: row[4], 1: row[5], 0: row[6]}
            rubric.append((current_section, row[1], descriptors))
    return rubric


def build_evaluation_record(student, scores, weighted_total, average_score, n_scored):
    """Create one browser-stored, wide-format evaluation record."""
    return {
        "student": student.strip(),
        "studentKey": student.strip().casefold(),
        "scores": scores,
        "weightedTotal": weighted_total,
        "averageScore": average_score,
        "categoriesScored": n_scored,
        "savedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
    }


def build_output_rows(records, metrics):
    """Turn stored evaluation records into rows for CSV and Excel downloads."""
    rows = []
    for record in records:
        if not isinstance(record, dict) or not isinstance(record.get("scores"), dict):
            continue
        row = {"Student": record.get("student", "")}
        row.update({metric: record["scores"].get(metric, "") for metric in metrics})
        row.update(
            {
                "Weighted Score": record.get("weightedTotal", ""),
                "Average Score": record.get("averageScore", ""),
                "Categories Scored": record.get("categoriesScored", ""),
                "Saved At": record.get("savedAt", ""),
            }
        )
        rows.append(row)
    return rows


def make_csv(rows, fieldnames):
    buffer = StringIO(newline="")
    writer = csv.DictWriter(buffer, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(rows)
    # UTF-8 BOM lets desktop Excel recognize names with accented characters.
    return buffer.getvalue().encode("utf-8-sig")


def make_workbook(rows, headers):
    workbook = openpyxl.Workbook()
    scores_sheet = workbook.active
    scores_sheet.title = "Evaluations"

    scores_sheet.append(headers)
    for row in rows:
        scores_sheet.append([row[header] for header in headers])
    scores_sheet.freeze_panes = "A2"
    scores_sheet.auto_filter.ref = scores_sheet.dimensions
    for cell in scores_sheet[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    for column_cells in scores_sheet.columns:
        width = max(len(str(cell.value or "")) for cell in column_cells)
        scores_sheet.column_dimensions[column_cells[0].column_letter].width = min(width + 2, 40)

    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def restore_saved_weights(saved_weights, allowed_keys):
    """Restore valid, known weights from browser storage into this session."""
    if not isinstance(saved_weights, dict):
        return

    for key in allowed_keys:
        value = saved_weights.get(key)
        if isinstance(value, (int, float)) and not isinstance(value, bool) and value >= 0:
            st.session_state[key] = float(value)


rubric = load_rubric()

# Group rubric entries by section, preserving order
sections_order = []
sections_map = {}
for section, subcat, descriptors in rubric:
    if section not in sections_map:
        sections_order.append(section)
        sections_map[section] = []
    sections_map[section].append((subcat, descriptors))

# Streamlit session state is tied to the current browser connection. Hydrate it
# once from localStorage before any weight widget is created, then save the
# current values to localStorage on each subsequent rerun.
weight_keys = {
    f"weight_section_{section}" for section in sections_order
} | {
    f"weight_cat_{subcat}" for _, subcat, _ in rubric
}
saved_weights = st.session_state.get("rubric_weight_storage", {}).get("savedWeights")
if saved_weights is not None and not st.session_state.get("weights_restored", False):
    restore_saved_weights(saved_weights, weight_keys)
    st.session_state["weights_restored"] = True

clear_saved_weights = st.session_state.pop("clear_saved_weights", False)
WEIGHT_STORAGE_COMPONENT(
    data={
        "storageKey": WEIGHTS_STORAGE_KEY,
        "clear": clear_saved_weights,
        "hydrated": st.session_state.get("weights_restored", False),
        "weights": {key: st.session_state.get(key, 1.0) for key in weight_keys},
    },
    default={"savedWeights": None},
    key="rubric_weight_storage",
    on_savedWeights_change=lambda: None,
    height=0,
)

scores = {}

for section in sections_order:
    st.divider()
    sec_col, sec_weight_label, sec_weight_col = st.columns([4, 1, 1])
    sec_col.subheader(section)
    sec_weight_label.markdown("<div style='padding-top:0.6rem; text-align:right'>Section weight:</div>", unsafe_allow_html=True)
    sec_weight_key = f"weight_section_{section}"
    if sec_weight_key not in st.session_state:
        st.session_state[sec_weight_key] = 1.0
    w_section = sec_weight_col.number_input(
        "Section weight",
        min_value=0.0,
        step=0.1,
        key=sec_weight_key,
        label_visibility="collapsed",
    )

    cat_weights = []
    for subcat, descriptors in sections_map[section]:
        score_key = f"score_{subcat}"
        weight_key = f"weight_cat_{subcat}"
        if score_key not in st.session_state:
            st.session_state[score_key] = None
        if weight_key not in st.session_state:
            st.session_state[weight_key] = 1.0

        cols = st.columns([2, 1, 1, 1, 1, 1, 1])
        cols[0].write(subcat)

        for col, score_val in zip(cols[1:6], [4, 3, 2, 1, 0]):
            is_selected = st.session_state[score_key] == score_val
            col.button(
                str(score_val),
                key=f"btn_{subcat}_{score_val}",
                help=descriptors.get(score_val) or "",
                type="primary" if is_selected else "secondary",
                use_container_width=True,
                on_click=lambda k=score_key, v=score_val: st.session_state.update({k: v}),
            )

        w_cat = cols[6].number_input(
            "Category weight",
            min_value=0.0,
            step=0.1,
            key=weight_key,
            label_visibility="collapsed",
        )
        cat_weights.append(w_cat)

        scores[subcat] = st.session_state[score_key]

    # Weight sum row for this section
    cat_weight_sum = sum(cat_weights)
    sum_cols = st.columns([2, 1, 1, 1, 1, 1, 1])
    color = "green" if abs(cat_weight_sum - 1.0) < 0.001 else "orange"
    sum_cols[0].markdown(
        "<div style='text-align:right; color:gray; font-size:0.85rem'>category weights sum:</div>",
        unsafe_allow_html=True,
    )
    sum_cols[6].markdown(
        f"<div style='text-align:center; color:{color}; font-weight:bold; font-size:0.85rem'>{cat_weight_sum:.2f}</div>",
        unsafe_allow_html=True,
    )

# ── Summary ──────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Summary")

filled = {k: v for k, v in scores.items() if v is not None}
n_scored = len(filled)
n_total = len(scores)
average_score = sum(filled.values()) / n_scored if n_scored else None

# Weighted score: sum_s( W_s * sum_c( W_c * score_c ) )
weighted_total = 0.0
for section in sections_order:
    w_section = st.session_state.get(f"weight_section_{section}", 1.0)
    section_score = 0.0
    for subcat, _ in sections_map[section]:
        score = scores.get(subcat)
        if score is not None:
            w_cat = st.session_state.get(f"weight_cat_{subcat}", 1.0)
            section_score += w_cat * score
    weighted_total += w_section * section_score

c1, c2, c3 = st.columns(3)
c1.metric("Weighted Score", f"{weighted_total:.2f}")
c2.metric("Average Score", f"{average_score:.2f}" if average_score is not None else "—")
c3.metric("Categories Scored", f"{n_scored} / {n_total}")

# ── Saved evaluation output ───────────────────────────────────────────────────
st.divider()
st.subheader("Saved evaluations")
st.caption(
    "Save one row per student, then download the accumulated results. Records stay "
    "in this browser until you clear them or clear this site's browser data."
)

metric_columns = [subcat for _, subcat, _ in rubric]
output_headers = [
    "Student",
    *metric_columns,
    "Weighted Score",
    "Average Score",
    "Categories Scored",
    "Saved At",
]
stored_records = st.session_state.get("rubric_evaluation_storage", {}).get("records")
if stored_records is not None and not st.session_state.get("evaluations_restored", False):
    st.session_state["evaluations_restored"] = True
if not isinstance(stored_records, list):
    stored_records = []

if st.session_state.pop("evaluation_save_notice", False):
    st.success("Evaluation saved in this browser. Saving the same student again updates their row.")

save_col, clear_col = st.columns(2)
if save_col.button("Save current evaluation", type="primary", use_container_width=True):
    if not student_name.strip():
        st.warning("Enter a student name or ID before saving an evaluation.")
    else:
        st.session_state["pending_evaluation"] = build_evaluation_record(
            student_name, scores, weighted_total, average_score, n_scored
        )
if clear_col.button(
    "Clear saved evaluations", type="secondary", use_container_width=True
):
    clear_saved_evaluations()

clear_saved_evaluations_request = st.session_state.pop("clear_saved_evaluations", False)
EVALUATION_STORAGE_COMPONENT(
    data={
        "storageKey": EVALUATIONS_STORAGE_KEY,
        "clear": clear_saved_evaluations_request,
        "hydrated": st.session_state.get("evaluations_restored", False),
        "record": st.session_state.get("pending_evaluation"),
    },
    default={"records": None},
    key="rubric_evaluation_storage",
    on_records_change=mark_evaluation_saved,
    height=0,
)

output_rows = build_output_rows(stored_records, metric_columns)
st.caption(f"{len(output_rows)} saved evaluation{'s' if len(output_rows) != 1 else ''}")
if output_rows:
    st.dataframe(output_rows, use_container_width=True, hide_index=True)

    csv_bytes = make_csv(output_rows, output_headers)
    workbook_bytes = make_workbook(output_rows, output_headers)
    download_csv, download_xlsx = st.columns(2)
    download_csv.download_button(
        "Download CSV",
        data=csv_bytes,
        file_name="scholarship-evaluations.csv",
        mime="text/csv",
        use_container_width=True,
    )
    download_xlsx.download_button(
        "Download Excel workbook",
        data=workbook_bytes,
        file_name="scholarship-evaluations.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
