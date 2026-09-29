import streamlit as st
import openpyxl
from io import BytesIO, StringIO
import csv
from pathlib import Path

st.set_page_config(page_title="Scholarship Rubric")

WEIGHTS_STORAGE_KEY = "scholarship-rubric-weights-v1"

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


col_title, col_forget_weights, col_reset = st.columns([5, 1, 1])
col_title.title("Scholarship Application Rubric")
col_title.caption("Your weight preferences are saved in this browser.")
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


def build_export_rows(rubric, scores):
    """Create one Excel-friendly row for each rubric category."""
    rows = []
    for section, subcat, _ in rubric:
        score = scores.get(subcat)
        section_weight = st.session_state.get(f"weight_section_{section}", 1.0)
        category_weight = st.session_state.get(f"weight_cat_{subcat}", 1.0)
        rows.append(
            {
                "Section": section,
                "Category": subcat,
                "Score": "" if score is None else score,
                "Category weight": category_weight,
                "Section weight": section_weight,
                "Weighted contribution": ""
                if score is None
                else section_weight * category_weight * score,
            }
        )
    return rows


def make_csv(rows, fieldnames):
    buffer = StringIO(newline="")
    writer = csv.DictWriter(buffer, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(rows)
    # UTF-8 BOM lets desktop Excel recognize names with accented characters.
    return buffer.getvalue().encode("utf-8-sig")


def make_workbook(rows, weighted_total, average_score, n_scored, n_total):
    workbook = openpyxl.Workbook()
    scores_sheet = workbook.active
    scores_sheet.title = "Scores"

    headers = list(rows[0].keys()) if rows else []
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

    summary_sheet = workbook.create_sheet("Summary")
    summary_sheet.append(["Metric", "Value"])
    summary_sheet.append(["Weighted score", weighted_total])
    summary_sheet.append(["Average score", average_score if average_score is not None else ""])
    summary_sheet.append(["Categories scored", n_scored])
    summary_sheet.append(["Categories available", n_total])
    for cell in summary_sheet[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    summary_sheet.column_dimensions["A"].width = 24
    summary_sheet.column_dimensions["B"].width = 18

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

# ── Export ────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Export scores")

export_rows = build_export_rows(rubric, scores)
export_headers = list(export_rows[0].keys()) if export_rows else []

# Tabs are the native column separator for a direct paste into Excel or Sheets.
tsv_buffer = StringIO(newline="")
tsv_writer = csv.DictWriter(
    tsv_buffer, fieldnames=export_headers, delimiter="\t", lineterminator="\n"
)
tsv_writer.writeheader()
tsv_writer.writerows(export_rows)
tab_separated_scores = tsv_buffer.getvalue()

st.caption(
    "Copy the complete block below (including its header row), then paste into cell A1 "
    "in Excel or Google Sheets. Tabs place each field in its own column."
)
st.code(tab_separated_scores, language=None)

csv_bytes = make_csv(export_rows, export_headers)
workbook_bytes = make_workbook(
    export_rows, weighted_total, average_score, n_scored, n_total
)
download_csv, download_xlsx = st.columns(2)
download_csv.download_button(
    "Download CSV",
    data=csv_bytes,
    file_name="scholarship-rubric-scores.csv",
    mime="text/csv",
    use_container_width=True,
)
download_xlsx.download_button(
    "Download Excel workbook",
    data=workbook_bytes,
    file_name="scholarship-rubric-scores.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
