import streamlit as st
import openpyxl
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="Scholarship Rubric", layout="wide")

col_title, col_reset = st.columns([6, 1])
col_title.title("Scholarship Application Rubric")
if col_reset.button("Reset", type="secondary", use_container_width=True):
    for key in list(st.session_state.keys()):
        if key.startswith("score_") or key == "n_terms":
            del st.session_state[key]
    st.rerun()

RUBRIC_PATH = Path(__file__).parent / "rubric_template.xlsx"

# Rows whose scores are multiplied by the N Terms weight
WEIGHTED_ROWS = {"Publications", "Other Contributions", "Conferences", "Research Awards"}


def terms_weight(n: int) -> float:
    if n <= 3:
        return 1.5
    elif n <= 5:
        return 1.25
    elif n <= 7:
        return 1.0
    elif n <= 10:
        return 0.85
    else:
        return 0.75


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


rubric = load_rubric()
scores = {}
n_terms = None
current_section = None

for section, subcat, descriptors in rubric:
    if section != current_section:
        st.divider()
        st.subheader(section)
        current_section = section

    # N Terms gets a number input instead of score buttons
    if subcat == "N Terms":
        col1, col2 = st.columns([2, 4])
        col1.write("N Terms (graduate)")
        n_terms = col2.number_input(
            "Number of completed graduate terms",
            min_value=1,
            max_value=20,
            value=6,
            step=1,
            label_visibility="collapsed",
            key="n_terms",
        )
        W = terms_weight(n_terms)
        col2.caption(
            f"Weight applied to Publications, Other Contributions, Conferences, "
            f"and Research Awards: **{W:.2f}×**"
        )
        continue

    key = f"score_{subcat}"
    if key not in st.session_state:
        st.session_state[key] = None

    cols = st.columns([2, 1, 1, 1, 1, 1])
    cols[0].write(subcat)

    for col, score_val in zip(cols[1:], [4, 3, 2, 1, 0]):
        is_selected = st.session_state[key] == score_val
        if col.button(
            str(score_val),
            key=f"btn_{subcat}_{score_val}",
            help=descriptors.get(score_val) or "",
            type="primary" if is_selected else "secondary",
            use_container_width=True,
        ):
            st.session_state[key] = score_val
            st.rerun()

    scores[subcat] = st.session_state[key]

# ── Summary ──────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Summary")

W = terms_weight(st.session_state.get("n_terms", 6))
filled = {k: v for k, v in scores.items() if v is not None}

raw_total = sum(filled.values())
weighted_total = sum(
    min(6.0, v * W) if k in WEIGHTED_ROWS else v
    for k, v in filled.items()
)

n_scored = len(filled)
n_total = len(scores)
raw_max = n_total * 4                          # 44
weighted_max = len(WEIGHTED_ROWS) * 6 + (n_total - len(WEIGHTED_ROWS)) * 4  # 52

c1, c2, c3, c4 = st.columns(4)
c1.metric("Raw Total", f"{raw_total} / {raw_max}")
c2.metric("Weighted Total", f"{weighted_total:.1f} / {weighted_max}")
c3.metric("Average Score", f"{weighted_total / n_scored:.2f}" if n_scored else "—")
c4.metric("Categories Scored", f"{n_scored} / {n_total}")

if n_scored > 0:
    st.progress(
        weighted_total / weighted_max,
        text=f"{weighted_total / weighted_max * 100:.0f}% of maximum (weighted)",
    )

# ── Export ────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Scores")

export = {}
for _, subcat, _ in rubric:
    if subcat == "N Terms":
        export["N Terms"] = st.session_state.get("n_terms", None)
    else:
        raw = scores.get(subcat)
        if subcat in WEIGHTED_ROWS:
            export[subcat] = round(min(6.0, raw * W), 1) if raw is not None else None
        else:
            export[subcat] = raw

st.dataframe(pd.DataFrame([export]), hide_index=True, use_container_width=True)
