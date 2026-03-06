import streamlit as st
import openpyxl
from pathlib import Path

st.set_page_config(page_title="Scholarship Rubric")

col_title, col_reset = st.columns([6, 1])
col_title.title("Scholarship Application Rubric")
if col_reset.button("Reset", type="secondary", use_container_width=True):
    for key in list(st.session_state.keys()):
        if key.startswith("score_"):
            del st.session_state[key]
    st.rerun()

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


rubric = load_rubric()
scores = {}
current_section = None

for section, subcat, descriptors in rubric:
    if section != current_section:
        st.divider()
        st.subheader(section)
        current_section = section

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

filled = {k: v for k, v in scores.items() if v is not None}
total = sum(filled.values())
n_scored = len(filled)
n_total = len(scores)
max_score = n_total * 4

c1, c2, c3 = st.columns(3)
c1.metric("Total", f"{total} / {max_score}")
c2.metric("Average Score", f"{total / n_scored:.2f}" if n_scored else "—")
c3.metric("Categories Scored", f"{n_scored} / {n_total}")

if n_scored > 0:
    st.progress(
        total / max_score,
        text=f"{total / max_score * 100:.0f}% of maximum",
    )

# ── Export ────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Scores")

export = {subcat: scores.get(subcat) for _, subcat, _ in rubric}

headers = ",".join(export.keys())
values = ",".join("" if v is None else str(v) for v in export.values())
st.caption("Copy/paste into Excel")
st.code(f"{values}", language=None)
st.code(f"{headers}", language=None)
