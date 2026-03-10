import streamlit as st
import openpyxl
from pathlib import Path

st.set_page_config(page_title="Scholarship Rubric")

col_title, col_reset = st.columns([6, 1])
col_title.title("Scholarship Application Rubric")
def reset_scores():
    for key in list(st.session_state.keys()):
        if key.startswith("score_"):
            del st.session_state[key]

col_reset.button("Reset", type="secondary", use_container_width=True, on_click=reset_scores)

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

# Group rubric entries by section, preserving order
sections_order = []
sections_map = {}
for section, subcat, descriptors in rubric:
    if section not in sections_map:
        sections_order.append(section)
        sections_map[section] = []
    sections_map[section].append((subcat, descriptors))

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
c2.metric("Average Score", f"{sum(filled.values()) / n_scored:.2f}" if n_scored else "—")
c3.metric("Categories Scored", f"{n_scored} / {n_total}")

# ── Export ────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Scores")

export = {subcat: scores.get(subcat) for _, subcat, _ in rubric}

headers = ",".join(export.keys())
values = ",".join("" if v is None else str(v) for v in export.values())
st.caption("Copy/paste into Excel")
st.code(f"{values}", language=None)
st.code(f"{headers}", language=None)
