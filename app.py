    return f"{float(value):,.2f}"


def boq_grid_df() -> pd.DataFrame:
    work = st.session_state.boq_df.copy()
    work.insert(0, "Open", "")
    if st.session_state.selected_article:
        mask = work["Article_ID"].astype(str) == str(st.session_state.selected_article)
        if mask.any():
            work.loc[mask, "Open"] = "+"
    return work[BOQ_COLUMNS]


def article_summary(article_id: str) -> tuple[float, float]:
    row = st.session_state.boq_df[st.session_state.boq_df["Article_ID"].astype(str) == str(article_id)]
    if row.empty:
        return 0.0, 0.0
    quantity = row.iloc[0]["Quantity"]
    total = row.iloc[0]["Total Price"]
    total_value = float(pd.to_numeric(total, errors="coerce") or 0.0)
    qty_value = float(pd.to_numeric(quantity, errors="coerce") or 0.0)
    unit_rate = total_value / qty_value if qty_value else 0.0
    return total_value, unit_rate


def article_breakdown_count(article_id: str) -> int:
    return len(get_breakdown(article_id))


init_state()

st.markdown(
    """
    <style>
    .excel-shell {
        border: 1px solid #d8dee9;
        border-radius: 14px;
        background: #ffffff;
        overflow: hidden;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
    }
    .toolbar {
        background: linear-gradient(90deg, #eef4ff 0%, #f7fbff 100%);
        border-bottom: 1px solid #d8dee9;
        padding: 12px 14px;
    }
    .sheet-title {
        font-size: 1.25rem;
        font-weight: 700;
        margin-bottom: 4px;
    }
    .sheet-subtitle {
        color: #5b6472;
        font-size: 0.92rem;
    }
    .panel {
        border: 1px solid #d8dee9;
        border-radius: 14px;
        background: #ffffff;
        padding: 16px;
        box-shadow: 0 6px 20px rgba(15, 23, 42, 0.05);
    }
    .active-article {
        background: #fff8e6;
        border: 1px solid #f3d47b;
        border-radius: 12px;
        padding: 10px 12px;
        margin-bottom: 10px;
    }
    .small-note {
        color: #5b6472;
        font-size: 0.88rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div class='excel-shell'>", unsafe_allow_html=True)
st.markdown("<div class='toolbar'>", unsafe_allow_html=True)
st.markdown("<div class='sheet-title'>BOQ Breakdown Calculator</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='sheet-subtitle'>Excel-style BOQ table with import, export, edit, save, and article breakdown opening from the + column.</div>",
    unsafe_allow_html=True,
)

toolbar_cols = st.columns([1.3, 1.3, 1.0, 1.0, 1.4, 1.4, 1.1, 3.5])
boq_file = toolbar_cols[0].file_uploader("Import BOQ", type=["xlsx"], label_visibility="collapsed", key="boq_upload")
lib_file = toolbar_cols[1].file_uploader("Import Library", type=["xlsx"], label_visibility="collapsed", key="lib_upload")

if toolbar_cols[2].button("Save", use_container_width=True):
    st.session_state.last_saved_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.save_message = f"Draft saved in session at {st.session_state.last_saved_at}"

if toolbar_cols[3].button("Edit On" if st.session_state.edit_mode else "Edit Off", use_container_width=True):
    st.session_state.edit_mode = not st.session_state.edit_mode

run_clicked = toolbar_cols[4].button("Run Calculation", type="primary", use_container_width=True)
load_sample_clicked = toolbar_cols[5].button("Load Sample", use_container_width=True)
add_row_clicked = toolbar_cols[6].button("Add Row", use_container_width=True)

toolbar_cols[7].download_button(
    "Export Excel",
    data=export_excel(),
    file_name="boq_breakdown_result.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.markdown("</div>", unsafe_allow_html=True)

if boq_file is not None or lib_file is not None:
    action_cols = st.columns([1.2, 5])
    if action_cols[0].button("Load Imported Files", use_container_width=True):
        try:
            if boq_file is not None:
                st.session_state.boq_df = normalize_boq_columns(pd.read_excel(boq_file))
                st.session_state.breakdowns = {}
                st.session_state.selected_article = None
            if lib_file is not None:
                st.session_state.library_df = normalize_library_columns(pd.read_excel(lib_file))
            st.success("Imported files loaded into the sheet.")
        except Exception as exc:
            st.error(str(exc))

if load_sample_clicked:
    st.session_state.boq_df = DEFAULT_BOQ.copy()
    st.session_state.library_df = DEFAULT_LIBRARY.copy()
    st.session_state.breakdowns = {}
    st.session_state.selected_article = None
    st.success("Sample BOQ and library loaded.")

if add_row_clicked:
    new_row = {
        "Article_ID": f"A{len(st.session_state.boq_df) + 1:03d}",
        "Description": "",
        "Quantity": 0.0,
        "Unit": "",
        "Unit Price": None,
        "Total Price": None,
        "Template_Name": "",
    }
    st.session_state.boq_df = pd.concat([st.session_state.boq_df, pd.DataFrame([new_row])], ignore_index=True)

if run_clicked:
    errors = run_all()
    if errors:
        st.warning("Calculation finished with some errors.")
        for error in errors[:10]:
            st.write(error)
    else:
        st.success("Calculation finished.")

if st.session_state.save_message:
    st.info(st.session_state.save_message)

st.markdown("<div style='padding: 16px;'>", unsafe_allow_html=True)
top_info_cols = st.columns([2.2, 1.3, 1.3, 2.2])
top_info_cols[0].markdown("**Main BOQ Sheet**")
top_info_cols[1].metric("Articles", len(st.session_state.boq_df))
filled_prices = pd.to_numeric(st.session_state.boq_df["Unit Price"], errors="coerce").notna().sum()
top_info_cols[2].metric("Priced Articles", int(filled_prices))
top_info_cols[3].markdown(
    f"<div class='small-note'>Edit mode: <b>{'ON' if st.session_state.edit_mode else 'OFF'}</b>. Type <b>+</b> in one Open cell to open that article's own breakdown sheet.</div>",
    unsafe_allow_html=True,
)

sheet_df = boq_grid_df()
edited_boq = st.data_editor(
    sheet_df,
    num_rows="dynamic" if st.session_state.edit_mode else "fixed",
    use_container_width=True,
    hide_index=True,
    key="boq_sheet",
    disabled=not st.session_state.edit_mode,
    column_config={
        "Open": st.column_config.TextColumn("Open", help="Type + in one row to open that article breakdown."),
        "Quantity": st.column_config.NumberColumn("Quantity", format="%.3f"),
        "Unit Price": st.column_config.NumberColumn("Unit Price", format="%.2f", disabled=True),
        "Total Price": st.column_config.NumberColumn("Total Price", format="%.2f", disabled=True),
        "Template_Name": st.column_config.SelectboxColumn(
            "Template_Name",
            options=sorted(st.session_state.library_df["Template_Name"].dropna().astype(str).unique().tolist()),
        ),
    },
)

open_rows = edited_boq[edited_boq["Open"].astype(str).str.strip() == "+"]
if not open_rows.empty:
    st.session_state.selected_article = str(open_rows.iloc[0]["Article_ID"])

st.session_state.boq_df = edited_boq[BOQ_STORAGE_COLUMNS].copy()

selected_article = st.session_state.selected_article
if selected_article:
    article_row = st.session_state.boq_df[
        st.session_state.boq_df["Article_ID"].astype(str) == str(selected_article)
    ]
    if not article_row.empty:
        article = article_row.iloc[0]
        ensure_article_breakdown(selected_article, article["Template_Name"])
        article_total, unit_rate = article_summary(selected_article)
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown(
            f"<div class='active-article'><b>Active article:</b> {selected_article} | {article['Description']} | Quantity: {article['Quantity']} {article['Unit']}<br><span class='small-note'>This breakdown belongs only to article {selected_article}. Editing it will not change other articles.</span></div>",
            unsafe_allow_html=True,
        )

        summary_cols = st.columns([1.1, 1.1, 1.6, 1.0, 1.0, 1.0])
        summary_cols[0].metric("Unit Price", fmt_money(unit_rate))
        summary_cols[1].metric("Total Price", fmt_money(article_total))
        summary_cols[2].write(f"Template: `{article['Template_Name'] or '-'}`")
        summary_cols[3].metric("Breakdown Rows", article_breakdown_count(selected_article))

        if summary_cols[4].button("Reload Template", use_container_width=True):
            if str(article["Template_Name"]).strip():
                load_template(selected_article, article["Template_Name"])
                st.success(f"Template copied again into article {selected_article}.")
            else:
                st.warning("Select a template name in the BOQ sheet first.")

        if summary_cols[5].button("Close", use_container_width=True):
            st.session_state.selected_article = None
            st.rerun()

        action_cols = st.columns([1.0, 1.2, 4.0])
        if action_cols[0].button("Add Breakdown Row", use_container_width=True):
            breakdown = get_breakdown(selected_article)
            new_row = {column: None for column in BREAKDOWN_COLUMNS}
            new_row["Type"] = "D"
            new_row["Category"] = "General"
            new_row["Norm"] = "F"
            new_row["Formula"] = "1"
            new_row["Unit Price"] = 0.0
            breakdown = pd.concat([breakdown, pd.DataFrame([new_row])], ignore_index=True)
            set_breakdown(selected_article, breakdown)

        if action_cols[1].button("Calculate Article", use_container_width=True):
            result_df, unit_price, total_price, errors = calculate_article(article, get_breakdown(selected_article))
            set_breakdown(selected_article, result_df)
            row_index = article_row.index[0]
            st.session_state.boq_df.at[row_index, "Unit Price"] = unit_price
            st.session_state.boq_df.at[row_index, "Total Price"] = total_price
            if errors:
                st.warning("Calculation finished with some errors.")
                for error in errors[:5]:
                    st.write(error)
            else:
                st.success("Selected article calculated.")
