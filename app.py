import io
import math
from datetime import datetime

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, DataReturnMode, GridOptionsBuilder, GridUpdateMode, JsCode

st.set_page_config(page_title="BOQ Breakdown Calculator", layout="wide")

BOQ_COLUMNS = [
    "Article_ID",
    "Description",
    "Quantity",
    "Unit",
    "Unit Price",
    "Total Price",
    "Template_Name",
]

BREAKDOWN_COLUMNS = [
    "Type",
    "Level",
    "Category",
    "Code",
    "Description",
    "Norm",
    "Formula",
    "Resultant",
    "Quantity",
    "Unit",
    "Unit Price",
    "Total Cost",
]

DEFAULT_BOQ = pd.DataFrame(
    [
        {
            "Article_ID": "A001",
            "Description": "Foundation footing",
            "Quantity": 100.0,
            "Unit": "m3",
            "Unit Price": None,
            "Total Price": None,
            "Template_Name": "FOUNDATION_FOOTING",
        },
        {
            "Article_ID": "A002",
            "Description": "Column concrete",
            "Quantity": 45.0,
            "Unit": "m3",
            "Unit Price": None,
            "Total Price": None,
            "Template_Name": "FOUNDATION_FOOTING",
        },
    ]
)[BOQ_COLUMNS]

DEFAULT_LIBRARY = pd.DataFrame(
    [
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "O",
            "Level": 0,
            "Category": "Main",
            "Code": "O-001",
            "Description": "FOUNDATION FOOTING",
            "Norm": "C",
            "Formula": "1",
            "Resultant": None,
            "Quantity": None,
            "Unit": "m3",
            "Unit Price": None,
            "Total Cost": None,
        },
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "S",
            "Level": 1,
            "Category": "Concrete",
            "Code": "S-001",
            "Description": "Concrete works",
            "Norm": "N",
            "Formula": "1",
            "Resultant": None,
            "Quantity": None,
            "Unit": "m3",
            "Unit Price": None,
            "Total Cost": None,
        },
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "M",
            "Level": 2,
            "Category": "Concrete",
            "Code": "0126120109",
            "Description": "Supply concrete C30/37",
            "Norm": "N",
            "Formula": "1.05",
            "Resultant": None,
            "Quantity": None,
            "Unit": "m3",
            "Unit Price": 132.50,
            "Total Cost": None,
        },
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "M",
            "Level": 2,
            "Category": "Concrete",
            "Code": "0490000411",
            "Description": "Concrete labor",
            "Norm": "C",
            "Formula": "30/60",
            "Resultant": None,
            "Quantity": None,
            "Unit": "hr",
            "Unit Price": 45.00,
            "Total Cost": None,
        },
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "S",
            "Level": 1,
            "Category": "Pump",
            "Code": "S-002",
            "Description": "Pump works",
            "Norm": "N",
            "Formula": "1.05",
            "Resultant": None,
            "Quantity": None,
            "Unit": "m3",
            "Unit Price": None,
            "Total Cost": None,
        },
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "M",
            "Level": 2,
            "Category": "Pump",
            "Code": "0126120308",
            "Description": "Concrete pumping 36m",
            "Norm": "N",
            "Formula": "1.05",
            "Resultant": None,
            "Quantity": None,
            "Unit": "m3",
            "Unit Price": 6.50,
            "Total Cost": None,
        },
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "S",
            "Level": 1,
            "Category": "Steel",
            "Code": "S-003",
            "Description": "Reinforcement works",
            "Norm": "N",
            "Formula": "80",
            "Resultant": None,
            "Quantity": None,
            "Unit": "kg",
            "Unit Price": None,
            "Total Cost": None,
        },
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "M",
            "Level": 2,
            "Category": "Steel",
            "Code": "0126111101",
            "Description": "Reinforcement bars",
            "Norm": "N",
            "Formula": "1",
            "Resultant": None,
            "Quantity": None,
            "Unit": "kg",
            "Unit Price": 0.90,
            "Total Cost": None,
        },
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "S",
            "Level": 1,
            "Category": "Formwork",
            "Code": "S-004",
            "Description": "Formwork works",
            "Norm": "N",
            "Formula": "6.67",
            "Resultant": None,
            "Quantity": None,
            "Unit": "m2",
            "Unit Price": None,
            "Total Cost": None,
        },
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "M",
            "Level": 2,
            "Category": "Formwork",
            "Code": "0126131003",
            "Description": "Formwork footing",
            "Norm": "N",
            "Formula": "1.05",
            "Resultant": None,
            "Quantity": None,
            "Unit": "m2",
            "Unit Price": 15.00,
            "Total Cost": None,
        },
    ]
)

SAFE_GLOBALS = {
    "__builtins__": {},
    "min": min,
    "max": max,
    "abs": abs,
    "round": round,
    "ceil": math.ceil,
    "floor": math.floor,
}


def normalize_type_value(value) -> str:
    value = str(value).strip().upper()
    return {"T": "O", "D": "M"}.get(value, value)


def normalize_norm_value(value) -> str:
    return str(value).strip().upper()


def init_state():
    if "boq_df" not in st.session_state:
        st.session_state.boq_df = DEFAULT_BOQ.copy()
    if "library_df" not in st.session_state:
        st.session_state.library_df = DEFAULT_LIBRARY.copy()
    if "breakdowns" not in st.session_state:
        st.session_state.breakdowns = {}
    if "selected_article" not in st.session_state:
        st.session_state.selected_article = None
    if "edit_mode" not in st.session_state:
        st.session_state.edit_mode = True
    if "save_message" not in st.session_state:
        st.session_state.save_message = ""


def empty_breakdown_df():
    return pd.DataFrame(columns=BREAKDOWN_COLUMNS)


def normalize_boq_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "article_id": "Article_ID",
        "description": "Description",
        "quantity": "Quantity",
        "unit": "Unit",
        "unit price": "Unit Price",
        "unit_price": "Unit Price",
        "total price": "Total Price",
        "total_price": "Total Price",
        "template_name": "Template_Name",
    }
    work = df.copy()
    work.columns = [rename_map.get(str(c).strip().lower(), c) for c in work.columns]

    for column in ["Article_ID", "Description", "Quantity", "Unit"]:
        if column not in work.columns:
            raise ValueError(f"BOQ file is missing required column: {column}")

    if "Unit Price" not in work.columns:
        work["Unit Price"] = None
    if "Total Price" not in work.columns:
        work["Total Price"] = None
    if "Template_Name" not in work.columns:
        work["Template_Name"] = ""

    return work[BOQ_COLUMNS].copy()


def normalize_library_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "template_name": "Template_Name",
        "type": "Type",
        "level": "Level",
        "category": "Category",
        "code": "Code",
        "description": "Description",
        "describtion": "Description",
        "norm": "Norm",
        "formula": "Formula",
        "resultant": "Resultant",
        "quantity": "Quantity",
        "unit": "Unit",
        "unit price": "Unit Price",
        "unit_price": "Unit Price",
        "total cost": "Total Cost",
        "total_cost": "Total Cost",
    }
    work = df.copy()
    work.columns = [rename_map.get(str(c).strip().lower(), c) for c in work.columns]

    required = ["Template_Name", "Type", "Category", "Code", "Description", "Norm", "Formula", "Unit"]
    for column in required:
        if column not in work.columns:
            raise ValueError(f"Library file is missing required column: {column}")

    defaults = {
        "Level": 0,
        "Resultant": None,
        "Quantity": None,
        "Unit Price": None,
        "Total Cost": None,
    }
    for column, value in defaults.items():
        if column not in work.columns:
            work[column] = value

    work["Type"] = work["Type"].apply(normalize_type_value)
    work["Norm"] = work["Norm"].apply(normalize_norm_value)
    work["Level"] = pd.to_numeric(work["Level"], errors="coerce").fillna(0).astype(int)

    return work[["Template_Name", *BREAKDOWN_COLUMNS]].copy()


def get_breakdown(article_id: str) -> pd.DataFrame:
    if article_id not in st.session_state.breakdowns:
        st.session_state.breakdowns[article_id] = empty_breakdown_df()
    return st.session_state.breakdowns[article_id].copy()


def set_breakdown(article_id: str, df: pd.DataFrame):
    work = df.copy()
    defaults = {
        "Type": "M",
        "Level": 2,
        "Category": "",
        "Code": "",
        "Description": "",
        "Norm": "N",
        "Formula": "1",
        "Resultant": None,
        "Quantity": None,
        "Unit": "",
        "Unit Price": 0.0,
        "Total Cost": None,
    }
    for column in BREAKDOWN_COLUMNS:
        if column not in work.columns:
            work[column] = defaults[column]

    work["Type"] = work["Type"].apply(normalize_type_value)
    work["Norm"] = work["Norm"].apply(normalize_norm_value)
    work["Level"] = pd.to_numeric(work["Level"], errors="coerce").fillna(0).astype(int)

    st.session_state.breakdowns[article_id] = work[BREAKDOWN_COLUMNS].copy()


def load_template(article_id: str, template_name: str):
    library = st.session_state.library_df
    rows = library[library["Template_Name"].astype(str) == str(template_name)].copy()
    if rows.empty:
        st.warning(f"Template '{template_name}' was not found.")
        return
    set_breakdown(article_id, rows[BREAKDOWN_COLUMNS].copy())


def ensure_article_breakdown(article_id: str, template_name: str):
    if not get_breakdown(article_id).empty:
        return
    if str(template_name).strip():
        load_template(article_id, template_name)


def eval_formula(formula_text: str) -> float:
    text = str(formula_text).strip()
    if not text:
        return 0.0
    return float(eval(text, SAFE_GLOBALS, {}))


def calculate_article(article_row: pd.Series, breakdown_df: pd.DataFrame):
    work = breakdown_df.copy().reset_index(drop=True)
    if work.empty:
        return work, 0.0, 0.0, []

    boq_qty = float(pd.to_numeric(article_row["Quantity"], errors="coerce") or 0.0)
    errors = []
    current_subgroup_qty = boq_qty
    current_subgroup_start = None
    current_overall_start = None

    def close_subgroup(end_idx: int):
        nonlocal current_subgroup_start
        if current_subgroup_start is None:
            return
        subtotal = pd.to_numeric(
            work.loc[current_subgroup_start + 1:end_idx, "Total Cost"], errors="coerce"
        ).fillna(0.0).sum()
        work.at[current_subgroup_start, "Unit Price"] = None
        work.at[current_subgroup_start, "Total Cost"] = subtotal

    def close_overall(end_idx: int):
        nonlocal current_overall_start
        if current_overall_start is None:
            return
        subtotal = pd.to_numeric(
            work.loc[current_overall_start + 1:end_idx, "Total Cost"], errors="coerce"
        ).fillna(0.0).sum()
        work.at[current_overall_start, "Unit Price"] = None
        work.at[current_overall_start, "Total Cost"] = subtotal

    for index in range(len(work)):
        row_type = normalize_type_value(work.at[index, "Type"])
        work.at[index, "Type"] = row_type
        norm = normalize_norm_value(work.at[index, "Norm"])
        work.at[index, "Norm"] = norm
        formula_text = work.at[index, "Formula"]
        unit_price = float(pd.to_numeric(work.at[index, "Unit Price"], errors="coerce") or 0.0)

        if row_type == "O":
            close_subgroup(index - 1)
            close_overall(index - 1)
            current_overall_start = index
            current_subgroup_qty = boq_qty
        elif row_type == "S":
            close_subgroup(index - 1)
            current_subgroup_start = index

        try:
            resultant = eval_formula(formula_text)
        except Exception as exc:
            resultant = 0.0
            errors.append(f"{article_row['Article_ID']} - row {index + 1}: {exc}")

        quantity = 0.0
        total_cost = 0.0

        if row_type == "O":
            quantity = boq_qty
        elif row_type == "S":
            if norm == "N":
                quantity = resultant * boq_qty
            elif norm == "C":
                quantity = resultant
            else:
                errors.append(f"{article_row['Article_ID']} - row {index + 1}: Norm must be N or C")
            current_subgroup_qty = quantity
        elif row_type == "M":
            if norm == "N":
                quantity = resultant * current_subgroup_qty
            elif norm == "C":
                quantity = resultant
            else:
                errors.append(f"{article_row['Article_ID']} - row {index + 1}: Norm must be N or C")
            total_cost = quantity * unit_price
        else:
            errors.append(f"{article_row['Article_ID']} - row {index + 1}: Type must be O, S, or M")

        work.at[index, "Resultant"] = resultant
        work.at[index, "Quantity"] = quantity
        work.at[index, "Total Cost"] = total_cost

    close_subgroup(len(work) - 1)
    close_overall(len(work) - 1)

    o_mask = work["Type"].astype(str).str.upper() == "O"
    article_total = float(pd.to_numeric(work["Total Cost"], errors="coerce").fillna(0.0).sum())
    if o_mask.any():
        article_total = float(pd.to_numeric(work.loc[o_mask, "Total Cost"], errors="coerce").fillna(0.0).iloc[-1])
    unit_rate = article_total / boq_qty if boq_qty else 0.0

    return work, unit_rate, article_total, errors


def run_all():
    boq = st.session_state.boq_df.copy()
    all_errors = []
    for index in boq.index:
        article_id = str(boq.at[index, "Article_ID"])
        breakdown = get_breakdown(article_id)
        if breakdown.empty:
            boq.at[index, "Unit Price"] = None
            boq.at[index, "Total Price"] = None
            continue
        result_df, unit_rate, article_total, errors = calculate_article(boq.loc[index], breakdown)
        set_breakdown(article_id, result_df)
        boq.at[index, "Unit Price"] = unit_rate
        boq.at[index, "Total Price"] = article_total
        all_errors.extend(errors)
    st.session_state.boq_df = boq
    return all_errors


def export_excel() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        st.session_state.boq_df.to_excel(writer, sheet_name="BOQ", index=False)
        st.session_state.library_df.to_excel(writer, sheet_name="Library", index=False)
        all_breakdowns = []
        for article_id, df in st.session_state.breakdowns.items():
            if not df.empty:
                temp = df.copy()
                temp.insert(0, "Article_ID", article_id)
                all_breakdowns.append(temp)
        if all_breakdowns:
            pd.concat(all_breakdowns, ignore_index=True).to_excel(writer, sheet_name="Breakdowns", index=False)
    return output.getvalue()


def fmt_money(value):
    if value is None or pd.isna(value):
        return ""
    return f"{float(value):,.2f}"


def make_new_breakdown_row():
    return {
        "Type": "M",
        "Level": 2,
        "Category": "",
        "Code": "",
        "Description": "",
        "Norm": "N",
        "Formula": "1",
        "Resultant": None,
        "Quantity": None,
        "Unit": "",
        "Unit Price": 0.0,
        "Total Cost": None,
    }


def get_selected_indexes(selected_rows):
    indexes = []
    for row in selected_rows or []:
        node_row_index = row.get("_selectedRowNodeInfo", {}).get("nodeRowIndex")
        if node_row_index is not None:
            indexes.append(int(node_row_index))
    return sorted(set(indexes))


def breakdown_grid(article_id: str):
    df = get_breakdown(article_id).copy()
    if df.empty:
        df = empty_breakdown_df()

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(editable=st.session_state.edit_mode, resizable=True, sortable=False, filter=False)
    gb.configure_selection("multiple", use_checkbox=True)
    gb.configure_column("Type", editable=st.session_state.edit_mode, cellEditor="agSelectCellEditor", cellEditorParams={"values": ["O", "S", "M"]})
    gb.configure_column("Norm", editable=st.session_state.edit_mode, cellEditor="agSelectCellEditor", cellEditorParams={"values": ["N", "C"]})
    gb.configure_column("Level", editable=st.session_state.edit_mode, type=["numericColumn"])
    gb.configure_column("Resultant", editable=False, type=["numericColumn"])
    gb.configure_column("Quantity", editable=False, type=["numericColumn"])
    gb.configure_column("Total Cost", editable=False, type=["numericColumn"])
    gb.configure_column("Unit Price", type=["numericColumn"])

    row_style = JsCode(
        """
        function(params) {
            if (!params.data || !params.data.Type) {
                return {};
            }
            const t = String(params.data.Type).toUpperCase();
            if (t === "O") {
                return {backgroundColor: "#fde68a", fontWeight: "700"};
            }
            if (t === "S") {
                return {backgroundColor: "#bfdbfe", fontWeight: "600"};
            }
            if (t === "M") {
                return {backgroundColor: "#dcfce7"};
            }
            return {};
        }
        """
    )

    grid_options = gb.build()
    grid_options["getRowStyle"] = row_style

    response = AgGrid(
        df,
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.AS_INPUT,
        update_mode=GridUpdateMode.VALUE_CHANGED | GridUpdateMode.SELECTION_CHANGED,
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True,
        theme="streamlit",
        height=420,
        reload_data=False,
    )

    updated_df = pd.DataFrame(response["data"])
    if updated_df.empty:
        updated_df = empty_breakdown_df()

    set_breakdown(article_id, updated_df)
    return response


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
        margin-top: 16px;
    }
    .active-article {
        background: #fff8e6;
        border: 1px solid #f3d47b;
        border-radius: 12px;
        padding: 10px 12px;
        margin-bottom: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div class='excel-shell'>", unsafe_allow_html=True)
st.markdown("<div class='toolbar'>", unsafe_allow_html=True)
st.markdown("<div class='sheet-title'>BOQ Breakdown Calculator</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='sheet-subtitle'>One colored editable breakdown table. Click an Article_ID like A001 to open its own breakdown.</div>",
    unsafe_allow_html=True,
)

toolbar_cols = st.columns([1.3, 1.3, 1.0, 1.0, 1.4, 1.4, 3.5])
boq_file = toolbar_cols[0].file_uploader("Import BOQ", type=["xlsx"], label_visibility="collapsed", key="boq_upload")
lib_file = toolbar_cols[1].file_uploader("Import Library", type=["xlsx"], label_visibility="collapsed", key="lib_upload")

if toolbar_cols[2].button("Save", use_container_width=True):
    st.session_state.save_message = f"Draft saved in session at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

if toolbar_cols[3].button("Edit On" if st.session_state.edit_mode else "Edit Off", use_container_width=True):
    st.session_state.edit_mode = not st.session_state.edit_mode

run_clicked = toolbar_cols[4].button("Run Calculation", type="primary", use_container_width=True)
load_sample_clicked = toolbar_cols[5].button("Load Sample", use_container_width=True)

try:
    export_data = export_excel()
    toolbar_cols[6].download_button(
        "Export Excel",
        data=export_data,
        file_name="boq_breakdown_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
except Exception:
    toolbar_cols[6].button("Export Excel", disabled=True, use_container_width=True)

st.markdown("</div>", unsafe_allow_html=True)

if boq_file is not None or lib_file is not None:
    load_cols = st.columns([1.2, 5])
    if load_cols[0].button("Load Imported Files", use_container_width=True):
        try:
            if boq_file is not None:
                st.session_state.boq_df = normalize_boq_columns(pd.read_excel(boq_file))
                st.session_state.breakdowns = {}
                st.session_state.selected_article = None
            if lib_file is not None:
                st.session_state.library_df = normalize_library_columns(pd.read_excel(lib_file))
            st.success("Imported files loaded.")
        except Exception as exc:
            st.error(str(exc))

if load_sample_clicked:
    st.session_state.boq_df = DEFAULT_BOQ.copy()
    st.session_state.library_df = DEFAULT_LIBRARY.copy()
    st.session_state.breakdowns = {}
    st.session_state.selected_article = None
    st.success("Sample data loaded.")

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
header_cols = st.columns([1.2, 3.2, 1.2, 1.0, 1.4, 1.4, 1.8])
header_cols[0].markdown("**Article_ID**")
header_cols[1].markdown("**Description**")
header_cols[2].markdown("**Quantity**")
header_cols[3].markdown("**Unit**")
header_cols[4].markdown("**Unit Price**")
header_cols[5].markdown("**Total Price**")
header_cols[6].markdown("**Template_Name**")

template_options = sorted(st.session_state.library_df["Template_Name"].dropna().astype(str).unique().tolist())
boq_df = st.session_state.boq_df.copy()

for idx in boq_df.index:
    row = boq_df.loc[idx]
    cols = st.columns([1.2, 3.2, 1.2, 1.0, 1.4, 1.4, 1.8])

    article_label = str(row["Article_ID"]) if pd.notna(row["Article_ID"]) else f"Row {idx + 1}"
    if cols[0].button(article_label, key=f"open_article_{idx}", use_container_width=True):
        st.session_state.selected_article = article_label
        st.rerun()

    if st.session_state.edit_mode:
        boq_df.at[idx, "Description"] = cols[1].text_input(
            "Description",
            value=str(row["Description"]) if pd.notna(row["Description"]) else "",
            key=f"description_{idx}",
            label_visibility="collapsed",
        )
        boq_df.at[idx, "Quantity"] = cols[2].number_input(
            "Quantity",
            value=float(pd.to_numeric(row["Quantity"], errors="coerce") or 0.0),
            key=f"quantity_{idx}",
            label_visibility="collapsed",
        )
        boq_df.at[idx, "Unit"] = cols[3].text_input(
            "Unit",
            value=str(row["Unit"]) if pd.notna(row["Unit"]) else "",
            key=f"unit_{idx}",
            label_visibility="collapsed",
        )
        cols[4].write(fmt_money(row["Unit Price"]))
        cols[5].write(fmt_money(row["Total Price"]))
        current_template = str(row["Template_Name"]) if pd.notna(row["Template_Name"]) else ""
        boq_df.at[idx, "Template_Name"] = cols[6].selectbox(
            "Template_Name",
            options=[""] + template_options,
            index=([""] + template_options).index(current_template) if current_template in template_options else 0,
            key=f"template_{idx}",
            label_visibility="collapsed",
        )
    else:
        cols[1].write(str(row["Description"]) if pd.notna(row["Description"]) else "")
        cols[2].write(f"{float(pd.to_numeric(row['Quantity'], errors='coerce') or 0.0):.3f}")
        cols[3].write(str(row["Unit"]) if pd.notna(row["Unit"]) else "")
        cols[4].write(fmt_money(row["Unit Price"]))
        cols[5].write(fmt_money(row["Total Price"]))
        cols[6].write(str(row["Template_Name"]) if pd.notna(row["Template_Name"]) else "")

st.session_state.boq_df = boq_df[BOQ_COLUMNS].copy()

selected_article = st.session_state.selected_article
if selected_article:
    article_row = st.session_state.boq_df[
        st.session_state.boq_df["Article_ID"].astype(str) == str(selected_article)
    ]
    if not article_row.empty:
        article = article_row.iloc[0]
        ensure_article_breakdown(selected_article, article["Template_Name"])

        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown(
            f"<div class='active-article'><b>{selected_article}</b><br>"
            f"Description: {article['Description']}<br>"
            f"Quantity: {article['Quantity']} {article['Unit']}</div>",
            unsafe_allow_html=True,
        )

        top_cols = st.columns([1.0, 1.0, 1.3, 1.0, 1.0])
        add_clicked = top_cols[0].button("Add Row", use_container_width=True)
        delete_clicked = top_cols[1].button("Delete Selected", use_container_width=True)
        reload_clicked = top_cols[2].button("Reload Template", use_container_width=True)
        calc_clicked = top_cols[3].button("Calculate This Article", use_container_width=True)
        close_clicked = top_cols[4].button("Close", use_container_width=True)

        grid_response = breakdown_grid(selected_article)
        selected_indexes = get_selected_indexes(grid_response.get("selected_rows", []))

        if add_clicked:
            breakdown = get_breakdown(selected_article)
            insert_at = selected_indexes[0] + 1 if selected_indexes else len(breakdown)
            new_row = pd.DataFrame([make_new_breakdown_row()])
            breakdown = pd.concat(
                [breakdown.iloc[:insert_at], new_row, breakdown.iloc[insert_at:]],
                ignore_index=True,
            )
            set_breakdown(selected_article, breakdown)
            st.rerun()

        if delete_clicked:
            if selected_indexes:
                breakdown = get_breakdown(selected_article)
                breakdown = breakdown.drop(index=selected_indexes).reset_index(drop=True)
                set_breakdown(selected_article, breakdown)
                st.rerun()
            else:
                st.warning("Select at least one row to delete.")

        if reload_clicked:
            if str(article["Template_Name"]).strip():
                load_template(selected_article, article["Template_Name"])
                st.success(f"Template reloaded for {selected_article}.")
                st.rerun()
            else:
                st.warning("Select a template first.")

        if calc_clicked:
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
            st.rerun()

        if close_clicked:
            st.session_state.selected_article = None
            st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

library_expander = st.expander("Template Library Sheet")
with library_expander:
    st.session_state.library_df = st.data_editor(
        st.session_state.library_df,
        num_rows="dynamic" if st.session_state.edit_mode else "fixed",
        use_container_width=True,
        hide_index=True,
        key="library_sheet",
        disabled=not st.session_state.edit_mode,
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=["O", "S", "M"]),
            "Level": st.column_config.NumberColumn("Level", format="%d"),
            "Norm": st.column_config.SelectboxColumn("Norm", options=["N", "C"]),
            "Unit Price": st.column_config.NumberColumn("Unit Price", format="%.4f"),
            "Resultant": st.column_config.NumberColumn("Resultant", format="%.3f"),
            "Quantity": st.column_config.NumberColumn("Quantity", format="%.3f"),
            "Total Cost": st.column_config.NumberColumn("Total Cost", format="%.2f"),
        },
    )

st.markdown("</div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)
