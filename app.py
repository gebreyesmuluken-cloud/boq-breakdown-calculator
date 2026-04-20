import io
import importlib.util
import math
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

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

BOQ_STORAGE_COLUMNS = BOQ_COLUMNS.copy()

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
)[BOQ_STORAGE_COLUMNS]

DEFAULT_LIBRARY = pd.DataFrame(
    [
        {
            "Template_Name": "FOUNDATION_FOOTING",
            "Type": "O",
            "Level": 0,
            "Category": "Main",
            "Code": "O-001",
            "Description": "FOUNDATION FOOTING",
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
            "Norm": "N",
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
    type_value = str(value).strip().upper()
    legacy_map = {
        "T": "O",
        "D": "M",
    }
    return legacy_map.get(type_value, type_value)


def normalize_norm_value(value) -> str:
    norm_value = str(value).strip().upper()
    legacy_map = {
        "F": "N",
    }
    return legacy_map.get(norm_value, norm_value)


def normalize_level_value(value, row_type: str) -> int:
    default_levels = {
        "O": 0,
        "S": 1,
        "M": 2,
    }
    parsed = pd.to_numeric(value, errors="coerce")
    if pd.notna(parsed):
        return max(0, int(parsed))
    return default_levels.get(row_type, 0)


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
    if "last_saved_at" not in st.session_state:
        st.session_state.last_saved_at = None


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
    work.columns = [rename_map.get(str(column).strip().lower(), column) for column in work.columns]
    for column in ["Article_ID", "Description", "Quantity", "Unit"]:
        if column not in work.columns:
            raise ValueError(f"BOQ file is missing required column: {column}")
    if "Unit Price" not in work.columns:
        work["Unit Price"] = None
    if "Total Price" not in work.columns:
        work["Total Price"] = None
    if "Template_Name" not in work.columns:
        work["Template_Name"] = ""
    return work[BOQ_STORAGE_COLUMNS].copy()


def normalize_library_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "template_name": "Template_Name",
        "type": "Type",
        "level": "Level",
        "catagory": "Category",
        "category": "Category",
        "code": "Code",
        "describtion": "Description",
        "description": "Description",
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
    work.columns = [rename_map.get(str(column).strip().lower(), column) for column in work.columns]
    required = [
        "Template_Name",
        "Type",
        "Level",
        "Category",
        "Code",
        "Description",
        "Norm",
        "Formula",
        "Unit",
        "Unit Price",
    ]
    for column in required:
        if column not in work.columns:
            raise ValueError(f"Library file is missing required column: {column}")
    for column in ["Resultant", "Quantity", "Total Cost"]:
        if column not in work.columns:
            work[column] = None
    return work[["Template_Name", *BREAKDOWN_COLUMNS]].copy()


def get_breakdown(article_id: str) -> pd.DataFrame:
    if article_id not in st.session_state.breakdowns:
        st.session_state.breakdowns[article_id] = empty_breakdown_df()
    return st.session_state.breakdowns[article_id].copy()


def set_breakdown(article_id: str, df: pd.DataFrame):
    work = df.copy()
    for column in BREAKDOWN_COLUMNS:
        if column not in work.columns:
            work[column] = None
    st.session_state.breakdowns[article_id] = work[BREAKDOWN_COLUMNS].copy()


def load_template(article_id: str, template_name: str):
    library = st.session_state.library_df
    rows = library[library["Template_Name"].astype(str) == str(template_name)].copy()
    if rows.empty:
        st.warning(f"Template '{template_name}' was not found.")
        return
    set_breakdown(article_id, rows[BREAKDOWN_COLUMNS].copy())


def ensure_article_breakdown(article_id: str, template_name: str):
    current = get_breakdown(article_id)
    if not current.empty:
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
    open_headers = []
    overall_total = 0.0

    def close_header():
        nonlocal overall_total
        if not open_headers:
            return

        header = open_headers.pop()
        subtotal = float(header["accumulated_cost"])
        quantity = float(header["quantity"])
        unit_price = subtotal / quantity if quantity else None

        work.at[header["index"], "Unit Price"] = unit_price
        work.at[header["index"], "Total Cost"] = subtotal

        if open_headers:
            open_headers[-1]["accumulated_cost"] += subtotal
        else:
            overall_total += subtotal

    for index in range(len(work)):
        row_type = normalize_type_value(work.at[index, "Type"])
        row_level = normalize_level_value(work.at[index, "Level"], row_type)
        is_header = row_type in {"O", "S"}

        work.at[index, "Type"] = row_type
        work.at[index, "Level"] = row_level

        norm = normalize_norm_value(work.at[index, "Norm"])
        work.at[index, "Norm"] = norm
        formula_text = work.at[index, "Formula"]
        unit_price = float(pd.to_numeric(work.at[index, "Unit Price"], errors="coerce") or 0.0)

        try:
            resultant = eval_formula(formula_text)
        except Exception as exc:
            resultant = 0.0
            errors.append(f"{article_row['Article_ID']} - row {index + 1}: {exc}")

        if is_header:
            while open_headers and row_level <= open_headers[-1]["level"]:
                close_header()

        quantity = 0.0
        total_cost = 0.0

        if row_type in {"O", "S", "M"}:
            if norm == "N":
                quantity = resultant * boq_qty
            elif norm == "C":
                quantity = resultant
            else:
                errors.append(f"{article_row['Article_ID']} - row {index + 1}: Norm must be N or C")

            if row_type == "M":
                total_cost = quantity * unit_price
                if open_headers:
                    open_headers[-1]["accumulated_cost"] += total_cost
                else:
                    overall_total += total_cost
            else:
                open_headers.append(
                    {
                        "index": index,
                        "level": row_level,
                        "quantity": quantity,
                        "accumulated_cost": 0.0,
                    }
                )
        else:
            errors.append(f"{article_row['Article_ID']} - row {index + 1}: Type must be O, S, or M")

        work.at[index, "Resultant"] = resultant
        work.at[index, "Quantity"] = quantity
        if row_type in {"O", "S"}:
            work.at[index, "Unit Price"] = None
        work.at[index, "Total Cost"] = total_cost

    while open_headers:
        close_header()

    unit_price = overall_total / boq_qty if boq_qty else 0.0
    return work, unit_price, overall_total, errors


def recalculate_all():
    errors = []
    boq = st.session_state.boq_df.copy()

    for idx, article in boq.iterrows():
        article_id = str(article["Article_ID"])
        template_name = str(article.get("Template_Name", "")).strip()
        ensure_article_breakdown(article_id, template_name)
        breakdown = get_breakdown(article_id)
        calculated, unit_price, total_price, article_errors = calculate_article(article, breakdown)
        set_breakdown(article_id, calculated)
        boq.at[idx, "Unit Price"] = unit_price if unit_price else None
        boq.at[idx, "Total Price"] = total_price if total_price else None
        errors.extend(article_errors)

    st.session_state.boq_df = boq
    return errors


def get_excel_engine():
    for engine in ["openpyxl", "xlsxwriter"]:
        if importlib.util.find_spec(engine) is not None:
            return engine
    return None


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    engine = get_excel_engine()
    if engine is None:
        raise ModuleNotFoundError("No Excel writer engine is installed.")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine=engine) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def build_export_workbook() -> bytes:
    engine = get_excel_engine()
    if engine is None:
        raise ModuleNotFoundError("No Excel writer engine is installed.")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine=engine) as writer:
        st.session_state.boq_df.to_excel(writer, index=False, sheet_name="BOQ")
        st.session_state.library_df.to_excel(writer, index=False, sheet_name="Library")
        for article_id, breakdown in st.session_state.breakdowns.items():
            sheet_name = f"BD_{article_id}"[:31]
            breakdown.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def build_export_zip() -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("BOQ.csv", st.session_state.boq_df.to_csv(index=False))
        archive.writestr("Library.csv", st.session_state.library_df.to_csv(index=False))
        for article_id, breakdown in st.session_state.breakdowns.items():
            archive.writestr(f"BD_{article_id}.csv", breakdown.to_csv(index=False))
    return output.getvalue()


def save_snapshot():
    st.session_state.last_saved_at = datetime.now()
    timestamp = st.session_state.last_saved_at.strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.save_message = f"Snapshot prepared at {timestamp}"


init_state()

st.title("BOQ Breakdown Calculator")
st.caption("Manage BOQ items, template libraries, and calculated cost breakdowns in one place.")

with st.sidebar:
    st.subheader("Data")
    boq_upload = st.file_uploader("Import BOQ", type=["xlsx", "xls", "csv"], key="boq_upload")
    library_upload = st.file_uploader(
        "Import Library", type=["xlsx", "xls", "csv"], key="library_upload"
    )

    if boq_upload is not None:
        try:
            if boq_upload.name.lower().endswith(".csv"):
                imported_boq = pd.read_csv(boq_upload)
            else:
                imported_boq = pd.read_excel(boq_upload)
            st.session_state.boq_df = normalize_boq_columns(imported_boq)
            st.session_state.breakdowns = {}
            st.success("BOQ imported successfully.")
        except Exception as exc:
            st.error(f"Failed to import BOQ: {exc}")

    if library_upload is not None:
        try:
            if library_upload.name.lower().endswith(".csv"):
                imported_library = pd.read_csv(library_upload)
            else:
                imported_library = pd.read_excel(library_upload)
            st.session_state.library_df = normalize_library_columns(imported_library)
            st.success("Library imported successfully.")
        except Exception as exc:
            st.error(f"Failed to import library: {exc}")

    if st.button("Recalculate", use_container_width=True):
        calc_errors = recalculate_all()
        if calc_errors:
            st.warning("\n".join(calc_errors))
        else:
            st.success("All article totals recalculated.")

    if st.button("Save Snapshot", use_container_width=True):
        save_snapshot()

    excel_engine = get_excel_engine()
    if excel_engine:
        st.download_button(
            "Download BOQ",
            data=dataframe_to_excel_bytes(st.session_state.boq_df, "BOQ"),
            file_name="boq.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.download_button(
            "Download Workbook",
            data=build_export_workbook(),
            file_name="boq_breakdown_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.info("Excel export is unavailable in this environment. CSV downloads are enabled instead.")
        st.download_button(
            "Download BOQ CSV",
            data=dataframe_to_csv_bytes(st.session_state.boq_df),
            file_name="boq.csv",
            mime="text/csv",
            use_container_width=True,
        )

        st.download_button(
            "Download Export ZIP",
            data=build_export_zip(),
            file_name="boq_breakdown_export.zip",
            mime="application/zip",
            use_container_width=True,
        )

    if st.session_state.save_message:
        st.info(st.session_state.save_message)

boq_tab, library_tab = st.tabs(["BOQ", "Library"])

with boq_tab:
    st.subheader("Bill of Quantities")
    st.caption("Click A001 or A002 in the BOQ table to open that article's own breakdown panel below.")
    boq_df = st.session_state.boq_df.copy()
    selection_event = st.dataframe(
        boq_df,
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key="boq_selector",
    )

    selected_rows = selection_event.selection.get("rows", []) if selection_event else []
    if selected_rows:
        selected_index = selected_rows[0]
        st.session_state.selected_article = str(boq_df.iloc[selected_index]["Article_ID"])
    elif (
        boq_df["Article_ID"].astype(str).tolist()
        and st.session_state.selected_article not in boq_df["Article_ID"].astype(str).tolist()
    ):
        st.session_state.selected_article = str(boq_df.iloc[0]["Article_ID"])

    st.divider()
    st.subheader("Breakdown Calculation")
    boq_df = st.session_state.boq_df
    article_options = boq_df["Article_ID"].astype(str).tolist()

    if article_options:
        if st.session_state.selected_article not in article_options:
            st.session_state.selected_article = article_options[0]

        selected_article = st.session_state.selected_article
        selected_row = boq_df[boq_df["Article_ID"].astype(str) == selected_article].iloc[0]
        template_name = str(selected_row.get("Template_Name", "")).strip()
        ensure_article_breakdown(selected_article, template_name)

        info_col, action_col = st.columns([3, 1])
        with info_col:
            st.markdown(f"### {selected_article}")
            st.write(f"Description: {selected_row['Description']}")
            st.write(f"Quantity: {selected_row['Quantity']} {selected_row['Unit']}")
            st.caption("Each article has its own separate breakdown calculation panel.")
        with action_col:
            if st.button("Calculate This Article", use_container_width=True):
                breakdown = get_breakdown(selected_article)
                calculated, unit_price, total_price, calc_errors = calculate_article(
                    selected_row, breakdown
                )
                set_breakdown(selected_article, calculated)
                row_index = boq_df[boq_df["Article_ID"].astype(str) == selected_article].index[0]
                st.session_state.boq_df.at[row_index, "Unit Price"] = unit_price if unit_price else None
                st.session_state.boq_df.at[row_index, "Total Price"] = (
                    total_price if total_price else None
                )
                if calc_errors:
                    st.warning("\n".join(calc_errors))
                else:
                    st.success(f"Breakdown calculated for {selected_article}.")

        edited_breakdown = st.data_editor(
            get_breakdown(selected_article),
            num_rows="dynamic",
            use_container_width=True,
            key=f"breakdown_editor_{selected_article}",
        )
        set_breakdown(selected_article, edited_breakdown)
    else:
        st.info("Add at least one BOQ article to manage breakdowns.")

with library_tab:
    st.subheader("Template Library")
    edited_library = st.data_editor(
        st.session_state.library_df,
        num_rows="dynamic",
        use_container_width=True,
        key="library_editor",
    )
    st.session_state.library_df = normalize_library_columns(edited_library)
