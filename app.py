import io
import math
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
            "Category": "Concrete",
            "Code": "S-001",
            "Description": "Concrete works",
            "Norm": "F",
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
            "Category": "Concrete",
            "Code": "0126120109",
            "Description": "Supply concrete C30/37",
            "Norm": "F",
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
            "Category": "Concrete",
            "Code": "0490000411",
            "Description": "Concrete labor",
            "Norm": "F",
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
            "Category": "Pump",
            "Code": "S-002",
            "Description": "Pump works",
            "Norm": "F",
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
            "Category": "Pump",
            "Code": "0126120308",
            "Description": "Concrete pumping 36m",
            "Norm": "F",
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
            "Category": "Steel",
            "Code": "S-003",
            "Description": "Reinforcement works",
            "Norm": "F",
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
            "Category": "Steel",
            "Code": "0126111101",
            "Description": "Reinforcement bars",
            "Norm": "F",
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
            "Category": "Formwork",
            "Code": "S-004",
            "Description": "Formwork works",
            "Norm": "F",
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
            "Category": "Formwork",
            "Code": "0126131003",
            "Description": "Formwork footing",
            "Norm": "F",
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
    required = ["Template_Name", "Type", "Category", "Code", "Description", "Norm", "Formula", "Unit", "Unit Price"]
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
        norm = str(work.at[index, "Norm"]).strip().upper()
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
            if norm == "F":
                quantity = resultant * boq_qty
            elif norm == "C":
                quantity = resultant
            else:
                errors.append(f"{article_row['Article_ID']} - row {index + 1}: Norm must be F or C")
            current_subgroup_qty = quantity
        elif row_type == "M":
            if norm == "F":
                quantity = resultant * current_subgroup_qty
            elif norm == "C":
                quantity = resultant
            else:
                errors.append(f"{article_row['Article_ID']}
