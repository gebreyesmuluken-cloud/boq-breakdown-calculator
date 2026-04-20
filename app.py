import io
import math
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="BOQ Breakdown Calculator", layout="wide")

BOQ_COLUMNS = [
    "Open",
    "Article_ID",
    "Description",
    "Quantity",
    "Unit",
    "Unit Price",
    "Total Price",
    "Template_Name",
]

BOQ_STORAGE_COLUMNS = [column for column in BOQ_COLUMNS if column != "Open"]

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
            "Type": "T",
            "Category": "Main",
            "Code": "T-001",
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
            "Type": "M",
            "Category": "Concrete",
            "Code": "M-001",
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
            "Type": "D",
            "Category": "Concrete",
            "Code": "0126120109",
            "Description": "Supply concrete C30/37",
            "Norm": "F",
            "Formula": "1.05",
            "Resultant": None,
            "Quantity": None,
            "Unit": "m3",
