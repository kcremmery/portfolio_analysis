#!/usr/bin/env python3
"""Self-contained Streamlit portfolio: Type 0 and IPD workbooks, charts and reports.

Run: .venv/bin/python -m streamlit run Portfolio_Analysis.py
Requires streamlit, pandas, numpy, openpyxl, matplotlib and plotly.
All workbook rules and chart renderers live here; no other project module is imported.
"""

from __future__ import annotations

from base64 import b64encode
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from hashlib import sha256
from io import BytesIO
import json
import math
from pathlib import Path
import re
import secrets
import threading
import warnings
from zipfile import ZIP_DEFLATED, ZipFile

import matplotlib.dates as mdates
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
from matplotlib.font_manager import FontProperties
from matplotlib.lines import Line2D
from matplotlib.offsetbox import AnnotationBbox, DrawingArea, HPacker, TextArea, VPacker
from matplotlib.ticker import FuncFormatter
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# Generate charts button colours. Set these to CSS colours, for example "#16803c".
GENERATE_BUTTON_COLOR = "#32AAFF"
GENERATE_BUTTON_TEXT_COLOR = "#ffffff"


@st.cache_resource
def render_lock():
    """Matplotlib is shared across reruns and browser sessions."""
    return threading.RLock()


def figure_png(figure, dpi=300):
    """Export and release an Agg figure without importing a GUI backend."""
    try:
        output = BytesIO()
        figure.savefig(output, format="png", dpi=dpi, facecolor="white")
        return output.getvalue()
    finally:
        figure.clear()



# Type 0 workbook rules and Gantt renderer

TYPE0_SHEET_NAME = "EMEA"
TYPE0_SKIP_ROWS = 5
TYPE0_SEGMENTS = ("Building", "Infrastructure", "Industry")
TYPE0_DEFAULT_STATUSES = ("In Progress", "To start")
TYPE0_MIN_PRIORITY = 10
TYPE0_MAX_PRIORITY = 19
TYPE0_LEFT_DATE_EXPANSION = 0.10
TYPE0_RIGHT_DATE_EXPANSION = 0.03
TYPE0_A3_WIDTH_MM = 420
TYPE0_A3_HEIGHT_MM = 297
TYPE0_PNG_DPI = 300
TYPE0_DATE_FIELDS = ("Creation Date", "Due Date", "Original Due Date", "Latest Due Date")
TYPE0_NUMERIC_CHECK_FIELDS = ("PAS (in k EUR)", "PIAS (in k EUR)", "CAPEX (in k EUR)", "CM% (as % of sales)")
TYPE0_K_EUR_MISSING_TOKENS = ("", "-", "nc", "na", "n/a")
TYPE0_ERROR_COLUMNS = ("Recipient", "Product_Manager", "Project_Manager", "Number", "Project_Name", "Action")
TYPE0_STATUS_COLOURS = {"On track": "#52DEAD", "At risk": "#FAA64A", "Overdue": "#FA6157"}
TYPE0_MISSING_DATA_LABEL = "missing and invalid data"
GREY = "#5E5E5E"
TODAY_COLOUR = "#3066E3"
MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def type0_cell_text(value) -> str:
    """Display/normalise one scalar cell without turning missing values into text."""
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (float, np.floating)) and np.isfinite(value) and value.is_integer():
        return str(int(value))
    return str(value).replace("\u00a0", " ").strip()


def type0_warn_invalid_values(values, field_name: str) -> None:
    unique_values = list(dict.fromkeys(values))
    if unique_values:
        warnings.warn(
            f"{field_name}: could not convert value(s): {', '.join(unique_values)}",
            UserWarning,
            stacklevel=2,
        )


def type0_parse_date(value):
    """Accept Excel dates/serials and the explicit formats used by the R script.

    Return NaT for impossible dates or unsupported text; never guess an ambiguous
    month/day order. Text dates use DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY or ISO.
    """
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, (date, datetime, pd.Timestamp, np.datetime64)):
        stamp = pd.Timestamp(value)
        if stamp.tzinfo is not None:
            stamp = stamp.tz_localize(None)
        return stamp.normalize()
    text = type0_cell_text(value).replace("\u00a0", "")
    if not text:
        return pd.NaT
    try:
        serial = float(text)
        if np.isfinite(serial) and 20000 < serial < 100000:
            return pd.Timestamp(datetime(1899, 12, 30) + timedelta(days=serial)).normalize()
    except (ValueError, OverflowError):
        pass
    formats = (
        (r"\d{1,2}/\d{1,2}/\d{4}", "%d/%m/%Y"),
        (r"\d{1,2}-\d{1,2}-\d{4}", "%d-%m-%Y"),
        (r"\d{1,2}\.\d{1,2}\.\d{4}", "%d.%m.%Y"),
        (r"\d{4}-\d{1,2}-\d{1,2}", "%Y-%m-%d"),
        (r"\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}:\d{2}(?:\.\d+)?", None),
    )
    for pattern, fmt in formats:
        if re.fullmatch(pattern, text):
            try:
                if fmt is None:
                    text = re.sub(r"\s+", " ", text)
                    fmt = "%Y-%m-%d %H:%M:%S.%f" if "." in text else "%Y-%m-%d %H:%M:%S"
                return pd.Timestamp(datetime.strptime(text, fmt)).normalize()
            except (ValueError, OverflowError):
                return pd.NaT
    return pd.NaT


def type0_convert_to_date(values: pd.Series, field_name="Date", warn_invalid=True) -> pd.Series:
    result = pd.Series([type0_parse_date(x) for x in values], index=values.index, dtype="datetime64[ns]")
    if warn_invalid:
        type0_warn_invalid_values(
            [type0_cell_text(raw) for raw, parsed in zip(values, result) if type0_cell_text(raw) and pd.isna(parsed)],
            field_name,
        )
    return result


def type0_numeric_value(value) -> float:
    text = type0_cell_text(value)
    if not text:
        return np.nan
    try:
        result = float(text)
        return result if np.isfinite(result) else np.nan
    except (ValueError, OverflowError):
        return np.nan


def type0__convert_numeric(values, parser, field_name, missing_tokens=("",)):
    result, invalid = [], []
    for value in values:
        text = type0_cell_text(value)
        parsed = parser(value)
        result.append(parsed)
        if text.lower() not in missing_tokens and pd.isna(parsed):
            invalid.append(text)
    type0_warn_invalid_values(invalid, field_name)
    return pd.Series(result, index=values.index, dtype=float)


def type0_convert_to_number(values, field_name="Number"):
    return type0__convert_numeric(values, type0_numeric_value, field_name)


def type0_convert_to_k_eur(values, field_name="Value"):
    def parse(value):
        text = type0_cell_text(value).lower()
        if text in TYPE0_K_EUR_MISSING_TOKENS or re.match(r"^[0-9,.]+\s*-\s*[0-9,.]+", text):
            return np.nan
        millions = re.search(r"m\s*(?:€|eur)?$", text) is not None and "k" not in text
        # Match the original convention: commas are thousands separators.
        text = text.replace(",", "")
        text = re.sub(r"k\s*(?:€|eur|e)?|m\s*(?:€|eur)?|€|eur", "", text).strip()
        parsed = type0_numeric_value(text)
        return parsed * 1000 if millions else parsed
    return type0__convert_numeric(values, parse, field_name, TYPE0_K_EUR_MISSING_TOKENS)


def type0_convert_to_percentage(values, field_name="Percentage"):
    def parse(value):
        text = type0_cell_text(value)
        if re.match(r"^[0-9,.]+\s*-\s*[0-9,.]+", text):
            return np.nan
        result = type0_numeric_value(text.replace("%", ""))
        return result / 100 if "%" in text else result
    return type0__convert_numeric(values, parse, field_name)


def type0_require_columns(df, columns, context):
    missing = [column for column in columns if column not in df.columns]
    if missing:
        raise ValueError(f"{context}. Missing column(s): {', '.join(missing)}")


def type0_convert_type0_data(df):
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(name).strip()) for name in df.columns]
    if df.columns.duplicated().any():
        raise ValueError("Excel contains duplicate column names after whitespace cleanup.")
    type0_require_columns(
        df,
        [*TYPE0_DATE_FIELDS, "PAS (in k EUR)", "PIAS (in k EUR)", "CM% (as % of sales)", "Priority proposal"],
        "Cannot prepare Type 0 data",
    )
    for field in TYPE0_DATE_FIELDS:
        raw_column = f".raw_date_{field}"
        if raw_column not in df:
            df[raw_column] = df[field].copy()
        df[field] = type0_convert_to_date(df[field], field)
    for field in TYPE0_NUMERIC_CHECK_FIELDS:
        raw_column = f".raw_numeric_{field}"
        if field in df and raw_column not in df:
            df[raw_column] = df[field].copy()
    for field in ("PAS (in k EUR)", "PIAS (in k EUR)"):
        df[field] = type0_convert_to_k_eur(df[field], field)
    if "CAPEX (in k EUR)" in df:
        df["CAPEX (in k EUR)"] = type0_convert_to_number(df["CAPEX (in k EUR)"], "CAPEX (in k EUR)")
    df["CM% (as % of sales)"] = type0_convert_to_percentage(df["CM% (as % of sales)"], "CM% (as % of sales)")
    df["Priority proposal"] = type0_convert_to_number(df["Priority proposal"], "Priority proposal")
    return df


def type0_is_completed_only(statuses):
    return {type0_cell_text(status).casefold() for status in statuses} == {'completed'}


def type0_parse_priority_limit(value):
    return float(value) if type0_cell_text(value) else None


def type0_priority_range_label(min_priority, max_priority):
    minimum, maximum = map(type0_parse_priority_limit, (min_priority, max_priority))
    if minimum is None and maximum is None:
        return 'All'
    if minimum is None:
        return f'Up to {type0_cell_text(maximum)}'
    if maximum is None:
        return f'From {type0_cell_text(minimum)}'
    return f'{type0_cell_text(minimum)}-{type0_cell_text(maximum)}'


def type0_get_priority_projects(df, segment, min_priority=TYPE0_MIN_PRIORITY, max_priority=TYPE0_MAX_PRIORITY,
                          statuses=TYPE0_DEFAULT_STATUSES, apply_priority=True):
    type0_require_columns(df, ["Segment", "Status", "Plant", "Priority proposal"], "Cannot select projects")
    mask = (
        df["Segment"].eq(segment)
        & df["Status"].isin(statuses)
    )
    if apply_priority and not type0_is_completed_only(statuses):
        minimum, maximum = map(type0_parse_priority_limit, (min_priority, max_priority))
        if minimum is not None:
            mask &= df["Priority proposal"].ge(minimum)
        if maximum is not None:
            mask &= df["Priority proposal"].le(maximum)
    selected = df.loc[mask].copy()
    # Preserve stable input order for ties and put missing plants last.
    selected["_plant_sort"] = selected["Plant"].map(lambda x: type0_cell_text(x) or None)
    selected = selected.sort_values(["_plant_sort", "Priority proposal"], kind="stable", na_position="last")
    return selected.drop(columns="_plant_sort").reset_index(drop=True)


def type0_original_date_value(row, field):
    return row.get(f".raw_date_{field}", row.get(field, ""))


def type0_tracking_entries(value):
    """Keep native dates intact; split only text lists of historical dates."""
    if not type0_cell_text(value):
        return []
    if isinstance(value, (date, datetime, pd.Timestamp, np.datetime64)):
        return [value]
    return [part.strip() for part in type0_cell_text(value).split(",") if part.strip()]


def type0_count_project_numbers(df):
    """Count nonempty project numbers, normalising whitespace and Excel numbers."""
    type0_require_columns(df, ["Number"], "Cannot check duplicate project numbers")
    numbers = df["Number"].map(type0_cell_text)
    return numbers.loc[numbers.ne("")].value_counts()


def type0_check_priority_projects(df_priority, segment, today=None, number_counts=None):
    today = pd.Timestamp(today if today is not None else date.today()).normalize()
    type0_require_columns(
        df_priority,
        ["Number", "Product Manager", "Project Manager", "Progress", "PAS (in k EUR)",
         "PIAS (in k EUR)", "CM% (as % of sales)", *TYPE0_DATE_FIELDS],
        f"Cannot perform priority project checks for {segment}",
    )
    if number_counts is None:
        number_counts = type0_count_project_numbers(df_priority)
    errors = []
    for _, row in df_priority.iterrows():
        number = type0_cell_text(row["Number"])
        product_manager = type0_cell_text(row["Product Manager"]).upper()
        project_manager = type0_cell_text(row["Project Manager"]).upper()
        requestor = type0_cell_text(row.get("Requestor Name", "")).upper()
        managers = list(dict.fromkeys(name for name in (product_manager, project_manager) if name))
        recipient = " / ".join(managers) or requestor or "Others"

        def add_error(action, action_recipient=None):
            errors.append({
                "Recipient": action_recipient or recipient,
                "Product_Manager": product_manager or "NO PRODUCT MANAGER",
                "Project_Manager": project_manager or "NO PROJECT MANAGER",
                "Number": number or "[No Number]", "Project_Name": type0_cell_text(row.get("Project Name", "")),
                "Action": action,
            })

        # Missing managers are assigned to the available manager, requestor or Others.
        if not product_manager:
            add_error("Add product manager")
        if not project_manager:
            add_error("Add project manager")
        if not number:
            add_error("Add project number")
        elif number_counts.get(number, 0) > 1:
            add_error(f"Duplicate project number '{number}' ({number_counts[number]} occurrences). Assign a unique project number.")
        if not type0_cell_text(row["Progress"]):
            add_error("Add progress")
        # Control 2: compare the financial values used in the activity label.
        pas, pias = row["PAS (in k EUR)"], row["PIAS (in k EUR)"]
        if pd.notna(pas) and pd.notna(pias):
            if pias > pas:
                add_error("PIAS cannot be larger than PAS")
            elif pias == pas:
                add_error(
                    f"PAS = PIAS ({type0_format_value(pas)} k EUR): "
                    "are all projected sales incremental?",
                    action_recipient=product_manager,
                )
        # Control 3: no due date before Creation Date.
        creation, due = row["Creation Date"], row["Due Date"]
        if pd.notna(creation) and any(
            pd.notna(row[field]) and row[field] < creation
            for field in ("Original Due Date", "Due Date", "Latest Due Date")
        ):
            add_error("Original Due Date, Due Date or Latest Due Date cannot be earlier than Creation Date")
        # Controls 4 and 5: genuine empty cells, not failed conversions.
        if not type0_cell_text(type0_original_date_value(row, "Creation Date")):
            add_error("Creation date cannot be empty")
        if not type0_cell_text(type0_original_date_value(row, "Due Date")):
            add_error("Due date cannot be empty")
        # Control 6: reschedule current due dates that are in the past.
        if pd.notna(due) and due < today:
            add_error("Reschedule Due Date")
        # Control 7: format/calendar validation, including each tracking entry.
        for field in (*TYPE0_DATE_FIELDS, "Due Date Tracking"):
            if field not in df_priority.columns:
                continue
            value = type0_original_date_value(row, field)
            if not type0_cell_text(value):
                continue
            values = type0_tracking_entries(value) if field == "Due Date Tracking" else [value]
            invalid = list(dict.fromkeys(type0_cell_text(v) for v in values if pd.isna(type0_parse_date(v))))
            if invalid:
                invalid_text = ", ".join(f"'{v}'" for v in invalid)
                add_error(
                    f"Date in '{field}' is not in the right format or is invalid: {invalid_text}. "
                    "Correct the date (e.g. DD/MM/YYYY)."
                )
        # PAS and CM% are mandatory; PIAS and CAPEX may be empty.
        for field in TYPE0_NUMERIC_CHECK_FIELDS:
            if field not in df_priority.columns:
                continue
            raw = type0_cell_text(row.get(f".raw_numeric_{field}", row[field]))
            missing_tokens = TYPE0_K_EUR_MISSING_TOKENS if field in ("PAS (in k EUR)", "PIAS (in k EUR)") else ("",)
            if raw.lower() in missing_tokens:
                if field in ("PAS (in k EUR)", "CM% (as % of sales)"):
                    add_error(f"Value in '{field}' cannot be empty. Enter a value.")
                continue
            if pd.notna(row[field]):
                continue
            if field == "CM% (as % of sales)":
                correction = "Enter a single percentage (e.g. 35% or 0.35); ranges are not supported."
            elif field == "CAPEX (in k EUR)":
                correction = "Enter a single number in k EUR without a unit suffix (e.g. 5 for 5 k€)."
            else:
                correction = "Enter a single amount in k EUR; ranges are not supported."
            add_error(f"Value in '{field}' could not be converted: '{raw}'. {correction}")
    return pd.DataFrame(errors, columns=TYPE0_ERROR_COLUMNS)


def type0_english_date(value):
    stamp = pd.Timestamp(value)
    return f"{stamp.day:02d} {MONTHS[stamp.month - 1]} {stamp.year}"


def type0_priority_checks_text(checks_by_segment, projects_by_segment, context=()):
    """Build the one consolidated report without writing shared server files."""
    now = datetime.now()
    lines = [
        "Priority Type 0 Project Data Checks - All Segments",
        f"Generated: {type0_english_date(now)} {now:%H:%M}", *context, "",
    ]
    for segment, errors in checks_by_segment.items():
        count = len(projects_by_segment[segment])
        lines += [f"Segment: {segment}", f"Projects checked: {count}", ""]
        if count == 0:
            lines += ["No priority projects found.", ""]
            continue
        if errors.empty:
            lines += ["No errors identified.", ""]
            continue
        errors = errors.sort_values(["Recipient", "Number", "Project_Name"], kind="stable")
        for recipient, group in errors.groupby("Recipient", sort=False):
            lines.append(f"Recipient: {recipient}")
            for row in group.itertuples(index=False):
                project_label = f"{row.Number} ({row.Project_Name})" if row.Project_Name else row.Number
                lines.append(f"- {project_label}: {row.Action}")
            lines.append("")
    return "\n".join(lines) + "\n"


def type0_extract_average_progress(progress_label):
    matches = re.findall(r"[0-9]+(?:[.,][0-9]+)?\s*%", type0_cell_text(progress_label))
    if not matches:
        return np.nan
    numbers = [float(part.replace("%", "").replace(",", ".").strip()) / 100 for part in matches]
    return float(np.clip(np.mean(numbers), 0, 1))


def type0_identify_top3_pias(df):
    result = pd.Series(False, index=df.index, dtype=bool)
    eligible = df.loc[df["Plant"].map(type0_cell_text).ne("") & df["PIAS (in k EUR)"].notna()]
    for _, group in eligible.groupby("Plant", sort=False):
        indices = group.sort_values("PIAS (in k EUR)", ascending=False, kind="stable").head(3).index
        result.loc[indices] = True
    return result


def type0_format_value(value):
    return "" if pd.isna(value) else f"{float(value):,.10f}".rstrip("0").rstrip(".")


def type0_create_activity_label(row, bold_pias=False):
    """Return (text, bold) chunks so only the top-three PIAS values are bold."""
    project = re.sub(r"^\s*[0-9]+[\W_]*", "", type0_cell_text(row["Project Name"]))
    project = "".join(ch if ch.isalnum() or ch == " " else " " for ch in project)
    project = re.sub(r"\s+", " ", project).strip()[:50].strip()
    plant = re.sub(r"\s+", " ", type0_cell_text(row["Plant"]))[:3].upper()
    number = type0_cell_text(row["Number"])
    parts = []
    if plant:
        parts.append((f"[{plant}]", False))
    if number:
        parts.append((f"[{number}]", False))
    if project:
        parts.append((project, False))
    for field, label in (("PAS (in k EUR)", "S"), ("PIAS (in k EUR)", "I"), ("Priority proposal", "P")):
        value = type0_format_value(row[field])
        if value:
            text = f"[P{value}]" if label == "P" else f"[{label}: {value}]"
            parts.append((text, bool(bold_pias) if label == "I" else False))
    return [(text + (" " if i < len(parts) - 1 else ""), bold) for i, (text, bold) in enumerate(parts)]


def type0_new_project_cutoff(new_projects_since=None):
    """Default to one calendar month before today, independently of reporting date."""
    if new_projects_since is None:
        return pd.Timestamp(date.today()) - pd.DateOffset(months=1)
    return pd.Timestamp(new_projects_since).normalize()


def type0_project_tag_masks(projects, new_projects_since=None):
    """Use the same K/C/N definitions for chart labels and portfolio filters."""
    capex = projects["CAPEX (in k EUR)"]
    return {
        "K": projects["Key Project"].map(type0_cell_text).str.casefold().eq("x"),
        "C": np.isfinite(capex) & capex.gt(0),
        "N": (projects["Creation Date"].ge(type0_new_project_cutoff(new_projects_since))
              if new_projects_since is not None else pd.Series(False, index=projects.index)),
    }


def type0_prepare_gantt_data(df_priority, today=None, new_projects_since=None):
    today = pd.Timestamp(today if today is not None else date.today()).normalize()
    type0_require_columns(df_priority, [
        "Number", "Project Name", "Plant", "Status", "Progress", "Key Project", "CAPEX (in k EUR)",
        "PAS (in k EUR)", "PIAS (in k EUR)", "Priority proposal", *TYPE0_DATE_FIELDS, "Due Date Tracking",
    ], "Cannot create Gantt chart")
    gantt = df_priority.copy().reset_index(drop=True)
    # Keep every selected project, including rows without a drawable timeline.
    gantt["Has_Valid_Dates"] = gantt["Creation Date"].notna() & gantt["Due Date"].notna()
    gantt["Top3_PIAS"] = type0_identify_top3_pias(gantt)
    gantt["Activity_Parts"] = [type0_create_activity_label(row, row["Top3_PIAS"]) for _, row in gantt.iterrows()]
    gantt["Activity"] = gantt["Activity_Parts"].map(lambda parts: "".join(t for t, _ in parts))
    gantt["Progress_Label"] = gantt["Progress"].map(type0_cell_text)
    tags = type0_project_tag_masks(gantt, new_projects_since)
    gantt["Key_Project_Label"] = np.where(tags["K"], "K", "")
    gantt["CAPEX_Label"] = np.where(tags["C"], "C", "")
    gantt["New_Project_Label"] = np.where(tags["N"], "N", "")
    gantt["Project_Label"] = [
        "|".join(label for label in labels if label)
        for labels in zip(gantt["Key_Project_Label"], gantt["CAPEX_Label"],
                          gantt["New_Project_Label"])
    ]
    gantt["Average_Progress"] = gantt["Progress_Label"].map(type0_extract_average_progress).astype(float)
    gantt["Original_Project_Duration_Days"] = (gantt["Original Due Date"] - gantt["Creation Date"]).dt.days
    gantt["Estimated_Remaining_Work_Days"] = gantt["Original_Project_Duration_Days"] * (1 - gantt["Average_Progress"])
    gantt["Days_Remaining_To_Due_Date"] = (gantt["Due Date"] - today).dt.days
    gantt["Timeline_Status"] = "On track"
    at_risk = gantt["Due Date"].ge(today) & gantt["Estimated_Remaining_Work_Days"].gt(gantt["Days_Remaining_To_Due_Date"])
    gantt.loc[at_risk, "Timeline_Status"] = "At risk"
    gantt.loc[gantt["Due Date"].lt(today), "Timeline_Status"] = "Overdue"
    gantt.loc[~gantt["Has_Valid_Dates"], "Timeline_Status"] = "Missing/invalid dates"
    gantt["Activity_ID"] = [f"Activity_{i + 1}" for i in range(len(gantt))]
    gantt["Y_Position"] = np.arange(len(gantt))
    return gantt


def type0_create_due_date_tracking_data(gantt):
    rows, invalid = [], []
    for _, row in gantt.iterrows():
        if not row["Has_Valid_Dates"]:
            continue
        parsed = []
        for value in type0_tracking_entries(row.get("Due Date Tracking", "")):
            tracking_date = type0_parse_date(value)
            if pd.isna(tracking_date):
                invalid.append(type0_cell_text(value))
                continue
            if tracking_date < row["Creation Date"]:
                continue
            latest = row["Latest Due Date"]
            if pd.notna(latest) and tracking_date > latest:
                continue
            if tracking_date not in parsed:
                parsed.append(tracking_date)
        for tracking_date in parsed:
            rows.append({
                "Activity_ID": row["Activity_ID"], "Y_Position": row["Y_Position"],
                "Tracking_Date": tracking_date, "Timeline_Status": row["Timeline_Status"],
            })
    type0_warn_invalid_values(invalid, "Due Date Tracking")
    return pd.DataFrame(rows, columns=["Activity_ID", "Y_Position", "Tracking_Date", "Timeline_Status"])


def type0_calculate_gantt_height_fraction(number_of_activities):
    return float(np.clip(0.30 + number_of_activities * 0.014, 0.30, 0.98))


def type0_text_area(text, size=9, bold=False, colour=GREY):
    return TextArea(text, textprops={
        "fontsize": size, "fontfamily": "DejaVu Sans", "fontweight": "bold" if bold else "normal",
        "color": colour, "parse_math": False,
    })


def type0_activity_box(parts):
    return HPacker(children=[type0_text_area(text, bold=bold) for text, bold in parts], align="baseline", pad=0, sep=0)


def type0_project_label_width_pt(row, renderer, dpi):
    """Required width to the left of the start marker, in physical points."""
    labels = [row[field] for field in ("Project_Label", "Progress_Label") if row[field]]
    if not labels:
        return 0.0
    font = FontProperties(family="DejaVu Sans", size=8.5)
    widths = [
        max(renderer.get_text_width_height_descent(line, font, False)[0] for line in label.split("\n")) * 72 / dpi
        for label in labels
    ]
    return 5.7 + sum(width + 3.0 for width in widths) + 3.0 * (len(labels) - 1)


def type0_add_project_labels(ax, row, renderer):
    # Keep progress closest to the start marker, with indicators in their own badge.
    labels = [row["Project_Label"], row["Progress_Label"]]
    font_size, padding, gap, offset = 8.5, 1.5, 3.0, 5.7
    font = FontProperties(family="DejaVu Sans", size=font_size)
    for label in reversed([value for value in labels if value]):
        width_px = max(renderer.get_text_width_height_descent(line, font, False)[0] for line in label.split("\n"))
        width_pt = width_px * 72 / ax.figure.dpi
        ax.annotate(
            label, xy=(mdates.date2num(row["Creation Date"]), row["Y_Position"]),
            xytext=(-offset - padding, 0), textcoords="offset points",
            ha="right", va="center", fontsize=font_size, color=GREY,
            fontfamily="DejaVu Sans", parse_math=False,
            bbox={"boxstyle": f"round,pad={padding / font_size}", "facecolor": "white", "edgecolor": GREY, "linewidth": 0.8},
            annotation_clip=False, zorder=6,
        )
        offset += width_pt + 2 * padding + gap


def type0_legend_item(text, colour=GREY, symbol="line"):
    drawing = DrawingArea(20, 11, 0, 0)
    if symbol == "triangle":
        line = Line2D([10], [5.5], linestyle="none", marker="^", markersize=7,
                      markerfacecolor="white", markeredgecolor=colour, markeredgewidth=0.8)
    elif symbol == "cross":
        line = Line2D([10], [5.5], linestyle="none", marker="x", markersize=5,
                      markeredgecolor=colour, markeredgewidth=0.8)
    else:
        line = Line2D([0, 20], [5.5, 5.5], color=colour,
                      linewidth=2 if symbol == "line" else 0.8,
                      linestyle="--" if symbol == "today" else "-")
    drawing.add_artist(line)
    return HPacker(children=[drawing, type0_text_area(text)], align="center", pad=0, sep=4)


def type0_add_two_line_legend(fig, y, today, date_label="Today", show_today=True):
    first = HPacker(children=[
        type0_text_area("S = PAS   |   I = PIAS   |   P = Priority   |   K = Key project   |   C = CAPEX   |   N = New project   |   Status"),
        *[type0_legend_item(name, colour) for name, colour in TYPE0_STATUS_COLOURS.items()],
    ], align="center", pad=0, sep=10)
    date_items = [
        type0_legend_item("Original due date", symbol="triangle"),
        type0_legend_item("Due date", symbol="cross"),
    ]
    if show_today:
        date_items.append(type0_legend_item(f"{date_label} ({type0_english_date(today)})", colour=TODAY_COLOUR, symbol="today"))
    second = HPacker(children=date_items, align="center", pad=0, sep=16)
    packed = VPacker(children=[first, second], align="left", pad=0, sep=2)
    artist = AnnotationBbox(packed, (0.5, y), xycoords=fig.transFigure,
                           frameon=False, box_alignment=(0.5, 1), annotation_clip=False)
    fig.add_artist(artist)
    return artist


def type0_timeline_limits(gantt, today):
    timelines = gantt.loc[gantt["Has_Valid_Dates"]]
    dates = [today]
    for field in ("Creation Date", "Due Date", "Original Due Date"):
        dates.extend(timelines[field].dropna().tolist())
    dates.extend(type0_create_due_date_tracking_data(gantt)["Tracking_Date"].tolist())
    if timelines.empty:
        dates.extend([today - pd.Timedelta(days=30), today + pd.Timedelta(days=30)])
    lo, hi = min(dates), max(dates)
    span_days = max((hi - lo).days, 1)
    return (mdates.date2num(lo) - TYPE0_LEFT_DATE_EXPANSION * span_days,
            mdates.date2num(hi) + TYPE0_RIGHT_DATE_EXPANSION * span_days)


def type0_plot_gantt(df_priority, segment, today=None, date_label="Today",
               new_projects_since=None, min_priority=TYPE0_MIN_PRIORITY, max_priority=TYPE0_MAX_PRIORITY,
               total_projects=None, total_pas=None, total_pias=None, show_today=True, prepared=None,
               limits=None, page_label=""):
    today = pd.Timestamp(today if today is not None else date.today()).normalize()
    gantt = prepared if prepared is not None else type0_prepare_gantt_data(df_priority, today, new_projects_since)
    if gantt.empty:
        return None
    timelines = gantt.loc[gantt["Has_Valid_Dates"]]
    tracking = type0_create_due_date_tracking_data(gantt)
    figure_settings = dict(figsize=(TYPE0_A3_WIDTH_MM / 25.4, TYPE0_A3_HEIGHT_MM / 25.4), dpi=100, facecolor="white")
    fig = Figure(**figure_settings)
    FigureCanvasAgg(fig)
    renderer = fig.canvas.get_renderer()
    width_pt = fig.get_figwidth() * 72
    height_pt = fig.get_figheight() * 72
    label_width_pt = 0
    for parts in gantt["Activity_Parts"]:
        width_px = sum(renderer.get_text_width_height_descent(
            text, FontProperties(family="DejaVu Sans", size=9, weight="bold" if bold else "normal"), False
        )[0] for text, bold in parts)
        label_width_pt = max(label_width_pt, width_px * 72 / fig.dpi)
    lower, upper = limits if limits is not None else type0_timeline_limits(gantt, today)

    base_left_pt = 20 + label_width_pt + 12
    right_pt = width_pt - 20
    available_pt = right_pt - base_left_pt
    gutter_pt = 0.0
    for _, row in timelines.iterrows():
        start_fraction = (mdates.date2num(row["Creation Date"]) - lower) / (upper - lower)
        needed_pt = type0_project_label_width_pt(row, renderer, fig.dpi)
        if start_fraction < 1:
            gutter_pt = max(gutter_pt, (needed_pt - available_pt * start_fraction) / (1 - start_fraction))
    # Reserve physical space outside the date axis when a long label extends
    # past it. This avoids overlap without increasing the 10% date expansion.
    gutter_pt = max(0.0, gutter_pt) + 2
    left = (base_left_pt + gutter_pt) / width_pt
    right = right_pt / width_pt
    if right - left < 0.20:
        fig.clear()
        raise ValueError(f"Activity labels for {segment} leave too little chart space; shorten the project names or numbers.")
    height_fraction = type0_calculate_gantt_height_fraction(len(gantt))
    title_y = 1 - 10 / height_pt
    subtitle_y = 1 - 30 / height_pt
    plot_top = 1 - 57 / height_pt
    block_bottom = 1 - height_fraction
    plot_bottom = block_bottom + 107 / height_pt
    ax = fig.add_axes([left, plot_bottom, right - left, plot_top - plot_bottom])
    ax.set_facecolor("white")
    ax.set_axisbelow(True)
    ax.grid(axis="x", color="#E5E5E5", linewidth=0.6)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.set_yticks([])
    ax.set_ylim(len(gantt) - 0.5, -0.5)
    ax.tick_params(axis="x", colors=GREY, labelsize=8, length=0)
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%y"))

    ax.set_xlim(lower, upper)
    year_positions, year_labels = [], []
    for year in range(mdates.num2date(lower).year, mdates.num2date(upper).year + 1):
        year_start = date(year, 1, 1)
        if lower <= mdates.date2num(year_start) <= upper:
            ax.axvline(year_start, color=GREY, alpha=0.55, linewidth=1, zorder=1,
                       label=f"1 Jan {year}")
        visible_start = max(lower, mdates.date2num(year_start))
        visible_end = min(upper, mdates.date2num(date(year + 1, 1, 1)))
        if visible_end > visible_start:
            year_positions.append((visible_start + visible_end) / 2)
            year_labels.append(str(year))
    for tick in ax.get_xticklabels():
        tick.set_rotation(45)
        tick.set_ha("right")

    year_axis = ax.secondary_xaxis("bottom")
    year_axis.spines["bottom"].set_position(("outward", 36))
    for spine in year_axis.spines.values():
        spine.set_visible(False)
    year_axis.set_xticks(year_positions, year_labels)
    year_axis.tick_params(axis="x", colors=GREY, labelsize=9, length=0, pad=3)
    for tick in year_axis.get_xticklabels():
        tick.set_fontweight("bold")

    previous_plant = None
    for _, row in gantt.iterrows():
        y = row["Y_Position"]
        plant = type0_cell_text(row["Plant"])
        if previous_plant is not None and plant != previous_plant:
            ax.axhline(y - 0.5, color=GREY, linewidth=0.7, linestyle=":")
        previous_plant = plant
        label = AnnotationBbox(
            type0_activity_box(row["Activity_Parts"]), (0, y), xycoords=("axes fraction", "data"),
            xybox=(-8 - gutter_pt, 0), boxcoords="offset points", box_alignment=(1, 0.5),
            frameon=False, annotation_clip=False,
        )
        ax.add_artist(label)
        if not row["Has_Valid_Dates"]:
            ax.text(
                0.02, y, TYPE0_MISSING_DATA_LABEL, transform=ax.get_yaxis_transform(),
                ha="left", va="center", fontsize=8.5, color="white",
                fontfamily="DejaVu Sans", parse_math=False,
                bbox={"boxstyle": "round,pad=0.15", "facecolor": TYPE0_STATUS_COLOURS["Overdue"],
                      "edgecolor": "none", "linewidth": 0},
                zorder=6,
            )
            continue
        colour = TYPE0_STATUS_COLOURS[row["Timeline_Status"]]
        ax.plot([row["Creation Date"], row["Due Date"]], [y, y], color=colour,
                linewidth=3.0, solid_capstyle="round", zorder=2)
        start_colour = TYPE0_STATUS_COLOURS["On track"] if row["Status"] == "In Progress" else TYPE0_STATUS_COLOURS["At risk"]
        ax.plot(row["Creation Date"], y, marker="o", markersize=7, markerfacecolor="white",
                markeredgecolor=start_colour, markeredgewidth=1.3, linestyle="none", zorder=4)
        original = row["Original Due Date"]
        if pd.notna(original):
            due_colour = TYPE0_STATUS_COLOURS["Overdue"] if original < today else TYPE0_STATUS_COLOURS["On track"]
            ax.plot(original, y, marker="^", markersize=8, markerfacecolor="white",
                    markeredgecolor=due_colour, markeredgewidth=0.8, linestyle="none", zorder=5)
        type0_add_project_labels(ax, row, renderer)
    for row in tracking.itertuples(index=False):
        due_colour = TYPE0_STATUS_COLOURS["Overdue"] if row.Tracking_Date < today else TYPE0_STATUS_COLOURS["On track"]
        ax.plot(row.Tracking_Date, row.Y_Position, marker="x", markersize=5,
                markeredgecolor=due_colour,
                markeredgewidth=0.8, linestyle="none", zorder=3)
    if show_today:
        ax.axvline(today, color=TODAY_COLOUR, linewidth=0.8, linestyle="--", zorder=1)

    # Compare displayed projects with the same segment scope across all priorities.
    selected_pas = gantt['PAS (in k EUR)'].sum()
    selected_pias = gantt['PIAS (in k EUR)'].sum()
    if total_projects is None:
        total_projects = len(gantt)
    if total_pas is None:
        total_pas = selected_pas
    if total_pias is None:
        total_pias = selected_pias
    subtitle = (
        f"[#Projects: {len(gantt)} / {total_projects}] - "
        f"[Total PAS: {selected_pas:,.0f} / {total_pas:,.0f} k€] "
        f"[Total PIAS: {selected_pias:,.0f} / {total_pias:,.0f} k€] - "
        f"[On track: {gantt['Timeline_Status'].eq('On track').mean():.0%}] "
        f"[At risk: {gantt['Timeline_Status'].eq('At risk').mean():.0%}] "
        f"[Overdue: {gantt['Timeline_Status'].eq('Overdue').mean():.0%}]"
    )
    missing_dates = int((~gantt['Has_Valid_Dates']).sum())
    if missing_dates:
        subtitle += f" - [Missing/invalid dates: {missing_dates}]"
    fig.text(20 / width_pt, title_y,
             f"Type 0 Projects {segment} [Priority: {type0_priority_range_label(min_priority, max_priority)}]{page_label}",
             fontsize=14, fontweight="bold", va="top", color="#222222")
    fig.text(20 / width_pt, subtitle_y, subtitle, fontsize=10, va="top", color=GREY)
    type0_add_two_line_legend(fig, plot_bottom - 68 / height_pt, today, date_label, show_today=show_today)
    return fig


TYPE0_REQUIRED_WEB_COLUMNS = (
    'Type', 'Segment', 'Status', 'Plant', 'Priority proposal', 'Number', 'Project Name',
    'Product Manager', 'Project Manager', 'Progress', 'Key Project',
    'CAPEX (in k EUR)', 'PAS (in k EUR)', 'PIAS (in k EUR)',
    'CM% (as % of sales)', *TYPE0_DATE_FIELDS, 'Due Date Tracking',
)


TYPE0_DATA_VALIDATION_WARNING = 'Data Validation extension is not supported and will be removed'


def type0_read_web_workbook(content, sheet_name=TYPE0_SHEET_NAME, header_row=6):
    """Read uploaded bytes, retaining original cells for validation messages."""
    if header_row < 1:
        raise ValueError('The header row must be 1 or greater.')
    raw = pd.read_excel(
        BytesIO(content), sheet_name=sheet_name, skiprows=header_row - 1,
        header=None, dtype=object, keep_default_na=False, engine='openpyxl',
    )
    if raw.empty:
        raise ValueError('The selected worksheet/header row contains no column headings.')
    headings = [re.sub(r'\s+', ' ', type0_cell_text(value)) for value in raw.iloc[0]]
    named = [value for value in headings if value]
    if len(named) != len(set(named)):
        raise ValueError('The header row has duplicate column names. Fix the column headings in the workbook.')
    raw = raw.iloc[1:].copy()
    raw.columns = [name or f'Unnamed column {index + 1}' for index, name in enumerate(headings)]
    type0_require_columns(raw, TYPE0_REQUIRED_WEB_COLUMNS, 'Check the worksheet and header row')
    raw = raw.loc[raw.apply(lambda col: col.map(type0_cell_text)).ne('').any(axis=1)].reset_index(drop=True)
    result = type0_convert_type0_data(raw)
    for field in ('Type', 'Segment', 'Status', 'Plant', 'Product Manager', 'Project Manager'):
        result[field] = result[field].map(type0_cell_text)
    return result


def type0_select_web_projects(df, segments, min_priority, max_priority, plants=None,
                        statuses=TYPE0_DEFAULT_STATUSES, key_projects_only=False,
                        projects_created_since=None, types=None, apply_priority=True,
                        project_tags=None, new_projects_since=None):
    completed_only = type0_is_completed_only(statuses)
    min_priority, max_priority = map(type0_parse_priority_limit, (min_priority, max_priority))
    if (apply_priority and not completed_only and min_priority is not None
            and max_priority is not None and min_priority > max_priority):
        raise ValueError('Minimum priority cannot exceed maximum priority.')
    if types is not None:
        df = df.loc[df['Type'].map(type0_cell_text).isin(types)]
    if projects_created_since is not None and completed_only:
        cutoff = pd.Timestamp(projects_created_since).normalize()
        df = df.loc[df['Creation Date'].ge(cutoff)]
    if project_tags is not None:
        tags = type0_project_tag_masks(df, new_projects_since)
        if set(project_tags) - tags.keys():
            raise ValueError('Project tags must be K, C, or N.')
        matches = pd.Series(False, index=df.index)
        for tag in project_tags:
            matches |= tags[tag]
        df = df.loc[matches]
    result = {}
    for segment in segments:
        selected = type0_get_priority_projects(df, segment, min_priority, max_priority,
                                         statuses=statuses, apply_priority=apply_priority)
        if plants is not None:
            selected = selected.loc[selected['Plant'].map(type0_cell_text).isin(plants)].reset_index(drop=True)
        if key_projects_only:
            selected = selected.loc[selected['Key Project'].map(type0_cell_text).ne('')].reset_index(drop=True)
        result[segment] = selected
    return result


def zip_outputs(files):
    buffer = BytesIO()
    with ZipFile(buffer, 'w', compression=ZIP_DEFLATED) as archive:
        for filename, content in files.items():
            archive.writestr(filename, content)
    return buffer.getvalue()


def type0_build_web_results(projects_by_segment, today, context=(), dpi=TYPE0_PNG_DPI,
                      new_projects_since=None, min_priority=TYPE0_MIN_PRIORITY, max_priority=TYPE0_MAX_PRIORITY,
                      all_priorities_by_segment=None, number_counts=None, show_today=True, rows_per_page=30):
    """Build in-memory outputs. The UI holds a shared Matplotlib render lock.

    show_today controls the reporting-date line and its legend entry only.
    """
    today = pd.Timestamp(today).normalize()
    if number_counts is None:
        selected = pd.concat(projects_by_segment.values(), ignore_index=True) if projects_by_segment else pd.DataFrame(columns=["Number"])
        number_counts = type0_count_project_numbers(selected)
    checks = {segment: type0_check_priority_projects(frame, segment, today, number_counts=number_counts)
              for segment, frame in projects_by_segment.items()}
    chart_projects = dict(projects_by_segment)
    if projects_by_segment:
        chart_projects = {"All": pd.concat(projects_by_segment.values(), ignore_index=True), **chart_projects}
        if all_priorities_by_segment is not None:
            all_priorities_by_segment = {
                "All": pd.concat(all_priorities_by_segment.values(), ignore_index=True),
                **all_priorities_by_segment,
            }
    charts, chart_pages, messages, summaries = {}, {}, {}, []
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter('always', UserWarning)
        for segment, selected in chart_projects.items():
            chart_pages[segment] = []
            gantt = type0_prepare_gantt_data(selected, today, new_projects_since)
            if segment in projects_by_segment:
                summaries.append({
                    'Segment': segment, 'Selected': len(selected), 'Chart rows': len(gantt),
                    'Missing/invalid dates': int((~gantt['Has_Valid_Dates']).sum()),
                    'At risk': int(gantt['Timeline_Status'].eq('At risk').sum()),
                    'Overdue': int(gantt['Timeline_Status'].eq('Overdue').sum()),
                    'PAS (k EUR)': float(selected['PAS (in k EUR)'].sum()),
                    'PIAS (k EUR)': float(selected['PIAS (in k EUR)'].sum()),
                    'Actions': len(checks[segment]),
                })
            if selected.empty:
                messages[segment] = 'No projects match the selected filters.'
                continue
            try:
                all_priorities = (all_priorities_by_segment[segment]
                                  if all_priorities_by_segment is not None else selected)
                limits = type0_timeline_limits(gantt, today)
                page_count = math.ceil(len(selected) / rows_per_page)
                segment_charts = {}
                for page in range(1, page_count + 1):
                    start = (page - 1) * rows_per_page
                    current = gantt.iloc[start:start + rows_per_page].copy()
                    current["Y_Position"] = np.arange(len(current))
                    chart = type0_plot_gantt(selected.iloc[start:start + rows_per_page], segment, today,
                                       date_label='As of', new_projects_since=new_projects_since,
                                       min_priority=min_priority, max_priority=max_priority,
                                       total_projects=len(all_priorities),
                                       total_pas=all_priorities['PAS (in k EUR)'].sum(),
                                       total_pias=all_priorities['PIAS (in k EUR)'].sum(),
                                       show_today=show_today, prepared=current, limits=limits,
                                       page_label=f" · Page {page} of {page_count}" if page_count > 1 else "")
                    suffix = f"_{page}" if page_count > 1 else ""
                    segment_charts[f'Priority_Projects_{segment}{suffix}.png'] = figure_png(chart, dpi)
                charts.update(segment_charts)
                chart_pages[segment] = list(segment_charts)
                missing_dates = int((~gantt['Has_Valid_Dates']).sum())
                if missing_dates:
                    messages[segment] = (
                        f'{missing_dates} project(s) shown with full activity labels and a red "{TYPE0_MISSING_DATA_LABEL}" label '
                        'because of missing/invalid creation or due dates. See Checks & actions for details.'
                    )
            except (ValueError, OverflowError) as exc:
                messages[segment] = f'Chart could not be generated: {exc}'
    chart_notes = [f'{segment}: {message}' for segment, message in messages.items()]
    notes = list(dict.fromkeys(str(item.message) for item in caught))
    report = type0_priority_checks_text(checks, projects_by_segment, context)
    if chart_notes or notes:
        report += '\nChart and conversion notes\n' + '\n'.join(chart_notes + notes) + '\n'
    issue_tables = [frame.assign(Segment=segment) for segment, frame in checks.items()]
    issues = pd.concat(issue_tables, ignore_index=True) if issue_tables else pd.DataFrame()
    return {
        'charts': charts, 'chart_pages': chart_pages, 'messages': messages, 'summary': pd.DataFrame(summaries),
        'issues': issues, 'report': report, 'warnings': notes,
    }


# IPD workbook rules and Gantt renderer

IPD_PROJECT = "Project"
IPD_DESCRIPTION = "Project description"
IPD_START = "Gate 1 (PS) Project"
IPD_END = "Commercial Launch date Project"
IPD_END_LABEL = "Launch date"
IPD_PHASE = "Current phase Project"
IPD_PHASE_ORDER = ("Scoping", "Opportunity Assessment", "Design", "Industrialize",
               "Launch Preparation", "Post Launch")
IPD_PHASE_ALIASES = {"opportunity assessment / desing": "Opportunity Assessment",
                 "opportunity assessment / design": "Opportunity Assessment",
                 "develop / industrialize": "Industrialize"}
IPD_ACTIVITY_METRICS = {"PAS": "PAS@5YM (UA) Project", "PIAS": "PIAS@5YM (UA) Project",
                    "PIAC": "PIAC@5YM (UA) Project"}
IPD_LAUNCH_METRICS = {"S": "PAS@5YM (UA) Project", "I": "PIAS@5YM (UA) Project"}
IPD_LAUNCH_VALUES_PER_MILLION = 1_000
IPD_PIPELINE_UNIT = "k EUR"
IPD_DEFAULT_HEADER_ROW = 2
IPD_FINANCIAL_METRICS = ("PAC", "PAS", "PIAC", "PIAS")
IPD_YEAR_COLUMNS = [f"Year {year}" for year in range(1, 12)]
IPD_FINANCIAL_COLUMNS = [f"{metric}_{year}" for metric in IPD_FINANCIAL_METRICS for year in IPD_YEAR_COLUMNS]
IPD_BAR_COLOUR = "#32AAFF"
IPD_EVENT_GREEN = "#52DEAD"
IPD_PAST_GATE_FILL = "#DCF8EF"
IPD_LAUNCH_OUTLINE = "#FA6157"
IPD_PAST_LAUNCH_FILL = "#FEDFDD"
IPD_EVENT_SIZE = 9
IPD_ERROR_COLOUR = "#C43C39"
IPD_OPTIONAL_GATES = (2, 3, 4, 5)
IPD_PROJECT_COLUMNS = ["Excel row", IPD_PROJECT, IPD_DESCRIPTION, "Activity", "Start", "End", "Valid dates"]
IPD_MILESTONE_COLUMNS = ["Excel row", "Activity", "Gate", "Column", "Date"]
IPD_ISSUE_COLUMNS = ["Excel row", "Activity", "Column", "Value", "Issue"]


def ipd_cell_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return str(int(value))
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ")).strip()


def ipd_format_activity_metric(value):
    text = ipd_cell_text(value)
    if not text:
        return "—"
    try:
        number = Decimal(text)
    except InvalidOperation:
        return text
    if not number.is_finite():
        return text
    rounded = number.to_integral_value(rounding=ROUND_HALF_UP)
    return format(rounded, "f") if rounded else "0"


def ipd_normalise_phase(value):
    name = re.sub(r"\s*/\s*", " / ", ipd_cell_text(value))
    if name.casefold() in IPD_PHASE_ALIASES:
        return IPD_PHASE_ALIASES[name.casefold()]
    return next((phase for phase in IPD_PHASE_ORDER if phase.casefold() == name.casefold()), name)


def ipd_activity_label(tag, description, row):
    display_tag = re.sub(r"^PRJ[\s_-]*", "", tag, flags=re.IGNORECASE)
    values = {ipd_cell_text(column).casefold(): value for column, value in row.items()}
    metrics = "".join(f"[{metric}: {ipd_format_activity_metric(values.get(column.casefold()))}]"
                      for metric, column in IPD_ACTIVITY_METRICS.items())
    priority = ipd_cell_text(values.get("priority project"))
    first_line = f"[{display_tag or 'Missing Project'}] {description or '(Missing project description)'}"
    if priority:
        first_line += f" [P{priority}]"
    return f"{first_line}\n{metrics}"


def ipd_default_start_date(today=None):
    return (pd.Timestamp(today or date.today()) - pd.DateOffset(years=1)).date()


def ipd_project_phases(projects):
    column = next((column for column in projects if ipd_cell_text(column).casefold() == IPD_PHASE.casefold()), None)
    if column is None:
        return pd.Series("", index=projects.index)

    return projects[column].map(ipd_normalise_phase)


def ipd_order_projects_by_phase(projects):
    phases = ipd_project_phases(projects).tolist()
    ranks = {phase: rank for rank, phase in enumerate(IPD_PHASE_ORDER)}
    ranks[""] = len(IPD_PHASE_ORDER) + 1
    order = sorted(range(len(projects)),
                   key=lambda index: (ranks.get(phases[index], len(IPD_PHASE_ORDER)), phases[index].casefold()))
    return projects.iloc[order]


def ipd_parse_date(value):
    """Parse explicit English month abbreviations without depending on locale."""
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, (date, datetime, pd.Timestamp)):
        stamp = pd.Timestamp(value)
        if stamp.tzinfo is not None:
            stamp = stamp.tz_localize(None)
        return stamp.normalize()
    match = re.fullmatch(r"(\d{1,2})/([A-Za-z]{3})/(\d{4})", ipd_cell_text(value))
    if match:
        day, month, year = match.groups()
        try:
            return pd.Timestamp(date(int(year), MONTHS.index(month.title()) + 1, int(day)))
        except (ValueError, OverflowError):
            pass
    return pd.NaT


def ipd_format_date(value) -> str:
    if pd.isna(value):
        return ""
    return f"{value.day:02d}/{MONTHS[value.month - 1]}/{value.year:04d}"


def ipd_resolve_columns(columns):
    """Allow whitespace, case, parentheses and suffixes on the gate headings."""
    lookup = {ipd_cell_text(column).casefold(): column for column in columns}
    if len(lookup) != len(columns):
        raise ValueError("Duplicate column headings. Give each column a unique name.")
    fields = {}
    for required in (IPD_PROJECT, IPD_DESCRIPTION, IPD_END):
        if required.casefold() not in lookup:
            raise ValueError(f"Missing column: {required}. Check the worksheet and header row.")
        fields[required] = lookup[required.casefold()]
    # Both 'Gate 1 (PS) Project' and the earlier 'Gate 1 PS (Project)' resolve alike.
    gate_headings = {column: ipd_cell_text(re.sub(r"[()]", " ", ipd_cell_text(column))).casefold()
                     for column in columns}
    start_columns = [column for column, heading in gate_headings.items()
                     if re.match(r"gate 1 ps\b", heading)]
    preferred = [column for column in start_columns
                 if gate_headings[column] == "gate 1 ps project"]
    if len(preferred) == 1:
        fields[IPD_START] = preferred[0]
    elif len(start_columns) == 1:
        fields[IPD_START] = start_columns[0]
    else:
        raise ValueError(f"Expected one '{IPD_START}' start-date column. Check the worksheet and header row.")
    gates = {gate: [column for column, heading in gate_headings.items()
                    if re.match(rf"gate {gate} ps\b", heading)]
             for gate in IPD_OPTIONAL_GATES}
    return fields, gates


def ipd_prepare_projects(raw, header_row=IPD_DEFAULT_HEADER_ROW, source_rows=None):
    """Keep workbook order and source row numbers, including incomplete projects."""
    raw = raw.copy()
    raw.columns = [ipd_cell_text(column) for column in raw.columns]
    fields, gates = ipd_resolve_columns(list(raw.columns))
    projects, milestones, issues = [], [], []
    extra_columns = [column for column in raw if column not in {*fields.values(), *IPD_PROJECT_COLUMNS}]
    source_rows = source_rows if source_rows is not None else range(header_row + 1, header_row + 1 + len(raw))
    for excel_row, (_, row) in zip(source_rows, raw.iterrows()):
        if not any(ipd_cell_text(value) for value in row):
            continue
        tag = ipd_cell_text(row[fields[IPD_PROJECT]])
        description = ipd_cell_text(row[fields[IPD_DESCRIPTION]])
        activity = ipd_activity_label(tag, description, row)

        def issue(column, value, message):
            issues.append({"Excel row": excel_row, "Activity": activity,
                           "Column": column, "Value": ipd_cell_text(value), "Issue": message})

        for field, value in ((IPD_PROJECT, tag), (IPD_DESCRIPTION, description)):
            if not value:
                issue(fields[field], value, f"Add {field}.")
        dates = {}
        for field in (IPD_START, IPD_END):
            column = fields[field]
            value = row[column]
            dates[field] = ipd_parse_date(value)
            if pd.isna(dates[field]):
                issue(column, value, "Missing date." if not ipd_cell_text(value)
                      else "Invalid date. Use DD/MMM/YYYY, for example 30/Sep/2026.")
        start, end = dates[IPD_START], dates[IPD_END]
        valid = pd.notna(start) and pd.notna(end)
        if valid and end < start:
            valid = False
            issue(fields[IPD_END], row[fields[IPD_END]], f"{IPD_END_LABEL} is before Gate 1 PS.")
        projects.append({"Excel row": excel_row, IPD_PROJECT: tag, IPD_DESCRIPTION: description,
                         "Activity": activity, "Start": start, "End": end, "Valid dates": valid,
                         **{column: row[column] for column in extra_columns}})
        for gate, columns in gates.items():
            for column in columns:
                value = row[column]
                if not ipd_cell_text(value):
                    continue
                milestone_date = ipd_parse_date(value)
                if pd.isna(milestone_date):
                    issue(column, value, "Invalid milestone date. Use DD/MMM/YYYY.")
                    continue
                milestones.append({"Excel row": excel_row, "Activity": activity,
                                   "Gate": gate, "Column": column, "Date": milestone_date})
                if valid and not start <= milestone_date <= end:
                    issue(column, value, "Milestone is outside the project dates; still shown in the chart.")
    return (pd.DataFrame(projects, columns=[*IPD_PROJECT_COLUMNS, *extra_columns]),
            pd.DataFrame(milestones, columns=IPD_MILESTONE_COLUMNS),
            pd.DataFrame(issues, columns=IPD_ISSUE_COLUMNS))


def ipd_section_metric(value):
    """Recognise a section title, optionally followed by units or a subtitle."""
    match = re.fullmatch(r"(PAC|PAS|PIAC|PIAS)(?:\s*[:(\-–].*)?", ipd_cell_text(value), re.IGNORECASE)
    return match.group(1).upper() if match else None


def ipd_financial_value(value):
    """Keep numeric values in the source units; never guess decimal separators."""
    if not ipd_cell_text(value):
        return pd.NA
    if isinstance(value, (bool, date, datetime, pd.Timestamp)):
        raise ValueError("Expected a number.")
    try:
        number = float(value)
        if math.isfinite(number):
            return number
    except (TypeError, ValueError, OverflowError):
        pass
    raise ValueError("Expected a number.")


def ipd_financial_year_mapping(headings):
    """Resolve labelled years without treating neighbouring metadata as amounts."""
    calendar = [(int(heading), index) for index, heading in enumerate(headings)
                if re.fullmatch(r"[12]\d{3}", heading)]
    relative = [(int(match.group(1)), index) for index, heading in enumerate(headings)
                if (match := re.fullmatch(r"Year\s+(\d+)", heading, re.IGNORECASE))]
    if not calendar and not relative:
        return {}  # Project-only workbooks need no yearly financial columns.
    years = sorted(calendar or relative)
    expected = list(range(years[0][0], years[0][0] + 11)) if calendar else list(range(1, 12))
    if (calendar and relative) or [year for year, _ in years] != expected:
        raise ValueError("Financial columns must have eleven unique consecutive calendar-year headings "
                         "(for example 2026–2036), or Year 1–Year 11. Check the header row.")
    return {index: year for (_, index), year in zip(years, IPD_YEAR_COLUMNS)}


def ipd_consolidate_sections(frame, header_row, initial_metric=None):
    """One row per Project; repeated metric rows are merged, never added together.

    Section titles apply to all following projects until the next title. Blank
    shared fields can be filled from later sections. Conflicts retain the first
    nonblank value and include the conflicting source cell in the issue table.
    """
    fields, gates = ipd_resolve_columns(list(frame.columns))
    date_columns = {fields[IPD_START], fields[IPD_END], *(column for columns in gates.values() for column in columns)}
    # Type identifies the financial section; Total sum belongs to that metric.
    # Neither is a shared project attribute or an individual year's amount.
    shared_columns = [column for column in frame if column not in IPD_YEAR_COLUMNS
                      and column.casefold() not in {"type", "total sum"}
                      and not column.startswith("__column_")]
    records, rows_by_project, problems = {}, {}, []
    current_metric = initial_metric

    for excel_row, (_, row) in enumerate(frame.iterrows(), start=header_row + 1):
        if not any(ipd_cell_text(value) for value in row):
            continue
        tag, description = ipd_cell_text(row[fields[IPD_PROJECT]]), ipd_cell_text(row[fields[IPD_DESCRIPTION]])
        # Some exports repeat the header at the start of every section.
        if tag.casefold() == IPD_PROJECT.casefold() and description.casefold() == IPD_DESCRIPTION.casefold():
            continue
        titles = {metric for column, value in row.items() if column not in IPD_YEAR_COLUMNS
                  and (metric := ipd_section_metric(value)) is not None}
        has_project_details = bool(
            (description and not ipd_section_metric(description))
            or any(ipd_cell_text(row[column]) and not ipd_section_metric(row[column]) for column in date_columns))
        if titles and not has_project_details and (not tag or ipd_section_metric(tag)):
            if len(titles) != 1:
                raise ValueError(f"Multiple financial section labels on Excel row {excel_row}.")
            current_metric = titles.pop()
            continue
        if not tag and not description and not any(ipd_cell_text(row[column]) for column in date_columns):
            continue  # Spacer or section total without a project identity.
        if (re.fullmatch(r"(?:grand total|subtotal|total)(?:\s+(?:PAC|PAS|PIAC|PIAS))?", tag or description,
                         flags=re.IGNORECASE)
                and not any(ipd_cell_text(row[column]) for column in date_columns)):
            continue
        metric = current_metric
        key = ("project", tag) if tag else ("row", excel_row)
        activity = ipd_activity_label(tag, description, row)

        def issue(column, value, message):
            problems.append({"Excel row": excel_row, "Activity": activity,
                             "Column": column, "Value": ipd_cell_text(value), "Issue": message})

        if key not in records:
            records[key] = {column: "" for column in shared_columns}
            records[key].update(dict.fromkeys(IPD_FINANCIAL_COLUMNS, pd.NA))
            rows_by_project[key] = []
        record = records[key]
        rows_by_project[key].append(excel_row)
        for column in shared_columns:
            incoming, previous = row[column], record[column]
            if not ipd_cell_text(incoming):
                continue
            if not ipd_cell_text(previous):
                record[column] = incoming
                continue
            same = ipd_cell_text(previous) == ipd_cell_text(incoming)
            if column in date_columns:
                first_date, next_date = ipd_parse_date(previous), ipd_parse_date(incoming)
                same = same or (pd.notna(first_date) and pd.notna(next_date) and first_date == next_date)
            if not same:
                issue(column, incoming, f"Conflicting project value across sections; kept '{ipd_cell_text(previous)}'.")
        for year in IPD_YEAR_COLUMNS:
            value = row[year]
            if not ipd_cell_text(value):
                continue
            if metric is None:
                issue(year, value, "Financial value has no PAC/PAS/PIAC/PIAS section label; not assigned to a metric.")
                continue
            column = f"{metric}_{year}"
            try:
                number = ipd_financial_value(value)
            except ValueError:
                issue(column, value, "Invalid financial value. Use a numeric Excel cell; value left empty.")
                continue
            previous = record[column]
            if pd.isna(previous):
                record[column] = number
            elif previous != number:
                issue(column, value, f"Conflicting value for this project and metric; kept {previous:g}. Values are not summed.")

    consolidated = pd.DataFrame(records.values(), columns=[*shared_columns, *IPD_FINANCIAL_COLUMNS])
    consolidated["Source rows"] = [", ".join(map(str, rows)) for rows in rows_by_project.values()]
    for column in IPD_FINANCIAL_COLUMNS:
        consolidated[column] = pd.array(consolidated[column], dtype="Float64")
    source_rows = [rows[0] for rows in rows_by_project.values()]
    return consolidated, source_rows, pd.DataFrame(problems, columns=IPD_ISSUE_COLUMNS)


def ipd_read_workbook(content, sheet_name, header_row=IPD_DEFAULT_HEADER_ROW):
    if header_row < 1:
        raise ValueError("Header row must be 1 or greater.")
    # Read the headings as data so pandas cannot silently rename duplicates.
    raw = pd.read_excel(BytesIO(content), sheet_name=sheet_name, header=None, dtype=object,
                        keep_default_na=False, engine="openpyxl")
    if len(raw) < header_row:
        raise ValueError("No column headings found at this header row.")
    headings = [ipd_cell_text(value) or f"__column_{index + 1}"
                for index, value in enumerate(raw.iloc[header_row - 1])]
    for index, year in ipd_financial_year_mapping(headings).items():
        headings[index] = year
    frame = raw.iloc[header_row:].copy()
    frame.columns = headings
    for year in IPD_YEAR_COLUMNS:
        if year not in frame:
            frame[year] = ""
    initial_titles = [metric for value in raw.iloc[:header_row - 1].to_numpy().flat
                      if (metric := ipd_section_metric(value)) is not None]
    consolidated, source_rows, section_issues = ipd_consolidate_sections(
        frame, header_row, initial_metric=initial_titles[-1] if initial_titles else None)
    projects, milestones, issues = ipd_prepare_projects(consolidated, header_row, source_rows)
    for column in IPD_FINANCIAL_COLUMNS:
        projects[column] = pd.array(projects[column], dtype="Float64")
    return projects, milestones, pd.concat([section_issues, issues], ignore_index=True)


def ipd_pipeline_values(projects):
    """Sum @5YM values and count contributing projects by launch year and segment."""
    columns = ["Year", "Metric", "Segment", f"Value ({IPD_PIPELINE_UNIT})", "Project count"]
    projects = projects.loc[projects["End"].notna()]
    if projects.empty:
        return pd.DataFrame(columns=columns)
    lookup = {ipd_cell_text(column).casefold(): column for column in projects}
    segment_column = lookup.get("market segment")
    segments = (projects[segment_column].map(ipd_cell_text).replace("", "(Blank)")
                if segment_column is not None else pd.Series("All projects", index=projects.index))
    amounts = {}
    for metric in ("PAS", "PIAS"):
        column = lookup.get(IPD_ACTIVITY_METRICS[metric].casefold())
        numbers = []
        for value in projects[column] if column is not None else [None] * len(projects):
            try:
                numbers.append(ipd_financial_value(value))
            except ValueError:
                numbers.append(pd.NA)
        amounts[metric] = pd.Series(numbers, index=projects.index, dtype="Float64")
    if not any(values.notna().any() for values in amounts.values()):
        return pd.DataFrame(columns=columns)
    launch_years = projects["End"].dt.year
    records = []
    for year in range(int(launch_years.min()), int(launch_years.max()) + 1):
        launched = launch_years.eq(year)
        for metric in ("PAS", "PIAS"):
            for segment in sorted(segments.unique(), key=str.casefold):
                selected = segments.eq(segment) & launched
                contributions = amounts[metric].loc[selected]
                total = contributions.sum(min_count=1) if launched.any() else 0.0
                records.append((year, metric, segment, total, int(contributions.count())))
    return pd.DataFrame(records, columns=columns)


def ipd_pipeline_segment_label(amount, metric, project_count=None, *, value_decimals=0):
    """Label a visible segment, adding the contributor count only for PAS."""
    if pd.isna(amount) or amount == 0:
        return ""
    label = f"{amount:,.{value_decimals}f}"
    if metric == "PAS" and project_count is not None and pd.notna(project_count):
        label += f" (# {int(project_count)})"
    return label


def ipd_plot_pipeline(values, *, segment_colours=None, value_decimals=2, label_segments=False,
                  legend_below=False, value_fontweight="normal", title="IPD Pipeline"):
    """Draw PAS beside PIAS for each year, stacked by market segment."""
    if values.empty:
        raise ValueError("There are no PAS/PIAS pipeline values to plot.")
    years = values["Year"].unique().tolist()
    segments = values["Segment"].unique().tolist()
    # Each segment has the same colour in both metrics and every year.
    segment_colours = {ipd_cell_text(segment).casefold(): colour
                       for segment, colour in (segment_colours or {}).items()}
    colours = ("#32AAFF", "#52B788", "#F4A261", "#9B72CF", "#E76F51",
               "#36A6A6", "#D6AD30", "#CF79A3", "#687DA8", "#8C9561")
    legend_rows = math.ceil(len(segments) / 5)
    fig = Figure(figsize=(max(12, len(years) * 1.7), 6 + legend_rows * 0.3), facecolor="white")
    FigureCanvasAgg(fig)
    ax = fig.add_subplot()
    fig.suptitle(title, fontsize=16, fontweight="bold", color=GREY, y=0.98)
    bars = [(year, metric) for year in years for metric in ("PAS", "PIAS")]
    positions = [index + offset for index in range(len(years)) for offset in (-0.2, 0.2)]
    positive, negative = [0.0] * len(bars), [0.0] * len(bars)
    totals, known = [0.0] * len(bars), [False] * len(bars)
    lookup = values.set_index(["Year", "Metric", "Segment"])
    for index, segment in enumerate(segments):
        heights, bottoms = [], []
        for bar, (year, metric) in enumerate(bars):
            amount = lookup.loc[(year, metric, segment), f"Value ({IPD_PIPELINE_UNIT})"]
            valid = pd.notna(amount)
            height = float(amount) if valid else 0.0
            heights.append(height)
            bottoms.append(positive[bar] if height >= 0 else negative[bar])
            if height >= 0:
                positive[bar] += height
            else:
                negative[bar] += height
            totals[bar] += height
            known[bar] |= valid
        ax.bar(positions, heights, width=0.35, bottom=bottoms, label=segment,
               color=segment_colours.get(ipd_cell_text(segment).casefold(), colours[index % len(colours)]),
               edgecolor="white", linewidth=0.5, zorder=3)
    counts = (values.groupby(["Year", "Metric"])["Project count"].sum()
              if "Project count" in values else None)
    for (year, metric), x, top, total, valid in zip(bars, positions, positive, totals, known):
        label = f"{total:,.{value_decimals}f}" if valid else "—"
        if value_decimals:
            label = label.rstrip("0").rstrip(".")
        if metric == "PAS" and counts is not None:
            label += f" (# {int(counts.loc[(year, metric)])})"
        ax.annotate(label, (x, top), xytext=(0, 6), textcoords="offset points",
                    ha="center", va="bottom", fontsize=8, color="black", fontweight=value_fontweight)
    if label_segments:
        for segment, container in zip(segments, ax.containers):
            labels = [ipd_pipeline_segment_label(
                bar.get_height(), metric,
                lookup.loc[(year, metric, segment), "Project count"] if "Project count" in lookup else None,
                value_decimals=value_decimals) for bar, (year, metric) in zip(container, bars)]
            ax.bar_label(container, labels=labels, label_type="center", fontsize=8, color="white",
                         fontweight=value_fontweight)
    ax.set_xticks(positions, [metric for _, metric in bars], fontsize=8, color=GREY)
    for index, year in enumerate(years):
        ax.text(index, -0.09, year, transform=ax.get_xaxis_transform(), ha="center",
                va="top", fontsize=10, color=GREY, fontweight="bold")
    ax.set_xlabel("Year", labelpad=38, color=GREY)
    ax.set_ylabel(f"Value ({IPD_PIPELINE_UNIT})", color=GREY)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:,.0f}"))
    ax.tick_params(axis="both", length=0, colors=GREY)
    ax.grid(axis="y", color="#E8EBEF", linewidth=0.8, zorder=0)
    ax.axhline(0, color="#B8BEC7", linewidth=0.8, zorder=2)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.set_xlim(-0.65, len(years) - 0.35)
    lower, upper = min(negative), max(positive)
    span = upper - lower or 1
    ax.set_ylim(lower - (0.08 * span if lower < 0 else 0), upper + 0.18 * span)
    if legend_below:
        legend_location, legend_anchor = "lower center", (0.5, 0.02)
        layout_rect = (0, 0.08 + (legend_rows - 1) * 0.035, 1, 0.96)
    else:
        legend_location, legend_anchor = "upper center", (0.5, 0.93)
        layout_rect = (0, 0, 1, 0.86 - (legend_rows - 1) * 0.035)
    fig.legend(*ax.get_legend_handles_labels(), loc=legend_location, bbox_to_anchor=legend_anchor,
               ncol=min(5, len(segments)), frameon=False, fontsize=9)
    fig.tight_layout(rect=layout_rect, pad=1.3)
    return fig


def ipd_projects_with_gates_since(projects, milestones, start_date):
    """Keep projects with a valid G1–G5 on or after the inclusive cutoff."""
    cutoff = pd.Timestamp(start_date).normalize()
    later_gates = milestones.loc[
        milestones["Gate"].isin(IPD_OPTIONAL_GATES) & milestones["Date"].ge(cutoff), "Excel row"]
    return projects.loc[projects["Start"].ge(cutoff) | projects["Excel row"].isin(later_gates)].copy()


def ipd_timeline_limits(projects, milestones, today=None, start_date=None):
    values = [*projects["Start"].dropna(), *projects["End"].dropna(), *milestones["Date"].dropna()]
    if today is not None:
        values.append(pd.Timestamp(today))
    if not values:
        values = [pd.Timestamp(date.today())]
    lower, upper = mdates.date2num(min(values)), mdates.date2num(max(values))
    if start_date is not None:
        lower = mdates.date2num(pd.Timestamp(start_date))
        upper = max(upper, lower + 30)
    padding = max((upper - lower) * 0.04, 15)
    return (lower if start_date is not None else lower - padding), upper + padding


def ipd_yearly_launch_labels(projects, years, limits, reference_date=None, historical_projects=None):
    """Summarise full past years and visible current/future launches.

    Historical projects retain the sidebar filters, before the gate cutoff and
    pagination, so a partly displayed past year still covers January–December.
    """
    current_year = pd.Timestamp(reference_date or date.today()).year
    historical_projects = projects if historical_projects is None else historical_projects
    dates = projects["End"].map(lambda stamp: mdates.date2num(stamp) if pd.notna(stamp) else math.nan)
    launches = projects.loc[dates.between(*limits)]
    labels = []
    for year in years:
        source = historical_projects if year < current_year else launches
        annual = source.loc[source["End"].dt.year.eq(year)]
        lookup = {ipd_cell_text(column).casefold(): column for column in annual}
        parts = [f"[#L: {len(annual)}]"]
        for metric, heading in IPD_LAUNCH_METRICS.items():
            column = lookup.get(heading.casefold())
            values = []
            if column is not None:
                for value in annual[column]:
                    try:
                        number = ipd_financial_value(value)
                    except ValueError:
                        continue
                    if pd.notna(number):
                        values.append(number)
            # Distinguish an unknown financial value from a year with no launches.
            total = f"{math.fsum(values) / IPD_LAUNCH_VALUES_PER_MILLION:.1f}M" if values or annual.empty else "—"
            parts.append(f"[{metric}: {total}]")
        labels.append(parts)
    return labels


def ipd_plot_gantt(projects, milestones, today=None, limits=None, page_label="", start_date=None,
               reference_date=None, group_by_phase=False, historical_projects=None):
    """Draw a headless chart in the visual style of Type_0_Gantt.py."""
    if projects.empty:
        raise ValueError("There are no projects to plot.")
    historical_projects = projects if historical_projects is None else historical_projects
    if start_date is not None:
        projects = ipd_projects_with_gates_since(projects, milestones, start_date)
        if projects.empty:
            raise ValueError("No projects have a G1–G5 milestone on or after the selected Start date.")
    milestones = milestones.loc[milestones["Excel row"].isin(projects["Excel row"])]
    if group_by_phase:
        projects = ipd_order_projects_by_phase(projects)
        phases = ipd_project_phases(projects).tolist()
        phase_starts = {index: phase or "Unspecified" for index, phase in enumerate(phases)
                        if index == 0 or phase.casefold() != phases[index - 1].casefold()}
    else:
        phase_starts = {}
    reference_date = pd.Timestamp(reference_date or today or date.today()).normalize()
    cutoff = pd.Timestamp(start_date).normalize() if start_date is not None else None
    event_groups = []
    for _, row in projects.iterrows():
        events = [(row["Start"], "G1", "D")]
        row_milestones = milestones.loc[milestones["Excel row"].eq(row["Excel row"])]
        events.extend((milestone.Date, f"G{milestone.Gate}", "D")
                      for milestone in row_milestones.itertuples(index=False))
        events.append((row["End"], "G6", "^"))
        groups = {}
        for stamp, label, marker in events:
            if pd.notna(stamp) and (cutoff is None or stamp >= cutoff):
                groups.setdefault(stamp, []).append((label, marker))
        for stamp, group in groups.items():
            gates = sorted({label for label, marker in group if marker == "D"},
                           key=lambda label: int(label[1:]))
            groups[stamp] = ([("/".join(gates), "D")] if gates else []) + [
                (label, marker) for label, marker in group if marker != "D"]
        event_groups.append(groups)
    # Preserve the two-line layout: project identity first, financial metrics second.
    labels = projects["Activity"].tolist()
    row_heights = [max(36, 12 * (label.count("\n") + 1) + 8,
                       24 * max((len(group) for group in groups.values()), default=1) + 12)
                   for label, groups in zip(labels, event_groups)]
    phase_header_height = 20
    plot_height = sum(row_heights) + phase_header_height * len(phase_starts)
    width_pt = 420 / 25.4 * 72
    height_pt = max(297 / 25.4 * 72, plot_height + 230)
    fig = Figure(figsize=(width_pt / 72, height_pt / 72), dpi=100, facecolor="white")
    FigureCanvasAgg(fig)
    renderer = fig.canvas.get_renderer()
    font = FontProperties(family="DejaVu Sans", size=9)
    label_width = max(renderer.get_text_width_height_descent(line, font, False)[0]
                      for label in labels for line in label.splitlines()) * 72 / fig.dpi
    header_font = FontProperties(family="DejaVu Sans", size=9, weight="bold")
    header_width = max((renderer.get_text_width_height_descent(phase, header_font, False)[0]
                        for phase in phase_starts.values()), default=0) * 72 / fig.dpi
    label_width = max(label_width, header_width)
    left = 20 + label_width + 16
    plot_top = height_pt - 65
    plot_bottom = plot_top - plot_height
    ax = fig.add_axes([left / width_pt, plot_bottom / height_pt,
                       (width_pt - left - 22) / width_pt, plot_height / height_pt])
    ax.set_facecolor("white")
    ax.set_axisbelow(True)
    ax.grid(axis="x", color="#E5E5E5", linewidth=0.6)
    for spine in ax.spines.values():
        spine.set_visible(False)
    limits = limits if limits is not None else ipd_timeline_limits(projects, milestones, today, cutoff)
    if cutoff is not None:
        lower = mdates.date2num(cutoff)
        limits = (lower, max(limits[1], lower + 30))
    ax.set_xlim(*limits)
    ax.set_ylim(plot_height, 0)
    year_positions, year_labels = [], []
    for year in range(mdates.num2date(limits[0]).year, mdates.num2date(limits[1]).year + 1):
        year_start = date(year, 1, 1)
        if limits[0] <= mdates.date2num(year_start) <= limits[1]:
            ax.axvline(year_start, color=GREY, alpha=0.55, linewidth=1, zorder=1,
                       label=f"1 Jan {year}")
        visible_start = max(limits[0], mdates.date2num(year_start))
        visible_end = min(limits[1], mdates.date2num(date(year + 1, 1, 1)))
        if visible_end > visible_start:
            year_positions.append((visible_start + visible_end) / 2)
            year_labels.append(str(year))
    positions, cursor = [], 0
    phase_boundaries = []
    phase_label_x = (20 - left) / (width_pt - left - 22)
    for index, height in enumerate(row_heights):
        if index in phase_starts:
            phase_boundaries.append(cursor)
            ax.text(phase_label_x, cursor + phase_header_height / 2, phase_starts[index],
                    transform=ax.get_yaxis_transform(), ha="left", va="center",
                    fontproperties=header_font, color=GREY, clip_on=False)
            cursor += phase_header_height
        positions.append(cursor + height / 2 - 4)
        cursor += height
    if phase_boundaries:
        for boundary in [*phase_boundaries, cursor]:
            ax.plot([phase_label_x, 1], [boundary, boundary], transform=ax.get_yaxis_transform(),
                    color="black", linestyle=":", linewidth=0.8, zorder=1, clip_on=False,
                    label="_phase_separator")
    ax.set_yticks(positions, labels, fontsize=9, color=GREY, fontfamily="DejaVu Sans")
    for tick in ax.get_yticklabels():
        tick.set_parse_math(False)
    ax.tick_params(axis="y", length=0, pad=12)
    interval = max(1, math.ceil((limits[1] - limits[0]) / (30.44 * 24)))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=interval))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%y"))
    ax.tick_params(axis="x", colors=GREY, labelsize=8, length=0)
    for tick in ax.get_xticklabels():
        tick.set_rotation(45)
        tick.set_ha("right")

    year_axis = ax.secondary_xaxis("bottom")
    year_axis.spines["bottom"].set_position(("outward", 36))
    for spine in year_axis.spines.values():
        spine.set_visible(False)
    year_axis.set_xticks(year_positions, year_labels)
    year_axis.tick_params(axis="x", colors=GREY, labelsize=9, length=0, pad=3)
    for tick in year_axis.get_xticklabels():
        tick.set_fontweight("bold")

    summary_parts = ipd_yearly_launch_labels(projects, map(int, year_labels), limits,
                                         reference_date, historical_projects)
    summary_font = FontProperties(family="DejaVu Sans", size=9, weight="normal")
    summary_positions = [ax.transData.transform((position, 0))[0] for position in year_positions]
    spacing = min((right - left for left, right in zip(summary_positions, summary_positions[1:])),
                  default=ax.bbox.width)
    summary_labels = []
    for parts in summary_parts:
        label = " ".join(parts)
        if renderer.get_text_width_height_descent(label, summary_font, False)[0] + 12 > spacing:
            label = "\n".join(parts)
        summary_labels.append(label)
    summary_axis = ax.secondary_xaxis("bottom")
    summary_axis.spines["bottom"].set_position(("outward", 52))
    for spine in summary_axis.spines.values():
        spine.set_visible(False)
    summary_axis.set_xticks(year_positions, summary_labels)
    summary_axis.tick_params(axis="x", colors=GREY, length=0, pad=3)
    for position, tick in zip(summary_positions, summary_axis.get_xticklabels()):
        tick.set_fontproperties(summary_font)
        text_width = max(renderer.get_text_width_height_descent(line, summary_font, False)[0]
                         for line in tick.get_text().splitlines())
        if position + text_width / 2 > fig.bbox.x1 - 10:
            tick.set_ha("right")
        elif position - text_width / 2 < fig.bbox.x0 + 10:
            tick.set_ha("left")

    gate_font = FontProperties(family="DejaVu Sans", size=7)
    for (_, row), y, groups in zip(projects.iterrows(), positions, event_groups):
        if row["Valid dates"]:
            ax.plot([row["Start"], row["End"]], [y, y], color=IPD_BAR_COLOUR,
                    linewidth=4.5, solid_capstyle="round", zorder=2)
        else:
            ax.text(0.01, y + 7, "Missing/invalid project dates", transform=ax.get_yaxis_transform(),
                    color=IPD_ERROR_COLOUR, fontsize=7, va="center", zorder=6,
                    bbox={"facecolor": "white", "edgecolor": "none", "pad": 1})
        # Keep the launch triangle separate when it shares a date with the combined gate diamond.
        for stamp, group in groups.items():
            for index, (label, marker) in enumerate(group):
                event_y = y + (index - (len(group) - 1) / 2) * 24
                past_colour = IPD_PAST_GATE_FILL if marker == "D" else IPD_PAST_LAUNCH_FILL
                outline_colour = IPD_EVENT_GREEN if marker == "D" else IPD_LAUNCH_OUTLINE
                ax.plot(stamp, event_y, marker=marker, markersize=IPD_EVENT_SIZE,
                        markerfacecolor=past_colour if stamp <= reference_date else "white",
                        markeredgecolor=outline_colour, markeredgewidth=1.3,
                        linestyle="none", zorder=5, label=label, clip_on=False)
                label_width = renderer.get_text_width_height_descent(label, gate_font, False)[0]
                event_x = ax.transData.transform((mdates.date2num(stamp), event_y))[0]
                alignment = ("left" if event_x - label_width / 2 < ax.bbox.x0 else
                             "right" if event_x + label_width / 2 > ax.bbox.x1 else "center")
                ax.annotate(label, (stamp, event_y), xytext=(0, -9), textcoords="offset points",
                            ha=alignment, va="top", fontproperties=gate_font, color=GREY)
    if today is not None and (cutoff is None or pd.Timestamp(today) >= cutoff):
        ax.axvline(pd.Timestamp(today), color=TODAY_COLOUR, linewidth=0.8, linestyle="--", zorder=1)

    segment_column = next((column for column in projects
                           if ipd_cell_text(column).casefold() == "market segment"), None)
    segments = (sorted({ipd_cell_text(value) or "(Blank)" for value in projects[segment_column]},
                       key=str.casefold) if segment_column is not None else [])
    segment_title = " - " + " - ".join(segments) if segments else ""
    fig.text(20 / width_pt, 1 - 12 / height_pt, f"IPD Projects{segment_title}{page_label}",
             fontsize=14, fontweight="bold", va="top", color="#222222")
    fig.text(20 / width_pt, 1 - 34 / height_pt,
             f"{len(projects)} projects · Gate 1 PS → {IPD_END_LABEL} · "
             f"{len(milestones.loc[milestones['Excel row'].isin(projects['Excel row'])])} milestones",
             fontsize=10, va="top", color=GREY)
    handles = [Line2D([], [], color=IPD_BAR_COLOUR, linewidth=4.5, label="Project duration"),
               Line2D([], [], color=IPD_EVENT_GREEN, marker="D", markersize=IPD_EVENT_SIZE,
                      markerfacecolor=IPD_PAST_GATE_FILL, linestyle="none", label="Past gates"),
               Line2D([], [], color=IPD_EVENT_GREEN, marker="D", markersize=IPD_EVENT_SIZE,
                      markerfacecolor="white", linestyle="none", label="Future gates"),
               Line2D([], [], color=IPD_LAUNCH_OUTLINE, marker="^", markersize=IPD_EVENT_SIZE,
                      markerfacecolor="white", linestyle="none", label=IPD_END_LABEL)]
    if today is not None:
        handles.append(Line2D([], [], color=TODAY_COLOUR, linestyle="--", linewidth=0.8,
                              label=f"Today ({ipd_format_date(pd.Timestamp(today))})"))
    fig.legend(handles=handles, loc="upper left", bbox_to_anchor=(20 / width_pt, (plot_bottom - 105) / height_pt),
               ncol=len(handles), frameon=False, fontsize=9, labelcolor=GREY)
    return fig


# Shared portfolio selection, analytics and exports



MODES = {"type0": "Type 0", "ipd": "IPD"}


SEGMENT_COLOURS = {"Building": "#32AAFF", "Infrastructure": "#D645CC",
                   "Industry": "#FAA64A", "Agriculture": "#EC7F07"}


MAX_UPLOAD_BYTES = 50 * 1024 * 1024


TIME_TO_MARKET_COLUMNS = ["Month end", "Segment", "Period", "Time to market (years)", "KPI",
                          "Launch count", "Actual launches", "Planned launches", "Launches this month"]


def load_data(mode, source, sheet):
    with render_lock(), warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always", UserWarning)
        data = (type0_read_web_workbook(source["content"], sheet, TYPE0_SKIP_ROWS + 1)
                if mode == "type0" else ipd_read_workbook(source["content"], sheet, IPD_DEFAULT_HEADER_ROW))
    notes = list(dict.fromkeys(str(item.message) for item in caught
                              if str(item.message) != TYPE0_DATA_VALIDATION_WARNING))
    if mode == "type0":
        unknown = sorted(set(data["Segment"].map(type0_cell_text)) - set(TYPE0_SEGMENTS) - {""})
        if unknown:
            notes.append("Other segment values outside the Type 0 scope: " + ", ".join(unknown))
    return data, notes


def build_pipeline_results(projects, mode):
    """Use one aggregation for the interactive pipeline and PNG export."""
    if mode == "Type 0":
        columns = {"Due Date": "End", "Segment": "Market segment",
                   "PAS (in k EUR)": IPD_ACTIVITY_METRICS["PAS"],
                   "PIAS (in k EUR)": IPD_ACTIVITY_METRICS["PIAS"]}
        projects = projects[list(columns)].rename(columns=columns)
    values = ipd_pipeline_values(projects)
    png = None
    if not values.empty:
        with render_lock():
            fig = ipd_plot_pipeline(values, title=f"{mode} Pipeline", value_decimals=0,
                                    label_segments=True, legend_below=True,
                                    value_fontweight="normal", segment_colours=SEGMENT_COLOURS)
            png = figure_png(fig, dpi=180)
    return dict(pipeline_values=values, pipeline_png=png)


def build_time_to_market_results(projects, mode, today=None):
    """Equal-weight project means in the 12 calendar months ending at each month-end."""
    today = pd.Timestamp(today if today is not None else date.today()).normalize()
    last_closed_month = today if today.is_month_end else today - pd.offsets.MonthEnd(1)
    start_column, end_column, segment_column, status_column = (
        ("Creation Date", "Due Date", "Segment", "Status") if mode == "type0"
        else ("Start", "End", "Market segment", "Project status"))
    start = pd.to_datetime(projects[start_column], errors="coerce").dt.normalize()
    end = pd.to_datetime(projects[end_column], errors="coerce").dt.normalize()
    missing = start.isna() | end.isna()
    reversed_dates = ~missing & end.lt(start)
    valid = ~(missing | reversed_dates)
    lookup = {str(column).casefold(): column for column in projects}
    segment_column = lookup.get(segment_column.casefold())
    segments = (projects[segment_column].map(type0_cell_text).replace("", "(Blank)")
                if segment_column is not None else pd.Series("All projects", index=projects.index))
    status_column = lookup.get(status_column.casefold())
    statuses = (projects[status_column].map(type0_cell_text).str.casefold()
                if status_column is not None else pd.Series("", index=projects.index))
    completed = statuses.isin(["completed", "launched"])
    cancelled = statuses.isin(["cancelled", "canceled"])
    actual = valid & completed & end.le(today)
    planned = valid & ~completed & ~cancelled & end.ge(today) & end.gt(last_closed_month)
    overdue = valid & ~completed & ~cancelled & ~planned
    future_completed = valid & completed & end.gt(today)
    eligible = actual | planned
    durations = pd.DataFrame({
        "End": end.loc[eligible],
        "Segment": segments.loc[eligible],
        "Actual": actual.loc[eligible],
        "Duration": (end.loc[eligible] - start.loc[eligible]).dt.days / 365.25,
    })
    # Always report at least the last 12 closed months, even with no launches.
    first_month = last_closed_month - pd.offsets.MonthEnd(11)
    final_month = last_closed_month
    if not durations.empty:
        first_month = min(first_month, durations["End"].min() + pd.offsets.MonthEnd(0))
        final_month = max(final_month, durations["End"].max() + pd.offsets.MonthEnd(0))
    records = []
    segment_names = list(dict.fromkeys([*segments, "Total"]))
    for month_end in pd.date_range(first_month, final_month, freq="ME"):
        # Period arithmetic also handles February/leap years: use 12 whole calendar months.
        window_start = (month_end.to_period("M") - 11).start_time
        window = durations.loc[durations["End"].between(window_start, month_end)]
        period = "Actual" if month_end <= last_closed_month else "Forecast"
        if period == "Actual":
            window = window.loc[window["Actual"]]
        for segment in segment_names:
            group = window if segment == "Total" else window.loc[window["Segment"].eq(segment)]
            count = len(group)
            actual_count = int(group["Actual"].sum())
            average = group["Duration"].mean()  # Each project contributes exactly once.
            this_month = int(group["End"].ge(month_end.to_period("M").start_time).sum())
            kpi = f"{average:.1f} years" if count else "N/A — no launches"
            records.append((month_end, segment, period, average, kpi, count, actual_count,
                            count - actual_count, this_month))
    values = pd.DataFrame(records, columns=TIME_TO_MARKET_COLUMNS)
    formula = ("Due Date − Creation Date" if mode == "type0"
               else "G6 (commercial launch date) − G1 (Gate 1 PS)")
    note = (f"Rolling 12 calendar months, recalculated at every month-end; duration = {formula}, in days ÷ 365.25. "
            "Every project has equal weight, regardless of sales or project size; Total averages individual projects "
            "across the selected segments. Values are displayed to one decimal; calculations retain full precision. "
            f"Solid lines: Completed/Launched projects through {last_closed_month:%d %b %Y}. "
            "Dashed lines: future month-ends using completed plus planned launches in the same rolling window. "
            "The current incomplete month is a forecast. n = launches in the 12-month window, including planned "
            "launches for forecasts; actual/planned counts and launches this month are in the hover and monthly table. "
            "Months without new launches are still recalculated; an empty window is N/A — no launches. "
            f"{int(eligible.sum()):,} project(s) included; {int(missing.sum()):,} excluded for missing or invalid dates; "
            f"{int(reversed_dates.sum()):,} excluded because completion precedes the start date. "
            f"{int(overdue.sum()):,} unfinished project(s) need an updated planned date for the forecast; "
            f"{int(future_completed.sum()):,} Completed/Launched project(s) with a future date excluded pending correction; "
            f"{int((valid & cancelled).sum()):,} cancelled project(s) excluded. "
            "Actuals require Completed/Launched status; include these in the Status filter to see actual history. "
            "Missing statuses do not establish a launch. Hover any month for its KPI and counts.")
    return dict(time_to_market_values=values, time_to_market_note=note)


def time_to_market_figure(values, mode):
    """Smooth actual/forecast curves; visual connections never fill missing KPI values."""
    figure = go.Figure()
    if values.empty or values["Time to market (years)"].isna().all():
        figure.add_annotation(text="N/A — no launches · 0 launches", x=0.5, y=0.5,
                              xref="paper", yref="paper", showarrow=False)
    if not values.empty:
        fallback = ["#52B788", "#9B72CF", "#36A6A6", "#E76F51"]
        for index, segment in enumerate(values["Segment"].unique()):
            rows = values.loc[values["Segment"].eq(segment)].sort_values("Month end")
            colour = "#18283e" if segment == "Total" else SEGMENT_COLOURS.get(segment, fallback[index % len(fallback)])
            width = 2.5 if segment == "Total" else 2.0
            actual = rows.loc[rows["Period"].eq("Actual")]
            forecast = rows.loc[rows["Period"].eq("Forecast")]
            known_actual = actual.dropna(subset=["Time to market (years)"])
            if not known_actual.empty and not forecast.empty:
                # Start the dashed curve at the last known actual so the two styles meet,
                # even if empty rolling windows separate actuals from planned launches.
                forecast = pd.concat([known_actual.tail(1), forecast])
            # Keep a long monthly history readable. Every point retains KPI/count hover details;
            # show a few Total labels, always pairing the displayed KPI with its denominator.
            labelled_months = set(rows["Month end"].iloc[::max(3, math.ceil(len(rows) / 10))])
            for period, points, dash in (("Actual", actual, "solid"), ("Forecast", forecast, "dash")):
                if points.empty:
                    continue
                y = [float(value) if pd.notna(value) else None for value in points["Time to market (years)"]]
                counts = points["Launch count"].tolist()
                figure.add_scatter(
                    name=segment, legendgroup=segment,
                    showlegend=(period == "Actual" and not known_actual.empty)
                               or (period == "Forecast" and known_actual.empty),
                    x=points["Month end"].tolist(), y=y, mode="lines+text", connectgaps=True,
                    line=dict(color=colour, width=width, dash=dash, shape="spline", smoothing=0.7),
                    text=[f"{value:.1f} · n={count}" if value is not None and segment == "Total"
                          and point_period == period
                          and (month in labelled_months or month == points["Month end"].iloc[-1]) else ""
                          for month, value, count, point_period in zip(points["Month end"], y, counts, points["Period"])],
                    textposition="top center" if period == "Actual" else "bottom center",
                    textfont=dict(color=colour, size=10),
                    customdata=points[["Period", "Launch count", "Actual launches", "Planned launches",
                                       "Launches this month"]].values.tolist(),
                    hovertemplate=(f"{segment} · %{{x|%d %b %Y}} · %{{customdata[0]}}<br>"
                                   "%{y:.1f} years · %{customdata[1]} launches in last 12 months<br>"
                                   "%{customdata[2]} actual + %{customdata[3]} planned<br>"
                                   "%{customdata[4]} launches this month<extra></extra>"))
        figure.update_xaxes(type="date", tickformat="%b\n%Y", hoverformat="%d %b %Y",
                            range=[values["Month end"].min() - pd.Timedelta(days=15),
                                   values["Month end"].max() + pd.Timedelta(days=15)])
    figure.update_layout(
        template="plotly_white", title=f"{MODES[mode]} time to market · rolling 12 months", height=550,
        margin=dict(l=60, r=35, t=70, b=105), font=dict(family="Arial, sans-serif", color="#35445a"),
        xaxis_title="Month-end · solid: actual · dashed: forecast · n: launches in rolling window",
        yaxis_title="Average time to market (years)",
        legend=dict(orientation="h", y=-0.25, x=0, groupclick="togglegroup"),
        hoverlabel=dict(bgcolor="white"), paper_bgcolor="white", uirevision="time-to-market-monthly")
    figure.update_yaxes(tickformat=".1f", rangemode="tozero")
    return figure


def time_to_market_png(chart):
    """Export the interactive chart's series, labels and actual/forecast styles."""
    with render_lock():
        figure = Figure(figsize=(12, 5.5), facecolor="white")
        FigureCanvasAgg(figure)
        ax = figure.subplots()
        figure.subplots_adjust(left=0.08, right=0.97, bottom=0.27, top=0.88)
        for trace in chart.data:
            # Match connectgaps=True without inventing values for missing months.
            points = [(x, y, text) for x, y, text in zip(trace.x, trace.y, trace.text)
                      if y is not None]
            if not points:
                continue
            x, y, _ = zip(*points)
            ax.plot(x, y, color=trace.line.color, linewidth=trace.line.width,
                    linestyle="--" if trace.line.dash == "dash" else "-",
                    label=trace.name if trace.showlegend else "_nolegend_")
            for stamp, value, label in points:
                if label:
                    above = trace.textposition == "top center"
                    ax.annotate(label, (stamp, value), xytext=(0, 7 if above else -7),
                                textcoords="offset points", ha="center", va="bottom" if above else "top",
                                fontsize=8, color=trace.line.color)
        for annotation in chart.layout.annotations:
            ax.text(annotation.x, annotation.y, annotation.text, transform=ax.transAxes,
                    ha="center", va="center", color="#35445a")
        if chart.layout.xaxis.range:
            ax.set_xlim(*chart.layout.xaxis.range)
        ax.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=4, maxticks=10))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:.1f}"))
        ax.set_ylim(bottom=0, top=max(1, ax.get_ylim()[1]) * 1.12)
        ax.set_title(chart.layout.title.text, loc="left", pad=20, color="#35445a")
        ax.set_xlabel(chart.layout.xaxis.title.text, labelpad=12, fontsize=9, color="#35445a")
        ax.set_ylabel(chart.layout.yaxis.title.text, color="#35445a")
        ax.tick_params(axis="both", length=0, colors="#35445a", labelsize=9)
        ax.set_axisbelow(True)
        ax.grid(color="#E8EBEF", linewidth=0.8)
        for spine in ax.spines.values():
            spine.set_visible(False)
        handles, labels = ax.get_legend_handles_labels()
        if handles:
            figure.legend(handles, labels, loc="lower left", bbox_to_anchor=(0.07, 0.01),
                          ncol=min(5, len(handles)), frameon=False, fontsize=9)
        return figure_png(figure)


def type0_report_context(source, sheet, selection, reporting_date, new_since):
    completed = type0_is_completed_only(selection["statuses"])
    priority = ("Not applied (Completed only)" if completed else
                type0_priority_range_label(selection["min_priority"], selection["max_priority"]))
    context = [
        f"Source: {source['name']} | Worksheet: {sheet} | Header row: {TYPE0_SKIP_ROWS + 1}",
        f"Reporting date: {type0_english_date(reporting_date)}",
        (f"Mark new projects created since: {type0_english_date(new_since)} (creation date, inclusive)"
         if new_since is not None else "Mark new projects created since: Not set (N labels disabled)"),
        f"Priority: {priority}",
    ]
    for label, key in (("Status", "statuses"), ("Types", "types"), ("Segments", "segments"), ("Plants", "plants")):
        context.append(f"{label}: {', '.join(value or '[Blank]' for value in selection[key])}")
    if completed:
        created = selection["projects_created_since"]
        context.append("Projects created since: " + (type0_english_date(created) if created else "All creation dates"))
    context.extend([
        f"Project tags (match any): {', '.join(selection['project_tags'])}",
        "Charts and checks include all selected projects; missing/invalid creation or due dates are labelled in red.",
        "Duplicate project numbers are checked across the full worksheet; actions cover selected projects.",
    ])
    return context


def select_ipd_projects(projects, milestones, issues, settings):
    """Retain the pre-cutoff scope used by historical launch totals."""
    mask = pd.Series(True, index=projects.index)
    for column, values in settings["filters"].items():
        mask &= projects[column].map(ipd_cell_text).isin(values)
    priority_column = settings["priority_column"]
    if priority_column is not None:
        priorities = pd.to_numeric(projects[priority_column].map(ipd_cell_text), errors="coerce")
        if settings["min_priority"] is not None:
            mask &= priorities.ge(settings["min_priority"])
        if settings["max_priority"] is not None:
            mask &= priorities.le(settings["max_priority"])
    historical = projects.loc[mask].copy()
    selected = ipd_projects_with_gates_since(historical, milestones, settings["start_date"])
    sort_column = {"Start date": "Start", IPD_END_LABEL: "End", "Activity label": "Activity"}.get(settings["order"])
    if sort_column:
        selected = selected.sort_values(sort_column, kind="stable", na_position="last")
    if settings["group_by_phase"]:
        selected = ipd_order_projects_by_phase(selected)
    source_rows = {int(row) for rows in selected["Source rows"] for row in rows.split(",")}
    return dict(projects=selected, historical_projects=historical,
                milestones=milestones.loc[milestones["Excel row"].isin(selected["Excel row"])].copy(),
                issues=issues.loc[issues["Excel row"].isin(source_rows)].copy(),
                settings=settings.copy(), charts={})


def ipd_chart(result, page):
    if page not in result["charts"]:
        settings = result["settings"]
        projects = result["projects"]
        rows = settings["rows_per_page"]
        pages = math.ceil(len(projects) / rows)
        current = projects.iloc[(page - 1) * rows:page * rows]
        today = settings["reference_date"] if settings["show_today"] else None
        with render_lock():
            fig = ipd_plot_gantt(
                current, result["milestones"], today=today,
                limits=result["limits"],
                page_label=f" · Page {page} of {pages}" if pages > 1 else "",
                start_date=settings["start_date"], reference_date=settings["reference_date"],
                group_by_phase=settings["group_by_phase"], historical_projects=result["historical_projects"],
            )
            result["charts"][page] = figure_png(fig)
    return result["charts"][page]


FILTER_FIELDS = {
    "type0": {"statuses": ("Status", "Status"), "types": ("Type", "Type"),
              "segments": ("Segment", "Segment"), "plants": ("Plant", "Plant")},
    "ipd": {"statuses": ("Status", "Project status"), "phases": ("Phase", IPD_PHASE),
            "segments": ("Segment", "Market segment")},
}


def defaults(mode, data=None):
    values = dict(minimum=None, maximum=None, show_today=True, rows=30)
    if mode == "type0":
        values.update(minimum=TYPE0_MIN_PRIORITY, maximum=TYPE0_MAX_PRIORITY, project_tags=["K", "C", "N"],
                      created_since=None, reporting_date=date.today().isoformat(),
                      new_since=type0_new_project_cutoff().date().isoformat())
    else:
        values.update(start_date=ipd_default_start_date(date.today()).isoformat(),
                      order="Excel order")
    projects = None if data is None else (data if mode == "type0" else data[0])
    lookup = {} if projects is None else {str(col).casefold(): col for col in projects}
    options = {}
    for key, (_, column) in FILTER_FIELDS[mode].items():
        actual = lookup.get(column.casefold())
        choices = (sorted(projects[actual].map(type0_cell_text).unique(), key=str.casefold)
                   if actual is not None else [])
        if mode == "type0" and key == "segments" and data is not None:
            choices = list(TYPE0_SEGMENTS)
        options[key] = choices
        selected = choices
        if key == "statuses":
            wanted = {"progress", "in progress", "to start"} if mode == "type0" else {"active", "launched"}
            selected = [choice for choice in choices if choice.casefold() in wanted]
        elif mode == "ipd" and key == "segments":
            selected = [choice for choice in choices if choice.casefold() != "agriculture"]
        values[key] = selected
    return values, options


def parse_day(value, label, *, required=False):
    if not value:
        if required:
            raise ValueError(f"Select a {label}.")
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError as exc:
        raise ValueError(f"Choose a valid {label}.") from exc


def select_scope(mode, data, values):
    """Validate controls and apply the portfolio selection rules."""
    minimum, maximum = values.get("minimum"), values.get("maximum")
    completed = mode == "type0" and type0_is_completed_only(values.get("statuses") or [])
    projects = data if mode == "type0" else data[0]
    lookup = {str(column).casefold(): column for column in projects}
    priority_column = lookup.get("priority project") if mode == "ipd" else "Priority proposal"
    if completed or priority_column is None:
        minimum = maximum = None
    for value in (minimum, maximum):
        if value is not None and (not isinstance(value, (int, float)) or not math.isfinite(value)
                                  or int(value) != value):
            raise ValueError("Priority limits must be whole numbers or empty.")
    if minimum is not None and maximum is not None and minimum > maximum:
        raise ValueError("Minimum priority cannot exceed maximum priority.")
    rows_per_page = values.get("rows", 30)
    if rows_per_page not in (10, 20, 30, 40, 50, 60, 70, 80, 90, 100):
        raise ValueError("Choose 10, 20, 30, 40, 50, 60, 70, 80, 90, or 100 projects per chart.")
    if mode == "type0":
        if not all(values.get(key) for key in FILTER_FIELDS[mode]):
            raise ValueError("Select at least one status, type, segment, and plant.")
        if not values.get("project_tags"):
            raise ValueError("Select at least one project tag (K, C, or N).")
        new_since = (parse_day(values.get("new_since"), "new-project date")
                     if "N" in values["project_tags"] else None)
        selection = dict(statuses=values["statuses"], types=values["types"], segments=values["segments"],
                         plants=values["plants"], min_priority=minimum, max_priority=maximum,
                         project_tags=values["project_tags"],
                         projects_created_since=parse_day(values.get("created_since"), "creation date") if completed else None)
        if "N" in values["project_tags"]:
            selection["new_projects_since"] = new_since
        options = dict(selection=selection,
                       reporting_date=parse_day(values.get("reporting_date"), "reporting date", required=True),
                       new_since=new_since,
                       show_today=bool(values.get("show_today")), rows_per_page=rows_per_page)
        return options, type0_select_web_projects(data, **selection)
    filters = {lookup[column.casefold()]: values.get(key) or []
               for key, (_, column) in FILTER_FIELDS[mode].items() if column.casefold() in lookup}
    options = dict(filters=filters, start_date=parse_day(values.get("start_date"), "start date", required=True),
                   reference_date=date.today(), priority_column=priority_column,
                   min_priority=minimum, max_priority=maximum,
                   group_by_phase=IPD_PHASE.casefold() in lookup, order=values.get("order"),
                   show_today=bool(values.get("show_today")), rows_per_page=rows_per_page)
    if options["order"] not in ("Excel order", "Start date", IPD_END_LABEL, "Activity label"):
        raise ValueError("Choose a valid project order.")
    return options, select_ipd_projects(*data, options)


def settings_key(source, sheet, options):
    # Pagination changes the chart layout, not the generated project selection.
    payload = [source["digest"], source["name"], sheet,
               {key: value for key, value in options.items() if key != "rows_per_page"}]
    return sha256(json.dumps(payload, sort_keys=True, default=str).encode()).hexdigest()


def build_ipd_results(result):
    """Render all pages once; the common builder owns tables and ZIP packaging."""
    projects, historical = result["projects"], result["historical_projects"]
    segment_column = next((column for column in projects
                           if ipd_cell_text(column).casefold() == "market segment"), None)
    segments = [None, *result["settings"]["filters"][segment_column]] if segment_column else [None]
    result["segments"] = []
    for index, segment in enumerate(segments):
        selected = (projects.loc[projects[segment_column].map(ipd_cell_text).eq(segment)]
                    if segment is not None else projects)
        past = (historical.loc[historical[segment_column].map(ipd_cell_text).eq(segment)]
                if segment is not None else historical)
        label = "All" if segment is None else segment or "(Blank)"
        filename_label = re.sub(r"[^\w .-]+", "_", label).strip(" .") or "Blank"
        group = dict(label=label, filename_prefix=f"Type_IPD_Gantt_{index:02d}_{filename_label}",
                     projects=selected, historical_projects=past,
                     milestones=result["milestones"].loc[
                         result["milestones"]["Excel row"].isin(selected["Excel row"])],
                     settings=result["settings"], charts={})
        today = result["settings"]["reference_date"] if result["settings"]["show_today"] else None
        group["limits"] = ipd_timeline_limits(selected, group["milestones"], today, result["settings"]["start_date"])
        result["segments"].append(group)
        for page in range(1, math.ceil(len(selected) / result["settings"]["rows_per_page"]) + 1):
            ipd_chart(group, page)
    return result


def planning_assets(result, mode):
    """Collect every segment's pages for navigation and downloads."""
    groups = []
    if mode == "type0":
        counts = result["summary"].set_index("Segment")["Selected"].to_dict()
        counts["All"] = len(result["projects"])
        for segment, pages in result["chart_pages"].items():
            groups.append(dict(label=segment, pages=pages, message=result["messages"].get(segment),
                               count=counts[segment]))
        assets = {**result["charts"], "Priority_Project_Checks.txt": result["report"].encode("utf-8")}
    else:
        assets = {}
        for group in result["segments"]:
            pages = []
            for page, png in group["charts"].items():
                filename = f"{group['filename_prefix']}_{page}.png"
                assets[filename] = png
                pages.append(filename)
            groups.append(dict(label=group["label"], pages=pages, count=len(group["projects"]),
                               message=None if pages else "No projects in this segment match the selected filters and G1–G5 cutoff."))
    return groups, assets


def repaginate_results(result, rows_per_page):
    """Replace Gantt pages and the ZIP while retaining the generated filter scope."""
    if rows_per_page not in (10, 20, 30, 40, 50, 60, 70, 80, 90, 100):
        raise ValueError("Choose 10, 20, 30, 40, 50, 60, 70, 80, 90, or 100 projects per chart.")
    updated = result.copy()
    if result["mode"] == "type0":
        with render_lock():
            updated.update(type0_build_web_results(**result["chart_request"], rows_per_page=rows_per_page))
    else:
        updated["settings"] = {**result["settings"], "rows_per_page": rows_per_page}
        build_ipd_results(updated)
    groups, charts = planning_assets(updated, result["mode"])
    previous_pages = {filename for group in result["groups"] for filename in group["pages"]}
    assets = {filename: content for filename, content in result["assets"].items() if filename not in previous_pages}
    assets.update(charts)
    archive = zip_outputs(assets)
    result.update(updated, groups=groups, assets=assets, zip=archive, rows_per_page=rows_per_page)


def build_results(mode, source, sheet, data, options, selected):
    """Generate each chart/CSV once and build one complete downloadable archive."""
    if mode == "type0":
        selection = options["selection"]
        chart_request = dict(
            projects_by_segment=selected, today=options["reporting_date"],
            context=type0_report_context(source, sheet, selection, options["reporting_date"], options["new_since"]),
            new_projects_since=options["new_since"], min_priority=selection["min_priority"],
            max_priority=selection["max_priority"], show_today=options["show_today"],
            all_priorities_by_segment=type0_select_web_projects(data, **selection, apply_priority=False),
            number_counts=type0_count_project_numbers(data))
        with render_lock():
            result = type0_build_web_results(**chart_request, rows_per_page=options["rows_per_page"])
        result["chart_request"] = chart_request
        projects = pd.concat(selected.values(), ignore_index=True)
        result.update(build_pipeline_results(projects, "Type 0"))
        projects = projects[[col for col in projects if not col.startswith((".raw_date_", ".raw_numeric_"))]].copy()
        for column in projects.select_dtypes(include=["object", "string"]):
            projects[column] = projects[column].map(type0_cell_text)
        result["projects"] = projects
        tables = {"Projects.csv": projects, "Segment_summary.csv": result["summary"],
                  "Checks_and_actions.csv": result["issues"]}
        result["missing_dates"] = int(projects["Due Date"].isna().sum())
    else:
        # The original renderer also supplies phase grouping and historical launch totals.
        result = build_ipd_results(selected)
        projects = result["projects"]
        result.update(build_pipeline_results(projects, "IPD"))
        tables = {"Projects.csv": projects.drop(columns=IPD_FINANCIAL_COLUMNS).rename(
                    columns={"Start": IPD_START, "End": IPD_END_LABEL}),
                  "Milestones.csv": result["milestones"], "Checks_and_actions.csv": result["issues"],
                  "Financial_metrics.csv": projects[[IPD_PROJECT, IPD_DESCRIPTION, *IPD_FINANCIAL_COLUMNS]]}
        result["missing_dates"] = int(projects["End"].isna().sum())
    groups, assets = planning_assets(result, mode)
    pipeline_name = "Type_0_Pipeline.png" if mode == "type0" else "Type_IPD_Pipeline.png"
    if result["pipeline_png"] is not None:
        assets[pipeline_name] = result["pipeline_png"]
    tables["Pipeline.csv"] = result["pipeline_values"]
    result.update(build_time_to_market_results(projects, mode))
    tables["Time_to_market.csv"] = result["time_to_market_values"]
    for filename, table in tables.items():
        assets[filename] = table.to_csv(index=False, date_format="%Y-%m-%d").encode("utf-8-sig")
    result.update(groups=groups, tables=tables, assets=assets, pipeline_name=pipeline_name,
                  generation=secrets.token_hex(16), fingerprint=settings_key(source, sheet, options),
                  source_name=source["name"], sheet=sheet, mode=mode, rows_per_page=options["rows_per_page"])
    result["zip"] = zip_outputs(assets)
    return result


def pipeline_figure(values, mode):
    """Interactive counterpart of the shared, exportable pipeline PNG."""
    figure = go.Figure()
    if values.empty:
        figure.add_annotation(text="No numeric PAS/PIAS values with valid dates in this selection.",
                              x=0.5, y=0.5, xref="paper", yref="paper", showarrow=False)
    else:
        pairs = list(dict.fromkeys(zip(values["Year"], values["Metric"])))
        x = [[str(year) for year, _ in pairs], [metric for _, metric in pairs]]
        fallback = ["#52B788", "#9B72CF", "#36A6A6", "#E76F51"]
        for index, segment in enumerate(values["Segment"].unique()):
            rows = values.loc[values["Segment"].eq(segment)].set_index(["Year", "Metric"])
            y = [None if pd.isna(rows.loc[pair, "Value (k EUR)"]) else float(rows.loc[pair, "Value (k EUR)"])
                 for pair in pairs]
            figure.add_bar(name=segment, x=x, y=y, marker_color=SEGMENT_COLOURS.get(segment, fallback[index % len(fallback)]),
                           text=[ipd_pipeline_segment_label(
                               value, pair[1], rows.loc[pair, "Project count"] if "Project count" in rows else None)
                               for value, pair in zip(y, pairs)], textposition="inside",
                           insidetextfont=dict(color="white"),
                           customdata=[[str(year), metric, segment] for year, metric in pairs],
                           hovertemplate="%{customdata[2]} · %{customdata[0]} %{customdata[1]}<br>%{y:,.0f} k EUR<extra></extra>")
        totals, tops = [], []
        for year, metric in pairs:
            rows = values.loc[values["Year"].eq(year) & values["Metric"].eq(metric)]
            numbers = rows["Value (k EUR)"]
            total = numbers.sum(min_count=1)
            label = "—" if pd.isna(total) else f"{total:,.0f}"
            if metric == "PAS" and "Project count" in rows:
                label += f" (# {int(rows['Project count'].sum())})"
            totals.append(label)
            tops.append(float(numbers.clip(lower=0).sum()))
        figure.add_scatter(x=x, y=tops, text=totals, mode="text", textposition="top center",
                           textfont=dict(color="black"), showlegend=False, hoverinfo="skip", cliponaxis=False)
    figure.update_layout(template="plotly_white", title=f"{MODES[mode]} pipeline", barmode="relative",
                         height=530, margin=dict(l=60, r=25, t=70, b=100), font=dict(family="Arial, sans-serif", color="#35445a"),
                         yaxis_title="Value (k EUR)", legend=dict(orientation="h", y=-0.22, x=0),
                         hoverlabel=dict(bgcolor="white"), paper_bgcolor="white", uirevision="pipeline")
    figure.update_xaxes(type="multicategory", title=None)
    figure.update_yaxes(tickformat=",.0f", rangemode="tozero")
    return figure


# Streamlit workspace and presentation

PRIORITY_HELP = "Clear both limits to include unassigned priorities."
LOGO_PATH = Path(__file__).resolve().parent / "assets" / "pipeline.svg"
LOGO_BASE64 = b64encode(LOGO_PATH.read_bytes()).decode("ascii")
APP_CSS = """
<style>
.stApp { background:#f3f6fb; }
[data-testid=stHeader] { display:none; }
.stMainBlockContainer { max-width:1900px; padding:2rem 30px 2.5rem; }
.st-key-brand { min-height:45px; }
.st-key-brand h1 { color:#32AAFF; font-size:25px; letter-spacing:-.7px;
  display:flex; align-items:center; line-height:1.2; margin:0; padding:0; }
.st-key-mode [role=radiogroup] { justify-content:center; }
.st-key-mode, .st-key-controls, .st-key-results { background:white; border-radius:14px; }
.st-key-mode h2, .st-key-controls h2 { font-size:16px; }
.st-key-workbook_header [data-testid=stMarkdownContainer] { margin:0; }
.st-key-workbook_header h2 { margin:0; padding:0; line-height:1.25; }
.st-key-controls h3 { font-size:14px; border-top:1px solid #edf0f5; padding-top:20px; }
.st-key-controls [data-testid=stWidgetLabel] p { font-size:12px; font-weight:600; color:#495b74; }
.st-key-controls [data-testid=stFileUploaderDropzone] { background:#f8faff;
  border:1.5px dashed #a8bad6; border-radius:9px; }
.st-key-controls [data-testid=stFileUploaderDropzone] small { font-size:10px; }
[data-testid=stMetric] { background:white; border:1px solid #dde4ee;
  border-radius:11px; padding:14px 18px; }
[data-testid=stMetricValue] { font-size:26px; color:#18283e; }
[data-testid=stMetricLabel], [data-testid=stCaptionContainer] { color:#65748b; }
[data-testid=stTabs] [role=tablist] { gap:16px; border-bottom:1px solid #dde4ee; }
[data-testid=stTabs] [role=tab][aria-selected=true] { color:#275ecf; }
.empty-portfolio { min-height:480px; display:flex; flex-direction:column; align-items:center;
  justify-content:center; text-align:center; background:white; border:1px solid #dde4ee;
  border-radius:14px; padding:45px; color:#18283e; }
.empty-portfolio .icon { width:72px; height:72px; margin-bottom:20px; }
.empty-portfolio h2 { font-size:18px; }
.empty-portfolio p { max-width:470px; color:#65748b; }
.empty-portfolio .steps { display:flex; gap:15px; margin:20px 0; flex-wrap:wrap; justify-content:center; }
.empty-portfolio .steps span { background:#f5f7fc; border-radius:7px; padding:9px 15px; font-size:12px; }
.portfolio-footer { color:#8290a4; text-align:center; margin-top:23px; font-size:11px; }
@media (max-width:760px) {
  .stMainBlockContainer { padding:1rem; }
  .empty-portfolio { min-height:320px; padding:25px; }
}
</style>
"""
APP_CSS += f"""
<style>
.st-key-brand h1::before, .empty-portfolio .icon {{
  background:transparent url("data:image/svg+xml;base64,{LOGO_BASE64}") center / contain no-repeat;
}}
.st-key-brand h1::before {{ content:''; display:block; flex-shrink:0;
  width:45px; height:45px; margin-right:14px; }}
.st-key-workbook_header [data-testid=stBaseButton-primary] {{
  background:{GENERATE_BUTTON_COLOR}; border-color:{GENERATE_BUTTON_COLOR};
  color:{GENERATE_BUTTON_TEXT_COLOR};
}}
.st-key-workbook_header [data-testid=stBaseButton-primary]:hover:not(:disabled) {{ filter:brightness(.92); }}
.st-key-workbook_header [data-testid=stBaseButton-primary]:disabled {{ opacity:.45; }}
</style>
"""


def read_upload(content, filename):
    """Validate raw upload bytes before replacing the current workbook."""
    if Path(filename or "").suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Choose an .xlsx or .xlsm workbook.")
    if len(content) > MAX_UPLOAD_BYTES:
        raise ValueError("The workbook exceeds the 50 MB upload limit.")
    with pd.ExcelFile(BytesIO(content), engine="openpyxl") as book:
        sheets = book.sheet_names
    if not sheets:
        raise ValueError("The workbook has no worksheets.")
    return dict(name=Path(filename).name, content=content, sheets=sheets,
                digest=sha256(content).hexdigest())


def portfolio_state(mode):
    """Durable non-widget state survives switching portfolios and result tabs."""
    views = st.session_state.setdefault("portfolio_views", {})
    return views.setdefault(mode, dict(preferences={}, revision=0, upload_revision=0))


def clear_workbook(mode):
    view = portfolio_state(mode)
    revision, upload_revision = view["revision"] + 1, view["upload_revision"] + 1
    view.clear()
    view.update(preferences={}, revision=revision, upload_revision=upload_revision)


def upload_changed(mode, key):
    upload = st.session_state.get(key)
    if upload is None:
        clear_workbook(mode)
        return
    view = portfolio_state(mode)
    try:
        source = read_upload(upload.getvalue(), upload.name)
    except Exception as exc:
        view["upload_error"] = f"Unable to read workbook: {exc}"
        return
    previous = view.get("source")
    if previous is None or (source["digest"], source["name"]) != (previous["digest"], previous["name"]):
        revision, upload_revision = view["revision"] + 1, view["upload_revision"]
        view.clear()
        view.update(preferences={}, revision=revision, upload_revision=upload_revision, source=source)
    view.pop("upload_error", None)


class Controls:
    """Separate widget keys from saved values so hidden controls retain settings."""

    def __init__(self, mode, scope):
        view = portfolio_state(mode)
        self.prefix = f"_{mode}_{view['revision']}_{scope}_"
        self.values = view["preferences"].setdefault(scope, {})

    def widget(self, kind, name, label, default, **kwargs):
        key = self.prefix + name
        if key not in st.session_state:
            st.session_state[key] = self.values.get(name, default)
        if kind == "selectbox" and st.session_state[key] not in kwargs["options"]:
            st.session_state[key] = default
        if kind == "multiselect":
            kwargs.setdefault("wrap", False)
            st.session_state[key] = [value for value in st.session_state[key] if value in kwargs["options"]]
        if kind in {"date_input", "number_input"}:
            kwargs["value"] = None
        value = getattr(st, kind)(label, key=key, **kwargs)
        self.values[name] = value
        return value


def load_workbook(mode):
    view = portfolio_state(mode)
    key = f"_upload_{mode}_{view['upload_revision']}"
    st.file_uploader("Choose a workbook", type=["xlsx", "xlsm"], key=key,
                     max_upload_size=50, on_change=upload_changed, args=(mode, key))
    if view.get("upload_error"):
        st.error(view["upload_error"])
    source = view.get("source")
    if source:
        with st.container(horizontal=True, wrap=False, vertical_alignment="center", gap="small"):
            st.caption(f"Loaded: {source['name']}", width="content")
            st.button("×", key=f"clear_{mode}", type="tertiary", help="Clear workbook",
                      on_click=clear_workbook, args=(mode,))
    if source is None:
        st.selectbox("Worksheet", [], disabled=True, placeholder="Upload a workbook first")
        return None
    sheets = source["sheets"]
    default = TYPE0_SHEET_NAME if mode == "type0" and TYPE0_SHEET_NAME in sheets else sheets[0]
    sheet = Controls(mode, "workbook").widget("selectbox", "sheet", "Worksheet", default, options=sheets)
    if view.get("sheet") != sheet:
        view["sheet"] = sheet
        view["preferences"] = {"workbook": view["preferences"]["workbook"]}
        view.pop("result", None)
        view.pop("data", None)
        view.pop("read_error", None)
        # Only selection widgets reset; the worksheet widget keeps its identity.
        view["selection_revision"] = view.get("selection_revision", 0) + 1
        try:
            with st.spinner("Reading projects…"):
                view["data"], view["notes"] = load_data(mode, source, sheet)
        except Exception as exc:
            row = 6 if mode == "type0" else 2
            other = "IPD" if mode == "type0" else "Type 0"
            view["read_error"] = (f"Unable to read worksheet: {exc} {MODES[mode]} expects headings on row {row}. "
                                  f"For a {other} workbook, switch to {other} and upload it there.")
    if view.get("read_error"):
        st.error(view["read_error"])
        return None
    if view.get("notes"):
        with st.expander(f"Workbook conversion warnings ({len(view['notes'])})"):
            st.caption("Warnings cover the full worksheet; checks cover the selected scope.")
            for note in view["notes"]:
                st.text(note)
    return source, sheet, view["data"]


def scope_controls(mode, data):
    initial, choices = defaults(mode, data)
    view = portfolio_state(mode)
    controls = Controls(mode, f"selection_{view.get('selection_revision', 0)}")
    values = {}
    planning_active = False
    active_tab = None
    result = view.get("result")
    if result is not None:
        navigation = Controls(mode, f"results_{result['generation']}")
        active_tab = st.session_state.get(navigation.prefix + "active_tab",
                                          navigation.values.get("active_tab", "Project planning"))
        planning_active = active_tab == "Project planning"
    projects = data if mode == "type0" else (data[0] if data is not None else None)
    columns = {str(column).casefold() for column in projects} if projects is not None else set()

    def multiselect(name):
        label, column = FILTER_FIELDS[mode][name]
        values[name] = controls.widget(
            "multiselect", name, label, initial[name], options=choices[name],
            format_func=lambda value: value or "(Blank)", disabled=column.casefold() not in columns,
            placeholder=f"Choose {label.lower()}…")

    def day(name, label, help_text=None):
        default = parse_day(initial[name], label)
        values[name] = controls.widget("date_input", name, label, default,
                                       min_value=date(1900, 1, 1), max_value=date(2200, 12, 31),
                                       format="DD/MM/YYYY", help=help_text, disabled=data is None)

    st.subheader("Portfolio filters", divider=False)
    multiselect("statuses")
    completed = mode == "type0" and type0_is_completed_only(values["statuses"])
    missing_priority = mode == "ipd" and "priority project" not in columns
    priority_help = "Add the 'Priority Project' column to enable priority limits." if missing_priority else PRIORITY_HELP
    for column, name, label in zip(st.columns(2), ("minimum", "maximum"), ("Min priority", "Max priority")):
        with column:
            values[name] = controls.widget("number_input", name, label, initial[name], step=1,
                                           placeholder="Any", help=priority_help,
                                           disabled=data is None or completed or missing_priority)
    if completed:
        st.caption("Priority limits are ignored for Completed-only selections.")
        day("created_since", "Projects created since")
    if mode == "type0":
        for name in ("types", "segments", "plants"):
            multiselect(name)
        values["project_tags"] = controls.widget(
            "multiselect", "project_tags", "Project tags", initial["project_tags"], options=["K", "C", "N"],
            format_func=lambda tag: {"K": "K — Key Project", "C": "C — Capex", "N": "N — New Project"}[tag],
            help="Include projects matching any selected tag. N uses the 'Mark new projects created since' date.",
            disabled=data is None)
        values["new_since"] = controls.values.get("new_since", initial["new_since"])
        if data is not None and "N" in values["project_tags"] and (result is None or planning_active):
            day("new_since", "Mark new projects created since",
                "N includes and labels projects created on or after this date. Clear to disable N.")
        values["reporting_date"] = controls.values.get("reporting_date", initial["reporting_date"])
        if active_tab not in {"Pipeline", "Time to market"}:
            st.subheader("Chart settings")
            day("reporting_date", "Reporting date")
    else:
        multiselect("phases")
        multiselect("segments")
        day("start_date", "Start date", "Include projects with a G1–G5 milestone on or after this date.")
        # Read toolbar changes before selecting projects and checking stale results.
        values["order"] = st.session_state.get(controls.prefix + "order",
                                               controls.values.get("order", initial["order"]))
        controls.values["order"] = values["order"]
        if planning_active:
            st.subheader("Chart settings")
    values["rows"] = Controls(mode, "planning").values.get("rows", initial["rows"])
    values["show_today"] = controls.values.get("show_today", True)
    if planning_active:
        values["show_today"] = controls.widget("checkbox", "show_today", "Show today", True, disabled=data is None)
    return values


def download(label, result, filename, stale, key, **kwargs):
    mime = {".zip": "application/zip", ".png": "image/png", ".csv": "text/csv", ".txt": "text/plain"}
    content = result["zip"] if filename.endswith(".zip") else result["assets"].get(filename)
    # Stale data is not registered with the download endpoint, even for a disabled button.
    st.download_button(label, data=content if content is not None and not stale else b"",
                       file_name=filename, mime=mime[Path(filename).suffix], key=key,
                       disabled=stale or content is None, on_click="ignore", width=180,
                       help="Generate charts to download results for the current filters." if stale else None,
                       **kwargs)


def data_table(frame, controls, name):
    """View-only search and multi-column sorting never alter generated exports."""
    if frame.empty:
        st.caption("No rows in this selection.")
        return
    with st.expander("Search and sort table"):
        search = controls.widget("text_input", f"{name}_search", "Search all columns", "")
        sort = controls.widget("multiselect", f"{name}_sort", "Sort columns", [], options=list(frame))
        descending = controls.widget("checkbox", f"{name}_descending", "Descending", False)
        column_filters = controls.widget("multiselect", f"{name}_filters", "Filter columns", [], options=list(frame))
        filters = {column: controls.widget("text_input", f"{name}_filter_{column}", f"Contains · {column}", "")
                   for column in column_filters}
    filtered = frame
    if search:
        matches = frame.astype("string").apply(lambda column: column.str.contains(search, case=False, regex=False, na=False))
        filtered = frame.loc[matches.any(axis=1)]
    for column, text in filters.items():
        if text:
            filtered = filtered.loc[filtered[column].astype("string").str.contains(text, case=False, regex=False, na=False)]
    if sort:
        try:
            filtered = filtered.sort_values(sort, ascending=not descending, kind="stable", na_position="last")
        except TypeError:
            filtered = filtered.sort_values(sort, ascending=not descending, kind="stable",
                                             key=lambda column: column.astype("string"), na_position="last")
    pages = max(1, math.ceil(len(filtered) / 15))
    page = controls.widget("selectbox", f"{name}_page", "Table page", 1,
                            options=list(range(1, pages + 1)), format_func=lambda value: f"{value} of {pages}",
                            width=150)
    st.caption(f"{len(filtered):,} of {len(frame):,} rows · exports include the full generated selection.")
    visible = filtered.iloc[(page - 1) * 15:page * 15].copy()
    # Mixed Excel types otherwise trigger Arrow conversion errors in dataframe display.
    for column in visible.select_dtypes(include=["object", "string"]):
        visible[column] = visible[column].map(type0_cell_text)
    dates = {column: st.column_config.DateColumn(str(column).replace("_", " "), format="YYYY-MM-DD")
             for column in visible.select_dtypes(include=["datetime", "datetimetz"])}
    st.dataframe(visible, hide_index=True, column_config=dates, width="stretch", height="auto")


def show_metrics(result):
    projects, mode = result["projects"], result["mode"]
    amounts = ("PAS (in k EUR)", "PIAS (in k EUR)") if mode == "type0" else (IPD_ACTIVITY_METRICS["PAS"], IPD_ACTIVITY_METRICS["PIAS"])
    lookup = {str(column).casefold(): column for column in projects}
    totals = []
    for column in amounts:
        actual = lookup.get(column.casefold())
        numbers = []
        for value in projects[actual] if actual is not None else []:
            try:
                numbers.append(ipd_financial_value(value))
            except ValueError:
                numbers.append(pd.NA)
        total = pd.Series(numbers, dtype="Float64").sum(min_count=1)
        totals.append("—" if pd.isna(total) else f"{total:,.0f}")
    metrics = (("Selected projects", f"{len(projects):,}"), ("PAS · k EUR", totals[0]),
               ("PIAS · k EUR", totals[1]), ("Checks & actions", f"{len(result['issues']):,}"))
    for column, (label, value) in zip(st.columns(4), metrics):
        column.metric(label, value)


def show_results(result, stale):
    mode = result["mode"]
    controls = Controls(mode, f"results_{result['generation']}")
    planning = Controls(mode, "planning")
    rows_per_page = st.session_state.get(planning.prefix + "rows", planning.values.get("rows", 30))
    if rows_per_page != result.get("rows_per_page", 30):
        with st.spinner("Updating chart pages and downloads…"):
            repaginate_results(result, rows_per_page)
    show_metrics(result)
    with st.container(border=True, key="results"):
        labels = ["Project planning", "Pipeline", "Time to market", "Projects", "Checks & actions"]
        if mode == "ipd":
            labels.append("Financial metrics")
        tab_key = controls.prefix + "active_tab"
        if tab_key not in st.session_state:
            st.session_state[tab_key] = controls.values.get("active_tab", labels[0])
        tabs = st.tabs(labels, key=tab_key, on_change="rerun")
        controls.values["active_tab"] = st.session_state[tab_key]
        for label, tab in zip(labels, tabs):
            if not tab.open:
                continue  # Avoid serializing large hidden tables/charts on every interaction.
            with tab:
                if label == "Project planning":
                    with st.container(horizontal=True, vertical_alignment="bottom", key="planning_toolbar"):
                        segment = controls.widget("selectbox", "segment", "Segment", 0,
                                                   options=list(range(len(result["groups"]))),
                                                   format_func=lambda index: result["groups"][index]["label"], width=200)
                        planning.widget("selectbox", "rows", "Projects per chart", 30,
                                        options=list(range(10, 101, 10)), width=180)
                        group = result["groups"][segment]
                        pages = list(range(1, len(group["pages"]) + 1))
                        page = controls.widget("selectbox", f"chart_page_{segment}", "Chart page", 1 if pages else None,
                                                options=pages, format_func=lambda value: f"{value} of {len(pages)}",
                                                disabled=not pages, width=180)
                        if mode == "ipd":
                            view = portfolio_state(mode)
                            selection = Controls(mode, f"selection_{view.get('selection_revision', 0)}")
                            selection.widget("selectbox", "order", "Project order", "Excel order",
                                             options=["Excel order", "Start date", IPD_END_LABEL, "Activity label"],
                                             width=180)
                        st.space("stretch")
                        filename = group["pages"][page - 1] if page else "Unavailable.png"
                        download("Download chart", result, filename, stale, f"chart_{mode}")
                    if group["message"]:
                        st.warning(group["message"])
                    if page:
                        st.caption(f"{group['label']} · {group['count']:,} projects · Page {page} of {len(pages)}")
                        st.image(result["assets"][filename], width="stretch")
                    if mode == "type0":
                        with st.expander("Segment summary"):
                            data_table(result["summary"], controls, "summary")
                elif label == "Pipeline":
                    with st.container(horizontal=True, horizontal_alignment="right"):
                        download("Download chart", result, result["pipeline_name"], stale, f"pipeline_png_{mode}")
                    st.plotly_chart(result["pipeline_figure"], theme=None, width="stretch", key=f"pipeline_{mode}",
                                     config={"displaylogo": False, "modeBarButtonsToRemove": ["toImage"], "responsive": True})
                elif label == "Time to market":
                    values = result["time_to_market_values"]
                    actual = values.loc[values["Period"].eq("Actual")]
                    if not actual.empty:
                        latest = actual.loc[actual["Month end"].eq(actual["Month end"].max())]
                        st.caption(f"As of {actual['Month end'].max():%d %b %Y} · last 12 months")
                        for start in range(0, len(latest), 4):
                            for column, (_, row) in zip(st.columns(min(4, len(latest) - start)), latest.iloc[start:start + 4].iterrows()):
                                with column:
                                    st.metric(row["Segment"], row["KPI"])
                                    st.caption(f"{row['Launch count']:,} " + ("launch" if row["Launch count"] == 1 else "launches"))
                    filename = "Type_0_Time_to_market.png" if mode == "type0" else "Type_IPD_Time_to_market.png"
                    if filename not in result["assets"]:
                        result["assets"][filename] = time_to_market_png(result["time_to_market_figure"])
                        result["zip"] = zip_outputs(result["assets"])
                    with st.container(horizontal=True, horizontal_alignment="right"):
                        download("Download chart", result, filename, stale, f"ttm_png_{mode}")
                    st.plotly_chart(result["time_to_market_figure"], theme=None, width="stretch", key=f"ttm_{mode}",
                                     config={"displaylogo": False, "modeBarButtonsToRemove": ["toImage"], "responsive": True})
                    with st.expander("Monthly KPI and launch counts"):
                        data_table(values.drop(columns="Time to market (years)").sort_values(
                            ["Month end", "Segment"], ascending=[False, True]), controls, "monthly")
                elif label == "Projects":
                    data_table(result["tables"]["Projects.csv"], controls, "projects")
                    if mode == "ipd":
                        st.subheader("Milestones")
                        data_table(result["milestones"], controls, "milestones")
                elif label == "Checks & actions":
                    st.caption(f"{len(result['issues']):,} checks & actions")
                    if result["issues"].empty:
                        if len(result["projects"]):
                            st.success("No control issues found.")
                        else:
                            st.info("No projects matched; no projects were checked.")
                    else:
                        data_table(result["issues"], controls, "checks")
                    if mode == "type0":
                        with st.expander("Full combined report"):
                            st.text(result["report"])
                    else:
                        st.caption("Blank optional milestones are ignored. Projects with invalid dates remain included when a valid "
                                   "G1–G5 meets the cutoff. Repeated projects are consolidated across financial sections; conflicting "
                                   "values retain the first nonblank value and appear here with their source Excel row.")
                else:
                    st.caption("One row per project. Year 1–Year 11 follow financial headings chronologically. "
                               "Workbook units are retained; blank values remain empty.")
                    data_table(result["tables"]["Financial_metrics.csv"], controls, "financial")
    file_slot, button_slot = st.columns([2, 1], vertical_alignment="bottom")
    with file_slot:
        files = [filename for filename in result["assets"] if not filename.endswith(".png")]
        filename = controls.widget("selectbox", "export", "Export a table or report", files[0], options=files,
                                    format_func=lambda value: value.replace("_", " "))
    with button_slot:
        download("Download file", result, filename, stale, f"file_{mode}")


def main():
    st.set_page_config(page_title="Portfolio Analysis", page_icon="▥", layout="wide")
    st.html(APP_CSS)
    page = st.container(key="portfolio_page", gap=2)
    with page.container(key="brand"):
        st.title("Portfolio Analysis")
    with page.container(key="workspace"):
        sidebar, content = st.columns([1, 3.6], gap="small")
        with sidebar, st.container(border=True, key="mode"):
            st.markdown("## Project portfolio")
            mode = st.radio("Portfolio", list(MODES), format_func=MODES.get, key="portfolio_mode",
                             horizontal=True, label_visibility="collapsed")
        with sidebar, st.container(border=True, key="controls"):
            view = portfolio_state(mode)
            with st.container(horizontal=True, wrap=False, vertical_alignment="center", key="workbook_header"):
                st.markdown(f"## {MODES[mode]} workbook", width="stretch")
                with st.container(width=150):
                    generate_slot = st.empty()
            loaded = load_workbook(mode)
            count_slot = st.empty()
            values = scope_controls(mode, loaded[2] if loaded else None)
            options = selected = None
            scope_error = None
            if loaded:
                source, sheet, data = loaded
                try:
                    options, selected = select_scope(mode, data, values)
                    count = sum(map(len, selected.values())) if mode == "type0" else len(selected["projects"])
                    total = len(data if mode == "type0" else data[0])
                    count_slot.caption(f"{count:,} of {total:,} projects selected")
                    if mode == "ipd" and count == 0:
                        scope_error = "No projects match the selected filters and G1–G5 cutoff. Adjust your scope."
                except ValueError as exc:
                    scope_error = str(exc)
                    count_slot.caption("Adjust the scope to continue")
            else:
                count_slot.caption("No workbook loaded")
            generate = generate_slot.button("Generate charts", type="primary", width=150, key=f"generate_{mode}",
                                             disabled=loaded is None or scope_error is not None)
        with content:
            if generate:
                try:
                    with st.spinner("Generating charts and downloads…"):
                        result = build_results(mode, source, sheet, data, options, selected)
                        result["pipeline_figure"] = pipeline_figure(result["pipeline_values"], mode)
                        result["time_to_market_figure"] = time_to_market_figure(result["time_to_market_values"], mode)
                    view["result"] = result
                    # Discard old navigation/search preferences after regeneration.
                    view["preferences"] = {key: value for key, value in view["preferences"].items()
                                           if not key.startswith("results_")}
                    # Refresh sidebar controls for the newly selected planning tab.
                    st.rerun()
                except Exception as exc:
                    st.error(f"Unable to generate charts: {exc}")
            result = view.get("result")
            stale = result is not None and (scope_error is not None or options is None or
                                            result["fingerprint"] != settings_key(source, sheet, options))
            if scope_error:
                st.warning(scope_error)
            if stale:
                st.warning("Settings changed. Generate charts to apply them. Previous results are shown; downloads are disabled.")
            if result is not None:
                show_results(result, stale)
            else:
                if loaded and not scope_error:
                    st.info("Your workbook is ready. Choose your scope and select Generate charts.")
                st.html(f'''<div class="empty-portfolio"><div class="icon" aria-hidden="true"></div>
                    <h2>Explore your {MODES[mode]} portfolio</h2>
                    <p>Bring your project workbook into one view. Review timelines, compare the pipeline,
                    and turn data checks into actions.</p><div class="steps"><span>1  Upload workbook</span>
                    <span>2  Choose your scope</span><span>3  Generate charts</span></div>
                    <p>Headings should be on row {6 if mode == 'type0' else 2}.</p></div>''')
    st.html('<div class="portfolio-footer">Portfolio Analysis · Type 0 &amp; IPD · Refreshing the page starts a new workspace</div>')


if __name__ == "__main__":
    main()
