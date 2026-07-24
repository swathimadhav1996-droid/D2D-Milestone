"""
OD2D Milestone Completeness — Streamlit app
-------------------------------------------
Upload the raw shipment-level export (the CSV produced by the OD2D shipment-level
SQL query) and download a formatted "Completeness Summary" workbook: overall
completeness + a breakdown by FFW / NVOCC / Carrier, one column per milestone,
with green/yellow/red traffic-light styling.

Sidebar filters:
  - Shipment grain (container-level / booking-level / every row)
  - Shipments to include (all rows / completed only), optional created-date range
  - Milestones to include — pick exactly which of the 24 milestones appear as
    columns in the report (Select all / Clear all shortcuts provided)

Run locally:   streamlit run streamlit_app.py
Deploy:        push this file + requirements.txt to GitHub, then deploy on
               https://share.streamlit.io  (main file path = streamlit_app.py)
"""

import io
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Milestone definitions — one column per individual milestone, in journey order.
# (name shown in the sheet, list of timestamp columns, journey category;
#  "present" = any of the listed timestamp columns is filled)
# ---------------------------------------------------------------------------
ALL_GROUPS = [
    ("Gate Out Empty At Terminal",            ["TS_GATE_OUT_EMPTY_AT_TERMINAL"],            "Empty Pickup"),
    # --- Pre Carriage ---
    ("Picked Up At Origin",                   ["TS_PICKED_UP_AT_ORIGIN"],                   "Pre Carriage"),
    ("Arrival At Origin CFS/WH",              ["TS_ARRIVAL_ORIGIN_CFS_OR_WH"],              "Pre Carriage"),
    ("Load At Origin CFS/WH",                 ["TS_LOAD_ORIGIN_CFS_OR_WH"],                 "Pre Carriage"),
    ("Arrival At Inland Export Terminal",     ["TS_ARRIVAL_INLAND_EXPORT_TERMINAL"],        "Pre Carriage"),
    ("Discharge At Inland Export Terminal",   ["TS_DISCHARGE_INLAND_EXPORT_TERMINAL"],      "Pre Carriage"),
    ("Load At Inland Export Terminal",        ["TS_LOAD_INLAND_EXPORT_TERMINAL"],           "Pre Carriage"),
    ("Departure From Inland Export Terminal", ["TS_DEPARTURE_INLAND_EXPORT_TERMINAL"],      "Pre Carriage"),
    # --- Port to port ---
    ("Gate In Full At POL",                   ["TS_GATE_IN_FULL_POL"],                      "Port to Port"),
    ("Load Onto Vessel At POL",               ["TS_LOAD_ONTO_VESSEL_POL"],                  "Port to Port"),
    ("Vessel Departure From POL",             ["TS_VESSEL_DEPARTURE_POL"],                  "Port to Port"),
    ("Vessel Arrival At POD",                 ["TS_VESSEL_ARRIVAL_POD"],                    "Port to Port"),
    ("Discharge From Vessel At POD",          ["TS_DISCHARGE_FROM_VESSEL_POD"],             "Port to Port"),
    ("Gate Out Full At POD",                  ["TS_GATE_OUT_FULL_POD"],                     "Port to Port"),
    # --- On Carriage ---
    ("Arrival (Rail) Inland Import Terminal", ["TS_ARRIVAL_RAIL_INLAND_IMPORT"],             "On Carriage"),
    ("Arrival At Inland Import Terminal",     ["TS_ARRIVAL_INLAND_IMPORT_TERMINAL"],        "On Carriage"),
    ("Discharge At Inland Import Terminal",   ["TS_DISCHARGE_INLAND_IMPORT_TERMINAL"],      "On Carriage"),
    ("Load At Inland Import Terminal",        ["TS_LOAD_INLAND_IMPORT_TERMINAL"],           "On Carriage"),
    ("Departure From Inland Import Terminal", ["TS_DEPARTURE_INLAND_IMPORT_TERMINAL"],      "On Carriage"),
    ("Out For Delivery",                      ["TS_OUT_FOR_DELIVERY"],                      "On Carriage"),
    ("Arrival Of Full Container At Consignee",["TS_ARRIVAL_AT_CONSIGNEE"],                  "On Carriage"),
    ("Proof Of Delivery",                     ["TS_PROOF_OF_DELIVERY"],                     "On Carriage"),
    ("Picked Up Empty From Consignee",        ["TS_EMPTY_PICKUP_FROM_CONSIGNEE"],           "On Carriage"),
    # --- Empty Return ---
    ("Gate In Empty At Terminal",             ["TS_GATE_IN_EMPTY_AT_TERMINAL"],             "Empty Return"),
]
ALL_GNAMES = [g for g, _, _ in ALL_GROUPS]
MILESTONE_CATEGORY = {g: cat for g, _, cat in ALL_GROUPS}

FFW_COL, NVOCC_COL, CARRIER_COL = "MASTER_FFW_NAME", "MASTER_NVOCC_NAME", "MASTER_CARRIER_NAME"

# colours (match the reference workbook)
NAVY, BLUE, F_LBL, F_CNT = "1F4E79", "2E75B6", "D6E4F7", "D9E1F2"
GREEN_F, GREEN_T = "C6EFCE", "276221"
YEL_F, YEL_T     = "FFEB9C", "7D6608"
RED_F, RED_T     = "FFC7CE", "9C0006"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(THIN, THIN, THIN, THIN)


def _nonempty(series: pd.Series) -> pd.Series:
    s = series.astype("string")
    return s.notna() & (s.str.strip() != "") & (s.str.lower() != "nan")


def compute(df: pd.DataFrame, groups=None):
    """Return (overall dict, entity DataFrame, ffw-subtotal DataFrame, N, pcols).

    `groups` is a list of (name, [timestamp_columns], category) tuples — pass a
    subset of ALL_GROUPS to restrict the report to only those milestones.
    Defaults to every milestone (ALL_GROUPS) if not given.
    """
    groups = groups if groups is not None else ALL_GROUPS
    gnames = [g for g, _, _ in groups]
    df = df.copy()
    # ensure every needed timestamp column exists
    for _, cols, _ in groups:
        for c in cols:
            if c not in df.columns:
                df[c] = pd.NA
    # present flag per milestone
    for g, cols, _ in groups:
        flags = pd.concat([_nonempty(df[c]) for c in cols], axis=1)
        df["_" + g] = flags.any(axis=1).astype(int)
    # clean party names
    for c in [FFW_COL, NVOCC_COL, CARRIER_COL]:
        if c not in df.columns:
            df[c] = pd.NA
        df[c] = df[c].where(_nonempty(df[c]), other=None)

    pcols = ["_" + g for g in gnames]
    N = len(df)
    overall = {g: df["_" + g].mean() for g in gnames}

    key_ffw = df[FFW_COL].fillna("—")
    ent = (df.groupby([key_ffw, df[NVOCC_COL].fillna("—"), df[CARRIER_COL].fillna("—")], dropna=False)
             .agg(cnt=(FFW_COL, "size"), **{p: (p, "mean") for p in pcols})
             .reset_index())
    ent.columns = ["FFW", "NVOCC", "CARRIER", "cnt"] + pcols

    fsub = (df.groupby(key_ffw).agg(cnt=(FFW_COL, "size"), **{p: (p, "mean") for p in pcols}).reset_index())
    fsub.columns = ["FFW", "cnt"] + pcols
    return overall, ent, fsub, N, pcols


def band(v):
    if v >= 0.80: return GREEN_F, GREEN_T
    if v >= 0.50: return YEL_F, YEL_T
    return RED_F, RED_T


def build_workbook(overall, ent, fsub, N, pcols, gnames) -> bytes:
    wb = Workbook(); ws = wb.active; ws.title = "Completeness Summary"
    NC = 4 + len(gnames)

    def pct(cell, v, fill=True, bold=False):
        cell.value = v; cell.number_format = "0%"
        f, t = band(v)
        cell.font = Font(name="Calibri", bold=bold, color=t)
        if fill: cell.fill = PatternFill("solid", fgColor=f)
        cell.alignment = Alignment(horizontal="center"); cell.border = BORDER

    r = 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    c = ws.cell(r, 1, f"Overall Completeness  |  Total Records: {N:,}")
    c.font = Font(name="Calibri", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=NAVY)
    r += 1
    for i, g in enumerate(gnames):
        cc = ws.cell(r, i + 1, g); cc.font = Font(name="Calibri", bold=True, color="FFFFFF")
        cc.fill = PatternFill("solid", fgColor=BLUE); cc.alignment = Alignment(horizontal="center", wrap_text=True); cc.border = BORDER
    r += 1
    for i, g in enumerate(gnames):
        cc = ws.cell(r, i + 1, overall[g]); cc.number_format = "0%"
        cc.font = Font(name="Calibri", bold=True, color=YEL_T); cc.fill = PatternFill("solid", fgColor=YEL_F)
        cc.alignment = Alignment(horizontal="center"); cc.border = BORDER
    r += 2

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    c = ws.cell(r, 1, "Completeness by Entity (FFW / NVOCC / Carrier)")
    c.font = Font(name="Calibri", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=NAVY)
    r += 1
    for i, h in enumerate(["FFW Name", "NVOCC Name", "Carrier Name", "Shipment Count"] + gnames):
        cc = ws.cell(r, i + 1, h); cc.font = Font(name="Calibri", bold=True, color="FFFFFF")
        cc.fill = PatternFill("solid", fgColor=BLUE); cc.alignment = Alignment(horizontal="center", wrap_text=True); cc.border = BORDER
    header_row = r
    r += 1

    def lbl(cell, val, bold=False, color="000000"):
        cell.value = val; cell.font = Font(name="Calibri", bold=bold, color=color)
        cell.fill = PatternFill("solid", fgColor=F_LBL); cell.border = BORDER; cell.alignment = Alignment(horizontal="left")

    def cntcell(cell, val, bold=True):
        cell.value = int(val); cell.number_format = "#,##0"; cell.font = Font(name="Calibri", bold=bold, color=NAVY)
        cell.fill = PatternFill("solid", fgColor=F_CNT); cell.alignment = Alignment(horizontal="center"); cell.border = BORDER

    ffws = sorted([f for f in ent["FFW"].unique() if f != "—"])
    order = (["—"] if "—" in ent["FFW"].values else []) + ffws
    for ffw in order:
        rows = ent[ent["FFW"] == ffw].sort_values("cnt", ascending=False)
        for _, row in rows.iterrows():
            lbl(ws.cell(r, 1), row["FFW"]); lbl(ws.cell(r, 2), row["NVOCC"]); lbl(ws.cell(r, 3), row["CARRIER"])
            cntcell(ws.cell(r, 4), row["cnt"], bold=False)
            for i, p in enumerate(pcols): pct(ws.cell(r, 5 + i), row[p])
            r += 1
        s = fsub[fsub["FFW"] == ffw].iloc[0]
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        lbl(ws.cell(r, 1), "Subtotal — No FFW" if ffw == "—" else f"Subtotal — {ffw}", bold=True, color=NAVY)
        ws.cell(r, 2).fill = PatternFill("solid", fgColor=F_LBL); ws.cell(r, 3).fill = PatternFill("solid", fgColor=F_LBL)
        cntcell(ws.cell(r, 4), s["cnt"], bold=True)
        for i, p in enumerate(pcols): pct(ws.cell(r, 5 + i), s[p], bold=True)
        r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    tc = ws.cell(r, 1, "Total"); tc.font = Font(name="Calibri", bold=True, color="FFFFFF"); tc.fill = PatternFill("solid", fgColor=NAVY); tc.border = BORDER
    for cc in (ws.cell(r, 2), ws.cell(r, 3)): cc.fill = PatternFill("solid", fgColor=NAVY); cc.border = BORDER
    tt = ws.cell(r, 4, N); tt.number_format = "#,##0"; tt.font = Font(name="Calibri", bold=True, color="FFFFFF")
    tt.fill = PatternFill("solid", fgColor=NAVY); tt.alignment = Alignment(horizontal="center"); tt.border = BORDER
    for i, g in enumerate(gnames):
        cc = ws.cell(r, 5 + i, overall[g]); cc.number_format = "0%"; cc.font = Font(name="Calibri", bold=True, color="FFFFFF")
        cc.fill = PatternFill("solid", fgColor=NAVY); cc.alignment = Alignment(horizontal="center"); cc.border = BORDER

    ws.freeze_panes = "E" + str(header_row + 1)
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 34; ws.column_dimensions["D"].width = 15
    for i in range(len(gnames)): ws.column_dimensions[get_column_letter(5 + i)].width = 15

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ============================== UI ==============================
st.set_page_config(page_title="OD2D Milestone Completeness", layout="wide")
st.title("OD2D Milestone Completeness Builder")
st.caption("Upload the raw shipment-level export and download the formatted completeness workbook.")

with st.sidebar:
    st.header("Options")
    grain = st.radio("Shipment grain (BOOKING_LEVEL)",
                     ["Container-level only", "Booking-level only", "Every row as exported"], index=0)
    scope = st.radio("Shipments to include",
                     ["All rows", "Completed only"], index=0)
    use_dates = st.checkbox("Filter by shipment created date", value=False)
    d_from = d_to = None
    if use_dates:
        d_from = st.date_input("Created from")
        d_to = st.date_input("Created to")

    st.markdown("---")
    st.subheader("Milestones to include")

    if "selected_milestones" not in st.session_state:
        st.session_state.selected_milestones = list(ALL_GNAMES)

    bcol1, bcol2 = st.columns(2)
    if bcol1.button("Select all", use_container_width=True):
        st.session_state.selected_milestones = list(ALL_GNAMES)
    if bcol2.button("Clear all", use_container_width=True):
        st.session_state.selected_milestones = []

    selected_milestones = st.multiselect(
        "Milestones",
        options=ALL_GNAMES,
        default=st.session_state.selected_milestones,
        key="selected_milestones",
        format_func=lambda m: f"{m}  ·  {MILESTONE_CATEGORY[m]}",
        label_visibility="collapsed",
    )
    st.caption(f"{len(selected_milestones)} of {len(ALL_GNAMES)} milestones selected.")

    st.markdown("---")
    st.caption("Green ≥ 80%  ·  Yellow 50–79%  ·  Red < 50%")

up = st.file_uploader("Raw shipment-level file (CSV or Excel)", type=["csv", "xlsx"])

if up is not None:
    df = pd.read_csv(up, dtype=str) if up.name.lower().endswith(".csv") else pd.read_excel(up, dtype=str)
    st.write(f"Loaded **{len(df):,}** rows.")

    if "BOOKING_LEVEL" in df.columns:
        if grain == "Container-level only":
            df = df[df["BOOKING_LEVEL"] == "container"]
        elif grain == "Booking-level only":
            df = df[df["BOOKING_LEVEL"] == "booking"]
    if scope == "Completed only" and "SHIPMENT_COMPLETED_DT" in df.columns:
        df = df[_nonempty(df["SHIPMENT_COMPLETED_DT"])]
    if use_dates and d_from and d_to and "SHIPMENT_CREATED_DT" in df.columns:
        cr = pd.to_datetime(df["SHIPMENT_CREATED_DT"], errors="coerce")
        df = df[(cr >= pd.Timestamp(d_from)) & (cr < pd.Timestamp(d_to) + pd.Timedelta(days=1))]

    if len(df) == 0:
        st.warning("No rows after filtering — adjust the options in the sidebar.")
    elif not selected_milestones:
        st.warning("No milestones selected — pick at least one milestone in the sidebar.")
    else:
        # keep ALL_GROUPS journey order, restricted to what the user picked
        groups = [g for g in ALL_GROUPS if g[0] in selected_milestones]
        gnames = [g for g, _, _ in groups]

        overall, ent, fsub, N, pcols = compute(df, groups)
        st.subheader(f"Overall completeness — {N:,} shipments  ({len(gnames)} milestones)")
        ov = pd.DataFrame({"Milestone": gnames, "Completeness": [overall[g] for g in gnames]})
        st.dataframe(ov.style.format({"Completeness": "{:.1%}"}), use_container_width=True, hide_index=True)

        xlsx = build_workbook(overall, ent, fsub, N, pcols, gnames)
        st.download_button("⬇️  Download completeness workbook (.xlsx)", data=xlsx,
                           file_name="OD2D_Completeness_By_Milestone.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success("Workbook ready — click the button above to download.")
else:
    st.info("Upload a file to begin. Expected columns include the TS_* milestone timestamps, "
            "MASTER_FFW_NAME, MASTER_NVOCC_NAME, MASTER_CARRIER_NAME, and BOOKING_LEVEL.")
