from __future__ import annotations

import csv
import io
import re
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from streamlit.runtime.scriptrunner import get_script_run_ctx


# Avoid confusing NameError cascades when this file is launched with
# `python app.py` instead of through the Streamlit runtime.
if get_script_run_ctx(suppress_warning=True) is None:
    print("This is a Streamlit app. Start it with: streamlit run app.py")
    raise SystemExit(0)


st.set_page_config(
    page_title="Options Brokerage Studio",
    page_icon="₹",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@600;700;800&display=swap');
      :root { --ink:#172033; --muted:#667085; --blue:#246bfd; --cyan:#06b6d4; }
      html, body, [class*="css"] { font-family:'DM Sans',sans-serif; }
      h1,h2,h3 { font-family:'Manrope',sans-serif !important; letter-spacing:-.035em; }
      .stApp { background:linear-gradient(145deg,#f5f8ff 0%,#fbfdff 46%,#f4fbfc 100%); color:var(--ink); }
      [data-testid="stSidebar"] { background:#101827; border-right:1px solid #263247; }
      [data-testid="stSidebar"] * { color:#eef4ff; }
      [data-testid="stSidebar"] input { color:#172033 !important; }
      .hero { padding:1.35rem 1.55rem; border-radius:22px; color:white; margin-bottom:1.1rem;
              background:radial-gradient(circle at 85% 15%,rgba(34,211,238,.32),transparent 28%),
                         linear-gradient(120deg,#172554,#1d4ed8 58%,#0891b2);
              box-shadow:0 20px 45px rgba(30,64,175,.18); }
      .hero h1 { margin:0 0 .25rem; font-size:2.15rem; }
      .hero p { margin:0; color:#dbeafe; max-width:760px; }
      [data-testid="stMetric"] { background:rgba(255,255,255,.88); border:1px solid #e3eaf5;
        padding:1rem 1.05rem; border-radius:16px; box-shadow:0 8px 24px rgba(30,50,90,.06); }
      [data-testid="stMetricLabel"] { color:#667085; }
      [data-testid="stMetricValue"] { color:#172033; font-family:'Manrope',sans-serif; }
      .hint { border:1px solid #dbe7ff; background:#f1f6ff; padding:.8rem 1rem; border-radius:12px; color:#344054; }
      .stDownloadButton button, .stButton button { border-radius:11px; font-weight:700; }
      div[data-testid="stDataFrame"] { border:1px solid #e4eaf3; border-radius:14px; overflow:hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)


MONEY = "₹{:,.2f}"
CORE_COLUMNS = [
    "Date", "Portfolio", "Leg", "Transaction", "Strike", "Option Type", "Expiry",
    "Entry/Sell Premium", "Exit/Buy Premium", "Quantity", "Orders", "Export Brokerage",
    "Exit Reason", "Start Time", "End Time",
]
MARKET_COLUMNS = ["VIX Start", "VIX End", "VIX Average", "Underlying Change", "Gap Change", "Gap Change %"]
BASE_COLUMNS = CORE_COLUMNS + MARKET_COLUMNS
VIX_BINS = [0, 12, 15, 18, 22, float("inf")]
VIX_BAND_LABELS = ["Very Low · Below 12", "Low · 12–15", "Medium · 15–18", "High · 18–22", "Very High · Above 22"]


def _first_number(value, default=float("nan")):
    match = re.search(r"[-+]?\d+(?:\.\d+)?", str(value).replace(",", ""))
    return float(match.group()) if match else default


def _percent_in_parentheses(value, default=float("nan")):
    match = re.search(r"\(([-+]?\d+(?:\.\d+)?)%\)", str(value))
    return float(match.group(1)) if match else default


def _vix_band(frame: pd.DataFrame, basis: str) -> pd.Series:
    return pd.cut(frame[basis], bins=VIX_BINS, labels=VIX_BAND_LABELS, right=False)


def _rows_from_upload(uploaded) -> list[list]:
    raw = uploaded.getvalue()
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        text = raw.decode("utf-8-sig", errors="replace")
        return list(csv.reader(io.StringIO(text)))
    if name.endswith((".xlsx", ".xlsm")):
        book = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        # Our generated workbook can be uploaded again directly.
        if "Trade Leg Details" in book.sheetnames:
            sheet = book["Trade Leg Details"]
            values = list(sheet.values)
            headers = list(values[0])
            records = []
            for row in values[1:]:
                if not row or row[0] in (None, "TOTAL"):
                    continue
                records.append(dict(zip(headers, row)))
            book.close()
            return ["__CALCULATED_WORKBOOK__", records]
        sheet = book[book.sheetnames[0]]
        rows = [["" if v is None else v for v in row] for row in sheet.iter_rows(values_only=True)]
        book.close()
        return rows
    raise ValueError("Please upload a CSV, XLSX, or XLSM file.")


def parse_trades(uploaded) -> pd.DataFrame:
    rows = _rows_from_upload(uploaded)
    if rows and rows[0] == "__CALCULATED_WORKBOOK__":
        frame = pd.DataFrame(rows[1])
        missing = set(CORE_COLUMNS) - set(frame.columns)
        if missing:
            raise ValueError(f"Workbook is missing columns: {', '.join(sorted(missing))}")
        for column in MARKET_COLUMNS:
            if column not in frame.columns:
                frame[column] = float("nan")
        return frame[BASE_COLUMNS].copy()

    current_date = None
    primary_portfolio = None
    current_market = {column: float("nan") for column in MARKET_COLUMNS}
    trades = []
    for row in rows:
        row = list(row) + [""] * max(0, 15 - len(row))
        if str(row[0]).strip() == "Portfolios" and str(row[1]).strip():
            primary_portfolio = str(row[1]).strip()
            continue
        if str(row[0]).strip() == "Date" and row[1]:
            value = row[1]
            if isinstance(value, datetime):
                current_date = value.date()
            else:
                current_date = pd.to_datetime(value, dayfirst=True).date()
            current_market = {column: float("nan") for column in MARKET_COLUMNS}
            continue
        portfolio = str(row[1]).strip()
        leg = str(row[2]).strip()
        # The first OTM2 summary row contains the day's primary market context.
        # Re-entry summary rows have their own intraday VIX windows, but all legs
        # are deliberately classified using the original day's VIX start/end.
        if current_date and portfolio == primary_portfolio and not leg.lower().startswith("leg") and row[9] != "":
            vix_start = _first_number(row[9])
            vix_end = _first_number(row[10])
            current_market = {
                "VIX Start": vix_start,
                "VIX End": vix_end,
                "VIX Average": (vix_start + vix_end) / 2,
                "Underlying Change": _first_number(row[8]),
                "Gap Change": _first_number(row[7]),
                "Gap Change %": _percent_in_parentheses(row[7]),
            }
            continue
        if current_date and primary_portfolio and portfolio.startswith(primary_portfolio) and leg.lower().startswith("leg"):
            strike = str(row[4]).strip()
            trades.append({
                "Date": current_date, "Portfolio": portfolio, "Leg": leg,
                "Transaction": str(row[3]).strip(), "Strike": strike,
                "Option Type": strike[-2:].upper(), "Expiry": str(row[5]).strip(),
                "Entry/Sell Premium": float(row[7]), "Exit/Buy Premium": float(row[8]),
                "Quantity": int(float(row[9])), "Orders": int(float(row[10])),
                "Export Brokerage": float(row[11] or 0), "Exit Reason": str(row[12]).strip(),
                "Start Time": str(row[13]).strip(), "End Time": str(row[14]).strip(),
                **current_market,
            })
    if not trades:
        raise ValueError("No MT Quant trade legs were found. Upload the complete export, including its Date rows.")
    return pd.DataFrame(trades, columns=BASE_COLUMNS)


def calculate(detail: pd.DataFrame, brokerage, stt, exchange, sebi, stamp, gst):
    d = detail.copy()
    d["Date"] = pd.to_datetime(d["Date"])
    for col in ["Entry/Sell Premium", "Exit/Buy Premium", "Quantity", "Orders"]:
        d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0)
    d["Sell Premium Value"] = d["Entry/Sell Premium"] * d["Quantity"]
    d["Buy Premium Value"] = d["Exit/Buy Premium"] * d["Quantity"]
    d["Premium Turnover"] = d["Sell Premium Value"] + d["Buy Premium Value"]
    d["Gross P&L"] = d["Sell Premium Value"] - d["Buy Premium Value"]
    d["Executed Orders"] = d["Orders"]
    d["Zerodha Brokerage"] = d["Executed Orders"] * brokerage
    d["STT"] = d["Sell Premium Value"] * stt / 100
    d["NSE Transaction Charges"] = d["Premium Turnover"] * exchange / 100
    d["SEBI Charges"] = d["Premium Turnover"] * sebi / 100
    d["Stamp Duty"] = d["Buy Premium Value"] * stamp / 100
    d["GST"] = (d["Zerodha Brokerage"] + d["NSE Transaction Charges"] + d["SEBI Charges"]) * gst / 100
    charge_cols = ["Zerodha Brokerage", "STT", "NSE Transaction Charges", "SEBI Charges", "Stamp Duty", "GST"]
    d["Total Brokerage (Including All Charges)"] = d[charge_cols].sum(axis=1)
    d["Net P&L After Charges"] = d["Gross P&L"] - d["Total Brokerage (Including All Charges)"]

    sums = [
        "Quantity", "Executed Orders", "Sell Premium Value", "Buy Premium Value", "Premium Turnover",
        "Gross P&L", *charge_cols, "Total Brokerage (Including All Charges)", "Net P&L After Charges",
    ]
    daily = d.groupby("Date", as_index=False)[sums].sum()
    market = d.groupby("Date", as_index=False)[MARKET_COLUMNS].first()
    daily = daily.merge(market, on="Date", how="left")
    daily["Trade Legs"] = d.groupby("Date").size().values
    daily["Cumulative Premium Turnover"] = daily["Premium Turnover"].cumsum()
    daily["Cumulative Net P&L"] = daily["Net P&L After Charges"].cumsum()
    order = ["Date", *MARKET_COLUMNS, "Trade Legs", "Quantity", "Executed Orders", "Sell Premium Value", "Buy Premium Value",
             "Premium Turnover", "Cumulative Premium Turnover", "Gross P&L", *charge_cols,
             "Total Brokerage (Including All Charges)", "Net P&L After Charges", "Cumulative Net P&L"]
    return d, daily[order]


def regime_summary(frame: pd.DataFrame, category: str) -> pd.DataFrame:
    summary = frame.groupby(category, dropna=False).agg(
        Days=("Date", "count"),
        Gross_PnL=("Gross P&L", "sum"),
        Charges=("Total Brokerage (Including All Charges)", "sum"),
        Net_PnL=("Net P&L After Charges", "sum"),
        Average_Net_PnL=("Net P&L After Charges", "mean"),
    ).reset_index()
    win_rates = frame.groupby(category, dropna=False)["Net P&L After Charges"].apply(
        lambda values: (values > 0).mean() * 100
    ).reset_index(name="Win Rate %")
    return summary.merge(win_rates, on=category, how="left")


def _trade_label(row: pd.Series) -> str:
    return f"{row['Date'].strftime('%d-%b-%Y')} · {row['Portfolio']} · {row['Leg']}"


def vix_regime_summary(
    frame: pd.DataFrame,
    band_column: str = "VIX Band",
    bands: list[str] | None = None,
) -> pd.DataFrame:
    rows = []
    for band in (bands if bands is not None else VIX_BAND_LABELS):
        band_frame = frame[frame[band_column].astype(str) == band]
        trade_legs = len(band_frame)
        gross_pnl = float(band_frame["Gross P&L"].sum()) if trade_legs else 0.0
        charges = float(band_frame["Total Brokerage (Including All Charges)"].sum()) if trade_legs else 0.0
        net_pnl = float(band_frame["Net P&L After Charges"].sum()) if trade_legs else 0.0
        win_legs = int((band_frame["Net P&L After Charges"] > 0).sum()) if trade_legs else 0
        loss_legs = int((band_frame["Net P&L After Charges"] < 0).sum()) if trade_legs else 0
        if net_pnl > 0:
            result = "Profit"
        elif net_pnl < 0:
            result = "Loss"
        else:
            result = "Break-even"
        rows.append({
            "VIX Band": band,
            "Trade Legs": trade_legs,
            "Winning Trades": win_legs,
            "Losing Trades": loss_legs,
            "Gross P&L": gross_pnl,
            "Charges": charges,
            "Net P&L": net_pnl,
            "Result": result,
        })
    return pd.DataFrame(rows)


def excel_bytes(detail, daily, rates):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format="dd-mmm-yyyy") as writer:
        daily.to_excel(writer, sheet_name="Daily Brokerage Summary", index=False)
        detail.to_excel(writer, sheet_name="Trade Leg Details", index=False)
        pd.DataFrame(rates.items(), columns=["Rate", "Value"]).to_excel(writer, sheet_name="Methodology", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
            for column in ws.columns:
                letter = column[0].column_letter
                ws.column_dimensions[letter].width = min(max(max(len(str(c.value or "")) for c in column) + 2, 12), 34)
    return out.getvalue()


def portfolio_excel_bytes(results, summary):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format="dd-mmm-yyyy") as writer:
        summary.to_excel(writer, sheet_name="All Files Summary", index=False)
        for number, (label, (file_detail, file_daily)) in enumerate(results.items(), start=1):
            file_daily.to_excel(writer, sheet_name=f"{number:02d} Daily", index=False)
            file_detail.to_excel(writer, sheet_name=f"{number:02d} Trades", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
            for column in ws.columns:
                letter = column[0].column_letter
                ws.column_dimensions[letter].width = min(max(max(len(str(c.value or "")) for c in column) + 2, 12), 34)
    return out.getvalue()


st.markdown("""
<section class="hero">
  <h1>Options Brokerage Studio</h1>
  <p>Upload an MT Quant export and turn every original, stop-loss, target and re-entry leg into an auditable brokerage report.</p>
</section>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Charge settings")
    st.caption("Current rates, editable for future changes")
    brokerage_rate = st.number_input("Brokerage per executed order (₹)", value=20.0, min_value=0.0, step=1.0)
    stt_rate = st.number_input("STT on sell premium (%)", value=0.15, min_value=0.0, format="%.5f")
    exchange_rate = st.number_input("NSE transaction charge (%)", value=0.03553, min_value=0.0, format="%.5f")
    sebi_rate = st.number_input("SEBI charge (%)", value=0.0001, min_value=0.0, format="%.5f")
    stamp_rate = st.number_input("Stamp duty on buy premium (%)", value=0.003, min_value=0.0, format="%.5f")
    gst_rate = st.number_input("GST (%)", value=18.0, min_value=0.0, format="%.2f")
    st.info("Rates are applied to every uploaded trade. This produces an estimate; contract-note rounding can differ slightly.")

uploaded_files = st.file_uploader(
    "Upload up to 10 MT Quant exports",
    type=["csv", "xlsx", "xlsm"],
    accept_multiple_files=True,
    help="Upload up to 10 original report-style exports or workbooks generated by this app.",
)

if not uploaded_files:
    st.markdown('<div class="hint">Drop one file for detailed analysis, or up to 10 files for combined portfolio analytics. The app automatically includes portfolios such as <b>SL_REX1</b>, <b>TGT_REX2</b>, and later re-entries.</div>', unsafe_allow_html=True)
    st.stop()
if len(uploaded_files) > 10:
    st.error("Please upload a maximum of 10 files at a time.")
    st.stop()

try:
    comparison_results = {}
    for position, uploaded in enumerate(uploaded_files, start=1):
        label = f"{position} · {uploaded.name}"
        base = parse_trades(uploaded)
        comparison_results[label] = calculate(base, brokerage_rate, stt_rate, exchange_rate, sebi_rate, stamp_rate, gst_rate)
except Exception as exc:
    st.error(f"Could not read this file: {exc}")
    st.stop()

with st.sidebar:
    st.divider()
    if len(comparison_results) > 1:
        st.header("Detailed dataset")
        active_label = st.radio("Use filters and detailed tabs for", list(comparison_results), index=0)
    else:
        active_label = next(iter(comparison_results))
    full_detail, full_daily = comparison_results[active_label]
    st.caption(active_label)
    st.divider()
    st.header("Market filters")
    market_available = full_daily["VIX Start"].notna().any()
    if market_available:
        vix_basis = st.selectbox("VIX basis", ["VIX Start", "VIX End", "VIX Average"], index=0)
        full_daily["VIX Band"] = _vix_band(full_daily, vix_basis)
        filter_mode = st.radio("VIX filter mode", ["Range", "Band"], horizontal=True)
        st.caption("Use Range for exact min/max values or Band to quickly isolate low, medium, or high VIX sessions.")
        vix_low = float(full_daily[vix_basis].min())
        vix_high = float(full_daily[vix_basis].max())
        selected_vix_bands = VIX_BAND_LABELS
        if filter_mode == "Range":
            vc1, vc2 = st.columns(2)
            with vc1:
                selected_vix_min = st.number_input("Min VIX", value=vix_low, min_value=0.0, step=0.1, format="%.2f")
            with vc2:
                selected_vix_max = st.number_input("Max VIX", value=vix_high, min_value=0.0, step=0.1, format="%.2f")
            vix_filter_summary = f"{selected_vix_min:.2f} to {selected_vix_max:.2f} ({vix_basis})"
            vix_filter_mask = full_daily[vix_basis].between(selected_vix_min, selected_vix_max, inclusive="both")
        else:
            selected_vix_bands = st.multiselect(
                "VIX bands",
                VIX_BAND_LABELS,
                default=VIX_BAND_LABELS,
                help="Choose one or more VIX buckets to isolate low-, medium-, or high-volatility trades.",
            )
            vix_filter_summary = f"{', '.join(selected_vix_bands) if selected_vix_bands else 'No bands selected'} ({vix_basis})"
            vix_filter_mask = full_daily["VIX Band"].isin(selected_vix_bands)

        underlying_low = float(full_daily["Underlying Change"].min())
        underlying_high = float(full_daily["Underlying Change"].max())
        uc1, uc2 = st.columns(2)
        with uc1:
            selected_underlying_min = st.number_input("Min underlying change", value=underlying_low, step=10.0, format="%.2f")
        with uc2:
            selected_underlying_max = st.number_input("Max underlying change", value=underlying_high, step=10.0, format="%.2f")
        absolute_moves = full_daily["Underlying Change"].abs().dropna()
        default_range_threshold = float(absolute_moves.quantile(0.35)) if len(absolute_moves) else 50.0
        default_large_threshold = float(absolute_moves.quantile(0.70)) if len(absolute_moves) else 100.0
        regime_range_threshold = st.number_input(
            "Range-bound threshold (points)", value=default_range_threshold, min_value=0.0,
            step=5.0, format="%.2f", help="Absolute underlying change at or below this value is classified as range-bound.",
        )
        large_move_threshold = st.number_input(
            "Large movement threshold (points)", value=default_large_threshold, min_value=0.0,
            step=5.0, format="%.2f", help="Absolute underlying change above this value is classified as large.",
        )
        full_daily["VIX Change"] = full_daily["VIX End"] - full_daily["VIX Start"]
        full_daily["VIX Change %"] = full_daily["VIX Change"].div(full_daily["VIX Start"]).mul(100)
        full_daily["VIX Direction"] = full_daily["VIX Change"].apply(
            lambda value: "VIX Rising" if value > 0.01 else ("VIX Falling" if value < -0.01 else "VIX Flat")
        )
        full_daily["Gap Direction"] = full_daily["Gap Change"].apply(
            lambda value: "Gap Up" if value > 0 else ("Gap Down" if value < 0 else "Flat Open")
        )
        full_daily["Underlying Movement"] = full_daily["Underlying Change"].abs()
        full_daily["Market Regime"] = full_daily["Underlying Movement"].apply(
            lambda value: "Range-bound" if value <= regime_range_threshold else "Trending"
        )
        full_daily["Movement Size"] = full_daily["Underlying Movement"].apply(
            lambda value: "Small movement" if value <= large_move_threshold else "Large movement"
        )
        filter_mask = (
            vix_filter_mask
            & full_daily["Underlying Change"].between(selected_underlying_min, selected_underlying_max, inclusive="both")
        )
        filtered_dates = full_daily.loc[filter_mask, "Date"]
        st.caption(f"Showing {len(filtered_dates)} of {len(full_daily)} trading days")
    else:
        st.warning("This workbook has no VIX fields. Upload the original MT Quant CSV to use market filters.")
        filtered_dates = full_daily["Date"]

daily = full_daily[full_daily["Date"].isin(filtered_dates)].copy()
detail = full_detail[full_detail["Date"].isin(filtered_dates)].copy()
if market_available:
    daily["VIX Band"] = _vix_band(daily, vix_basis)
    detail["VIX Band"] = _vix_band(detail, vix_basis)
daily["Cumulative Premium Turnover"] = daily["Premium Turnover"].cumsum()
daily["Cumulative Net P&L"] = daily["Net P&L After Charges"].cumsum()
if daily.empty:
    st.warning("No trading days match these filters. Widen the VIX or underlying-change range.")
    st.stop()

total = detail.sum(numeric_only=True)
portfolio_count = detail["Portfolio"].nunique()
if len(comparison_results) > 1:
    all_summary_rows = []
    combined_daily_parts = []
    for file_label, (file_detail, file_daily) in comparison_results.items():
        file_total = file_detail.sum(numeric_only=True)
        winning = file_daily[file_daily["Net P&L After Charges"] > 0]
        losing = file_daily[file_daily["Net P&L After Charges"] < 0]
        equity = file_daily["Net P&L After Charges"].cumsum()
        max_drawdown = float((equity - equity.cummax()).min()) if len(equity) else 0.0
        gross_profit = float(winning["Net P&L After Charges"].sum())
        gross_loss = abs(float(losing["Net P&L After Charges"].sum()))
        profit_factor = gross_profit / gross_loss if gross_loss else float("inf")
        all_summary_rows.append({
            "Dataset": file_label,
            "Days": len(file_daily),
            "Trade Legs": len(file_detail),
            "Orders": int(file_total["Executed Orders"]),
            "Premium Turnover": file_total["Premium Turnover"],
            "Gross MTM / P&L": file_total["Gross P&L"],
            "Zerodha Brokerage": file_total["Zerodha Brokerage"],
            "Other Charges": file_total["Total Brokerage (Including All Charges)"] - file_total["Zerodha Brokerage"],
            "Total Charges": file_total["Total Brokerage (Including All Charges)"],
            "Net P&L": file_total["Net P&L After Charges"],
            "Winning Days": len(winning),
            "Losing Days": len(losing),
            "Win Rate %": len(winning) / len(file_daily) * 100 if len(file_daily) else 0,
            "Profit Factor": profit_factor,
            "Best Day": file_daily["Net P&L After Charges"].max(),
            "Worst Day": file_daily["Net P&L After Charges"].min(),
            "Max Drawdown": max_drawdown,
            "VIX Minimum": file_daily["VIX Start"].min(),
            "VIX Average": file_daily["VIX Start"].mean(),
            "VIX Maximum": file_daily["VIX Start"].max(),
        })
        tagged = file_daily.copy()
        tagged["Dataset"] = file_label
        combined_daily_parts.append(tagged)
    all_summary = pd.DataFrame(all_summary_rows)
    combined_daily = pd.concat(combined_daily_parts, ignore_index=True)

    st.markdown("## All-files portfolio analysis")
    grand_turnover = all_summary["Premium Turnover"].sum()
    grand_charges = all_summary["Total Charges"].sum()
    grand_brokerage = all_summary["Zerodha Brokerage"].sum()
    grand_other_charges = all_summary["Other Charges"].sum()
    grand_gross = all_summary["Gross MTM / P&L"].sum()
    grand_net = all_summary["Net P&L"].sum()
    ac1, ac2, ac3, ac4, ac5 = st.columns(5)
    ac1.metric("Files analysed", len(all_summary))
    ac2.metric("Combined turnover", MONEY.format(grand_turnover))
    ac3.metric("Combined gross MTM", MONEY.format(grand_gross))
    ac4.metric("Combined charges", MONEY.format(grand_charges))
    ac5.metric("Combined net P&L", MONEY.format(grand_net))
    bc1, bc2, bc3 = st.columns(3)
    bc1.metric("Zerodha brokerage", MONEY.format(grand_brokerage), help="₹20 × total executed option orders")
    bc2.metric("Other charges", MONEY.format(grand_other_charges), help="STT + NSE + SEBI + stamp duty + GST")
    bc3.metric("All combined charges", MONEY.format(grand_charges), help="Zerodha brokerage + all other charges")

    portfolio_tab, monthly_tab, vix_portfolio_tab, regime_portfolio_tab = st.tabs([
        "Dataset scorecard", "Monthly MTM", "Combined VIX", "Market-regime comparison"
    ])
    with portfolio_tab:
        score_money = ["Premium Turnover", "Gross MTM / P&L", "Zerodha Brokerage", "Other Charges", "Total Charges", "Net P&L", "Best Day", "Worst Day", "Max Drawdown"]
        st.dataframe(
            all_summary.style.format({
                **{column: "₹{:,.2f}" for column in score_money},
                "Win Rate %": "{:.1f}%", "Profit Factor": "{:.2f}",
                "VIX Minimum": "{:.2f}", "VIX Average": "{:.2f}", "VIX Maximum": "{:.2f}",
            }),
            width="stretch", height=390, hide_index=True,
        )
        score_fig = go.Figure()
        score_fig.add_trace(go.Bar(x=all_summary["Dataset"], y=all_summary["Net P&L"], name="Net P&L", marker_color="#2563eb"))
        score_fig.add_trace(go.Bar(x=all_summary["Dataset"], y=all_summary["Total Charges"], name="Charges", marker_color="#f59e0b"))
        score_fig.update_layout(title="Net P&L and charges by dataset", barmode="group", template="plotly_white", height=400, margin=dict(l=20,r=20,t=55,b=80), xaxis_tickangle=-20)
        st.plotly_chart(score_fig, width="stretch")
        charge_split_fig = go.Figure()
        charge_split_fig.add_trace(go.Bar(x=all_summary["Dataset"], y=all_summary["Zerodha Brokerage"], name="Zerodha brokerage", marker_color="#2563eb"))
        charge_split_fig.add_trace(go.Bar(x=all_summary["Dataset"], y=all_summary["Other Charges"], name="Other charges", marker_color="#f59e0b"))
        charge_split_fig.update_layout(title="Brokerage versus all other charges", barmode="stack", template="plotly_white", height=390, margin=dict(l=20,r=20,t=55,b=80), xaxis_tickangle=-20)
        st.plotly_chart(charge_split_fig, width="stretch")
    with monthly_tab:
        combined_monthly = combined_daily.copy()
        combined_monthly["Month"] = combined_monthly["Date"].dt.to_period("M").astype(str)
        combined_monthly = combined_monthly.groupby(["Dataset", "Month"], as_index=False).agg(
            Gross_MTM=("Gross P&L", "sum"), Charges=("Total Brokerage (Including All Charges)", "sum"), Net_PnL=("Net P&L After Charges", "sum")
        )
        monthly_fig = go.Figure()
        for dataset_name, month_group in combined_monthly.groupby("Dataset"):
            monthly_fig.add_trace(go.Scatter(x=month_group["Month"], y=month_group["Net_PnL"], mode="lines+markers", name=dataset_name))
        monthly_fig.add_hline(y=0, line_dash="dot", line_color="#94a3b8")
        monthly_fig.update_layout(title="Monthly net MTM / P&L across all files", template="plotly_white", height=440, margin=dict(l=20,r=20,t=55,b=70), xaxis_tickangle=-45)
        st.plotly_chart(monthly_fig, width="stretch")
        st.dataframe(combined_monthly.style.format({"Gross_MTM":"₹{:,.2f}", "Charges":"₹{:,.2f}", "Net_PnL":"₹{:,.2f}"}), width="stretch", height=330, hide_index=True)
    with vix_portfolio_tab:
        combined_daily["VIX Band"] = pd.cut(combined_daily["VIX Start"], bins=VIX_BINS, labels=VIX_BAND_LABELS, right=False)
        all_vix = combined_daily.groupby(["Dataset", "VIX Band"], observed=False).agg(
            Days=("Date", "count"), Gross_MTM=("Gross P&L", "sum"), Charges=("Total Brokerage (Including All Charges)", "sum"), Net_PnL=("Net P&L After Charges", "sum")
        ).reset_index()
        st.dataframe(all_vix.style.format({"Gross_MTM":"₹{:,.2f}", "Charges":"₹{:,.2f}", "Net_PnL":"₹{:,.2f}"}), width="stretch", height=430, hide_index=True)
        vix_fig = go.Figure()
        for dataset_name, vix_group in all_vix.groupby("Dataset"):
            vix_fig.add_trace(go.Bar(x=vix_group["VIX Band"].astype(str), y=vix_group["Net_PnL"], name=dataset_name))
        vix_fig.update_layout(title="Net P&L by VIX band and dataset", barmode="group", template="plotly_white", height=430, margin=dict(l=20,r=20,t=55,b=70))
        st.plotly_chart(vix_fig, width="stretch")
    with regime_portfolio_tab:
        if not market_available:
            st.warning("Upload original MT Quant CSV exports to compare market regimes.")
        else:
            regime_data = combined_daily.copy()
            regime_data["VIX Change"] = regime_data["VIX End"] - regime_data["VIX Start"]
            regime_data["VIX Change %"] = regime_data["VIX Change"].div(regime_data["VIX Start"]).mul(100)
            regime_data["VIX Direction"] = regime_data["VIX Change"].apply(
                lambda value: "VIX Rising" if value > 0.01 else ("VIX Falling" if value < -0.01 else "VIX Flat")
            )
            regime_data["Gap Direction"] = regime_data["Gap Change"].apply(
                lambda value: "Gap Up" if value > 0 else ("Gap Down" if value < 0 else "Flat Open")
            )
            regime_data["Underlying Movement"] = regime_data["Underlying Change"].abs()
            regime_data["Market Regime"] = regime_data["Underlying Movement"].apply(
                lambda value: "Range-bound" if value <= regime_range_threshold else "Trending"
            )
            regime_data["Movement Size"] = regime_data["Underlying Movement"].apply(
                lambda value: "Small movement" if value <= large_move_threshold else "Large movement"
            )
            regime_choices = {
                "VIX rising versus falling": "VIX Direction",
                "Gap-up versus gap-down": "Gap Direction",
                "Trending versus range-bound": "Market Regime",
                "Large versus small movement": "Movement Size",
            }
            selected_regime_name = st.selectbox("Choose a market regime to compare", list(regime_choices))
            selected_regime_column = regime_choices[selected_regime_name]
            comparison_table = regime_data.groupby(["Dataset", selected_regime_column], dropna=False).agg(
                Days=("Date", "count"),
                Gross_PnL=("Gross P&L", "sum"),
                Charges=("Total Brokerage (Including All Charges)", "sum"),
                Net_PnL=("Net P&L After Charges", "sum"),
                Average_Net_PnL=("Net P&L After Charges", "mean"),
            ).reset_index()
            comparison_wins = regime_data.groupby(["Dataset", selected_regime_column], dropna=False)["Net P&L After Charges"].apply(
                lambda values: (values > 0).mean() * 100
            ).reset_index(name="Win Rate %")
            comparison_table = comparison_table.merge(
                comparison_wins, on=["Dataset", selected_regime_column], how="left"
            )
            st.dataframe(
                comparison_table.style.format({
                    "Gross_PnL": "₹{:,.2f}", "Charges": "₹{:,.2f}", "Net_PnL": "₹{:,.2f}",
                    "Average_Net_PnL": "₹{:,.2f}", "Win Rate %": "{:.1f}%",
                }),
                width="stretch", height=420, hide_index=True,
            )
            regime_compare_fig = go.Figure()
            for dataset_name, dataset_group in comparison_table.groupby("Dataset"):
                regime_compare_fig.add_trace(go.Bar(
                    x=dataset_group[selected_regime_column].astype(str), y=dataset_group["Net_PnL"],
                    name=dataset_name,
                ))
            regime_compare_fig.add_hline(y=0, line_color="#94a3b8")
            regime_compare_fig.update_layout(
                title=f"Strategy profit comparison · {selected_regime_name}", barmode="group",
                template="plotly_white", height=450, margin=dict(l=20,r=20,t=55,b=70),
                yaxis_title="Net P&L after charges",
            )
            st.plotly_chart(regime_compare_fig, width="stretch")
            profit_pivot = comparison_table.pivot(
                index="Dataset", columns=selected_regime_column, values="Net_PnL"
            ).reset_index()
            st.markdown("#### Net-profit matrix")
            st.dataframe(
                profit_pivot.style.format({column: "₹{:,.2f}" for column in profit_pivot.columns if column != "Dataset"}),
                width="stretch", hide_index=True,
            )

st.markdown(f"## Detailed analysis · {active_label}")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Trading days", f"{daily.shape[0]:,}")
c2.metric("Trade legs", f"{detail.shape[0]:,}", help=f"Across {portfolio_count} portfolio blocks")
c3.metric("Premium turnover", MONEY.format(total["Premium Turnover"]))
c4.metric("Total brokerage", MONEY.format(total["Total Brokerage (Including All Charges)"]))
c5.metric("Net P&L", MONEY.format(total["Net P&L After Charges"]))
selected_brokerage = total["Zerodha Brokerage"]
selected_all_charges = total["Total Brokerage (Including All Charges)"]
selected_other_charges = selected_all_charges - selected_brokerage
sc1, sc2, sc3 = st.columns(3)
sc1.metric("Zerodha brokerage", MONEY.format(selected_brokerage), help="₹20 × executed option orders")
sc2.metric("Other charges", MONEY.format(selected_other_charges), help="STT + NSE + SEBI + stamp duty + GST")
sc3.metric("All charges", MONEY.format(selected_all_charges), help="Zerodha brokerage + all other charges")

if len(comparison_results) == 2:
    st.markdown("## Side-by-side comparison")
    st.caption("Both files use the same charge rates. Use the sidebar selector to open either file in the detailed tabs below.")
    compare_columns = st.columns(2, gap="large")
    for container, (label, (compare_detail, compare_daily)) in zip(compare_columns, comparison_results.items()):
        with container:
            compare_total = compare_detail.sum(numeric_only=True)
            st.markdown(f"### {label}")
            sm1, sm2, sm3 = st.columns(3)
            sm1.metric("Days", f"{len(compare_daily):,}")
            sm2.metric("Legs", f"{len(compare_detail):,}")
            sm3.metric("Orders", f"{int(compare_total['Executed Orders']):,}")
            st.metric("Premium turnover", MONEY.format(compare_total["Premium Turnover"]))
            pn1, pn2 = st.columns(2)
            pn1.metric("Total charges", MONEY.format(compare_total["Total Brokerage (Including All Charges)"]))
            pn2.metric("Net P&L", MONEY.format(compare_total["Net P&L After Charges"]))

            st.markdown("#### India VIX comparison")
            vx1, vx2, vx3 = st.columns(3)
            vx1.metric("Minimum VIX", f"{compare_daily['VIX Start'].min():.2f}")
            vx2.metric("Average VIX", f"{compare_daily['VIX Start'].mean():.2f}")
            vx3.metric("Maximum VIX", f"{compare_daily['VIX Start'].max():.2f}")
            compare_bands = pd.cut(
                compare_daily["VIX Start"], bins=VIX_BINS, labels=VIX_BAND_LABELS, right=False
            )
            compare_vix_table = compare_daily.assign(**{"VIX Band": compare_bands}).groupby(
                "VIX Band", observed=False
            ).agg(
                Days=("Date", "count"),
                Charges=("Total Brokerage (Including All Charges)", "sum"),
                Net_PnL=("Net P&L After Charges", "sum"),
            ).reset_index()
            st.dataframe(
                compare_vix_table.style.format({"Charges": "₹{:,.2f}", "Net_PnL": "₹{:,.2f}"}),
                width="stretch", hide_index=True,
            )

            compare_charge_view = pd.DataFrame({
                "Charge": ["Brokerage", "STT", "NSE", "SEBI", "Stamp", "GST", "TOTAL"],
                "Amount": [
                    compare_total["Zerodha Brokerage"], compare_total["STT"],
                    compare_total["NSE Transaction Charges"], compare_total["SEBI Charges"],
                    compare_total["Stamp Duty"], compare_total["GST"],
                    compare_total["Total Brokerage (Including All Charges)"],
                ],
            })
            st.dataframe(compare_charge_view.style.format({"Amount": "₹{:,.2f}"}), width="stretch", hide_index=True)

            compare_dates = compare_daily[[
                "Date", "VIX Start", "VIX End", "Underlying Change", "Premium Turnover",
                "Gross P&L", "Total Brokerage (Including All Charges)", "Net P&L After Charges",
            ]].copy()
            compare_dates["Date"] = compare_dates["Date"].dt.strftime("%d-%b-%Y")
            st.dataframe(
                compare_dates.style.format({
                    "VIX Start": "{:.2f}", "VIX End": "{:.2f}", "Underlying Change": "{:+,.2f}",
                    "Premium Turnover": "₹{:,.2f}", "Gross P&L": "₹{:,.2f}",
                    "Total Brokerage (Including All Charges)": "₹{:,.2f}", "Net P&L After Charges": "₹{:,.2f}",
                }),
                width="stretch", height=390, hide_index=True,
            )

st.caption(f"Filtered result · {len(daily)} of {len(full_daily)} days · {detail.shape[0]} trade legs")
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Overview", "VIX analysis", "Market regimes", "Daily breakdown", "Trade legs", "Charges audit"])
with tab1:
    left, right = st.columns([1.6, 1])
    with left:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=daily["Date"], y=daily["Gross P&L"], name="Gross P&L", marker_color="#60a5fa"))
        fig.add_trace(go.Scatter(x=daily["Date"], y=daily["Cumulative Net P&L"], name="Cumulative net P&L", yaxis="y2", line=dict(color="#0f766e", width=3)))
        fig.update_layout(title="Performance after all charges", template="plotly_white", height=410, margin=dict(l=20,r=20,t=55,b=20),
                          legend=dict(orientation="h", y=1.1), yaxis2=dict(overlaying="y", side="right", showgrid=False))
        st.plotly_chart(fig, width="stretch")
    with right:
        charges = pd.Series({c: total[c] for c in ["Zerodha Brokerage", "STT", "NSE Transaction Charges", "SEBI Charges", "Stamp Duty", "GST"]})
        pie = go.Figure(go.Pie(labels=charges.index, values=charges.values, hole=.62, marker_colors=["#2563eb","#06b6d4","#8b5cf6","#64748b","#f59e0b","#14b8a6"]))
        pie.update_layout(title="Charge composition", template="plotly_white", height=410, margin=dict(l=10,r=10,t=55,b=15), showlegend=True)
        st.plotly_chart(pie, width="stretch")

with tab2:
    vleft, vright = st.columns([1.45, 1])
    with vleft:
        scatter = go.Figure(go.Scatter(
            x=daily["VIX Start"], y=daily["Net P&L After Charges"], mode="markers",
            text=daily["Date"].dt.strftime("%d-%b-%Y"),
            customdata=daily[["Underlying Change", "Total Brokerage (Including All Charges)"]],
            marker=dict(size=11, color=daily["Underlying Change"], colorscale="RdBu", colorbar=dict(title="Underlying<br>change"), line=dict(width=1, color="white")),
            hovertemplate="%{text}<br>VIX start: %{x:.2f}<br>Net P&L: ₹%{y:,.2f}<br>Underlying: %{customdata[0]:,.2f}<br>Charges: ₹%{customdata[1]:,.2f}<extra></extra>",
        ))
        scatter.add_hline(y=0, line_dash="dot", line_color="#94a3b8")
        scatter.update_layout(title="VIX versus net P&L", template="plotly_white", height=430, margin=dict(l=20,r=20,t=55,b=20), xaxis_title="India VIX start", yaxis_title="Net P&L after charges")
        st.plotly_chart(scatter, width="stretch")
    with vright:
        band_labels = VIX_BAND_LABELS
        bands = _vix_band(full_daily, vix_basis)
        band_view = full_daily.assign(**{"VIX Band": bands}).groupby("VIX Band", observed=True).agg(
            Days=("Date", "count"), Gross_PnL=("Gross P&L", "sum"), Charges=("Total Brokerage (Including All Charges)", "sum"), Net_PnL=("Net P&L After Charges", "sum")
        ).reset_index()
        st.subheader("All-days VIX bands")
        st.dataframe(band_view.style.format({"Gross_PnL":"₹{:,.2f}", "Charges":"₹{:,.2f}", "Net_PnL":"₹{:,.2f}"}), width="stretch", hide_index=True)
        st.caption(f"This band table uses the selected VIX basis ({vix_basis}); sidebar filters control the other results.")

    st.subheader("Dates by VIX band")
    selected_band = st.selectbox("Choose a VIX band", ["All VIX bands", *band_labels], index=0)
    date_view = full_daily.assign(**{"VIX Band": bands})
    if selected_band != "All VIX bands":
        date_view = date_view[date_view["VIX Band"].astype(str) == selected_band]
    date_view = date_view[[
        "Date", "VIX Band", "VIX Start", "VIX End", "VIX Average", "Underlying Change",
        "Gross P&L", "Total Brokerage (Including All Charges)", "Net P&L After Charges",
    ]].copy()
    date_view["Date"] = date_view["Date"].dt.strftime("%d-%b-%Y")
    st.caption(f"{len(date_view)} trading dates in {selected_band.lower()}")
    st.dataframe(
        date_view.style.format({
            "VIX Start": "{:.2f}", "VIX End": "{:.2f}", "VIX Average": "{:.2f}",
            "Underlying Change": "{:+,.2f}", "Gross P&L": "₹{:,.2f}",
            "Total Brokerage (Including All Charges)": "₹{:,.2f}",
            "Net P&L After Charges": "₹{:,.2f}",
        }),
        width="stretch", height=430, hide_index=True,
    )

with tab3:
    if not market_available:
        st.warning("Upload the original MT Quant CSV to analyse market regimes.")
    else:
        st.subheader("Market-condition performance")
        st.caption(
            f"Current definitions: range-bound ≤ {regime_range_threshold:.2f} points; "
            f"large movement > {large_move_threshold:.2f} points. Change these thresholds in the sidebar."
        )
        regime_specs = [
            ("VIX Direction", "VIX rising versus falling"),
            ("Gap Direction", "Gap-up versus gap-down"),
            ("Market Regime", "Trending versus range-bound"),
            ("Movement Size", "Large versus small underlying movement"),
        ]
        regime_columns = st.columns(2, gap="large")
        for index, (category, title) in enumerate(regime_specs):
            with regime_columns[index % 2]:
                st.markdown(f"#### {title}")
                regime_view = regime_summary(daily, category)
                st.dataframe(
                    regime_view.style.format({
                        "Gross_PnL": "₹{:,.2f}", "Charges": "₹{:,.2f}",
                        "Net_PnL": "₹{:,.2f}", "Average_Net_PnL": "₹{:,.2f}",
                        "Win Rate %": "{:.1f}%",
                    }),
                    width="stretch", hide_index=True,
                )
                regime_chart = go.Figure(go.Bar(
                    x=regime_view[category].astype(str), y=regime_view["Net_PnL"],
                    marker_color=["#2563eb" if value >= 0 else "#ef4444" for value in regime_view["Net_PnL"]],
                    text=[f"₹{value:,.0f}" for value in regime_view["Net_PnL"]], textposition="auto",
                ))
                regime_chart.add_hline(y=0, line_color="#94a3b8")
                regime_chart.update_layout(template="plotly_white", height=300, margin=dict(l=10,r=10,t=20,b=45), showlegend=False)
                st.plotly_chart(regime_chart, width="stretch")

        st.markdown("#### VIX percentage-change analysis")
        vix_change_left, vix_change_right = st.columns([1.3, 1])
        with vix_change_left:
            vix_change_chart = go.Figure(go.Scatter(
                x=daily["VIX Change %"], y=daily["Net P&L After Charges"], mode="markers",
                text=daily["Date"].dt.strftime("%d-%b-%Y"),
                marker=dict(size=11, color=daily["Underlying Movement"], colorscale="Viridis", colorbar=dict(title="Absolute<br>move")),
                hovertemplate="%{text}<br>VIX change: %{x:+.2f}%<br>Net P&L: ₹%{y:,.2f}<extra></extra>",
            ))
            vix_change_chart.add_vline(x=0, line_dash="dot", line_color="#94a3b8")
            vix_change_chart.add_hline(y=0, line_dash="dot", line_color="#94a3b8")
            vix_change_chart.update_layout(template="plotly_white", height=420, margin=dict(l=20,r=20,t=30,b=50), xaxis_title="VIX change %", yaxis_title="Net P&L after charges")
            st.plotly_chart(vix_change_chart, width="stretch")
        with vix_change_right:
            condition_dates = daily[[
                "Date", "VIX Start", "VIX End", "VIX Change %", "VIX Direction",
                "Gap Direction", "Market Regime", "Movement Size", "Underlying Change",
                "Net P&L After Charges",
            ]].copy()
            condition_dates["Date"] = condition_dates["Date"].dt.strftime("%d-%b-%Y")
            st.dataframe(
                condition_dates.style.format({
                    "VIX Start": "{:.2f}", "VIX End": "{:.2f}", "VIX Change %": "{:+.2f}%",
                    "Underlying Change": "{:+,.2f}", "Net P&L After Charges": "₹{:,.2f}",
                }),
                width="stretch", height=420, hide_index=True,
            )

with tab4:
    display_daily = daily.copy()
    display_daily["Date"] = display_daily["Date"].dt.strftime("%d-%b-%Y")
    money_cols = [c for c in display_daily.columns if any(term in c for term in ["Premium", "P&L", "Brokerage", "STT", "Transaction Charges", "SEBI Charges", "Stamp Duty", "GST"])]
    st.dataframe(display_daily.style.format({c: "₹{:,.2f}" for c in money_cols}), width="stretch", height=540, hide_index=True)

with tab5:
    st.subheader("Trade legs by VIX regime")
    st.caption("Pick the VIX bands you want to inspect. Select `Low` and `Medium` to see the trades taken in lower-volatility sessions.")
    vix_band_selection = st.multiselect(
        "VIX band filter",
        VIX_BAND_LABELS,
        default=["Low · 12–15", "Medium · 15–18"],
        help="Choose one or more VIX bands. Low and Medium are selected by default so you can focus on those trades first.",
    )
    portfolios = sorted(detail["Portfolio"].unique())
    selected = st.multiselect("Portfolio blocks", portfolios, default=portfolios)
    filtered = detail[detail["Portfolio"].isin(selected)].copy()
    if market_available:
        filtered["VIX Band"] = _vix_band(filtered, vix_basis)
        if vix_band_selection:
            filtered = filtered[filtered["VIX Band"].astype(str).isin(vix_band_selection)]
        st.caption("This table is now filtered to the VIX bands you picked above.")
        regime_summary = vix_regime_summary(filtered, bands=vix_band_selection or None)
        st.markdown("#### VIX regime summary")
        st.caption("Each row shows how many trades were taken in that VIX band and whether the total net P&L was a profit or a loss.")
        st.dataframe(
            regime_summary.style.format({
                "Gross P&L": MONEY,
                "Charges": MONEY,
                "Net P&L": MONEY,
            }),
            width="stretch",
            hide_index=True,
        )
    filtered["Date"] = filtered["Date"].dt.strftime("%d-%b-%Y")
    money_cols = [
        column for column in filtered.columns
        if any(term in column for term in ["Premium", "P&L", "Brokerage", "STT", "Transaction Charges", "SEBI Charges", "Stamp Duty", "GST"])
    ]
    st.dataframe(
        filtered.style.format({column: MONEY for column in money_cols}),
        width="stretch",
        height=560,
        hide_index=True,
    )

with tab6:
    audit = pd.DataFrame({
        "Charge": ["Zerodha brokerage", "STT", "NSE transaction charges", "SEBI charges", "Stamp duty", "GST", "FINAL TOTAL"],
        "Calculation basis": ["Executed orders × rate", "Sell premium value", "Buy + sell premium turnover", "Buy + sell premium turnover", "Buy premium value", "Brokerage + NSE + SEBI", "Sum of every charge"],
        "Amount": [total["Zerodha Brokerage"], total["STT"], total["NSE Transaction Charges"], total["SEBI Charges"], total["Stamp Duty"], total["GST"], total["Total Brokerage (Including All Charges)"]],
    })
    st.dataframe(audit.style.format({"Amount": "₹{:,.2f}"}), width="stretch", hide_index=True)
    st.success(f"P&L reconciliation: sell premium − buy premium = {MONEY.format(total['Gross P&L'])}")

rates = {
    "Rate basis": "Current user-selected rates applied to all uploaded trades",
    "Brokerage per executed order": brokerage_rate, "STT % on sell premium": stt_rate,
    "NSE transaction charge %": exchange_rate, "SEBI charge %": sebi_rate,
    "Stamp duty % on buy premium": stamp_rate, "GST %": gst_rate,
    "Applied VIX filter": vix_filter_summary if market_available else "Not available",
    "Applied underlying-change filter": f"{selected_underlying_min:.2f} to {selected_underlying_max:.2f}" if market_available else "Not available",
    "Range-bound threshold": f"{regime_range_threshold:.2f} points" if market_available else "Not available",
    "Large-movement threshold": f"{large_move_threshold:.2f} points" if market_available else "Not available",
    "Included trading days": len(daily),
}
download = excel_bytes(detail, daily, rates)
st.download_button("Download complete Excel report", download, file_name="Brokerage Calculation.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
if len(comparison_results) > 1:
    portfolio_download = portfolio_excel_bytes(comparison_results, all_summary)
    st.download_button(
        "Download all-files analysis workbook", portfolio_download,
        file_name="All Files Brokerage MTM VIX Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

