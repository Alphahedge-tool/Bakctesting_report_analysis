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


def _first_number(value, default=float("nan")):
    match = re.search(r"[-+]?\d+(?:\.\d+)?", str(value).replace(",", ""))
    return float(match.group()) if match else default


def _percent_in_parentheses(value, default=float("nan")):
    match = re.search(r"\(([-+]?\d+(?:\.\d+)?)%\)", str(value))
    return float(match.group(1)) if match else default


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

uploaded = st.file_uploader("Upload MT Quant export", type=["csv", "xlsx", "xlsm"], help="Upload the original report-style export or a workbook generated by this app.")

if not uploaded:
    st.markdown('<div class="hint">Drop your export above. The app automatically includes portfolios such as <b>SL_REX1</b>, <b>TGT_REX2</b>, and later re-entries.</div>', unsafe_allow_html=True)
    st.stop()

try:
    base = parse_trades(uploaded)
    full_detail, full_daily = calculate(base, brokerage_rate, stt_rate, exchange_rate, sebi_rate, stamp_rate, gst_rate)
except Exception as exc:
    st.error(f"Could not read this file: {exc}")
    st.stop()

with st.sidebar:
    st.divider()
    st.header("Market filters")
    market_available = full_daily["VIX Start"].notna().any()
    if market_available:
        vix_basis = st.selectbox("VIX basis", ["VIX Start", "VIX End", "VIX Average"], index=0)
        vix_low = float(full_daily[vix_basis].min())
        vix_high = float(full_daily[vix_basis].max())
        vc1, vc2 = st.columns(2)
        with vc1:
            selected_vix_min = st.number_input("Min VIX", value=vix_low, min_value=0.0, step=0.1, format="%.2f")
        with vc2:
            selected_vix_max = st.number_input("Max VIX", value=vix_high, min_value=0.0, step=0.1, format="%.2f")

        underlying_low = float(full_daily["Underlying Change"].min())
        underlying_high = float(full_daily["Underlying Change"].max())
        uc1, uc2 = st.columns(2)
        with uc1:
            selected_underlying_min = st.number_input("Min underlying change", value=underlying_low, step=10.0, format="%.2f")
        with uc2:
            selected_underlying_max = st.number_input("Max underlying change", value=underlying_high, step=10.0, format="%.2f")
        filter_mask = (
            full_daily[vix_basis].between(selected_vix_min, selected_vix_max, inclusive="both")
            & full_daily["Underlying Change"].between(selected_underlying_min, selected_underlying_max, inclusive="both")
        )
        filtered_dates = full_daily.loc[filter_mask, "Date"]
        st.caption(f"Showing {len(filtered_dates)} of {len(full_daily)} trading days")
    else:
        st.warning("This workbook has no VIX fields. Upload the original MT Quant CSV to use market filters.")
        filtered_dates = full_daily["Date"]

daily = full_daily[full_daily["Date"].isin(filtered_dates)].copy()
detail = full_detail[full_detail["Date"].isin(filtered_dates)].copy()
daily["Cumulative Premium Turnover"] = daily["Premium Turnover"].cumsum()
daily["Cumulative Net P&L"] = daily["Net P&L After Charges"].cumsum()
if daily.empty:
    st.warning("No trading days match these filters. Widen the VIX or underlying-change range.")
    st.stop()

total = detail.sum(numeric_only=True)
portfolio_count = detail["Portfolio"].nunique()
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Trading days", f"{daily.shape[0]:,}")
c2.metric("Trade legs", f"{detail.shape[0]:,}", help=f"Across {portfolio_count} portfolio blocks")
c3.metric("Premium turnover", MONEY.format(total["Premium Turnover"]))
c4.metric("Total brokerage", MONEY.format(total["Total Brokerage (Including All Charges)"]))
c5.metric("Net P&L", MONEY.format(total["Net P&L After Charges"]))

st.caption(f"Filtered result · {len(daily)} of {len(full_daily)} days · {detail.shape[0]} trade legs")
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Overview", "VIX analysis", "Daily breakdown", "Trade legs", "Charges audit"])
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
        band_labels = ["Very Low · Below 12", "Low · 12–15", "Medium · 15–18", "High · 18–22", "Very High · Above 22"]
        bands = pd.cut(full_daily["VIX Start"], bins=[0, 12, 15, 18, 22, float("inf")], labels=band_labels, right=False)
        band_view = full_daily.assign(**{"VIX Band": bands}).groupby("VIX Band", observed=True).agg(
            Days=("Date", "count"), Gross_PnL=("Gross P&L", "sum"), Charges=("Total Brokerage (Including All Charges)", "sum"), Net_PnL=("Net P&L After Charges", "sum")
        ).reset_index()
        st.subheader("All-days VIX bands")
        st.dataframe(band_view.style.format({"Gross_PnL":"₹{:,.2f}", "Charges":"₹{:,.2f}", "Net_PnL":"₹{:,.2f}"}), width="stretch", hide_index=True)
        st.caption("This band table uses all uploaded days; sidebar filters control the other results.")

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
    display_daily = daily.copy()
    display_daily["Date"] = display_daily["Date"].dt.strftime("%d-%b-%Y")
    money_cols = [c for c in display_daily.columns if any(term in c for term in ["Premium", "P&L", "Brokerage", "STT", "Transaction Charges", "SEBI Charges", "Stamp Duty", "GST"])]
    st.dataframe(display_daily.style.format({c: "₹{:,.2f}" for c in money_cols}), width="stretch", height=540, hide_index=True)

with tab4:
    portfolios = sorted(detail["Portfolio"].unique())
    selected = st.multiselect("Portfolio blocks", portfolios, default=portfolios)
    filtered = detail[detail["Portfolio"].isin(selected)].copy()
    filtered["Date"] = filtered["Date"].dt.strftime("%d-%b-%Y")
    st.dataframe(filtered, width="stretch", height=560, hide_index=True)

with tab5:
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
    "Applied VIX filter": f"{selected_vix_min:.2f} to {selected_vix_max:.2f} ({vix_basis})" if market_available else "Not available",
    "Applied underlying-change filter": f"{selected_underlying_min:.2f} to {selected_underlying_max:.2f}" if market_available else "Not available",
    "Included trading days": len(daily),
}
download = excel_bytes(detail, daily, rates)
st.download_button("Download complete Excel report", download, file_name="Brokerage Calculation.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
