import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import requests
import plotly.graph_objects as go

st.set_page_config(page_title="Income Tracker", page_icon="💰", layout="wide")

st.markdown("""
<style>
.metric-card {
    background: #ffffff;
    border: 1px solid #e6e6e6;
    border-radius: 12px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    padding: 16px 14px;
    text-align: center;
    margin-bottom: 14px;
}
.metric-card .metric-label {
    font-size: 13px;
    color: #888888;
    margin-bottom: 6px;
}
.metric-card .metric-value {
    font-size: 21px;
    font-weight: 700;
    color: #333333;
}
</style>
""", unsafe_allow_html=True)


def metric_card(label, value):
    st.markdown(
        f'<div class="metric-card"><div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div></div>',
        unsafe_allow_html=True,
    )

# ============================================
WEBSITES = ["Shutterstock", "Adobe", "Getty", "Dreamtime", "123RF", "Deposit", "Freepik", "Colorbox"]
WEBSITE_CURRENCY = {
    "Shutterstock": "USD", "Adobe": "USD", "Getty": "USD", "Dreamtime": "USD",
    "123RF": "USD", "Deposit": "USD", "Freepik": "USD", "Colorbox": "EUR",
}
# Pastel palette — soft, muted, easy on the eyes
WEBSITE_COLORS = {
    "Shutterstock": "#AEDFF7", "Adobe": "#FFD8B8", "Getty": "#C3EDC0", "Dreamtime": "#FFB6B9",
    "123RF": "#D7C4F2", "Deposit": "#E3C9C1", "Freepik": "#FFCCE1", "Colorbox": "#D6D6D6",
}
TREND_LINE_COLOR = "#F6A9A9"
INCOME_HEADERS = ["Year", "Month", "Website", "Currency", "Amount", "Entry Date", "Source", "Amount_THB"]
MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

SHEET_NAME = st.secrets.get("INCOME_SHEET_NAME", "PhotoStockIncome")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

CHART_LAYOUT_DEFAULTS = dict(
    template="plotly_white",
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Helvetica, Arial, sans-serif", size=13, color="#444444"),
    margin=dict(t=55, b=40, l=40, r=30),
)


# ============================================
# Google Sheets connection
# ============================================

@st.cache_resource
def get_gspread_client():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    return gspread.authorize(creds)


def get_worksheet():
    client = get_gspread_client()
    sheet = client.open(SHEET_NAME)
    try:
        ws = sheet.worksheet("Income")
    except gspread.WorksheetNotFound:
        ws = sheet.add_worksheet(title="Income", rows=2000, cols=10)
        ws.append_row(INCOME_HEADERS)
    return ws


def load_records():
    ws = get_worksheet()
    values = ws.get_all_values()
    if len(values) < 2:
        return pd.DataFrame(columns=INCOME_HEADERS)

    header = values[0]
    rows = [list(r) for r in values[1:]]

    # Migrate older sheets that don't have the Amount_THB column yet
    if "Amount_THB" not in header:
        ws.update_cell(1, len(header) + 1, "Amount_THB")
        header = header + ["Amount_THB"]

    thb_col = header.index("Amount_THB")
    amount_col = header.index("Amount")
    currency_col = header.index("Currency")

    usd_rate, eur_rate, _ = fetch_rates()
    changed = False
    for r in rows:
        while len(r) <= thb_col:
            r.append("")
        # Any row missing a locked THB value gets one computed now, then frozen permanently
        if not str(r[thb_col]).strip():
            try:
                amount = float(r[amount_col])
                currency = r[currency_col]
                r[thb_col] = round(to_thb(amount, currency, usd_rate, eur_rate), 2)
                changed = True
            except (ValueError, IndexError):
                pass

    if changed and rows:
        a1 = gspread.utils.rowcol_to_a1(1, thb_col + 1)
        col_letter = "".join(ch for ch in a1 if ch.isalpha())
        col_values = [[header[thb_col]]] + [[r[thb_col]] for r in rows]
        ws.update(f"{col_letter}1:{col_letter}{len(rows) + 1}", col_values)

    df = pd.DataFrame(rows, columns=header)
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df["Month"] = pd.to_numeric(df["Month"], errors="coerce")
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    df["Amount_THB"] = pd.to_numeric(df["Amount_THB"], errors="coerce")
    df = df.dropna(subset=["Year", "Month", "Amount", "Amount_THB"])
    df["Year"] = df["Year"].astype(int)
    df["Month"] = df["Month"].astype(int)
    return df


def save_entry(year, month, website, currency, amount):
    usd_rate, eur_rate, _ = fetch_rates()
    thb_value = to_thb(float(amount), currency, usd_rate, eur_rate)
    ws = get_worksheet()
    ws.append_row([
        int(year), int(month), website, currency, float(amount),
        datetime.now().strftime("%Y-%m-%d %H:%M"), "Manual", round(thb_value, 2),
    ])


def already_has_imported_data():
    ws = get_worksheet()
    values = ws.get_all_values()
    return any(row[6] == "Imported" for row in values[1:] if len(row) > 6)


def import_legacy_excel(uploaded_file):
    import openpyxl
    usd_rate, eur_rate, _ = fetch_rates()
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws_src = wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws_src[1]]
    col_index = {h: i for i, h in enumerate(headers)}

    if "Year" not in col_index or "Month" not in col_index:
        return None, "Could not find 'Year' or 'Month' columns in the header row."

    website_cols = [w for w in WEBSITES if w in col_index]
    if not website_cols:
        return None, "Could not find any recognized website columns in the header row."

    rows_to_add = []
    for row in ws_src.iter_rows(min_row=2, values_only=True):
        year = row[col_index["Year"]]
        month = row[col_index["Month"]]
        if year is None or month is None:
            continue
        if hasattr(month, "month"):
            month = month.month
        try:
            year = int(year)
            month = int(month)
        except (ValueError, TypeError):
            continue
        for site in website_cols:
            amount = row[col_index[site]]
            if isinstance(amount, (int, float)) and amount != 0:
                thb_value = to_thb(amount, WEBSITE_CURRENCY[site], usd_rate, eur_rate)
                rows_to_add.append([
                    year, month, site, WEBSITE_CURRENCY[site], amount,
                    datetime.now().strftime("%Y-%m-%d %H:%M"), "Imported", round(thb_value, 2),
                ])

    if not rows_to_add:
        return None, "No importable income data found in this file."

    ws = get_worksheet()
    ws.append_rows(rows_to_add)
    return len(rows_to_add), website_cols


# ============================================
# Exchange rates (free, no API key, cached for 24h)
# ============================================

@st.cache_data(ttl=86400)
def fetch_rates():
    try:
        usd = requests.get("https://api.frankfurter.dev/v2/rate/USD/THB", timeout=8).json()["rate"]
        eur = requests.get("https://api.frankfurter.dev/v2/rate/EUR/THB", timeout=8).json()["rate"]
        return float(usd), float(eur), True
    except Exception:
        return 33.0, 36.0, False


def to_thb(amount, currency, usd_rate, eur_rate):
    if currency == "USD":
        return amount * usd_rate
    if currency == "EUR":
        return amount * eur_rate
    return amount


# ============================================
# Chart builders (Plotly — pastel, interactive, native fullscreen)
# ============================================

def make_stacked_bar(categories, pivot_df, title):
    fig = go.Figure()
    for site in WEBSITES:
        if site in pivot_df.columns and pivot_df[site].sum() > 0:
            fig.add_trace(go.Bar(
                x=categories,
                y=pivot_df[site],
                name=site,
                marker_color=WEBSITE_COLORS.get(site, "#CCCCCC"),
                hovertemplate=f"<b>{site}</b>: %{{y:,.0f}} THB<extra></extra>",
            ))

    totals = pivot_df.sum(axis=1)
    fig.add_trace(go.Scatter(
        x=categories, y=totals, mode="markers",
        marker=dict(opacity=0, size=1),
        showlegend=False,
        hovertemplate="<b>Total: %{y:,.0f} THB</b><extra></extra>",
        name="Total",
    ))

    fig.update_layout(
        barmode="stack",
        title=title,
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=10)),
        hovermode="x unified",
        **{**CHART_LAYOUT_DEFAULTS, "margin": dict(t=55, b=40, l=40, r=140)},
    )
    fig.update_xaxes(tickangle=-45, showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="#EEEEEE")
    return fig


def make_line(x, y, title):
    fig = go.Figure(go.Scatter(
        x=x, y=y, mode="lines+markers",
        line=dict(color=TREND_LINE_COLOR, width=3),
        marker=dict(size=8, color=TREND_LINE_COLOR),
        fill="tozeroy", fillcolor="rgba(246, 169, 169, 0.12)",
        hovertemplate="<b>%{x}</b><br>%{y:,.0f} THB<extra></extra>",
    ))
    fig.update_layout(title=title, **CHART_LAYOUT_DEFAULTS)
    fig.update_xaxes(tickangle=-45, showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="#EEEEEE")
    return fig


def make_donut(labels, values, title):
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.55,
        marker=dict(colors=[WEBSITE_COLORS.get(w, "#CCCCCC") for w in labels], line=dict(color="#FFFFFF", width=2)),
        textinfo="percent",
        textfont=dict(size=12, color="#444444"),
        hovertemplate="<b>%{label}</b><br>%{value:,.0f} THB<br>%{percent}<extra></extra>",
    ))
    fig.update_layout(
        title=title,
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=11)),
        **{**CHART_LAYOUT_DEFAULTS, "margin": dict(t=55, b=20, l=20, r=140)},
    )
    return fig


# ============================================
# UI
# ============================================

st.title("💰 Stock Photo Income Tracker")

tab_dashboard, tab_log, tab_import = st.tabs(["📊 Dashboard", "📝 Log Income", "📂 Import Legacy Data"])

# ---- Tab: Log Income ----
with tab_log:
    st.subheader("Log This Month's Income")

    col1, col2 = st.columns(2)
    with col1:
        year = st.number_input("Year", min_value=2000, max_value=2100, value=datetime.now().year, step=1)
        website = st.selectbox("Website", WEBSITES)
    with col2:
        month = st.selectbox("Month", list(range(1, 13)), index=datetime.now().month - 1)
        currency = WEBSITE_CURRENCY.get(website, "USD")
        st.text_input("Currency (auto)", value=currency, disabled=True)

    amount = st.number_input("Amount", min_value=0.0, step=0.01, format="%.2f")

    if st.button("💾 Save Entry", type="primary"):
        if amount <= 0:
            st.warning("Please enter an amount greater than 0.")
        else:
            try:
                save_entry(year, month, website, currency, amount)
                st.success(f"Saved: {year}/{month} — {website} — {amount} {currency}")
                st.cache_data.clear()
            except Exception as e:
                st.error(f"Failed to save: {e}")

# ---- Tab: Import Legacy Data ----
with tab_import:
    st.subheader("Import Existing Excel File (one-time)")
    st.caption("Upload your existing income Excel file (Year, Month, and website columns). It will be converted and added to the Google Sheet automatically.")

    if already_has_imported_data():
        st.info("It looks like legacy data was already imported before. Importing again may create duplicates.")

    uploaded = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    if uploaded and st.button("📥 Import Now"):
        with st.spinner("Importing..."):
            count, info = import_legacy_excel(uploaded)
        if count is None:
            st.error(info)
        else:
            st.success(f"Imported {count} entries successfully.")
            st.write(f"Website columns found: {', '.join(info)}")
            st.cache_data.clear()

# ---- Tab: Dashboard ----
with tab_dashboard:
    st.subheader("📊 Income Dashboard")

    df = load_records()

    if df.empty:
        st.info("No income data yet — import your legacy file or log income first.")
    else:
        usd_today, eur_today, fetched_ok = fetch_rates()

        colf1, colf2 = st.columns(2)
        with colf1:
            years_available = sorted(df["Year"].unique(), reverse=True)
            year_choice = st.selectbox("Year", ["All Years"] + [str(y) for y in years_available])
        with colf2:
            website_choice = st.selectbox("Website", ["All Websites"] + WEBSITES)

        if fetched_ok:
            st.caption(f"💱 THB values are locked in at the rate on the day each entry was logged. New entries today use 1 USD = {usd_today:.2f} THB, 1 EUR = {eur_today:.2f} THB (source: frankfurter.dev)")
        else:
            st.caption("💱 THB values are locked in at the rate on the day each entry was logged.")

        filtered = df if website_choice == "All Websites" else df[df["Website"] == website_choice]

        if filtered.empty:
            st.warning(f"No data found for {website_choice}.")
        else:
            filtered = filtered.copy()
            filtered["THB"] = filtered["Amount_THB"]

            show_all_years = year_choice == "All Years"

            if show_all_years:
                scope = filtered
                period_totals = scope.groupby("Year")["THB"].sum().sort_index()
                total_income = period_totals.sum()
                avg_per_month = total_income / (len(period_totals) * 12) if len(period_totals) else 0
            else:
                selected_year = int(year_choice)
                scope = filtered[filtered["Year"] == selected_year]
                period_totals = scope.groupby("Month")["THB"].sum().reindex(range(1, 13), fill_value=0)
                total_income = period_totals.sum()
                months_with_data = (period_totals > 0).sum()
                avg_per_month = total_income / months_with_data if months_with_data else 0

            website_totals = scope.groupby("Website")["THB"].sum().sort_values(ascending=False)
            top_website = website_totals.idxmax() if not website_totals.empty else "-"

            row1_c1, row1_c2 = st.columns(2)
            with row1_c1:
                metric_card("Total Income (THB)", f"{total_income:,.0f}")
            with row1_c2:
                metric_card("Average per Month (THB)", f"{avg_per_month:,.0f}")

            if website_choice == "All Websites" and not website_totals.empty:
                ranked = website_totals.sort_values(ascending=False)
                top_group = ranked.iloc[:4]
                rest_group = ranked.iloc[4:8]

                if len(top_group) > 0:
                    row2_cols = st.columns(len(top_group))
                    for col, (site, val) in zip(row2_cols, top_group.items()):
                        with col:
                            metric_card(site, f"{val:,.0f}")

                if len(rest_group) > 0:
                    row3_cols = st.columns(len(rest_group))
                    for col, (site, val) in zip(row3_cols, rest_group.items()):
                        with col:
                            metric_card(site, f"{val:,.0f}")

            st.divider()

            row1_col1, row1_col2 = st.columns(2)
            row2_col1, row2_col2 = st.columns(2)

            # Panel 1: main stacked bar
            if show_all_years:
                pivot = filtered.pivot_table(index="Year", columns="Website", values="THB", aggfunc="sum", fill_value=0)
                pivot = pivot.reindex(sorted(pivot.index))
                cats = pivot.index.astype(str)
                fig1 = make_stacked_bar(cats, pivot, "Total Income by Year (THB)")
                trend_x, trend_y, trend_title = cats, period_totals.values, "Yearly Trend (THB)"
            else:
                pivot = scope.pivot_table(index="Month", columns="Website", values="THB", aggfunc="sum", fill_value=0)
                pivot = pivot.reindex(range(1, 13), fill_value=0)
                fig1 = make_stacked_bar(MONTH_NAMES, pivot, f"Monthly Income {selected_year} (THB)")
                cumulative = period_totals.cumsum()
                trend_x, trend_y, trend_title = MONTH_NAMES, cumulative.values, f"Cumulative Income {selected_year} (THB)"

            with row1_col1:
                st.plotly_chart(fig1, use_container_width=True)

            # Panel 2: trend line
            fig2 = make_line(trend_x, trend_y, trend_title)
            with row1_col2:
                st.plotly_chart(fig2, use_container_width=True)

            # Panel 3: latest year monthly (always pinned)
            latest_year = int(filtered["Year"].max())
            latest_scope = filtered[filtered["Year"] == latest_year]
            latest_pivot = latest_scope.pivot_table(index="Month", columns="Website", values="THB", aggfunc="sum", fill_value=0)
            latest_pivot = latest_pivot.reindex(range(1, 13), fill_value=0)
            fig3 = make_stacked_bar(MONTH_NAMES, latest_pivot, f"Latest Year Monthly ({latest_year}, THB)")
            with row2_col1:
                st.plotly_chart(fig3, use_container_width=True)

            # Panel 4: website share donut
            scope_label = "All-Time" if show_all_years else year_choice
            if not website_totals.empty:
                fig4 = make_donut(website_totals.index, website_totals.values, f"Income Share by Website ({scope_label})")
                with row2_col2:
                    st.plotly_chart(fig4, use_container_width=True)

            # Panel 5 & 6: average income per month, broken down by year (always all-years, pinned)
            st.divider()
            months_active_per_year = filtered.groupby("Year")["Month"].nunique()
            total_per_year = filtered.groupby("Year")["THB"].sum()
            avg_per_year = (total_per_year / months_active_per_year).sort_index()
            avg_years_x = avg_per_year.index.astype(str)

            fig5 = go.Figure(go.Bar(
                x=avg_years_x, y=avg_per_year.values,
                marker_color="#B8E0D2",
                hovertemplate="<b>%{x}</b><br>Avg: %{y:,.0f} THB/month<extra></extra>",
            ))
            fig5.update_layout(title="Average Income per Month, by Year (THB)", **CHART_LAYOUT_DEFAULTS)
            fig5.update_xaxes(tickangle=-45, showgrid=False)
            fig5.update_yaxes(showgrid=True, gridcolor="#EEEEEE")

            fig6 = make_line(avg_years_x, avg_per_year.values, "Average Income per Month, by Year — Trend (THB)")

            row3_col1, row3_col2 = st.columns(2)
            with row3_col1:
                st.plotly_chart(fig5, use_container_width=True)
            with row3_col2:
                st.plotly_chart(fig6, use_container_width=True)
