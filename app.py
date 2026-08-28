import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import requests
import matplotlib.pyplot as plt

st.set_page_config(page_title="Income Tracker", page_icon="💰", layout="wide")

# ============================================
WEBSITES = ["Shutterstock", "Adobe", "Getty", "Dreamtime", "123RF", "Deposit", "Freepik", "Colorbox"]
WEBSITE_CURRENCY = {
    "Shutterstock": "USD", "Adobe": "USD", "Getty": "USD", "Dreamtime": "USD",
    "123RF": "USD", "Deposit": "USD", "Freepik": "USD", "Colorbox": "EUR",
}
WEBSITE_COLORS = {
    "Shutterstock": "#1f77b4", "Adobe": "#ff7f0e", "Getty": "#2ca02c", "Dreamtime": "#d62728",
    "123RF": "#9467bd", "Deposit": "#8c564b", "Freepik": "#e377c2", "Colorbox": "#7f7f7f",
}
INCOME_HEADERS = ["Year", "Month", "Website", "Currency", "Amount", "Entry Date", "Source"]
MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

SHEET_NAME = st.secrets.get("INCOME_SHEET_NAME", "PhotoStockIncome")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]


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
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=INCOME_HEADERS)
    df = pd.DataFrame(data)
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df["Month"] = pd.to_numeric(df["Month"], errors="coerce")
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    df = df.dropna(subset=["Year", "Month", "Amount"])
    df["Year"] = df["Year"].astype(int)
    df["Month"] = df["Month"].astype(int)
    return df


def save_entry(year, month, website, currency, amount):
    ws = get_worksheet()
    ws.append_row([
        int(year), int(month), website, currency, float(amount),
        datetime.now().strftime("%Y-%m-%d %H:%M"), "Manual",
    ])


def already_has_imported_data():
    ws = get_worksheet()
    values = ws.get_all_values()
    return any(row[6] == "Imported" for row in values[1:] if len(row) > 6)


def import_legacy_excel(uploaded_file):
    import openpyxl
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
                rows_to_add.append([
                    year, month, site, WEBSITE_CURRENCY[site], amount,
                    datetime.now().strftime("%Y-%m-%d %H:%M"), "Imported",
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
# UI
# ============================================

st.title("💰 Stock Photo Income Tracker")

tab_log, tab_import, tab_dashboard = st.tabs(["📝 Log Income", "📂 Import Legacy Data", "📊 Dashboard"])

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
        usd_default, eur_default, fetched_ok = fetch_rates()

        colf1, colf2, colf3, colf4 = st.columns(4)
        with colf1:
            years_available = sorted(df["Year"].unique(), reverse=True)
            year_choice = st.selectbox("Year", ["All Years"] + [str(y) for y in years_available])
        with colf2:
            website_choice = st.selectbox("Website", ["All Websites"] + WEBSITES)
        with colf3:
            usd_rate = st.number_input("USD → THB", value=float(usd_default), step=0.01, format="%.4f")
        with colf4:
            eur_rate = st.number_input("EUR → THB", value=float(eur_default), step=0.01, format="%.4f")

        if fetched_ok:
            st.caption(f"✅ Rates auto-updated today — source: frankfurter.dev")
        else:
            st.caption("⚠️ Couldn't reach frankfurter.dev — using default rates. Edit the fields above if needed.")

        filtered = df if website_choice == "All Websites" else df[df["Website"] == website_choice]

        if filtered.empty:
            st.warning(f"No data found for {website_choice}.")
        else:
            filtered = filtered.copy()
            filtered["THB"] = filtered.apply(lambda r: to_thb(r["Amount"], r["Currency"], usd_rate, eur_rate), axis=1)

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

            c1, c2, c3 = st.columns(3)
            c1.metric("Total Income (THB)", f"{total_income:,.0f}")
            c2.metric("Average per Month (THB)", f"{avg_per_month:,.0f}")
            c3.metric("Top Website", top_website)

            fig, axes = plt.subplots(2, 2, figsize=(13, 8))

            # Panel 1: main stacked bar
            ax1 = axes[0, 0]
            if show_all_years:
                pivot = filtered.pivot_table(index="Year", columns="Website", values="THB", aggfunc="sum", fill_value=0)
                pivot = pivot.reindex(sorted(pivot.index))
                bottom = pd.Series(0, index=pivot.index, dtype=float)
                for site in WEBSITES:
                    if site in pivot.columns and pivot[site].sum() > 0:
                        ax1.bar(pivot.index.astype(str), pivot[site], bottom=bottom, label=site, color=WEBSITE_COLORS.get(site))
                        bottom += pivot[site]
                ax1.set_title("Total Income by Year (THB)")
                trend_x, trend_y, trend_title = pivot.index.astype(str), period_totals.values, "Yearly Trend (THB)"
            else:
                pivot = scope.pivot_table(index="Month", columns="Website", values="THB", aggfunc="sum", fill_value=0)
                pivot = pivot.reindex(range(1, 13), fill_value=0)
                bottom = pd.Series(0, index=pivot.index, dtype=float)
                for site in WEBSITES:
                    if site in pivot.columns and pivot[site].sum() > 0:
                        ax1.bar(MONTH_NAMES, pivot[site], bottom=bottom, label=site, color=WEBSITE_COLORS.get(site))
                        bottom += pivot[site]
                ax1.set_title(f"Monthly Income {selected_year} (THB)")
                cumulative = period_totals.cumsum()
                trend_x, trend_y, trend_title = MONTH_NAMES, cumulative.values, f"Cumulative Income {selected_year} (THB)"
            ax1.legend(fontsize=7, loc="upper left", ncol=2)
            ax1.tick_params(axis="x", rotation=45)

            # Panel 2: trend line
            ax2 = axes[0, 1]
            ax2.plot(trend_x, trend_y, marker="o", color="#e74c3c")
            ax2.set_title(trend_title)
            ax2.tick_params(axis="x", rotation=45)

            # Panel 3: latest year monthly (always pinned)
            ax3 = axes[1, 0]
            latest_year = int(filtered["Year"].max())
            latest_scope = filtered[filtered["Year"] == latest_year]
            latest_pivot = latest_scope.pivot_table(index="Month", columns="Website", values="THB", aggfunc="sum", fill_value=0)
            latest_pivot = latest_pivot.reindex(range(1, 13), fill_value=0)
            bottom = pd.Series(0, index=latest_pivot.index, dtype=float)
            for site in WEBSITES:
                if site in latest_pivot.columns and latest_pivot[site].sum() > 0:
                    ax3.bar(MONTH_NAMES, latest_pivot[site], bottom=bottom, label=site, color=WEBSITE_COLORS.get(site))
                    bottom += latest_pivot[site]
            ax3.set_title(f"Latest Year Monthly ({latest_year}, THB)")
            ax3.legend(fontsize=7, loc="upper left", ncol=2)
            ax3.tick_params(axis="x", rotation=45)

            # Panel 4: website share pie
            ax4 = axes[1, 1]
            if not website_totals.empty:
                colors = [WEBSITE_COLORS.get(w, "#999999") for w in website_totals.index]
                ax4.pie(website_totals.values, labels=website_totals.index, autopct="%1.0f%%", startangle=90, colors=colors)
            scope_label = "All-Time" if show_all_years else year_choice
            ax4.set_title(f"Income Share by Website ({scope_label})")

            fig.tight_layout()
            st.pyplot(fig)
