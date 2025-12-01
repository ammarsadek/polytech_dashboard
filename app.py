import streamlit as st
import pandas as pd
import numpy as np
import os
import openpyxl

# -----------------------------
#  Helper: Load & Prepare Data
# -----------------------------
@st.cache_data
def load_data(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)

    # Strip spaces from column names
    df.columns = df.columns.str.strip()

    # Ensure Date is datetime
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'])

    # Standardize machine & product text a bit
    if 'Machine' in df.columns:
        df['Machine'] = df['Machine'].astype(str).str.strip()

    if 'Product' in df.columns:
        df['Product'] = df['Product'].astype(str).str.strip()

    # Calculate Good Units, Quality, Performance, Availability, OEE
    # Assumed columns (from your file):
    # 'Production per unit', 'Reject per unit', 'Reject %', 'Performance %',
    # 'Working days', 'downtime'

    # Good units
    df['Good units'] = df['Production per unit'] - df['Reject per unit']

    # Quality = good / total
    df['Quality'] = np.where(
        df['Production per unit'] > 0,
        df['Good units'] / df['Production per unit'],
        np.nan
    )

    # Performance is already a ratio (e.g. 0.89)
    df['Performance'] = df['Performance %']

    # Availability: we’ll compute later based on hours/day slider
    # For now, just store planned days
    df['Planned days'] = df['Working days']

    return df


def add_oee_metrics(df: pd.DataFrame, hours_per_day: float) -> pd.DataFrame:
    df = df.copy()

    # Planned hours = working days * hours_per_day
    df['Planned hours'] = df['Planned days'] * hours_per_day

    # Availability = 1 - (downtime / planned hours)
    df['Availability'] = np.where(
        df['Planned hours'] > 0,
        1 - (df['downtime'] / df['Planned hours']),
        np.nan
    )

    # Cap between 0 and 1
    df['Availability'] = df['Availability'].clip(lower=0, upper=1)

    # OEE = A * P * Q
    df['OEE'] = df['Availability'] * df['Performance'] * df['Quality']

    return df


# -----------------------------
#  Helper: Aggregations
# -----------------------------
def aggregate_oee(df: pd.DataFrame, group_cols):
    agg = (
        df.groupby(group_cols)
        .agg(
            Production_units=('Production per unit', 'sum'),
            Good_units=('Good units', 'sum'),
            Reject_units=('Reject per unit', 'sum'),
            Downtime=('downtime', 'sum'),
            Planned_hours=('Planned hours', 'sum'),
        )
        .reset_index()
    )

    agg['Quality'] = np.where(
        agg['Production_units'] > 0,
        agg['Good_units'] / agg['Production_units'],
        np.nan,
    )
    # Weighted averages for A and P using hours and production as weights
    # to avoid simple mean bias
    # Availability
    agg['Availability'] = np.where(
        agg['Planned_hours'] > 0,
        1 - (agg['Downtime'] / agg['Planned_hours']),
        np.nan,
    )
    agg['Availability'] = agg['Availability'].clip(0, 1)

    # Performance: approximate using total production vs total target?
    # If you want to use your original 'Performance %', you can instead
    # compute a weighted mean by production. For now, keep it simple:
    # we will recompute Performance as Good_units / Production_units * some factor
    # BUT better: take mean from original df:
    tmp = (
        df.groupby(group_cols)
        .agg(Performance=('Performance', 'mean'))
        .reset_index()
    )
    agg = agg.merge(tmp, on=group_cols, how='left')

    agg['OEE'] = agg['Availability'] * agg['Performance'] * agg['Quality']
    return agg


def format_pct(x):
    if pd.isna(x):
        return "-"
    return f"{x*100:.1f}%"


# -----------------------------
#  Streamlit App
# -----------------------------
def main():
    st.set_page_config(
        page_title="Factory OEE & Production Dashboard",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.title("Polytech Production Dashboard")

    st.markdown(
        """
        This dashboard analyzes production data, KPIs and **OEE** for all machines and molds.  
        Use the filters on the left to drill down by date, machine, and product.
        """
    )

    # -------------------------
    # Sidebar: File & Settings
    # -------------------------
    st.sidebar.header("📂 Data & Settings")

    default_file = "Production data.xlsx"
    uploaded_file = st.sidebar.file_uploader(
        "Upload production Excel file",
        type=["xlsx", "xls"],
        help="If empty, the app will try to load 'Production data.xlsx' from the working directory.",
    )

    if uploaded_file is not None:
        df_raw = load_data(uploaded_file)
    else:
        if os.path.exists(default_file):
            df_raw = load_data(default_file)
        else:
            st.error(
                "No file uploaded and 'Production data.xlsx' not found in working directory."
            )
            st.stop()

    hours_per_day = st.sidebar.number_input(
        "Planned hours per working day",
        min_value=1.0,
        max_value=24.0,
        value=24.0,
        step=1.0,
        help="Used to calculate Availability = 1 - (downtime / planned hours).",
    )

    df = add_oee_metrics(df_raw, hours_per_day)

    # -------------------------
    # Sidebar: Filters
    # -------------------------
    st.sidebar.header("🔍 Filters")

    # Date filter
    if "Date" in df.columns:
        min_date = df["Date"].min()
        max_date = df["Date"].max()
        date_range = st.sidebar.date_input(
            "Date range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )

        if isinstance(date_range, tuple) or isinstance(date_range, list):
            start_date, end_date = date_range
        else:
            start_date = date_range
            end_date = date_range

        df = df[(df["Date"] >= pd.to_datetime(start_date)) & (df["Date"] <= pd.to_datetime(end_date))]

    # Machine filter
    machines = sorted(df["Machine"].dropna().unique())
    selected_machines = st.sidebar.multiselect(
        "Machines",
        options=machines,
        default=machines,
    )
    if selected_machines:
        df = df[df["Machine"].isin(selected_machines)]

    # Product filter
    products = sorted(df["Product"].dropna().unique())
    selected_products = st.sidebar.multiselect(
        "Products / Molds",
        options=products,
        default=products,
    )
    if selected_products:
        df = df[df["Product"].isin(selected_products)]

    if df.empty:
        st.warning("No data after applying filters. Adjust filters and try again.")
        st.stop()

    # Add Month-Year for trends
    if "Date" in df.columns:
        df["MonthYear"] = df["Date"].dt.to_period("M").astype(str)

    # -------------------------
    # Top-Level KPIs
    # -------------------------
    st.subheader("📊 Overall KPIs (Filtered Data)")

    total_prod = df["Production per unit"].sum()
    total_good = df["Good units"].sum()
    total_reject = df["Reject per unit"].sum()
    total_downtime = df["downtime"].sum()
    overall_quality = total_good / total_prod if total_prod > 0 else np.nan

    # Weighted availability & performance
    total_planned_hours = df["Planned hours"].sum()
    overall_availability = (
        1 - total_downtime / total_planned_hours if total_planned_hours > 0 else np.nan
    )
    overall_availability = (
        np.clip(overall_availability, 0, 1) if not pd.isna(overall_availability) else np.nan
    )

    # Performance: mean weighted by production
    if total_prod > 0:
        overall_performance = np.average(df["Performance"], weights=df["Production per unit"])
    else:
        overall_performance = np.nan

    overall_oee = (
        overall_availability * overall_performance * overall_quality
        if not any(pd.isna([overall_availability, overall_performance, overall_quality]))
        else np.nan
    )

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.metric("OEE", format_pct(overall_oee))
    with col2:
        st.metric("Availability", format_pct(overall_availability))
    with col3:
        st.metric("Performance", format_pct(overall_performance))
    with col4:
        st.metric("Quality", format_pct(overall_quality))
    with col5:
        st.metric("Total Downtime (hrs)", f"{total_downtime:,.1f}")

    st.markdown("---")

    # -------------------------
    # Tabs for Analysis
    # -------------------------
    tab_overview, tab_machines, tab_products, tab_trends, tab_data = st.tabs(
        ["🔭 Overview", "🛠 Machines", "🧩 Products/Molds", "📈 Trends", "📄 Raw Data"]
    )

    # ---------- Overview Tab ----------
    with tab_overview:
        st.subheader("OEE by Machine")

        machine_agg = aggregate_oee(df, ["Machine"])
        machine_agg_sorted = machine_agg.sort_values("OEE", ascending=False)

        # Top machines table
        st.dataframe(
            machine_agg_sorted.assign(
                OEE_pct=lambda d: d["OEE"].apply(format_pct),
                Availability_pct=lambda d: d["Availability"].apply(format_pct),
                Performance_pct=lambda d: d["Performance"].apply(format_pct),
                Quality_pct=lambda d: d["Quality"].apply(format_pct),
            )[["Machine", "OEE_pct", "Availability_pct", "Performance_pct", "Quality_pct",
               "Production_units", "Downtime"]],
            use_container_width=True,
        )

        # Simple bar chart using Streamlit
        st.bar_chart(
            machine_agg_sorted.set_index("Machine")["OEE"],
            use_container_width=True,
        )

        st.subheader("OEE by Product / Mold")

        product_agg = aggregate_oee(df, ["Product"])
        product_agg_sorted = product_agg.sort_values("OEE", ascending=False)

        st.dataframe(
            product_agg_sorted.assign(
                OEE_pct=lambda d: d["OEE"].apply(format_pct),
                Availability_pct=lambda d: d["Availability"].apply(format_pct),
                Performance_pct=lambda d: d["Performance"].apply(format_pct),
                Quality_pct=lambda d: d["Quality"].apply(format_pct),
            )[["Product", "OEE_pct", "Availability_pct", "Performance_pct", "Quality_pct",
               "Production_units", "Downtime"]],
            use_container_width=True,
        )

    # ---------- Machines Tab ----------
    with tab_machines:
        st.subheader("Machine Drill-down")

        machine_list = sorted(df["Machine"].unique())
        selected_machine = st.selectbox("Select machine", machine_list)

        df_m = df[df["Machine"] == selected_machine].copy()
        if df_m.empty:
            st.info("No data for selected machine.")
        else:
            # KPI for this machine
            m_agg = aggregate_oee(df_m, ["Machine"]).iloc[0]

            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("Machine OEE", format_pct(m_agg["OEE"]))
            with col2:
                st.metric("Availability", format_pct(m_agg["Availability"]))
            with col3:
                st.metric("Performance", format_pct(m_agg["Performance"]))
            with col4:
                st.metric("Quality", format_pct(m_agg["Quality"]))
            with col5:
                st.metric("Prod units", f"{m_agg['Production_units']:,.0f}")

            if "Date" in df_m.columns:
                # Sort by date
                df_m = df_m.sort_values("Date")

                st.line_chart(
                    df_m.set_index("Date")[["OEE", "Availability", "Performance", "Quality"]],
                    use_container_width=True,
                )

            st.markdown("#### Detailed records")
            st.dataframe(df_m, use_container_width=True)

    # ---------- Products Tab ----------
    with tab_products:
        st.subheader("Product / Mold Drill-down")

        product_list = sorted(df["Product"].unique())
        selected_product = st.selectbox("Select product / mold", product_list)

        df_p = df[df["Product"] == selected_product].copy()
        if df_p.empty:
            st.info("No data for selected product.")
        else:
            p_agg = aggregate_oee(df_p, ["Product"]).iloc[0]

            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("Product OEE", format_pct(p_agg["OEE"]))
            with col2:
                st.metric("Availability", format_pct(p_agg["Availability"]))
            with col3:
                st.metric("Performance", format_pct(p_agg["Performance"]))
            with col4:
                st.metric("Quality", format_pct(p_agg["Quality"]))
            with col5:
                st.metric("Prod units", f"{p_agg['Production_units']:,.0f}")

            if "Date" in df_p.columns:
                df_p = df_p.sort_values("Date")
                st.line_chart(
                    df_p.set_index("Date")[["OEE", "Availability", "Performance", "Quality"]],
                    use_container_width=True,
                )

            st.markdown("#### Detailed records")
            st.dataframe(df_p, use_container_width=True)

    # ---------- Trends Tab ----------
    with tab_trends:
        st.subheader("OEE Trend Over Time")

        if "Date" in df.columns:
            # Daily trend
            daily_agg = aggregate_oee(df, ["Date"]).sort_values("Date")
            st.markdown("**Daily OEE**")
            st.line_chart(
                daily_agg.set_index("Date")[["OEE", "Availability", "Performance", "Quality"]],
                use_container_width=True,
            )

        if "MonthYear" in df.columns:
            monthly_agg = aggregate_oee(df, ["MonthYear"]).sort_values("MonthYear")
            st.markdown("**Monthly OEE**")
            st.line_chart(
                monthly_agg.set_index("MonthYear")[["OEE", "Availability", "Performance", "Quality"]],
                use_container_width=True,
            )

    # ---------- Raw Data Tab ----------
    with tab_data:
        st.subheader("Raw Data (with Calculated Metrics)")
        st.dataframe(df, use_container_width=True)


if __name__ == "__main__":
    main()
