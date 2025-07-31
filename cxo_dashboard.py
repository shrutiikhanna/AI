import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm

# Inject Nunito font via custom CSS
st.markdown("""
<link href="https://fonts.googleapis.com/css?family=Nunito:400,700&display=swap" rel="stylesheet">
<style>
body, div, p, span, label, h1, h2, h3, h4, h5, h6,
.stTextInput, .stDataFrame, .stMarkdown, .stButton, .stSelectbox, .stSubheader,
.stRadio, .stSlider, .stNumberInput, .stFileUploader, .stDateInput, .stTimeInput,
.stTable, .stAlert, .stForm, .stFormSubmitButton, .stMetric, .stSidebar, .stTabs,
.stPlotlyChart, .stPyplot, .stImage, .stTextArea, .stCheckbox, .stColorPicker,
.stExpander, .stJson, .stCode, .stDownloadButton, .stMultiSelect, .stDataEditor {
    font-family: 'Nunito', sans-serif !important;
    font-size: 12px !important;
}
h1, h2, h3, h4, h5, h6 {
    font-size: 14px !important;
}
.stSubheader {
    font-size: 12px !important;
}
</style>
""", unsafe_allow_html=True)

# CXO Dashboard Template
st.set_page_config(page_title="CXO AI Dashboard", page_icon="📈", layout="wide")
st.title("CXO Dashboard")

st.markdown("""
Welcome to the **CXO Dashboard**. This dashboard provides a high-level overview of key business metrics for executive decision-making. Upload your Excel file to get started, or use the sample data provided for demonstration.
""")


# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file (Excel format: .xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # Read all sheets
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    # Use actual sheet names for tabs
    tab_names = sheet_names
    tabs = st.tabs(tab_names)

    # Profitability Tab (if exists)
    if "Profitability" in sheet_names:
        with tabs[sheet_names.index("Profitability")]:
            df = pd.read_excel(xls, sheet_name="Profitability")
            st.write("Columns found in your file:", list(df.columns))
            st.subheader("Data Preview")
            st.dataframe(df)

            # Sidebar menu with new options
            menu = st.sidebar.selectbox(
                "Select Visualization",
                [
                    "KPI Card",
                    "Product Wise",
                    "Zone Wise",
                    "BD Wise",
                    "AM Wise",
                    "Segment Wise",
                    "State Wise"
                ]
            )
            if menu == "KPI Card":
                kpi_options = [
                    "Number of Transaction",
                    "GMV",
                    "Gross Revenue",
                    "Bank & PG Charges",
                    "Referral Charges",
                    "Net Earnings"
                ]
                kpi_values = {
                    "Number of Transaction": 315432,
                    "GMV": 2306914235,
                    "Gross Revenue": 4081767,
                    "Bank & PG Charges": 755049,
                    "Referral Charges": 1255277,
                    "Net Earnings": 2071441
                }
                st.subheader("KPI Cards")
                # Display KPIs in a grid with border and comma formatting
                card_style = """
                    display: flex;
                    flex-wrap: wrap;
                    gap: 16px;
                """
                st.markdown(f"<div style='{card_style}'>", unsafe_allow_html=True)
                for kpi in kpi_options:
                    value = kpi_values[kpi]
                    card_html = f"""
                        <div style='border:2px solid #1f77b4; border-radius:8px; padding:16px; min-width:180px; margin-bottom:8px; background:#f9f9f9;'>
                            <div style='font-size:16px; font-weight:bold;'>{kpi}</div>
                            <div style='font-size:20px; color:#1f77b4; font-weight:bold;'>{value:,}</div>
                        </div>
                    """
                    st.markdown(card_html, unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
            elif menu == "Product Wise":
                product_col = None
                for col in df.columns:
                    if "Product" in col:
                        product_col = col
                        break
                if product_col:
                    grouped = df.groupby(product_col).agg({
                        "Number of Transaction": "sum",
                        "GMV": "sum",
                        "Gross Revenue": "sum",
                        "Net Earnings": "sum"
                    }).reset_index()
                    # Format numeric columns with commas
                    for col in ["Number of Transaction", "GMV", "Gross Revenue", "Net Earnings"]:
                        if col in grouped.columns:
                            grouped[col] = grouped[col].apply(lambda x: f"{int(x):,}")
                    st.dataframe(grouped)
                    # Pie Charts for Product Wise KPIs with legend and color
                    kpi_list = ["Number of Transaction", "GMV", "Gross Revenue", "Net Earnings"]
                    colors = plt.cm.tab20.colors
                    for idx, kpi in enumerate(kpi_list):
                        if kpi in grouped.columns:
                            # Convert formatted string numbers back to int for plotting
                            values = grouped[kpi].apply(lambda x: int(str(x).replace(",", ""))).values
                            fig, ax = plt.subplots(figsize=(2, 2))
                            wedges, texts = ax.pie(
                                values,
                                labels=None,
                                startangle=90,
                                colors=colors[:len(grouped[product_col])],
                            )
                            ax.set_title(f"Product wise {kpi} (Pie Chart)")
                            total = sum(values)
                            legend_labels = [f"{label}: {value/total*100:.1f}%" for label, value in zip(grouped[product_col], values)]
                            ax.legend(wedges, legend_labels, title="Products", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                            st.pyplot(fig)
                else:
                    st.warning("No 'Product' column found in your data.")
            elif menu == "Zone Wise":
                zone_col = None
                for col in df.columns:
                    if "Zone" in col:
                        zone_col = col
                        break
                if zone_col:
                    grouped = df.groupby(zone_col).agg({
                        "Number of Transaction": "sum",
                        "GMV": "sum",
                        "Gross Revenue": "sum",
                        "Net Earnings": "sum"
                    }).reset_index()
                    # Format numeric columns with commas
                    for col in ["Number of Transaction", "GMV", "Gross Revenue", "Net Earnings"]:
                        if col in grouped.columns:
                            grouped[col] = grouped[col].apply(lambda x: f"{int(x):,}")
                    st.dataframe(grouped)
                    # Pie Charts for Zone Wise KPIs with legend and color
                    colors = plt.cm.tab20.colors
                    kpi_list = ["Number of Transaction", "GMV", "Gross Revenue", "Net Earnings"]
                    for idx, kpi in enumerate(kpi_list):
                        if kpi in grouped.columns:
                            values = grouped[kpi].apply(lambda x: int(str(x).replace(",", ""))).values
                            fig, ax = plt.subplots(figsize=(2, 3))
                            wedges, texts = ax.pie(
                                values,
                                labels=None,
                                startangle=90,
                                colors=colors[:len(grouped[zone_col])],
                            )
                            ax.set_title(f"Zone wise {kpi} (Pie Chart)")
                            total = sum(values)
                            legend_labels = [f"{label}: {value/total*100:.1f}%" for label, value in zip(grouped[zone_col], values)]
                            ax.legend(wedges, legend_labels, title="Zones", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                            st.pyplot(fig)
                else:
                    st.warning("No 'Zone' column found in your data.")
            elif menu == "BD Wise":
                bd_col = None
                for col in df.columns:
                    if "BD" in col:
                        bd_col = col
                        break
                if bd_col:
                    grouped = df.groupby(bd_col).agg({
                        "Number of Transaction": "sum",
                        "GMV": "sum",
                        "Gross Revenue": "sum",
                        "Net Earnings": "sum"
                    }).reset_index()
                    # Format numeric columns with commas
                    for col in ["Number of Transaction", "GMV", "Gross Revenue", "Net Earnings"]:
                        if col in grouped.columns:
                            grouped[col] = grouped[col].apply(lambda x: f"{int(x):,}")
                    st.dataframe(grouped)
                    # Simple Bar Charts for BD Wise KPIs (Y axis in Lacs)
                    kpi_list = ["Number of Transaction", "GMV", "Gross Revenue", "Net Earnings"]
                    for kpi in kpi_list:
                        if kpi in grouped.columns:
                            x_labels = grouped[bd_col].astype(str)
                            y_values = grouped[kpi].apply(lambda x: int(str(x).replace(",", "")) / 100000)
                            fig, ax = plt.subplots(figsize=(4, 3))
                            ax.bar(x_labels, y_values, color='#1f77b4')
                            ax.set_title(f"BD wise {kpi} (in Lacs)")
                            ax.set_xlabel(bd_col)
                            ax.set_ylabel(f"{kpi} (Lacs)")
                            plt.xticks(rotation=45, ha='right')
                            st.pyplot(fig)
                else:
                    st.warning("No 'BD' column found in your data.")
            elif menu == "AM Wise":
                am_col = None
                for col in df.columns:
                    if "AM" in col:
                        am_col = col
                        break
                if am_col:
                    grouped = df.groupby(am_col).agg({
                        "Number of Transaction": "sum",
                        "GMV": "sum",
                        "Gross Revenue": "sum",
                        "Net Earnings": "sum"
                    }).reset_index()
                    # Format numeric columns with commas
                    for col in ["Number of Transaction", "GMV", "Gross Revenue", "Net Earnings"]:
                        if col in grouped.columns:
                            grouped[col] = grouped[col].apply(lambda x: f"{int(x):,}")
                    st.dataframe(grouped)
                    # Simple Bar Charts for AM Wise KPIs (Y axis in Lacs)
                    kpi_list = ["Number of Transaction", "GMV", "Gross Revenue", "Net Earnings"]
                    for kpi in kpi_list:
                        if kpi in grouped.columns:
                            x_labels = grouped[am_col].astype(str)
                            y_values = grouped[kpi].apply(lambda x: int(str(x).replace(",", "")) / 100000)
                            fig, ax = plt.subplots()
                            ax.bar(x_labels, y_values, color='#ff7f0e')
                            ax.set_title(f"AM wise {kpi} (in Lacs)")
                            ax.set_xlabel(am_col)
                            ax.set_ylabel(f"{kpi} (Lacs)")
                            plt.xticks(rotation=45, ha='right')
                            st.pyplot(fig)
                else:
                    st.warning("No 'AM' column found in your data.")
            elif menu == "Segment Wise":
                segment_col = None
                for col in df.columns:
                    if "Segment" in col:
                        segment_col = col
                        break
                if segment_col:
                    grouped = df.groupby(segment_col).agg({
                        "GMV": "sum",
                        "Gross Revenue": "sum",
                        "Net Earnings": "sum"
                    }).reset_index()
                    # Format numeric columns with commas
                    for col in ["GMV", "Gross Revenue", "Net Earnings"]:
                        if col in grouped.columns:
                            grouped[col] = grouped[col].apply(lambda x: f"{int(x):,}")
                    st.dataframe(grouped)
                    # Pie Charts for Segment Wise KPIs with legend and color
                    colors = plt.cm.tab20.colors
                    kpi_list = ["GMV", "Gross Revenue", "Net Earnings"]
                    plt.rcParams['font.family'] = 'Nunito'
                    for idx, kpi in enumerate(kpi_list):
                        if kpi in grouped.columns:
                            # Convert formatted string numbers back to int for plotting
                            values = grouped[kpi].apply(lambda x: int(str(x).replace(",", ""))).values
                            fig, ax = plt.subplots()
                            wedges, texts = ax.pie(
                                values,
                                labels=None,
                                startangle=90,
                                colors=colors[:len(grouped[segment_col])],
                            )
                            ax.set_title(f"Segment wise {kpi} (Pie Chart)")
                            total = sum(values)
                            legend_labels = [f"{label}: {value/total*100:.1f}%" for label, value in zip(grouped[segment_col], values)]
                            ax.legend(wedges, legend_labels, title="Segments", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                            st.pyplot(fig)
                else:
                    st.warning("No 'Segment' column found in your data.")
            elif menu == "State Wise":
                state_col = None
                for col in df.columns:
                    if "State" in col:
                        state_col = col
                        break
                if state_col:
                    grouped = df.groupby(state_col).agg({
                        "GMV": "sum",
                        "Gross Revenue": "sum",
                        "Net Earnings": "sum"
                    }).reset_index()
                    # Format numeric columns with commas
                    for col in ["GMV", "Gross Revenue", "Net Earnings"]:
                        if col in grouped.columns:
                            grouped[col] = grouped[col].apply(lambda x: f"{int(x):,}")
                    st.dataframe(grouped)
                    # Bar Charts for State Wise KPIs (Y axis in Lacs)
                    kpi_list = ["GMV", "Gross Revenue", "Net Earnings"]
                    colors = ['#1f77b4', '#ff7f0e', '#2ca02c']
                    for idx, kpi in enumerate(kpi_list):
                        if kpi in grouped.columns:
                            x_labels = grouped[state_col].astype(str)
                            y_values = grouped[kpi].apply(lambda x: int(str(x).replace(",", "")) / 100000)
                            fig, ax = plt.subplots(figsize=(6, 3))
                            ax.bar(x_labels, y_values, color=colors[idx % len(colors)])
                            ax.set_title(f"State wise {kpi} (in Lacs)")
                            ax.set_xlabel(state_col)
                            ax.set_ylabel(f"{kpi} (Lacs)")
                            plt.xticks(rotation=45, ha='right')
                            st.pyplot(fig)
                else:
                    st.warning("No 'State' column found in your data.")

    # Dashboard Summary Tab (if exists)
    if "Dashboard Summary" in sheet_names:
        with tabs[sheet_names.index("Dashboard Summary")]:
            st.subheader("Dashboard Summary Data Table")
            # Try reading with default header
            df_dash = pd.read_excel(xls, sheet_name="Dashboard Summary")
            # If empty or only one column, try reading with no header
            if df_dash.empty or df_dash.shape[1] <= 1:
                df_dash_raw = pd.read_excel(xls, sheet_name="Dashboard Summary", header=None)
                # If still empty, show warning
                if df_dash_raw.empty or df_dash_raw.shape[1] == 0:
                    st.warning("No data found in Dashboard Summary sheet.")
                else:
                    # Format numeric columns with commas
                    df_dash_fmt = df_dash_raw.copy()
                    for col in df_dash_fmt.columns:
                        if pd.api.types.is_numeric_dtype(df_dash_fmt[col]):
                            df_dash_fmt[col] = df_dash_fmt[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else x)
                    st.dataframe(df_dash_fmt)
            else:
                # Format numeric columns with commas
                df_dash_fmt = df_dash.copy()
                for col in df_dash_fmt.columns:
                    if pd.api.types.is_numeric_dtype(df_dash_fmt[col]):
                        df_dash_fmt[col] = df_dash_fmt[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else x)
                st.dataframe(df_dash_fmt)

            # Use whichever dataframe has data for charting
            chart_df = df_dash if not df_dash.empty and df_dash.shape[1] > 1 else (df_dash_raw if 'df_dash_raw' in locals() else None)
            if chart_df is not None and not chart_df.empty and chart_df.shape[1] > 0:
                # Handle case where months are column headers and metrics are rows
                metrics = [
                    "Revenue from Operations (A+B+C)",
                    "Direct Expenses",
                    "Indirect Expenses",
                    "EBITDA"
                ]
                # Detect if first column is 'Particulars' and months are columns
                if 'Particulars' in chart_df.columns:
                    month_cols = [col for col in chart_df.columns if col != 'Particulars']
                    available_metrics = [m for m in metrics if m in chart_df['Particulars'].values]
                    if not month_cols:
                        st.warning("No month columns found for line chart.")
                    elif not available_metrics:
                        st.warning("None of the required metrics found for line chart: Revenue from Operations (A+B+C), Direct Expenses, Indirect Expenses, EBITDA.")
                    else:
                        st.subheader("Month-wise KPI Bar Graphs")
                        # Build a dataframe: rows=months, columns=metrics
                        df_bar = chart_df.set_index('Particulars').loc[available_metrics]
                        # Convert all values to numeric and to Lacs
                        df_bar = df_bar.applymap(lambda x: pd.to_numeric(str(x).replace(",", ""), errors='coerce') / 100000 if pd.notnull(x) else x)
                        df_bar = df_bar.T  # Transpose: now index=months, columns=metrics
                        if df_bar.empty:
                            st.warning("No data available to plot month-wise KPI bar graphs.")
                        else:
                            for m in available_metrics:
                                if m in df_bar.columns:
                                    st.subheader(f"{m} - Month-wise Bar Graph")
                                    fig, ax = plt.subplots(figsize=(7, 3))
                                    ax.bar(df_bar.index.astype(str), df_bar[m], color='#1f77b4')
                                    ax.set_xlabel("Month")
                                    ax.set_ylabel(f"{m} (Lacs)")
                                    ax.set_title(f"Month-wise {m}")
                                    plt.xticks(rotation=45, ha='right')
                                    st.pyplot(fig)
                            # Bar chart for 'Net Worth as on' if present
                            if 'Net Worth as on' in chart_df['Particulars'].values:
                                net_worth_row = chart_df[chart_df['Particulars'] == 'Net Worth as on']
                                # Remove 'Particulars' column for plotting
                                net_worth_vals = net_worth_row.drop('Particulars', axis=1).iloc[0]
                                # Convert to numeric and to Lacs
                                net_worth_vals = net_worth_vals.apply(lambda x: pd.to_numeric(str(x).replace(",", ""), errors='coerce') / 100000 if pd.notnull(x) else x)
                                st.subheader("Net Worth as on - Month-wise Bar Chart")
                                fig, ax = plt.subplots(figsize=(7, 3))
                                ax.bar(net_worth_vals.index.astype(str), net_worth_vals.values, color='#2ca02c')
                                ax.set_xlabel("Month")
                                ax.set_ylabel("Net Worth (Lacs)")
                                ax.set_title("Month-wise Net Worth as on")
                                plt.xticks(rotation=45, ha='right')
                                st.pyplot(fig)
                        # ...existing code...
                else:
                    # Fallback to previous logic if 'Month' column exists
                    month_col = None
                    for col in chart_df.columns:
                        col_clean = str(col).strip().lower().replace(' ', '')
                        if col_clean.startswith('month') or 'month' in col_clean:
                            month_col = col
                            break
                    available_metrics = [m for m in metrics if m in chart_df.columns]
                    if not month_col:
                        st.warning("No month column found for line chart. Please ensure a column named 'Month' exists.")
                    elif not available_metrics:
                        st.warning("None of the required metrics found for line chart: Revenue from Operations (A+B+C), Direct Expenses, Indirect Expenses, EBITDA.")
                    else:
                        st.subheader("Month-wise KPI Bar Graphs")
                        # Convert metrics to numeric and to Lacs if needed
                        for m in available_metrics:
                            chart_df[m] = pd.to_numeric(chart_df[m].astype(str).str.replace(",", ""), errors='coerce') / 100000
                        month_data = chart_df.groupby(month_col)[available_metrics].sum().sort_index()
                        if month_data.empty:
                            st.warning("No data available to plot month-wise KPI bar graphs.")
                        else:
                            for m in available_metrics:
                                st.subheader(f"{m} - Month-wise Bar Graph")
                                fig, ax = plt.subplots(figsize=(7, 3))
                                ax.bar(month_data.index.astype(str), month_data[m], color='#1f77b4')
                                ax.set_xlabel("Month")
                                ax.set_ylabel(f"{m} (Lacs)")
                                ax.set_title(f"Month-wise {m}")
                                plt.xticks(rotation=45, ha='right')
                                st.pyplot(fig)

    # P&L Summary Tab (if exists)
    if "P&L Summary" in sheet_names:
        with tabs[sheet_names.index("P&L Summary")]:
            st.subheader("P&L Summary Data Table")
            try:
                df_pl = pd.read_excel(xls, sheet_name="P&L Summary")
                # Drop rows where all columns are empty or NaN
                df_pl = df_pl.dropna(how='all')
                # Format numeric columns with commas for readability
                df_pl_fmt = df_pl.copy()
                for col in df_pl_fmt.columns:
                    if pd.api.types.is_numeric_dtype(df_pl_fmt[col]):
                        df_pl_fmt[col] = df_pl_fmt[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else x)
                st.dataframe(df_pl_fmt)
            except Exception as e:
                st.warning(f"Could not read P&L Summary sheet: {e}")

if 'Nunito' in [f.name for f in fm.fontManager.ttflist]:
    plt.rcParams['font.family'] = 'Nunito'
else:
    plt.rcParams['font.family'] = 'sans-serif'
