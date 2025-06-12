# ============================================
# IMPORTS AND DEPENDENCIES
# ============================================
import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
# import plotly.graph_objects as go # Not explicitly used, can be removed if not needed elsewhere
import calendar
import traceback
from typing import Optional, Tuple, Dict, List, Union

# ============================================
# CONFIGURATION AND CONSTANTS
# ============================================
COLUMN_MAPPING = {
    'Employee': ['Employee Name', 'sheet name'],
    'Team': ['Team'],
    'Month': ['Month'],
    'Week': ['Week'],
    'Total Score': ['Total Score'],
    'Weekly Target (Sales)': [
        'Weekly Target (Sales)',
    ],
    'Weekly Achievement (Sales)': [
        'Weekly Achievement (Sales)',
    ],
    'Sales Achievement %': ['Sales Achievement %'],
    'Sales Score': ['Sales Score'],
    # Sales team
    'Weekly Target (Selling Cost %)': ['Weekly Target (Selling Cost %)'],
    'Weekly Achievement (Selling Cost %)': ['Weekly Achievement (Selling Cost %)'],
    'Selling Cost % Score': ['Selling Cost % Score'],
    # Ads team
    'Weekly Target (ACOS %)': ['Weekly Target (ACOS %)'],
    'Weekly Achievement (ACOS %)': ['Weekly Achievement (ACOS %)'],
    'ACOS % Score': ['ACOS % Score', 'ACOS Cost % Score'],
    # Website Ads team
    'Weekly Target (ROAS %)': ['Weekly Target (ROAS %)'],
    'Weekly Achievement (ROAS %)': ['Weekly Achievement (ROAS %)'],
    'ROAS % Score': ['ROAS % Score'],
    # Portfolio Holders team
    'Weekly Target (Sales Trend %)': ['Weekly Target (Sales Trend %)'],
    'Weekly Achievement (Sales Trend %)': ['Weekly Achievement (Sales Trend %)'],
    'Sales Trend % Score': ['Sales Trend % Score'],
    'Weekly Target (Conversion Rate %)': ['Weekly Target (Conversion Rate %)'],
    'Weekly Achievement (Conversion Rate %)': ['Weekly Achievement (Conversion Rate %)'],
    'Conversion Rate % Score': ['Conversion Rate % Score'],
    # Common
    'Weekly Target (AOV)': ['Weekly Target (AOV)'],
    'Weekly Achievement (AOV)': ['Weekly Achievement (AOV)'],
    'AOV Achievement %': ['AOV Achievement %', 'AOV Achievement  %'],
    'AOV Score': ['AOV Score'],
    'Year': ['Year'],
}

STANDARD_COLUMNS = list(COLUMN_MAPPING.keys())

CRITICAL_COLUMNS = [ 
    'Employee', 'Team', 'Month', 'Week', 'Total Score',
    'Weekly Target (Sales)', 'Weekly Achievement (Sales)', 'Sales Achievement %', 'Sales Score',
    'Weekly Target (Selling Cost %)', 'Weekly Achievement (Selling Cost %)', 'Selling Cost % Score',
    'Weekly Target (AOV)', 'Weekly Achievement (AOV)', 'AOV Achievement %', 'AOV Score',
    # Ads team
    'Weekly Target (ACOS %)', 'Weekly Achievement (ACOS %)', 'ACOS % Score',
    # Website ads team
    'Weekly Target (ROAS %)', 'Weekly Achievement (ROAS %)', 'ROAS % Score',
    # PH team
    'Weekly Target (Sales Trend %)', 'Weekly Achievement (Sales Trend %)', 'Sales Trend % Score',
    'Weekly Target (Conversion Rate %)', 'Weekly Achievement (Conversion Rate %)', 'Conversion Rate % Score',
    'Year',
]

# ============================================
# HELPER FUNCTIONS
# ============================================
def standardize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    reverse_mapping = {}
    for standard_name, variations in COLUMN_MAPPING.items():
        for var in variations:
            reverse_mapping[var.strip().lower()] = standard_name
    
    column_mapping_dict = {}
    for col_original in df.columns:
        col_normalized = str(col_original).strip().lower()
        standard_name = reverse_mapping.get(col_normalized, str(col_original).strip())
        column_mapping_dict[col_original] = standard_name
    return df.rename(columns=column_mapping_dict)

# get_current_month_name is not currently used, but can be kept.
# def get_current_month_name() -> str:
# return datetime.now().strftime("%B")

# ============================================
# DATA PROCESSING FUNCTIONS
# ============================================
@st.cache_data(ttl=18000) # Cache for 5 hours
def load_and_process_data(uploaded_file) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    def clean_currency(value):
        if isinstance(value, (int, float)): return float(value)
        if isinstance(value, str):
            cleaned_value = ''.join(c for c in value if c.isdigit() or c in '.-')
            try: return float(cleaned_value) if cleaned_value and cleaned_value != '-' else None
            except (ValueError, TypeError): return None
        return None

    def clean_percentage(value):
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            cleaned_value = value.replace('%', '').strip()
            if cleaned_value.lower() in ['n/a', '', 'nan', 'none']:
                return None
            try:
                # Accept numbers like '0.12', '12.5', '12.5%'
                num = float(cleaned_value)
                return num
            except (ValueError, TypeError):
                return None
        return None

    def clean_score(value):
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            cleaned_value = value.replace('%', '').replace(',', '').strip()
            # Remove any non-numeric except dot and minus
            cleaned_value = ''.join(c for c in cleaned_value if c.isdigit() or c in '.-')
            try:
                return float(cleaned_value)
            except (ValueError, TypeError):
                return None
        return None

    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names_original = xls.sheet_names
        
        # --- Sheet Skipping Logic ---
        # Default: skip first 3 sheets (indices 0, 1, 2)
        data_sheet_start_index = 3 
        # Example of more dynamic skipping:
        # skip_until_sheet = 'Skipped KPI Rows' 
        # if skip_until_sheet in sheet_names_original:
        #     data_sheet_start_index = sheet_names_original.index(skip_until_sheet) + 1
        #     print(f"Found '{skip_until_sheet}', data processing starts from sheet index {data_sheet_start_index}.")
        # else:
        #     print(f"'{skip_until_sheet}' not found, using default start index {data_sheet_start_index} for data sheets.")
        
        sheets_to_process_original = sheet_names_original[data_sheet_start_index:] if len(sheet_names_original) > data_sheet_start_index else []
        
        if not sheets_to_process_original:
            return None, f"Error: No data sheets found to process based on current skipping logic (starts after sheet at index {data_sheet_start_index-1})."

        all_employee_dfs = []
        for sheet_name_iter in sheets_to_process_original:
            sheet_name_for_data = str(sheet_name_iter).strip()
            try:
                df_sheet = pd.read_excel(xls, sheet_name=sheet_name_iter, header=0, dtype=str) # Read all as str
                
                if df_sheet.empty:
                    continue
                
                df_sheet = standardize_column_names(df_sheet)
                
                # Employee Name Handling
                if 'Employee' in df_sheet.columns and not df_sheet['Employee'].astype(str).str.strip().replace('', pd.NA).isna().all():
                    df_sheet['Employee'] = df_sheet['Employee'].astype(str).str.strip()
                else: 
                    df_sheet['Employee'] = sheet_name_for_data # Fallback to sheet name
                
                # Ensure all CRITICAL_COLUMNS (standard names) exist
                for std_col_name in CRITICAL_COLUMNS:
                    if std_col_name not in df_sheet.columns:
                        df_sheet[std_col_name] = pd.NA # Use pd.NA for better type handling later
                
                # Define which columns get which cleaning function (using STANDARD names)
                currency_cols_std = ['Weekly Target (Sales)', 'Weekly Achievement (Sales)', 'Weekly Target (AOV)', 'Weekly Achievement (AOV)']
                percentage_cols_std = [
                    'Sales Achievement %',
                    'Weekly Target (Selling Cost %)', 'Weekly Achievement (Selling Cost %)',
                    'AOV Achievement %',
                    'Weekly Target (ACOS %)', 'Weekly Achievement (ACOS %)',  # <-- Added for Ads team
                    'Weekly Target (ROAS %)', 'Weekly Achievement (ROAS %)',  # <-- For Website Ads team
                    'Weekly Target (Sales Trend %)', 'Weekly Achievement (Sales Trend %)',  # <-- For Portfolio Holders
                    'Weekly Target (Conversion Rate %)', 'Weekly Achievement (Conversion Rate %)'  # <-- For Portfolio Holders
                ]
                score_cols_std = [
                    'Sales Score', 'Selling Cost % Score', 'AOV Score', 'Total Score',
                    'ACOS % Score', 'ROAS % Score', 'Sales Trend % Score', 'Conversion Rate % Score'
                ]
                
                for col_name in df_sheet.columns:
                    if col_name in currency_cols_std: df_sheet[col_name] = df_sheet[col_name].apply(clean_currency)
                    elif col_name in percentage_cols_std: df_sheet[col_name] = df_sheet[col_name].apply(clean_percentage)
                    elif col_name in score_cols_std: df_sheet[col_name] = df_sheet[col_name].apply(clean_score)

                # Debug: Show first few values after cleaning for ACOS % columns
                if 'Weekly Target (ACOS %)' in df_sheet.columns:
                    print('DEBUG: Weekly Target (ACOS %) after cleaning:', df_sheet['Weekly Target (ACOS %)'].head(3).tolist())
                if 'Weekly Achievement (ACOS %)' in df_sheet.columns:
                    print('DEBUG: Weekly Achievement (ACOS %) after cleaning:', df_sheet['Weekly Achievement (ACOS %)'].head(3).tolist())
                
                if 'Week' in df_sheet.columns: df_sheet['Week'] = pd.to_numeric(df_sheet['Week'], errors='coerce').fillna(0).astype(int) # fillna(0) for week might be debatable
                if 'Month' in df_sheet.columns: df_sheet['Month'] = df_sheet['Month'].astype(str).str.strip().str.title()

                # Drop rows where all key value columns are NaN.
                value_cols_for_dropna = [c for c in CRITICAL_COLUMNS if c not in ['Employee', 'Team', 'Month', 'Week']]
                df_sheet.dropna(how='all', subset=value_cols_for_dropna, inplace=True)
                
                if not df_sheet.empty:
                    cols_to_keep = [col for col in STANDARD_COLUMNS if col in df_sheet.columns]
                    df_sheet = df_sheet[cols_to_keep] # Ensure consistent columns
                    all_employee_dfs.append(df_sheet)
                    print(f"  Successfully processed sheet: '{sheet_name_for_data}' with {len(df_sheet)} valid rows.")
                else: print(f"  Warning: No valid data rows in sheet '{sheet_name_for_data}' after cleaning.")
            except Exception as e: print(f"  Error processing sheet '{sheet_name_for_data}': {str(e)}\n{traceback.format_exc()}")
        
        if not all_employee_dfs: return None, "Error: No data could be successfully extracted from any sheets."
            
        try:
            combined_df = pd.concat(all_employee_dfs, ignore_index=True, sort=False)
            combined_df.dropna(how='all', inplace=True)
            combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()] 
            
            if combined_df.empty: return None, "Error: Combined data is empty after processing all sheets."

            if 'Employee' in combined_df.columns: combined_df['Employee'] = combined_df['Employee'].astype(str).str.strip()

            if 'Month' in combined_df.columns:
                valid_month_names = [m for m in calendar.month_name[1:] if m]
                month_categories = pd.CategoricalDtype(categories=valid_month_names, ordered=True)
                
                unique_months_in_data = combined_df['Month'].unique()
                problem_months = [m for m in unique_months_in_data if m not in valid_month_names and pd.notna(m) and m != '']
                if problem_months: print(f"WARNING: Non-standard month names found: {problem_months}. These will become NaT.")
                
                combined_df['Month'] = combined_df['Month'].astype(month_categories)
                # combined_df.dropna(subset=['Month'], inplace=True) # Optional: remove rows with invalid months

            if 'Week' in combined_df.columns: combined_df['Week'] = pd.to_numeric(combined_df['Week'], errors='coerce').fillna(0).astype(int)
            if 'Total Score' in combined_df.columns: combined_df['Total Score'] = pd.to_numeric(combined_df['Total Score'], errors='coerce')


            sort_cols_list = []
            if 'Employee' in combined_df.columns: sort_cols_list.append('Employee')
            if 'Month' in combined_df.columns and combined_df['Month'].dtype.name == 'category': sort_cols_list.append('Month')
            if 'Week' in combined_df.columns: sort_cols_list.append('Week')
            
            if sort_cols_list:
                try: combined_df.sort_values(by=sort_cols_list, ascending=True, inplace=True)
                except Exception as e: print(f"Warning: Could not sort combined_df: {e}")
            
            print(f"\nSuccessfully processed {len(all_employee_dfs)} sheets. Final combined DataFrame has {len(combined_df)} rows.")
            return combined_df, None
        except Exception as e: return None, f"Error during final combination/cleanup: {str(e)}\n{traceback.format_exc()}"
    except Exception as e: return None, f"Fatal error in load_and_process_data: {str(e)}\n{traceback.format_exc()}"

# ============================================
# UI COMPONENTS
# ============================================
def display_employee_metrics(employee_data: pd.DataFrame) -> None:
    if employee_data.empty: return
    def format_metric(value, is_percent: bool = False, is_currency: bool = False) -> str:
        if pd.isna(value) or (isinstance(value, str) and value.strip().lower() in ['n/a', '', 'nan', 'none']):
            return 'N/A'
        try:
            num_value = float(value)
            if is_percent:
                if 0 < abs(num_value) < 1 and num_value != 0:
                    num_value *= 100
                return f"{num_value:.1f}%"
            elif is_currency:
                return f"€{num_value:,.2f}"
            return f"{int(num_value)}" if num_value == int(num_value) else f"{num_value:.2f}"
        except (ValueError, TypeError): return 'N/A'

    latest_data = employee_data.iloc[-1] if not employee_data.empty else pd.Series(dtype='object')
    team_name = str(latest_data.get('Team', '')).strip().lower()

    # Default: Sales team (Selling Cost %)
    if 'ads' in team_name and 'website' not in team_name:
        cost_label = 'ACOS %'
        target_col = 'Weekly Target (ACOS %)' 
        achieve_col = 'Weekly Achievement (ACOS %)' 
        score_col = 'ACOS % Score'
    elif 'website' in team_name:
        cost_label = 'ROAS %'
        target_col = 'Weekly Target (ROAS %)' 
        achieve_col = 'Weekly Achievement (ROAS %)' 
        score_col = 'ROAS % Score'
    elif 'portfolio' in team_name:
        cost_label = 'Sales Trend %'
        target_col = 'Weekly Target (Sales Trend %)' 
        achieve_col = 'Weekly Achievement (Sales Trend %)' 
        score_col = 'Sales Trend % Score'
    else:
        cost_label = 'Selling Cost %'
        target_col = 'Weekly Target (Selling Cost %)' 
        achieve_col = 'Weekly Achievement (Selling Cost %)' 
        score_col = 'Selling Cost % Score'

    # Debug print for dynamic cost metric value
    print(f"DEBUG: For team '{team_name}', using score_col '{score_col}'. Value in latest_data: {latest_data.get(score_col)}")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Weekly Target (Sales)", format_metric(latest_data.get('Weekly Target (Sales)'), is_currency=True))
        st.metric(f"Weekly Target ({cost_label})", format_metric(latest_data.get(target_col), is_percent=True))
        st.metric("Weekly Target (AOV)", format_metric(latest_data.get('Weekly Target (AOV)'), is_currency=True))
    with col2:
        st.metric("Weekly Achievement (Sales)", format_metric(latest_data.get('Weekly Achievement (Sales)'), is_currency=True))
        st.metric(f"Weekly Achievement ({cost_label})", format_metric(latest_data.get(achieve_col), is_percent=True))
        st.metric("Weekly Achievement (AOV)", format_metric(latest_data.get('Weekly Achievement (AOV)'), is_currency=True))
    with col3:
        st.metric("Sales Achievement %", format_metric(latest_data.get('Sales Achievement %'), is_percent=True))
        st.metric(f"{cost_label} Score", format_metric(latest_data.get(score_col)))
        st.metric("AOV Achievement %", format_metric(latest_data.get('AOV Achievement %'), is_percent=True))
    with col4:
        st.metric("Sales Score", format_metric(latest_data.get('Sales Score')))
        st.metric("AOV Score", format_metric(latest_data.get('AOV Score')))
        st.metric("Total Score", format_metric(latest_data.get('Total Score')))

def display_employee_card(employee_data_for_card: pd.DataFrame, employee_name: str, selected_month_display: str, selected_week_display: str) -> None:
    with st.container(border=True):
        st.markdown(f"#### {employee_name}")
        if employee_data_for_card.empty:
            period_str_list = []
            if selected_month_display != "All Months": period_str_list.append(selected_month_display)
            if selected_week_display != "All Weeks": period_str_list.append(f"Week {selected_week_display}")
            period_message = ", ".join(period_str_list) if period_str_list else "the selected period"
            st.info(f"No data available for {employee_name} for {period_message}.")
            return

        latest_row = employee_data_for_card.iloc[-1]
        col_info1, col_info2, col_info3 = st.columns(3)
        col_info1, col_info2, col_info3 = st.columns(3)
        with col_info1: st.markdown(f"**Team:** {latest_row.get('Team', 'N/A') if pd.notna(latest_row.get('Team')) else 'N/A'}")
        with col_info2: st.markdown(f"**Month (Latest):** {str(latest_row.get('Month', selected_month_display))}") # Ensure month is str for display
        with col_info3: st.markdown(f"**Week (Latest):** {latest_row.get('Week', selected_week_display if selected_week_display != 'All Weeks' else 'N/A')}")
        
        st.markdown("---", help="Metrics below are from the latest record within the selection for this employee.")
        display_employee_metrics(employee_data_for_card)
        
        with st.expander("View Detailed Data for this Selection (Card)", expanded=False):
            # Show only value columns, drop identifying columns already displayed
            cols_to_drop_in_table = ['Employee','Team','Month','Week']
            display_df_card = employee_data_for_card.drop(columns=cols_to_drop_in_table, errors='ignore').reset_index(drop=True)
            st.dataframe(display_df_card, use_container_width=True, hide_index=True)

def display_achievements_and_performers_chart(data_for_view: pd.DataFrame, selected_week_filter_val: Union[str, int], selected_month_filter_val: str) -> None:
    st.header("🏆 Achievements & Top Performers")
    
    if data_for_view.empty:
        st.info("No data for the selected period to determine achievements or top performers.")
        return

    perfect_score = 10 
    # Ensure 'Total Score' is numeric for reliable comparison and aggregation
    if 'Total Score' in data_for_view.columns:
        # Make a copy to avoid SettingWithCopyWarning if data_for_view is a slice
        data_for_view = data_for_view.copy() 
        data_for_view['Total Score'] = pd.to_numeric(data_for_view['Total Score'], errors='coerce')
    else:
        st.warning("'Total Score' column not found. Cannot determine perfect scores or display performer chart.")
        return # Exit if no Total Score column
    
    # --- Star of the Month and Performer of the Week logic ---
    # Only show Star of the Month if filtering by month only (not week)
    if selected_month_filter_val != "All Months" and selected_week_filter_val == "All Weeks" and 'Month' in data_for_view.columns:
        month_df = data_for_view[data_for_view['Month'] == selected_month_filter_val].copy()
        if not month_df.empty and 'Total Score' in month_df.columns and 'Weekly Achievement (Sales)' in month_df.columns:
            month_df['Total Score'] = pd.to_numeric(month_df['Total Score'], errors='coerce')
            month_df['Weekly Achievement (Sales)'] = pd.to_numeric(month_df['Weekly Achievement (Sales)'], errors='coerce')
            perfect_employees = month_df[month_df['Total Score'] == perfect_score]['Employee'].dropna().unique().tolist()
            if perfect_employees:
                sales_sum = month_df[month_df['Employee'].isin(perfect_employees)].groupby('Employee')['Weekly Achievement (Sales)'].sum().reset_index()
                if not sales_sum.empty:
                    max_sales = sales_sum['Weekly Achievement (Sales)'].max()
                    stars = sales_sum[sales_sum['Weekly Achievement (Sales)'] == max_sales]['Employee'].tolist()
                    if stars:
                        st.success(f"⭐ Star of the Month ({selected_month_filter_val}): {', '.join(stars)} (Total Net Sales: €{max_sales:,.2f})")
    # Show Performer of the Week if filtering by week only (not month)
    elif selected_week_filter_val != "All Weeks" and selected_month_filter_val == "All Months" and 'Week' in data_for_view.columns:
        week_df = data_for_view[data_for_view['Week'] == selected_week_filter_val].copy()
        if not week_df.empty and 'Total Score' in week_df.columns and 'Weekly Achievement (Sales)' in week_df.columns:
            week_df['Total Score'] = pd.to_numeric(week_df['Total Score'], errors='coerce')
            week_df['Weekly Achievement (Sales)'] = pd.to_numeric(week_df['Weekly Achievement (Sales)'], errors='coerce')
            perfect_employees = week_df[week_df['Total Score'] == perfect_score]['Employee'].dropna().unique().tolist()
            if perfect_employees:
                sales_sum = week_df[week_df['Employee'].isin(perfect_employees)].groupby('Employee')['Weekly Achievement (Sales)'].sum().reset_index()
                if not sales_sum.empty:
                    max_sales = sales_sum['Weekly Achievement (Sales)'].max()
                    stars = sales_sum[sales_sum['Weekly Achievement (Sales)'] == max_sales]['Employee'].tolist()
                    if stars:
                        st.success(f"🏅 Performer of the Week (Week {selected_week_filter_val}): {', '.join(stars)} (Net Sales: €{max_sales:,.2f})")

    top_performers_df = data_for_view.dropna(subset=['Total Score'])
    top_performers_df = top_performers_df[top_performers_df['Total Score'] == perfect_score]
    top_performers_employees = top_performers_df['Employee'].dropna().unique()

    period_desc_list = []
    if selected_month_filter_val != "All Months": period_desc_list.append(selected_month_filter_val)
    if selected_week_filter_val != "All Weeks": period_desc_list.append(f"Week {selected_week_filter_val}")
    current_period_str = ", ".join(period_desc_list) if period_desc_list else "the overall data"

    if len(top_performers_employees) > 0:
        st.success(f"{len(top_performers_employees)} employee(s) achieved perfect scores ({perfect_score} points) in {current_period_str}!")
        st.markdown("**Congratulations to:**")
        # Show team name next to each performer
        for performer_name in top_performers_employees:
            # Find the team for this performer (use the latest record for that employee)
            team = None
            if 'Team' in data_for_view.columns:
                emp_rows = data_for_view[data_for_view['Employee'] == performer_name]
                if not emp_rows.empty:
                    team = emp_rows.iloc[-1]['Team']
            team_str = f" ({team})" if team and pd.notna(team) else ""
            st.markdown(f"- {performer_name}{team_str}")
    else:
        st.info(f"No employees achieved a perfect score of {perfect_score} for {current_period_str}.")
    
    st.markdown("---")
    st.subheader(f"Employee Scores for {current_period_str}")

    # Prepare data for the chart, ensuring 'Total Score' is numeric and handling NaNs
    if 'Employee' in data_for_view.columns and data_for_view['Total Score'].notna().any():
        chart_data = data_for_view[['Employee', 'Total Score']].copy()
        chart_data.dropna(subset=['Total Score'], inplace=True)

        agg_scores_df: pd.DataFrame
        chart_title: str

        if selected_week_filter_val == "All Weeks" and selected_month_filter_val != "All Months":
            agg_scores_df = chart_data.groupby('Employee')['Total Score'].mean().reset_index()
            chart_title = f"Average Total Scores in {selected_month_filter_val}"
        elif selected_week_filter_val == "All Weeks" and selected_month_filter_val == "All Months":
            agg_scores_df = chart_data.groupby('Employee')['Total Score'].mean().reset_index()
            chart_title = "Average Total Scores (All Time)"
        else:
            agg_scores_df = chart_data.drop_duplicates(subset=['Employee'], keep='last')
            chart_title = f"Total Scores for {current_period_str}"

        if not agg_scores_df.empty:
            agg_scores_df = agg_scores_df.sort_values('Total Score', ascending=False)
            fig = px.bar(agg_scores_df, 
                         x='Employee', y='Total Score', title=chart_title,
                         text='Total Score', color='Total Score',
                         color_continuous_scale=px.colors.sequential.Tealgrn)
            fig.update_traces(texttemplate='%{text:.1f}', textposition='outside')
            fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', xaxis_tickangle=-45,
                              height=500)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No scores to display in chart after aggregation for the current selection.")
    else:
        st.info("Required data ('Employee' or valid 'Total Score' values) missing for the performers chart.")

# ============================================
# MAIN APPLICATION
# ============================================
def main():
    st.set_page_config(page_title="Employee Performance Dashboard", page_icon="📊", layout="wide")
    st.title("📊 Employee Performance Dashboard")
    
    # Initialize session state variables if they don't exist
    if 'master_df' not in st.session_state: st.session_state.master_df = None
    if 'error_msg' not in st.session_state: st.session_state.error_msg = None
    if 'current_file_name' not in st.session_state: st.session_state.current_file_name = None

    uploaded_file = st.sidebar.file_uploader("Upload KPI Excel File", type=["xlsx"], help="Upload your Excel file with employee KPI data.")
    
    if uploaded_file:
        # Process file if it's new or if data hasn't been loaded successfully yet
        if st.session_state.current_file_name != uploaded_file.name or st.session_state.master_df is None:
            with st.spinner("Processing data, please wait..."):
                print(f"\n--- File Upload Detected. Processing: {uploaded_file.name} ---")
                # load_and_process_data is cached, so it only recomputes if input changes
                df_loaded, error_loaded = load_and_process_data(uploaded_file) 
                
                st.session_state.master_df = df_loaded
                st.session_state.error_msg = error_loaded
                st.session_state.current_file_name = uploaded_file.name
                
                # Rerun to update UI based on new data/error state.
                # This is important after the first successful load or if an error occurs.
                if error_loaded or (df_loaded is not None and st.session_state.master_df is not None):
                     st.rerun() 

    if st.session_state.error_msg:
        st.error(f"Error processing file '{st.session_state.current_file_name}': {st.session_state.error_msg}")
        st.sidebar.warning("Please check the file for issues or upload a corrected version.")
        return 
    
    if st.session_state.master_df is None:
        st.info("👋 Welcome! Please upload your KPI Excel file to get started.")
        return

    df = st.session_state.master_df
    # Success message can be shown once after initial load, or conditionally
    # st.success(f"Successfully loaded and processed data from '{st.session_state.current_file_name}'. Displaying {len(df)} records.")
        
    all_employees_from_master = sorted(df['Employee'].dropna().unique().tolist()) if 'Employee' in df.columns else []

    # --- Sidebar Filters ---
    with st.sidebar:
        st.subheader("Filters")
        # Team Filter
        # Dynamically build team options from the data, mapping to display names if possible
        raw_teams = list(df['Team'].dropna().astype(str).unique()) if 'Team' in df.columns else []
        # Map raw team values to display names
        team_display_map = {
            'sales': 'Sales team',
            'ads': 'Ads team',
            'website ads': 'Websit Ads team',
            'portfolio holders': 'Portfolio Holders team',
        }
        # Lowercase and strip for matching
        mapped_teams = []
        for t in raw_teams:
            t_clean = t.lower().replace(' ', '').replace('-', '').replace('_', '')
            if 'sales' in t_clean:
                mapped_teams.append('Sales team')
            elif 'ads' in t_clean and 'website' in t_clean:
                mapped_teams.append('Websit Ads team')
            elif 'ads' in t_clean:
                mapped_teams.append('Ads team')
            elif 'portfolio' in t_clean:
                mapped_teams.append('Portfolio Holders team')
        mapped_teams = sorted(set(mapped_teams))
        team_options = ['All Teams'] + mapped_teams
        if len(team_options) == 1:
            st.warning("No teams found in the data to select. Please check your data file.")
        selected_team = st.selectbox("Team", team_options, index=0, key="selectbox_team")

        # Year Filter
        year_options = ["All Years"]
        if 'Year' in df.columns and df['Year'].notna().any():
            years_sorted = sorted(df['Year'].dropna().unique())
            year_options.extend(years_sorted)
        selected_year = st.selectbox("Year", year_options, index=0, key="selectbox_year")

        # Month Filter
        month_options = ["All Months"]
        if 'Month' in df.columns and df['Month'].notna().any():
            # Use calendar.month_name[1:] for correct order
            valid_months = [m for m in calendar.month_name[1:] if m in df['Month'].dropna().unique()]
            month_options.extend(valid_months)
        selected_month = st.selectbox("Month", month_options, index=0, key="selectbox_month")

        # Week Filter
        week_options = ["All Weeks"]
        if 'Week' in df.columns and df['Week'].notna().any():
            week_values = pd.to_numeric(df['Week'], errors='coerce').dropna()
            week_values = week_values[week_values > 0].astype(int).unique().tolist()
            week_options.extend(sorted(week_values))
        selected_week_str = st.selectbox("Week", week_options, index=0, 
                                         format_func=lambda x: f"Week {x}" if x != "All Weeks" else "All Weeks", 
                                         key="selectbox_week")
        selected_week_value = int(selected_week_str) if selected_week_str != "All Weeks" else "All Weeks"

        # Employee Filter
        employee_options_cards = ["All Employees"] + all_employees_from_master
        selected_employee_cards = st.selectbox("Employee", employee_options_cards, key="selectbox_employee_cards")

    # --- Apply Global Filters to Create `data_view` ---
    data_view = df.copy()
    # Map selected_team back to raw team values for filtering
    if selected_team != "All Teams" and 'Team' in data_view.columns:
        # Find all raw team values that map to the selected display name
        team_match = []
        for t in data_view['Team'].dropna().astype(str).unique():
            t_clean = t.lower().replace(' ', '').replace('-', '').replace('_', '')
            if selected_team == 'Sales team' and 'sales' in t_clean:
                team_match.append(t)
            elif selected_team == 'Websit Ads team' and 'ads' in t_clean and 'website' in t_clean:
                team_match.append(t)
            elif selected_team == 'Ads team' and 'ads' in t_clean and 'website' not in t_clean:
                team_match.append(t)
            elif selected_team == 'Portfolio Holders team' and 'portfolio' in t_clean:
                team_match.append(t)
        if team_match:
            data_view = data_view[data_view['Team'].isin(team_match)]
        else:
            data_view = data_view.iloc[0:0]  # Empty DataFrame if no match
    if selected_year != "All Years" and 'Year' in data_view.columns:
        data_view = data_view[data_view['Year'] == selected_year]
    if selected_month != "All Months" and 'Month' in data_view.columns:
        data_view = data_view[data_view['Month'].astype(str) == selected_month]
    if selected_week_value != "All Weeks" and 'Week' in data_view.columns:
        data_view = data_view[data_view['Week'] == selected_week_value]

    # --- Display Filter & Compare Section ---
    st.markdown("## Filter & Compare")
    st.markdown(
        """
        Filter by team, product, marketplace, or time period using the selections in the sidebar.\
        Compare performance across individuals or teams below.
        """
    )
    st.markdown("---")
    st.markdown("## Sales Metrics & Trends")

    # --- Handle single employee filter feedback ---
    if selected_employee_cards != "All Employees":
        if data_view.empty:
            st.warning(f"No data available for employee '{selected_employee_cards}' in the selected period.")
            return
        else:
            st.info(f"Showing metrics and trends for employee: {selected_employee_cards}")

    # --- Weekly Sales Breakdown ---
    if 'Week' in data_view.columns and 'Weekly Achievement (Sales)' in data_view.columns:
        weekly_sales_breakdown = data_view.groupby('Week')['Weekly Achievement (Sales)'].sum().reset_index()
        weekly_sales_breakdown = weekly_sales_breakdown[weekly_sales_breakdown['Week'] > 0]
        weekly_sales_breakdown = weekly_sales_breakdown.sort_values('Week')  # Ensure weeks are in ascending order
        st.subheader("Weekly Sales Breakdown")
        if not weekly_sales_breakdown.empty:
            cols = st.columns(min(5, len(weekly_sales_breakdown)))
            for idx, row in weekly_sales_breakdown.iterrows():
                with cols[idx % len(cols)]:
                    st.metric(f"Week {int(row['Week'])}", f"€{row['Weekly Achievement (Sales)']:,.2f}")
        else:
            st.info("No weekly sales data available for the current filter.")
    else:
        st.info("No weekly sales data available for the current filter.")

    # --- Monthly Sales Breakdown ---
    if 'Month' in data_view.columns and 'Weekly Achievement (Sales)' in data_view.columns:
        monthly_sales_breakdown = data_view.groupby('Month')['Weekly Achievement (Sales)'].sum().reset_index()
        st.subheader("Monthly Sales Breakdown")
        if not monthly_sales_breakdown.empty:
            cols = st.columns(min(5, len(monthly_sales_breakdown)))
            for idx, row in monthly_sales_breakdown.iterrows():
                with cols[idx % len(cols)]:
                    st.metric(f"{row['Month']}", f"€{row['Weekly Achievement (Sales)']:,.2f}")
        else:
            st.info("No monthly sales data available for the current filter.")
    else:
        st.info("No monthly sales data available for the current filter.")

    # --- Sales Trend Chart ---
    if 'Week' in data_view.columns and 'Weekly Achievement (Sales)' in data_view.columns:
        sales_trend_df = data_view.groupby('Week')['Weekly Achievement (Sales)'].sum().reset_index()
        fig_sales_trend = px.line(sales_trend_df, x='Week', y='Weekly Achievement (Sales)', title='Sales Trend by Week')
        st.plotly_chart(fig_sales_trend, use_container_width=True)

    # --- Overall Performance Trend Chart ---
    if 'month' in data_view.columns and 'weekly achievement (sales)' in data_view.columns:
        perf_trend_df = data_view.groupby('month')['weekly achievement (sales)'].mean().reset_index()
        fig_perf_trend = px.line(perf_trend_df, x='month', y='weekly achievement (sales)', title='Overall Performance Trend by Month')
        st.plotly_chart(fig_perf_trend, use_container_width=True)

    # --- Weekly Sales Trends (This Year vs. Last Year) ---
    if 'Week' in data_view.columns and 'Weekly Achievement (Sales)' in data_view.columns and 'Month' in data_view.columns:
        # Try to infer year from Month if possible
        if 'Year' in data_view.columns:
            weekly_year_df = data_view.groupby(['Year', 'Week'])['Weekly Achievement (Sales)'].sum().reset_index()
            fig_weekly_yoy = px.line(weekly_year_df, x='Week', y='Weekly Achievement (Sales)', color='Year', title='Weekly Sales Trends (This Year vs. Last Year)')
            st.plotly_chart(fig_weekly_yoy, use_container_width=True)
    # --- Monthly Sales Trends (This Year vs. Last Year) ---
    if 'Month' in data_view.columns and 'Weekly Achievement (Sales)' in data_view.columns:
        if 'Year' in data_view.columns:
            monthly_year_df = data_view.groupby(['Year', 'Month'])['Weekly Achievement (Sales)'].sum().reset_index()
            fig_monthly_yoy = px.line(monthly_year_df, x='Month', y='Weekly Achievement (Sales)', color='Year', title='Monthly Sales Trends (This Year vs. Last Year)')
            st.plotly_chart(fig_monthly_yoy, use_container_width=True)

    # --- Sales Trends This month vs. Last month (Target vs. Achieved) ---
    st.subheader("Sales Trends This month vs. Last month (Target vs. Achieved)")
    if (
        'Month' in data_view.columns and
        'Weekly Target (Sales)' in data_view.columns and
        'Weekly Achievement (Sales)' in data_view.columns
    ):
        # Get months with data, sorted by calendar order, for the current filtered data_view (including employee filter)
        filtered_months = data_view['Month'].dropna().unique().tolist()
        valid_months = [m for m in calendar.month_name[1:] if m in filtered_months]
        if len(valid_months) >= 2:
            last_two_months = valid_months[-2:]
            month_trend_df = data_view[data_view['Month'].isin(last_two_months)]
            # Group by Month, sum target and achieved
            month_trend_grouped = month_trend_df.groupby('Month').agg({
                'Weekly Target (Sales)': 'sum',
                'Weekly Achievement (Sales)': 'sum'
            }).reset_index()
            # Melt for plotting
            month_trend_melted = pd.melt(
                month_trend_grouped,
                id_vars=['Month'],
                value_vars=['Weekly Target (Sales)', 'Weekly Achievement (Sales)'],
                var_name='Metric',
                value_name='Sales (€)'
            )
            fig_month_vs_last = px.bar(
                month_trend_melted,
                x='Month',
                y='Sales (€)',
                color='Metric',
                barmode='group',
                text='Sales (€)',
                title='Sales Trends This month vs. Last month (Target vs. Achieved)'
            )
            fig_month_vs_last.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            fig_month_vs_last.update_layout(height=400)
            st.plotly_chart(fig_month_vs_last, use_container_width=True)
        else:
            st.info("Not enough monthly data to compare this month vs. last month.")
    else:
        st.info("No data found for Sales Trends This month vs. Last month (Target vs. Achieved) for the current filter.")

    # --- Overall Performance Trend Chart ---
    st.subheader("Sales Trends This Year vs. Last Year (Target vs. Achieved)")
    if (
        'Year' in data_view.columns and
        'Month' in data_view.columns and
        'Weekly Target (Sales)' in data_view.columns and
        'Weekly Achievement (Sales)' in data_view.columns
    ):
        # Prepare data for both years, month-wise
        sales_trend_yoy = data_view.groupby(['Year', 'Month']).agg({
            'Weekly Target (Sales)': 'sum',
            'Weekly Achievement (Sales)': 'sum'
        }).reset_index()
        if not sales_trend_yoy.empty:
            # Melt for plotting: one line for Target, one for Achieved, colored by Year
            sales_trend_melted = pd.melt(
                sales_trend_yoy,
                id_vars=['Year', 'Month'],
                value_vars=['Weekly Target (Sales)', 'Weekly Achievement (Sales)'],
                var_name='Metric',
                value_name='Sales (€)'
            )
            fig_yoy = px.line(
                sales_trend_melted,
                x='Month',
                y='Sales (€)',
                color='Year',
                line_dash='Metric',
                markers=True,
                title='Sales Trends This Year vs. Last Year (Target vs. Achieved)',
                labels={'Sales (€)': 'Sales (€)', 'Month': 'Month', 'Year': 'Year', 'Metric': 'Metric'}
            )
            fig_yoy.update_layout(legend_title_text='Year / Metric', height=500)
            st.plotly_chart(fig_yoy, use_container_width=True)
        else:
            st.info("No data found for Sales Trends This Year vs. Last Year (Target vs. Achieved) for the current filter.")
    else:
        st.info("No data found for Sales Trends This Year vs. Last Year (Target vs. Achieved) for the current filter.")

    st.markdown("---")
    
    if df.empty: # Should be caught by master_df is None, but as a safeguard
        st.warning("No data available to display.")
        return

    if data_view.empty and (selected_month != "All Months" or selected_week_value != "All Weeks"):
        st.warning(f"No data matches the selected filters: {selected_month}, {selected_week_str}.")
    
    display_achievements_and_performers_chart(data_view.copy(), selected_week_value, selected_month) # Pass copies
    
    st.markdown("---")
    st.subheader("Employee Performance Cards")
    
    # For All Employees, include rows with missing/blank Employee as 'Unknown' for card display
    if selected_employee_cards != "All Employees":
        employee_card_data = data_view[data_view['Employee'] == selected_employee_cards] if 'Employee' in data_view.columns else pd.DataFrame()
        display_employee_card(employee_card_data, selected_employee_cards, selected_month, selected_week_str)
    else: # Display cards for all employees, including blanks as 'Unknown'
        # Use all unique, non-empty employee names in the current filtered data_view
        if 'Employee' in data_view.columns:
            all_employees_in_view = data_view['Employee'].fillna('').astype(str).str.strip()
            # Compute max Total Score per employee for sorting
            if 'Total Score' in data_view.columns:
                score_per_employee = data_view.groupby('Employee', dropna=False)['Total Score'].max().reset_index()
                score_per_employee['Employee'] = score_per_employee['Employee'].fillna('Unknown').astype(str).str.strip()
                # Remove empty string employees except 'Unknown'
                score_per_employee = score_per_employee[score_per_employee['Employee'] != '']
                # Sort by Total Score descending, then by Employee name for tie-breaker
                score_per_employee = score_per_employee.sort_values(['Total Score', 'Employee'], ascending=[False, True])
                unique_employees = [e for e in score_per_employee['Employee'] if e.lower() != 'nan']
            else:
                unique_employees = sorted(set([e for e in all_employees_in_view if e and e.lower() != 'nan']))
            # Ensure 'Unknown' is last if present
            if 'Unknown' in unique_employees:
                unique_employees = [e for e in unique_employees if e != 'Unknown'] + ['Unknown']
            employee_order_list = unique_employees
        else:
            employee_order_list = []
        for employee_name_iter in employee_order_list:
            if employee_name_iter == 'Unknown':
                employee_card_data_iter = data_view[data_view['Employee'].isna() | (data_view['Employee'].astype(str).str.strip().isin(['', 'nan', 'NaN', 'None']))]
            else:
                employee_card_data_iter = data_view[data_view['Employee'].astype(str).str.strip() == employee_name_iter]
            display_employee_card(employee_card_data_iter, employee_name_iter, selected_month, selected_week_str)

    with st.expander("View Filtered Raw Data Table", expanded=False):
        st.dataframe(data_view, use_container_width=True)

if __name__ == "__main__":
    main()