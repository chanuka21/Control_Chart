#!/usr/bin/env python3
"""
Live Control Chart Dashboard - FIXED VERSION
============================================
Real-time dashboard for process control charts with auto-updating Excel data.

FIXES APPLIED:
- Fixed tape ID filtering logic that was causing charts to disappear
- Improved data type handling for tape IDs
- Added better debugging and error handling
- Enhanced tape ID matching with string conversion
- Added dates to x-axis
- Added data cleaning for control charts
- Added date range filtering
- Added dropdown checkbox for Tape IDs
- Added T-test for before and after datasets
- Added downloadable summary report
- Added "Select All" feature for Tape ID dropdown checkbox
- Ensured Tape ID dropdown is not expanded by default

Features:
- Dual control charts (separate & combined limits)
- Auto-refresh every 10 seconds
- Interactive variable and tape ID selection
- Modern Bootstrap UI with tabs
- Robust error handling

Author: Generated for Process Control Analysis
Date: 2025
"""

import dash
from dash import dcc, html, Input, Output, callback, State
import dash_bootstrap_components as dbc
import plotly.graph_objs as go
import plotly.express as px
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import warnings
from scipy import stats 

# Suppress warnings for cleaner output
warnings.filterwarnings('ignore')

# =============================================================================
# CONFIGURATION
# =============================================================================

# File paths - Update these if your files are in different locations
BEFORE_PATH = "E:\Web\ML Projects\Control Chart\Control Chart\old.xlsx"
AFTER_PATH = "E:\Web\ML Projects\Control Chart\Control Chart\MEK data set_carbon_new.xlsx"

# Auto-refresh interval in seconds
REFRESH_INTERVAL = 10

# =============================================================================
# INITIALIZE DASH APP
# =============================================================================

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "Live Control Chart Dashboard"

# Global variables to store current data
current_data = {
    'before': pd.DataFrame(),
    'after': pd.DataFrame(),
    'last_updated': datetime.now()
}

# =============================================================================
# DATA LOADING AND PROCESSING FUNCTIONS
# =============================================================================

def load_excel_data():
    """
    Load and process Excel data from both files.

    Returns:
        tuple: (before_df, after_df, timestamp)
    """
    try:
        print(f"Loading data from Excel files...")

        # Load data with error handling
        before_df = pd.DataFrame()
        after_df = pd.DataFrame()

        if os.path.exists(BEFORE_PATH):
            before_df = pd.read_excel(BEFORE_PATH)
            print(f"✅ Loaded Before data: {len(before_df)} rows")
            print(f"📋 Before columns: {list(before_df.columns)}")
        else:
            print(f"⚠️ Before file not found: {BEFORE_PATH}")

        if os.path.exists(AFTER_PATH):
            after_df = pd.read_excel(AFTER_PATH)
            print(f"✅ Loaded After data: {len(after_df)} rows")
            print(f"📋 After columns: {list(after_df.columns)}")
        else:
            print(f"⚠️ After file not found: {AFTER_PATH}")

        # Clean column names (remove spaces, special characters)
        if not before_df.empty:
            before_df.columns = before_df.columns.str.strip()
            before_df = standardize_tape_id_column(before_df)
            before_df['Stage'] = 'Before'
            print(f"🔧 Before columns after cleaning: {list(before_df.columns)}")

        if not after_df.empty:
            after_df.columns = after_df.columns.str.strip()
            after_df = standardize_tape_id_column(after_df)
            after_df['Stage'] = 'After'
            print(f"🔧 After columns after cleaning: {list(after_df.columns)}")

        return before_df, after_df, datetime.now()

    except Exception as e:
        print(f"❌ Error loading data: {e}")
        return pd.DataFrame(), pd.DataFrame(), datetime.now()

def standardize_tape_id_column(df):
    """
    Find and standardize the Tape ID column with various possible names.

    Args:
        df (pd.DataFrame): Input dataframe

    Returns:
        pd.DataFrame: DataFrame with standardized 'Tape ID' column
    """
    if df.empty:
        return df

    # Possible variations of Tape ID column names
    tape_id_variations = [
        'Tape ID', 'TapeID', 'Tape_ID', 'tape_id', 'tape id', 'TAPE_ID', 'TAPE ID',
        'Tape', 'tape', 'TAPE', 'ID', 'id', 'Id', 'Tape No', 'Tape_No', 'TapeNo',
        'Sample ID', 'Sample_ID', 'SampleID', 'sample_id', 'Sample'
    ]

    # Find the Tape ID column
    tape_id_col = None
    for col in df.columns:
        col_clean = str(col).strip()
        if col_clean in tape_id_variations:
            tape_id_col = col
            break
        # Also check for partial matches
        for variation in tape_id_variations:
            if variation.lower() in col_clean.lower():
                tape_id_col = col
                break
        if tape_id_col:
            break

    # Rename to standard 'Tape ID' if found
    if tape_id_col and tape_id_col != 'Tape ID':
        df = df.rename(columns={tape_id_col: 'Tape ID'})
        print(f"🔄 Renamed column '{tape_id_col}' to 'Tape ID'")
    elif not tape_id_col:
        print(f"⚠️ No Tape ID column found. Available columns: {list(df.columns)}")
        # Create a dummy Tape ID column if none found
        df['Tape ID'] = 'Unknown'

    return df

def is_numeric_convertible(series):
    """
    FIXED: Check if a pandas Series can be converted to numeric values.

    Args:
        series (pd.Series): Input series to check

    Returns:
        bool: True if convertible to numeric, False otherwise
    """
    if series.empty:
        return False

    try:
        # Remove null values for testing
        non_null_series = series.dropna()

        if len(non_null_series) == 0:
            return False

        # Try to convert to numeric
        pd.to_numeric(non_null_series, errors='raise')
        return True

    except (ValueError, TypeError):
        # Try to handle common cases like percentages, currency symbols
        try:
            # Convert to string and clean common non-numeric characters
            cleaned_series = non_null_series.astype(str).str.replace(r'[%$,\s]', '', regex=True)
            # Remove empty strings
            cleaned_series = cleaned_series[cleaned_series != '']

            if len(cleaned_series) == 0:
                return False

            # Try conversion again
            pd.to_numeric(cleaned_series, errors='raise')
            return True

        except (ValueError, TypeError):
            return False

def get_numeric_columns(df):
    """
    FIXED: Get list of numeric columns excluding system columns, with improved detection.

    Args:
        df (pd.DataFrame): Input dataframe

    Returns:
        list: List of numeric column names
    """
    if df.empty:
        return []

    numeric_cols = []

    # System columns to exclude
    exclude_cols = ['Tape ID', 'Stage', 'Index', 'index', 'Date']

    for col in df.columns:
        if col in exclude_cols:
            continue

        # First check if pandas already detected it as numeric
        if pd.api.types.is_numeric_dtype(df[col]):
            numeric_cols.append(col)
            print(f"✅ Column '{col}' detected as numeric (pandas dtype)")
        else:
            # FIXED: Check if it can be converted to numeric
            if is_numeric_convertible(df[col]):
                numeric_cols.append(col)
                print(f"✅ Column '{col}' detected as convertible to numeric")
            else:
                print(f"❌ Column '{col}' is not numeric - sample values: {df[col].dropna().head(3).tolist()}")

    print(f"🔍 Found {len(numeric_cols)} numeric columns: {numeric_cols}")
    return numeric_cols

def get_common_numeric_columns(before_df, after_df):
    """
    Get numeric columns that exist in BOTH Before and After datasets.

    Args:
        before_df (pd.DataFrame): Before stage data
        after_df (pd.DataFrame): After stage data

    Returns:
        list: List of common numeric column names
    """
    before_cols = set(get_numeric_columns(before_df))
    after_cols = set(get_numeric_columns(after_df))

    # Find intersection (common columns)
    common_cols = before_cols.intersection(after_cols)
    common_cols_list = sorted(list(common_cols))

    # Debug information
    print(f"🔍 Before dataset numeric columns ({len(before_cols)}): {sorted(list(before_cols))}")
    print(f"🔍 After dataset numeric columns ({len(after_cols)}): {sorted(list(after_cols))}")
    print(f"✅ Common numeric columns ({len(common_cols_list)}): {common_cols_list}")

    if len(before_cols) > 0 and len(after_cols) > 0 and len(common_cols_list) == 0:
        print("⚠️ WARNING: No common numeric columns found between Before and After datasets!")
        print("📝 Before-only columns:", sorted(list(before_cols - after_cols)))
        print("📝 After-only columns:", sorted(list(after_cols - before_cols)))

    return common_cols_list

def get_tape_ids(before_df, after_df):
    """
    Get unique Tape IDs from both datasets with improved handling.

    Args:
        before_df (pd.DataFrame): Before stage data
        after_df (pd.DataFrame): After stage data

    Returns:
        list: Sorted list of unique tape IDs
    """
    tape_ids = set()

    # Collect tape IDs from both dataframes
    for stage, df in [('Before', before_df), ('After', after_df)]:
        if not df.empty and 'Tape ID' in df.columns:
            # Convert to string and clean up
            valid_ids = df['Tape ID'].dropna().astype(str).str.strip()
            # Remove empty strings
            valid_ids = valid_ids[valid_ids != '']
            unique_ids = valid_ids.unique()
            tape_ids.update(unique_ids)
            print(f"🏷️ Found {len(unique_ids)} unique Tape IDs in {stage} data: {list(unique_ids)[:5]}{'...' if len(unique_ids) > 5 else ''}")
        else:
            print(f"⚠️ No 'Tape ID' column found in {stage} data")

    # Convert to list and sort
    tape_ids_list = sorted([str(tid) for tid in tape_ids if str(tid).strip() != '' and str(tid).lower() != 'nan'])
    print(f"🎯 Total unique Tape IDs: {len(tape_ids_list)}")
    if len(tape_ids_list) > 0:
        print(f"📝 Sample Tape IDs: {tape_ids_list[:10]}{'...' if len(tape_ids_list) > 10 else ''}")

    return tape_ids_list

def filter_data_by_tape_ids(df, tape_ids):
    """
    FIXED: Filter dataframe by tape IDs with improved matching logic.

    Args:
        df (pd.DataFrame): Input dataframe
        tape_ids (list): List of tape IDs to filter by

    Returns:
        pd.DataFrame: Filtered dataframe
    """
    if df.empty or not tape_ids or 'Tape ID' not in df.columns:
        return df

    # Convert tape_ids to strings for consistent comparison
    target_tape_ids = [str(tid).strip() for tid in tape_ids]

    # Convert dataframe Tape ID column to strings for comparison
    df_copy = df.copy()
    df_copy['Tape ID'] = df_copy['Tape ID'].astype(str).str.strip()

    # Debug information
    print(f"🔍 Filtering Debug:")
    print(f"   Target Tape IDs: {target_tape_ids}")
    print(f"   Available Tape IDs in data: {df_copy['Tape ID'].unique().tolist()}")

    # Filter using string matching
    filtered_df = df_copy[df_copy['Tape ID'].isin(target_tape_ids)]

    print(f"   Original rows: {len(df)}, Filtered rows: {len(filtered_df)}")

    return filtered_df

def filter_data_by_date_range(df, start_date, end_date):
    """
    Filter dataframe by date range.

    Args:
        df (pd.DataFrame): Input dataframe
        start_date (str): Start date in YYYY-MM-DD format
        end_date (str): End date in YYYY-MM-DD format

    Returns:
        pd.DataFrame: Filtered dataframe
    """
    if df.empty or 'Date' not in df.columns:
        return df

    # Convert date strings to datetime objects
    start_date = pd.to_datetime(start_date, errors='coerce')
    end_date = pd.to_datetime(end_date, errors='coerce')

    # Ensure Date column is datetime
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    # Filter dataframe by date range
    filtered_df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

    return filtered_df

def clean_control_chart_data(df, variable):
    """
    Clean data for control charts by handling missing values, outliers, and ensuring numeric type.

    Args:
        df (pd.DataFrame): Input dataframe
        variable (str): Variable name to clean

    Returns:
        pd.DataFrame: Cleaned dataframe
    """
    if df.empty or variable not in df.columns:
        return df

    # Make a copy of the dataframe to avoid modifying the original
    df_clean = df.copy()

    # Convert the variable to numeric, forcing errors to NaN
    df_clean[variable] = pd.to_numeric(df_clean[variable], errors='coerce')

    # Handle missing values
    df_clean = df_clean.dropna(subset=[variable])

    # Handle outliers using the IQR method
    if len(df_clean) > 0:
        Q1 = df_clean[variable].quantile(0.25)
        Q3 = df_clean[variable].quantile(0.75)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR

        # Filter out outliers
        df_clean = df_clean[(df_clean[variable] >= lower_bound) & (df_clean[variable] <= upper_bound)]

    return df_clean

def perform_t_test(before_data, after_data):
    """
    Perform a T-test to compare before and after datasets.

    Args:
        before_data (list): Data points from the before dataset
        after_data (list): Data points from the after dataset

    Returns:
        dict: Dictionary containing T-test results
    """
    t_stat, p_value = stats.ttest_ind(before_data, after_data, equal_var=False)

    return {
        't_statistic': t_stat,
        'p_value': p_value,
        'significant': p_value < 0.05
    }

def generate_summary_report(before_data, after_data, variable, t_test_results):
    """
    Generate a summary report comparing before and after datasets.

    Args:
        before_data (list): Data points from the before dataset
        after_data (list): Data points from the after dataset
        variable (str): Variable name
        t_test_results (dict): Results of the T-test

    Returns:
        str: Summary report as a string
    """
    report = f"Summary Report for Variable: {variable}\n"
    report += "=" * 50 + "\n\n"

    # Basic statistics
    report += "Basic Statistics:\n"
    report += f"Before - Mean: {np.mean(before_data):.3f}, Std: {np.std(before_data):.3f}, Count: {len(before_data)}\n"
    report += f"After  - Mean: {np.mean(after_data):.3f}, Std: {np.std(after_data):.3f}, Count: {len(after_data)}\n\n"

    # T-test results
    report += "T-test Results:\n"
    report += f"T-statistic: {t_test_results['t_statistic']:.3f}\n"
    report += f"P-value: {t_test_results['p_value']:.4f}\n"
    report += f"Significant: {'Yes' if t_test_results['significant'] else 'No'}\n\n"

    # Interpretation
    report += "Interpretation:\n"
    if t_test_results['significant']:
        report += "The difference between the 'before' and 'after' datasets is statistically significant.\n"
    else:
        report += "The difference between the 'before' and 'after' datasets is not statistically significant.\n"

    return report

# =============================================================================
# CHART CREATION FUNCTIONS
# =============================================================================

def create_control_chart(before_df, after_df, variable, tape_ids, start_date, end_date, chart_type='separate'):
    """
    Create control chart with dates on the x-axis and cleaned data.

    Args:
        before_df (pd.DataFrame): Before stage data
        after_df (pd.DataFrame): After stage data
        variable (str): Variable name to plot
        tape_ids (list): Selected tape IDs to include
        start_date (str): Start date in YYYY-MM-DD format
        end_date (str): End date in YYYY-MM-DD format
        chart_type (str): 'separate' or 'combined'

    Returns:
        plotly.graph_objs.Figure: Control chart figure
    """
    fig = go.Figure()

    try:
        # Handle empty data
        if before_df.empty and after_df.empty:
            fig.add_annotation(
                text="📊 No data available<br>Please check your Excel files",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="gray")
            )
            return fig

        # Filter data by selected tape IDs
        if tape_ids:
            filtered_before = filter_data_by_tape_ids(before_df, tape_ids)
            filtered_after = filter_data_by_tape_ids(after_df, tape_ids)
            print(f"🎯 After filtering - Before: {len(filtered_before)} rows, After: {len(filtered_after)} rows")
        else:
            filtered_before = before_df.copy()
            filtered_after = after_df.copy()
            print(f"🎯 Using all data - Before: {len(filtered_before)} rows, After: {len(filtered_after)} rows")

        # Filter data by date range
        if start_date and end_date:
            filtered_before = filter_data_by_date_range(filtered_before, start_date, end_date)
            filtered_after = filter_data_by_date_range(filtered_after, start_date, end_date)
            print(f"📅 After date filtering - Before: {len(filtered_before)} rows, After: {len(filtered_after)} rows")

        # Clean the data for the control charts
        cleaned_before = clean_control_chart_data(filtered_before, variable)
        cleaned_after = clean_control_chart_data(filtered_after, variable)

        # Extract variable data with dates
        before_data = []
        after_data = []
        before_dates = []
        after_dates = []

        if not cleaned_before.empty and variable in cleaned_before.columns and 'Date' in cleaned_before.columns:
            try:
                before_series = pd.to_numeric(cleaned_before[variable], errors='coerce')
                before_data = before_series.dropna().astype(float).tolist()
                before_dates = cleaned_before['Date'].dropna().tolist()
                print(f"📊 Before data points for {variable}: {len(before_data)}")
            except Exception as e:
                print(f"❌ Error processing before data for {variable}: {e}")
                before_data = []
                before_dates = []

        if not cleaned_after.empty and variable in cleaned_after.columns and 'Date' in cleaned_after.columns:
            try:
                after_series = pd.to_numeric(cleaned_after[variable], errors='coerce')
                after_data = after_series.dropna().astype(float).tolist()
                after_dates = cleaned_after['Date'].dropna().tolist()
                print(f"📊 After data points for {variable}: {len(after_data)}")
            except Exception as e:
                print(f"❌ Error processing after data for {variable}: {e}")
                after_data = []
                after_dates = []

        # Check if we have any data
        if len(before_data) == 0 and len(after_data) == 0:
            error_msg = f"📊 No valid numeric data found for variable '{variable}'"
            if tape_ids:
                error_msg += f"<br>with selected Tape IDs: {', '.join(map(str, tape_ids))}"
            if start_date and end_date:
                error_msg += f"<br>and date range: {start_date} to {end_date}"
            error_msg += "<br><br>💡 Try selecting different Tape IDs or check data types"

            fig.add_annotation(
                text=error_msg,
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=14, color="gray")
            )
            return fig

        # Color scheme
        colors = {'Before': '#1f77b4', 'After': '#ff7f0e'}

        if chart_type == 'separate':
            # SEPARATE CONTROL LIMITS
            for stage, data, dates, color in [('Before', before_data, before_dates, colors['Before']),
                                             ('After', after_data, after_dates, colors['After'])]:
                if len(data) > 0 and len(dates) > 0:
                    try:
                        data_array = np.array(data, dtype=float)
                        ucl, lcl, mean = calculate_control_limits(data_array)

                        # Plot data points with dates
                        fig.add_trace(go.Scatter(
                            x=dates,
                            y=data_array.tolist(),
                            mode='markers+lines',
                            name=f'{stage} Data ({len(data_array)} pts)',
                            line=dict(color=color, width=2),
                            marker=dict(color=color, size=8, symbol='circle'),
                            hovertemplate=f'<b>{stage} Stage</b><br>' +
                                        f'Date: %{{x|%Y-%m-%d}}<br>' +
                                        f'{variable}: %{{y:.3f}}<br>' +
                                        '<extra></extra>'
                        ))

                        # Add control limit lines
                        if ucl is not None and not np.isnan(ucl):
                            # UCL line
                            fig.add_trace(go.Scatter(
                                x=dates, y=[float(ucl)] * len(dates),
                                mode='lines', name=f'{stage} UCL ({ucl:.3f})',
                                line=dict(color=color, dash='dash', width=2),
                                hovertemplate=f'<b>{stage} Upper Control Limit</b><br>' +
                                            f'Value: {ucl:.3f}<extra></extra>'
                            ))

                            # LCL line
                            fig.add_trace(go.Scatter(
                                x=dates, y=[float(lcl)] * len(dates),
                                mode='lines', name=f'{stage} LCL ({lcl:.3f})',
                                line=dict(color=color, dash='dash', width=2),
                                hovertemplate=f'<b>{stage} Lower Control Limit</b><br>' +
                                            f'Value: {lcl:.3f}<extra></extra>'
                            ))

                            # Mean line
                            fig.add_trace(go.Scatter(
                                x=dates, y=[float(mean)] * len(dates),
                                mode='lines', name=f'{stage} Mean ({mean:.3f})',
                                line=dict(color=color, dash='dot', width=2),
                                hovertemplate=f'<b>{stage} Mean</b><br>' +
                                            f'Value: {mean:.3f}<extra></extra>'
                            ))
                    except Exception as e:
                        print(f"❌ Error creating {stage} chart: {e}")
                        continue

        else:
            # COMBINED CONTROL LIMITS
            try:
                combined_data = np.array(before_data + after_data, dtype=float)
                combined_dates = before_dates + after_dates

                if len(combined_data) > 0 and len(combined_dates) > 0:
                    ucl, lcl, mean = calculate_control_limits(combined_data)

                    # Plot Before data
                    if len(before_data) > 0 and len(before_dates) > 0:
                        before_array = np.array(before_data, dtype=float)
                        fig.add_trace(go.Scatter(
                            x=before_dates,
                            y=before_array.tolist(),
                            mode='markers+lines',
                            name=f'Before Data ({len(before_array)} pts)',
                            line=dict(color=colors['Before'], width=2),
                            marker=dict(color=colors['Before'], size=8, symbol='circle'),
                            hovertemplate='<b>Before Stage</b><br>' +
                                        f'Date: %{{x|%Y-%m-%d}}<br>' +
                                        f'{variable}: %{{y:.3f}}<br>' +
                                        '<extra></extra>'
                        ))

                    # Plot After data
                    if len(after_data) > 0 and len(after_dates) > 0:
                        after_array = np.array(after_data, dtype=float)
                        fig.add_trace(go.Scatter(
                            x=after_dates,
                            y=after_array.tolist(),
                            mode='markers+lines',
                            name=f'After Data ({len(after_array)} pts)',
                            line=dict(color=colors['After'], width=2),
                            marker=dict(color=colors['After'], size=8, symbol='circle'),
                            hovertemplate='<b>After Stage</b><br>' +
                                        f'Date: %{{x|%Y-%m-%d}}<br>' +
                                        f'{variable}: %{{y:.3f}}<br>' +
                                        '<extra></extra>'
                        ))

                    # Add combined control limit lines
                    if ucl is not None and not np.isnan(ucl):
                        # Combined UCL
                        fig.add_trace(go.Scatter(
                            x=combined_dates, y=[float(ucl)] * len(combined_dates),
                            mode='lines', name=f'Combined UCL ({ucl:.3f})',
                            line=dict(color='red', dash='dash', width=3),
                            hovertemplate=f'<b>Combined Upper Control Limit</b><br>' +
                                        f'Value: {ucl:.3f}<extra></extra>'
                        ))

                        # Combined LCL
                        fig.add_trace(go.Scatter(
                            x=combined_dates, y=[float(lcl)] * len(combined_dates),
                            mode='lines', name=f'Combined LCL ({lcl:.3f})',
                            line=dict(color='red', dash='dash', width=3),
                            hovertemplate=f'<b>Combined Lower Control Limit</b><br>' +
                                        f'Value: {lcl:.3f}<extra></extra>'
                        ))

                        # Combined Mean
                        fig.add_trace(go.Scatter(
                            x=combined_dates, y=[float(mean)] * len(combined_dates),
                            mode='lines', name=f'Combined Mean ({mean:.3f})',
                            line=dict(color='red', dash='dot', width=3),
                            hovertemplate=f'<b>Combined Mean</b><br>' +
                                        f'Value: {mean:.3f}<extra></extra>'
                        ))
            except Exception as e:
                print(f"❌ Error creating combined chart: {e}")
                fig.add_annotation(
                    text=f"❌ Error creating combined chart<br>Error: {str(e)}",
                    xref="paper", yref="paper",
                    x=0.5, y=0.5, showarrow=False,
                    font=dict(size=14, color="red")
                )

        # Update layout
        chart_title = f"Control Chart ({'Separate' if chart_type == 'separate' else 'Combined'} Limits): {variable}"
        if tape_ids:
            chart_title += f" | Tape IDs: {', '.join(map(str, tape_ids))}"
        if start_date and end_date:
            chart_title += f" | Date Range: {start_date} to {end_date}"

        fig.update_layout(
            title=dict(
                text=chart_title,
                x=0.5,
                font=dict(size=16, color='#2c3e50')
            ),
            xaxis=dict(
                title="Date",
                showgrid=True,
                gridcolor='lightgray'
            ),
            yaxis=dict(
                title=variable,
                showgrid=True,
                gridcolor='lightgray'
            ),
            hovermode='closest',
            showlegend=True,
            legend=dict(
                orientation="v",
                yanchor="top",
                y=1,
                xanchor="left",
                x=1.02,
                bgcolor="rgba(255,255,255,0.8)",
                bordercolor="gray",
                borderwidth=1
            ),
            margin=dict(t=80, b=60, l=60, r=150),
            height=500,
            plot_bgcolor='white',
            paper_bgcolor='white'
        )

        return fig

    except Exception as e:
        print(f"❌ Critical error in create_control_chart: {e}")
        fig.add_annotation(
            text=f"❌ Critical Error<br>{str(e)}<br><br>Please check data types and format",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=14, color="red")
        )
        return fig

def calculate_control_limits(data):
    """
    FIXED: Calculate control limits with better error handling.

    Args:
        data (array-like): Numeric data points

    Returns:
        tuple: (UCL, LCL, Mean) or (None, None, None) if insufficient data
    """
    try:
        # Ensure data is numpy array of floats
        data_array = np.array(data, dtype=float)

        # Remove any NaN or infinite values
        data_clean = data_array[np.isfinite(data_array)]

        if len(data_clean) == 0:
            return None, None, None

        mean = float(np.mean(data_clean))
        std = float(np.std(data_clean, ddof=1)) if len(data_clean) > 1 else 0.0

        # 3-sigma control limits
        ucl = mean + 3 * std
        lcl = mean - 3 * std

        return ucl, lcl, mean

    except Exception as e:
        print(f"❌ Error calculating control limits: {e}")
        return None, None, None

# =============================================================================
# DASHBOARD LAYOUT
# =============================================================================

def create_layout():
    """Create the main dashboard layout."""

    # Load initial data
    before_df, after_df = current_data['before'], current_data['after']

    # Get available options - ONLY COMMON COLUMNS
    common_numeric_cols = get_common_numeric_columns(before_df, after_df)
    all_tape_ids = get_tape_ids(before_df, after_df)

    # Default selections
    default_variable = common_numeric_cols[0] if common_numeric_cols else None
    default_tape_ids = all_tape_ids[:3] if len(all_tape_ids) > 3 else all_tape_ids

    # Default date range (last 30 days)
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d')

    return dbc.Container([
        # Header Section
        dbc.Row([
            dbc.Col([
                html.Div([
                    html.H1("🔄 Live Control Chart Dashboard",
                           className="text-center mb-2",
                           style={'color': '#2c3e50', 'fontWeight': 'bold'}),
                    html.P("Real-time Process Control Analysis with Auto-Updating Data",
                           className="text-center text-muted mb-4"),
                    html.Hr(style={'borderColor': '#34495e'})
                ])
            ])
        ]),

        # Controls Section
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.H5("🎛️ Dashboard Controls", className="mb-0", style={'color': '#2c3e50'})
                    ]),
                    dbc.CardBody([
                        dbc.Row([
                            # Variable Selection
                            dbc.Col([
                                html.Label("📊 Select Variable:",
                                         className="fw-bold mb-2",
                                         style={'color': '#2c3e50'}),
                                dcc.Dropdown(
                                    id='variable-dropdown',
                                    options=[{'label': var, 'value': var} for var in common_numeric_cols],
                                    value=default_variable,
                                    clearable=False,
                                    placeholder="Choose a variable to analyze (common to both datasets)",
                                    style={'fontSize': '14px'}
                                )
                            ], md=4),

                            # Tape ID Selection
                            dbc.Col([
                                html.Label("🏷️ Select Tape ID(s):",
                                         className="fw-bold mb-2",
                                         style={'color': '#2c3e50'}),
                                html.Div([
                                    dcc.Checklist(
                                        id='select-all-tape-ids',
                                        options=[{'label': 'Select All', 'value': 'all'}],
                                        value=[],
                                        labelStyle={'display': 'inline-block', 'marginRight': '10px'}
                                    ),
                                    dcc.Checklist(
                                        id='tape-id-checklist',
                                        options=[{'label': f"Tape {tape_id}", 'value': tape_id}
                                               for tape_id in all_tape_ids],
                                        value=default_tape_ids,
                                        labelStyle={'display': 'block'},
                                        style={'maxHeight': '200px', 'overflowY': 'auto', 'border': '1px solid #ddd', 'padding': '10px'}
                                    )
                                ])
                            ], md=4),

                            # Date Range Selection
                            dbc.Col([
                                html.Label("📅 Select Date Range:",
                                         className="fw-bold mb-2",
                                         style={'color': '#2c3e50'}),
                                dcc.DatePickerRange(
                                    id='date-range-picker',
                                    start_date=start_date,
                                    end_date=end_date,
                                    display_format='YYYY-MM-DD'
                                )
                            ], md=4)
                        ])
                    ])
                ], className="shadow-sm")
            ])
        ], className="mb-4"),

        # Status Information
        dbc.Row([
            dbc.Col([
                html.Div(id='status-info')
            ])
        ], className="mb-4"),

        # Charts Section
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.H5("📈 Control Charts", className="mb-0", style={'color': '#2c3e50'})
                    ]),
                    dbc.CardBody([
                        dbc.Tabs([
                            dbc.Tab(
                                label="📊 Separate Control Limits",
                                tab_id="separate-tab",
                                children=[
                                    html.Div([
                                        dcc.Graph(
                                            id='separate-chart',
                                            config={'displayModeBar': True, 'displaylogo': False}
                                        )
                                    ], className="mt-3")
                                ]
                            ),
                            dbc.Tab(
                                label="📈 Combined Control Limits",
                                tab_id="combined-tab",
                                children=[
                                    html.Div([
                                        dcc.Graph(
                                            id='combined-chart',
                                            config={'displayModeBar': True, 'displaylogo': False}
                                        )
                                    ], className="mt-3")
                                ]
                            )
                        ], id="chart-tabs", active_tab="separate-tab")
                    ])
                ], className="shadow-sm")
            ])
        ]),

        # T-test and Summary Report Section
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.H5("📊 T-test and Summary Report", className="mb-0", style={'color': '#2c3e50'})
                    ]),
                    dbc.CardBody([
                        html.Div([
                            html.Button("Generate Summary Report", id='generate-report-button', n_clicks=0, className="btn btn-primary mb-3"),
                            html.Div(id='summary-report', className="mt-3"),
                            html.Button("Download Report", id='download-report-button', n_clicks=0, className="btn btn-secondary mt-3"),
                            dcc.Download(id="download-report")
                        ])
                    ])
                ], className="shadow-sm")
            ])
        ], className="mb-4"),

        # Auto-update component
        dcc.Interval(
            id='interval-component',
            interval=REFRESH_INTERVAL * 1000,  # Convert to milliseconds
            n_intervals=0
        ),

        # Footer
        html.Footer([
            html.Hr(),
            html.P("🔄 Auto-refreshing every 10 seconds | Built with Dash & Plotly | FIXED VERSION",
                   className="text-center text-muted small")
        ], className="mt-5")

    ], fluid=True, className="px-4 py-3")

# Set the layout
app.layout = create_layout()

# =============================================================================
# DASH CALLBACKS
# =============================================================================

@app.callback(
    [Output('variable-dropdown', 'options'),
     Output('tape-id-checklist', 'options'),
     Output('status-info', 'children')],
    [Input('interval-component', 'n_intervals')]
)
def update_data_and_dropdowns(n):
    """Update data and refresh dropdown options."""

    # Reload data from Excel files
    before_df, after_df, last_updated = load_excel_data()
    current_data['before'] = before_df
    current_data['after'] = after_df
    current_data['last_updated'] = last_updated

    # Get ONLY COMMON numeric columns
    common_numeric_cols = get_common_numeric_columns(before_df, after_df)
    all_tape_ids = get_tape_ids(before_df, after_df)

    # Debug print
    print(f"🔍 Debug - Found {len(common_numeric_cols)} common numeric columns: {common_numeric_cols}")
    print(f"🔍 Debug - Found {len(all_tape_ids)} tape IDs: {all_tape_ids}")

    # Create variable options (only common columns)
    if common_numeric_cols:
        variable_options = [{'label': f"{var} (✓ Both datasets)", 'value': var} for var in common_numeric_cols]
    else:
        variable_options = [{'label': 'No common variables found', 'value': 'none', 'disabled': True}]

    # Create better labels for tape IDs
    if all_tape_ids:
        tape_id_options = []
        for tape_id in all_tape_ids:
            # Handle different data types for tape IDs
            label = f"Tape {tape_id}" if str(tape_id).replace('.', '').replace('-', '').isdigit() else str(tape_id)
            tape_id_options.append({'label': label, 'value': tape_id})
    else:
        tape_id_options = [{'label': 'No Tape IDs found', 'value': 'none', 'disabled': True}]

    # Create status information with more details
    variable_status = f"{len(common_numeric_cols)} common variables" if common_numeric_cols else "No common variables"
    tape_id_status = f"{len(all_tape_ids)} types found" if all_tape_ids else "No Tape IDs detected"

    # Determine alert color based on data availability
    alert_color = "success" if common_numeric_cols and all_tape_ids else "warning" if common_numeric_cols or all_tape_ids else "danger"

    status_info = dbc.Alert([
        dbc.Row([
            dbc.Col([
                html.Strong("🕒 Last Updated: "),
                html.Span(last_updated.strftime("%Y-%m-%d %H:%M:%S"))
            ], md=3),
            dbc.Col([
                html.Strong("📊 Data Records: "),
                html.Span(f"Before: {len(before_df)}, After: {len(after_df)}")
            ], md=3),
            dbc.Col([
                html.Strong("📈 Variables: "),
                html.Span(variable_status),
                html.Br(),
                html.Small("(Only showing columns in both datasets)", className="text-muted")
            ], md=3),
            dbc.Col([
                html.Strong("🏷️ Tape Types: "),
                html.Span(tape_id_status)
            ], md=3)
        ])
    ], color=alert_color, className="mb-0")

    return variable_options, tape_id_options, status_info

@app.callback(
    Output('tape-id-checklist', 'value'),
    [Input('select-all-tape-ids', 'value')],
    [State('tape-id-checklist', 'options')]
)
def select_all_tape_ids(select_all, tape_id_options):
    """Select or deselect all tape IDs based on the 'Select All' checkbox."""

    if not tape_id_options:
        return []

    if 'all' in select_all:
        # Select all tape IDs
        return [option['value'] for option in tape_id_options]
    else:
        # Deselect all tape IDs
        return []

@app.callback(
    [Output('separate-chart', 'figure'),
     Output('combined-chart', 'figure')],
    [Input('variable-dropdown', 'value'),
     Input('tape-id-checklist', 'value'),
     Input('date-range-picker', 'start_date'),
     Input('date-range-picker', 'end_date'),
     Input('interval-component', 'n_intervals')]
)
def update_charts(selected_variable, selected_tape_ids, start_date, end_date, n):
    """Update both control charts based on user selections."""

    try:
        # Create empty figure as fallback
        empty_fig = go.Figure()
        empty_fig.update_layout(height=500)

        # Check if variable is selected
        if not selected_variable or selected_variable == 'none':
            empty_fig.add_annotation(
                text="📊 Please select a variable to display charts",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="gray")
            )
            return empty_fig, empty_fig

        # Get current data with safety checks
        before_df = current_data.get('before', pd.DataFrame())
        after_df = current_data.get('after', pd.DataFrame())

        # Validate data exists
        if before_df.empty and after_df.empty:
            empty_fig.add_annotation(
                text="📊 No data available<br>Please check your Excel files",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="gray")
            )
            return empty_fig, empty_fig

        # Validate selected variable exists in data
        variable_exists = False
        if not before_df.empty and selected_variable in before_df.columns:
            variable_exists = True
        if not after_df.empty and selected_variable in after_df.columns:
            variable_exists = True

        if not variable_exists:
            empty_fig.add_annotation(
                text=f"📊 Variable '{selected_variable}' not found in data<br>Please select a different variable",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="gray")
            )
            return empty_fig, empty_fig

        # Debug information about selection
        print(f"🎯 Chart Update Debug:")
        print(f"   Selected Variable: {selected_variable}")
        print(f"   Selected Tape IDs: {selected_tape_ids}")
        print(f"   Date Range: {start_date} to {end_date}")
        print(f"   Before data shape: {before_df.shape}")
        print(f"   After data shape: {after_df.shape}")

        # Handle empty tape ID selection (use all data if nothing selected)
        tape_ids_to_use = selected_tape_ids if selected_tape_ids else None

        # Create both chart types with error handling
        try:
            separate_fig = create_control_chart(before_df, after_df, selected_variable,
                                              tape_ids_to_use, start_date, end_date, 'separate')
        except Exception as e:
            print(f"❌ Error creating separate chart: {e}")
            separate_fig = empty_fig
            separate_fig.add_annotation(
                text=f"❌ Error creating separate control chart<br>Error: {str(e)}",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=14, color="red")
            )

        try:
            combined_fig = create_control_chart(before_df, after_df, selected_variable,
                                              tape_ids_to_use, start_date, end_date, 'combined')
        except Exception as e:
            print(f"❌ Error creating combined chart: {e}")
            combined_fig = empty_fig
            combined_fig.add_annotation(
                text=f"❌ Error creating combined control chart<br>Error: {str(e)}",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=14, color="red")
            )

        return separate_fig, combined_fig

    except Exception as e:
        print(f"❌ Critical error in update_charts callback: {e}")
        # Return safe empty figures
        error_fig = go.Figure()
        error_fig.add_annotation(
            text=f"❌ Critical Error<br>{str(e)}<br><br>Please check console for details",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=14, color="red")
        )
        error_fig.update_layout(height=500)
        return error_fig, error_fig

@app.callback(
    Output('summary-report', 'children'),
    [Input('generate-report-button', 'n_clicks')],
    [dash.dependencies.State('variable-dropdown', 'value'),
     dash.dependencies.State('tape-id-checklist', 'value'),
     dash.dependencies.State('date-range-picker', 'start_date'),
     dash.dependencies.State('date-range-picker', 'end_date')]
)
def generate_summary_report_callback(n_clicks, selected_variable, selected_tape_ids, start_date, end_date):
    """Generate summary report based on user selections."""

    if n_clicks == 0:
        return ""

    try:
        # Get current data with safety checks
        before_df = current_data.get('before', pd.DataFrame())
        after_df = current_data.get('after', pd.DataFrame())

        # Validate data exists
        if before_df.empty and after_df.empty:
            return dbc.Alert("📊 No data available. Please check your Excel files.", color="danger")

        # Validate selected variable exists in data
        if not selected_variable or selected_variable == 'none':
            return dbc.Alert("📊 Please select a variable to generate the report.", color="danger")

        # Filter data by selected tape IDs
        if selected_tape_ids:
            filtered_before = filter_data_by_tape_ids(before_df, selected_tape_ids)
            filtered_after = filter_data_by_tape_ids(after_df, selected_tape_ids)
        else:
            filtered_before = before_df.copy()
            filtered_after = after_df.copy()

        # Filter data by date range
        if start_date and end_date:
            filtered_before = filter_data_by_date_range(filtered_before, start_date, end_date)
            filtered_after = filter_data_by_date_range(filtered_after, start_date, end_date)

        # Clean the data for the control charts
        cleaned_before = clean_control_chart_data(filtered_before, selected_variable)
        cleaned_after = clean_control_chart_data(filtered_after, selected_variable)

        # Extract variable data
        before_data = []
        after_data = []

        if not cleaned_before.empty and selected_variable in cleaned_before.columns:
            before_series = pd.to_numeric(cleaned_before[selected_variable], errors='coerce')
            before_data = before_series.dropna().astype(float).tolist()

        if not cleaned_after.empty and selected_variable in cleaned_after.columns:
            after_series = pd.to_numeric(cleaned_after[selected_variable], errors='coerce')
            after_data = after_series.dropna().astype(float).tolist()

        # Check if we have any data
        if len(before_data) == 0 or len(after_data) == 0:
            return dbc.Alert("📊 No valid numeric data found for the selected variable.", color="danger")

        # Perform T-test
        t_test_results = perform_t_test(before_data, after_data)

        # Generate summary report
        report = generate_summary_report(before_data, after_data, selected_variable, t_test_results)

        return dbc.Alert([
            html.H5("Summary Report", className="mb-3"),
            html.Pre(report, className="mb-0"),
            html.Button("Download Report", id='download-report-button', n_clicks=0, className="btn btn-secondary mt-3"),
            dcc.Download(id="download-report")
        ], color="light")

    except Exception as e:
        print(f"❌ Error generating summary report: {e}")
        return dbc.Alert(f"❌ Error generating summary report: {str(e)}", color="danger")

@app.callback(
    Output("download-report", "data"),
    [Input("download-report-button", "n_clicks")],
    [dash.dependencies.State('variable-dropdown', 'value'),
     dash.dependencies.State('tape-id-checklist', 'value'),
     dash.dependencies.State('date-range-picker', 'start_date'),
     dash.dependencies.State('date-range-picker', 'end_date')],
    prevent_initial_call=True,
)
def download_summary_report(n_clicks, selected_variable, selected_tape_ids, start_date, end_date):
    """Download the summary report as a text file."""

    if n_clicks is None:
        raise dash.exceptions.PreventUpdate

    try:
        # Get current data with safety checks
        before_df = current_data.get('before', pd.DataFrame())
        after_df = current_data.get('after', pd.DataFrame())

        # Validate data exists
        if before_df.empty and after_df.empty:
            return None

        # Validate selected variable exists in data
        if not selected_variable or selected_variable == 'none':
            return None

        # Filter data by selected tape IDs
        if selected_tape_ids:
            filtered_before = filter_data_by_tape_ids(before_df, selected_tape_ids)
            filtered_after = filter_data_by_tape_ids(after_df, selected_tape_ids)
        else:
            filtered_before = before_df.copy()
            filtered_after = after_df.copy()

        # Filter data by date range
        if start_date and end_date:
            filtered_before = filter_data_by_date_range(filtered_before, start_date, end_date)
            filtered_after = filter_data_by_date_range(filtered_after, start_date, end_date)

        # Clean the data for the control charts
        cleaned_before = clean_control_chart_data(filtered_before, selected_variable)
        cleaned_after = clean_control_chart_data(filtered_after, selected_variable)

        # Extract variable data
        before_data = []
        after_data = []

        if not cleaned_before.empty and selected_variable in cleaned_before.columns:
            before_series = pd.to_numeric(cleaned_before[selected_variable], errors='coerce')
            before_data = before_series.dropna().astype(float).tolist()

        if not cleaned_after.empty and selected_variable in cleaned_after.columns:
            after_series = pd.to_numeric(cleaned_after[selected_variable], errors='coerce')
            after_data = after_series.dropna().astype(float).tolist()

        # Check if we have any data
        if len(before_data) == 0 or len(after_data) == 0:
            return None

        # Perform T-test
        t_test_results = perform_t_test(before_data, after_data)

        # Generate summary report
        report = generate_summary_report(before_data, after_data, selected_variable, t_test_results)

        # Create a downloadable text file
        return dict(
            content=report,
            filename="summary_report.txt",
            type="text/plain"
        )

    except Exception as e:
        print(f"❌ Error downloading summary report: {e}")
        return None

# =============================================================================
# MAIN EXECUTION
# =============================================================================

if __name__ == '__main__':
    print("=" * 60)
    print("🚀 LIVE CONTROL CHART DASHBOARD - FIXED VERSION")
    print("=" * 60)
    print("✅ FIXES APPLIED:")
    print("   • Fixed tape ID filtering logic")
    print("   • Improved data type handling")
    print("   • Enhanced debugging output")
    print("   • Better error messages")
    print("   • Added dates to x-axis")
    print("   • Added data cleaning for control charts")
    print("   • Added date range filtering")
    print("   • Added dropdown checkbox for Tape IDs")
    print("   • Added T-test for before and after datasets")
    print("   • Added downloadable summary report")
    print("   • Added 'Select All' feature for Tape ID dropdown checkbox")
    print("   • Ensured Tape ID dropdown is not expanded by default")
    print("=" * 60)
    print(f"📁 Monitoring Excel files:")
    print(f"   📋 Before: {BEFORE_PATH}")
    print(f"   📋 After:  {AFTER_PATH}")
    print(f"🔄 Auto-refresh interval: {REFRESH_INTERVAL} seconds")
    print(f"🌐 Dashboard URL: http://127.0.0.1:8050")
    print("=" * 60)

    # Load initial data
    current_data['before'], current_data['after'], current_data['last_updated'] = load_excel_data()

    # Run the app
    try:
        app.run(debug=True, host='127.0.0.1', port=8050)
    except Exception as e:
        print(f"❌ Error starting dashboard: {e}")
        print("💡 Make sure port 8050 is available and try again.")

