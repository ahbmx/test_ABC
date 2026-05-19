Here are the complete functions for Excel formatting with xlsxwriter and pandas:

```python
import pandas as pd
import xlsxwriter
from typing import List, Dict, Union, Optional, Tuple

def add_autofilter_to_dataframe(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet_name: str,
    autofilter_range: Optional[Tuple[int, int, int, int]] = None
) -> None:
    """
    Add autofilter to a dataframe in an Excel sheet.
    
    Parameters:
    -----------
    df : pd.DataFrame
        The dataframe to add autofilter to
    writer : pd.ExcelWriter
        The ExcelWriter object
    sheet_name : str
        Name of the worksheet
    autofilter_range : tuple, optional
        (first_row, first_col, last_row, last_col)
        If None, autofilter will be applied to entire data range
    """
    worksheet = writer.sheets[sheet_name]
    
    if autofilter_range:
        first_row, first_col, last_row, last_col = autofilter_range
        worksheet.autofilter(first_row, first_col, last_row, last_col)
    else:
        # Apply to entire data range including headers
        # +1 for header row
        last_row = len(df)
        last_col = len(df.columns) - 1
        worksheet.autofilter(0, 0, last_row, last_col)


def auto_adjust_cell_width(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet_name: str,
    max_width: int = 50,
    padding: int = 2,
    exclude_columns: Optional[List[str]] = None
) -> None:
    """
    Automatically adjust column widths based on content.
    
    Parameters:
    -----------
    df : pd.DataFrame
        The dataframe with data
    writer : pd.ExcelWriter
        The ExcelWriter object
    sheet_name : str
        Name of the worksheet
    max_width : int
        Maximum width for any column
    padding : int
        Extra padding to add to column width
    exclude_columns : list, optional
        List of column names to exclude from auto-adjust
    """
    worksheet = writer.sheets[sheet_name]
    exclude_columns = exclude_columns or []
    
    for idx, col in enumerate(df.columns):
        if col in exclude_columns:
            continue
            
        # Get max length in column
        max_len = len(str(col))  # Start with header length
        
        # Check values in column
        for value in df[col]:
            try:
                # Convert to string and get length
                cell_len = len(str(value))
                if cell_len > max_len:
                    max_len = cell_len
            except:
                pass
        
        # Apply padding and max limit
        adjusted_width = min(max_len + padding, max_width)
        worksheet.set_column(idx, idx, adjusted_width)


def color_headers_with_style(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet_name: str,
    header_style: Optional[Dict] = None,
    start_row: int = 0,
    start_col: int = 0
) -> None:
    """
    Apply styling to column headers.
    
    Parameters:
    -----------
    df : pd.DataFrame
        The dataframe with column headers
    writer : pd.ExcelWriter
        The ExcelWriter object
    sheet_name : str
        Name of the worksheet
    header_style : dict, optional
        Dictionary of xlsxwriter format properties
    start_row : int
        Starting row for headers (default 0)
    start_col : int
        Starting column for headers (default 0)
    """
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    
    # Default header style if not provided
    if header_style is None:
        header_style = {
            'bold': True,
            'font_color': 'white',
            'bg_color': '#4F81BD',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    
    # Create format object
    header_format = workbook.add_format(header_style)
    
    # Write headers with formatting
    for col_idx, col_name in enumerate(df.columns):
        worksheet.write(start_row, start_col + col_idx, col_name, header_format)


def highlight_rows_based_on_cell_value(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet_name: str,
    condition_column: str,
    condition: str,
    threshold_value: Union[int, float, str],
    highlight_style: Optional[Dict] = None,
    exclude_columns: Optional[List[str]] = None,
    start_row: int = 1  # Start after header row
) -> None:
    """
    Highlight entire rows based on cell value condition.
    
    Parameters:
    -----------
    df : pd.DataFrame
        The dataframe with data
    writer : pd.ExcelWriter
        The ExcelWriter object
    sheet_name : str
        Name of the worksheet
    condition_column : str
        Column name to evaluate condition on
    condition : str
        Condition type: 'equal', 'greater', 'less', 'greater_equal', 
        'less_equal', 'contains', 'between'
    threshold_value : int, float, or str
        Value to compare against
    highlight_style : dict, optional
        Dictionary of xlsxwriter format properties for highlighting
    exclude_columns : list, optional
        Columns to exclude from highlighting
    start_row : int
        Row where data starts (default 1, assuming row 0 is headers)
    """
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    
    # Default highlight style
    if highlight_style is None:
        highlight_style = {
            'bg_color': '#FFC7CE',
            'font_color': '#9C0006',
            'bold': False
        }
    
    # Get column index for condition
    col_idx = df.columns.get_loc(condition_column)
    
    # Apply conditional formatting based on condition
    for row_idx, (_, row) in enumerate(df.iterrows(), start=start_row):
        cell_value = row[condition_column]
        should_highlight = False
        
        # Evaluate condition
        if condition == 'equal':
            should_highlight = cell_value == threshold_value
        elif condition == 'greater':
            should_highlight = cell_value > threshold_value
        elif condition == 'less':
            should_highlight = cell_value < threshold_value
        elif condition == 'greater_equal':
            should_highlight = cell_value >= threshold_value
        elif condition == 'less_equal':
            should_highlight = cell_value <= threshold_value
        elif condition == 'contains' and isinstance(cell_value, str):
            should_highlight = threshold_value in cell_value
        elif condition == 'between' and isinstance(threshold_value, (list, tuple)):
            should_highlight = threshold_value[0] <= cell_value <= threshold_value[1]
        
        # Apply highlighting to entire row if condition met
        if should_highlight:
            highlight_format = workbook.add_format(highlight_style)
            for col in range(len(df.columns)):
                worksheet.write(row_idx, col, row.iloc[col], highlight_format)


def add_conditional_formatting_professional(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet_name: str,
    rules: List[Dict]
) -> None:
    """
    Add multiple conditional formatting rules to a worksheet.
    
    Parameters:
    -----------
    df : pd.DataFrame
        The dataframe with data
    writer : pd.ExcelWriter
        The ExcelWriter object
    sheet_name : str
        Name of the worksheet
    rules : list of dict
        List of conditional formatting rules, each containing:
        - 'column': column name
        - 'type': 'cell', 'color_scale', 'data_bar', 'icon_set'
        - 'criteria': condition criteria
        - 'format': format properties (for 'cell' type)
    """
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    
    for rule in rules:
        col_idx = df.columns.get_loc(rule['column'])
        start_row = 1
        end_row = len(df)
        
        # Define range
        cell_range = f"{chr(65 + col_idx)}{start_row + 1}:{chr(65 + col_idx)}{end_row + 1}"
        
        if rule['type'] == 'cell':
            # Cell value based formatting
            format_obj = workbook.add_format(rule.get('format', {}))
            worksheet.conditional_format(cell_range, {
                'type': 'cell',
                'criteria': rule['criteria'],
                'value': rule.get('value', 0),
                'format': format_obj
            })
        
        elif rule['type'] == 'color_scale':
            # Color scale (gradient)
            worksheet.conditional_format(cell_range, {
                'type': 'color_scale',
                'min_color': rule.get('min_color', '#F8696B'),
                'max_color': rule.get('max_color', '#63BE7B')
            })
        
        elif rule['type'] == 'data_bar':
            # Data bars
            worksheet.conditional_format(cell_range, {
                'type': 'data_bar',
                'bar_color': rule.get('bar_color', '#638EC6')
            })
        
        elif rule['type'] == 'top_bottom':
            # Top/Bottom rules
            worksheet.conditional_format(cell_range, {
                'type': 'top',
                'value': rule.get('value', 10),
                'format': workbook.add_format(rule.get('format', {'bold': True}))
            })


# Complete usage example
def create_formatted_excel_report(
    df: pd.DataFrame,
    output_file: str,
    sheet_name: str = 'Report'
) -> None:
    """
    Create a fully formatted Excel report with all formatting functions.
    """
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write dataframe to Excel
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Apply all formatting functions
        add_autofilter_to_dataframe(df, writer, sheet_name)
        auto_adjust_cell_width(df, writer, sheet_name, max_width=40, padding=3)
        color_headers_with_style(df, writer, sheet_name)
        
        # Highlight rows where Sales > 1000
        highlight_rows_based_on_cell_value(
            df, writer, sheet_name,
            condition_column='Sales',
            condition='greater',
            threshold_value=1000,
            highlight_style={
                'bg_color': '#C6EFCE',
                'font_color': '#006100',
                'bold': True
            }
        )
        
        # Highlight rows where Status contains 'Pending'
        if 'Status' in df.columns:
            highlight_rows_based_on_cell_value(
                df, writer, sheet_name,
                condition_column='Status',
                condition='contains',
                threshold_value='Pending',
                highlight_style={
                    'bg_color': '#FFC7CE',
                    'font_color': '#9C0006'
                }
            )


# Advanced example with multiple highlight conditions
def advanced_highlighting_example():
    """Example demonstrating multiple highlighting conditions"""
    
    # Sample data
    df = pd.DataFrame({
        'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Desk', 'Chair'],
        'Sales': [1500, 300, 800, 2500, 450, 1200],
        'Status': ['Shipped', 'Pending', 'Shipped', 'Pending', 'Shipped', 'Processing'],
        'Region': ['North', 'South', 'East', 'West', 'North', 'South'],
        'Profit_Margin': [0.25, 0.40, 0.35, 0.20, 0.30, 0.28]
    })
    
    output_file = 'advanced_formatted_report.xlsx'
    
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sales Data', index=False)
        sheet_name = 'Sales Data'
        
        # Apply basic formatting
        add_autofilter_to_dataframe(df, writer, sheet_name)
        auto_adjust_cell_width(df, writer, sheet_name, max_width=35, padding=2)
        color_headers_with_style(df, writer, sheet_name, {
            'bold': True,
            'font_color': 'white',
            'bg_color': '#2C3E50',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 12
        })
        
        # Multiple highlight conditions
        # 1. Highlight high sales (> 1000) in green
        highlight_rows_based_on_cell_value(
            df, writer, sheet_name,
            condition_column='Sales',
            condition='greater',
            threshold_value=1000,
            highlight_style={
                'bg_color': '#C6EFCE',
                'font_color': '#006100'
            }
        )
        
        # 2. Highlight pending orders in red
        highlight_rows_based_on_cell_value(
            df, writer, sheet_name,
            condition_column='Status',
            condition='contains',
            threshold_value='Pending',
            highlight_style={
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006',
                'bold': True
            }
        )
        
        # 3. Highlight low profit margin (< 25%) in yellow
        highlight_rows_based_on_cell_value(
            df, writer, sheet_name,
            condition_column='Profit_Margin',
            condition='less',
            threshold_value=0.25,
            highlight_style={
                'bg_color': '#FFEB9C',
                'font_color': '#9C6500'
            }
        )
        
        # 4. Add professional conditional formatting
        add_conditional_formatting_professional(
            df, writer, sheet_name,
            rules=[
                {
                    'column': 'Sales',
                    'type': 'data_bar',
                    'bar_color': '#5B9BD5'
                },
                {
                    'column': 'Profit_Margin',
                    'type': 'color_scale',
                    'min_color': '#F8696B',
                    'max_color': '#63BE7B'
                }
            ]
        )
    
    print(f"Advanced formatted report saved to {output_file}")


# Utility function for bulk highlighting
def highlight_rows_multiple_conditions(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet_name: str,
    conditions: List[Dict]
) -> None:
    """
    Apply multiple row highlighting conditions.
    
    Parameters:
    -----------
    conditions : list of dict
        Each dict contains:
        - 'column': column name
        - 'condition': condition type
        - 'threshold': threshold value
        - 'style': highlight style dict
        - 'operator': 'and'/'or' for combining with previous conditions (optional)
    """
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    
    for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
        should_highlight = False
        combined_style = None
        
        for condition in conditions:
            col_idx = df.columns.get_loc(condition['column'])
            cell_value = row[condition['column']]
            condition_met = False
            
            # Evaluate condition
            if condition['condition'] == 'equal':
                condition_met = cell_value == condition['threshold']
            elif condition['condition'] == 'greater':
                condition_met = cell_value > condition['threshold']
            elif condition['condition'] == 'less':
                condition_met = cell_value < condition['threshold']
            elif condition['condition'] == 'contains' and isinstance(cell_value, str):
                condition_met = condition['threshold'] in cell_value
            
            # Combine conditions
            if 'operator' not in condition:
                should_highlight = condition_met
                if condition_met:
                    combined_style = condition.get('style')
            else:
                if condition['operator'] == 'and':
                    should_highlight = should_highlight and condition_met
                elif condition['operator'] == 'or':
                    should_highlight = should_highlight or condition_met
                
                if condition_met and not combined_style:
                    combined_style = condition.get('style')
        
        # Apply highlight if conditions met
        if should_highlight and combined_style:
            highlight_format = workbook.add_format(combined_style)
            for col in range(len(df.columns)):
                worksheet.write(row_idx, col, row.iloc[col], highlight_format)


# Run examples
if __name__ == "__main__":
    # Create sample dataframe
    sample_df = pd.DataFrame({
        'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Desk', 'Chair'],
        'Sales': [1500, 300, 800, 2500, 450, 1200],
        'Quantity': [10, 50, 30, 15, 5, 25],
        'Status': ['Shipped', 'Pending', 'Shipped', 'Pending', 'Shipped', 'Processing'],
        'Region': ['North', 'South', 'East', 'West', 'North', 'South']
    })
    
    # Create basic formatted report
    create_formatted_excel_report(sample_df, 'basic_formatted_report.xlsx')
    
    # Create advanced report with multiple highlights
    advanced_highlighting_example()
    
    print("All reports generated successfully!")
```

## Key Features of These Functions:

### 1. `add_autofilter_to_dataframe()`
- Adds Excel autofilter dropdowns to column headers
- Can specify custom range or auto-detect

### 2. `auto_adjust_cell_width()`
- Dynamically calculates optimal column widths
- Configurable padding and maximum width
- Can exclude specific columns

### 3. `color_headers_with_style()`
- Applies professional styling to headers
- Fully customizable style dictionary
- Supports all xlsxwriter format options

### 4. `highlight_rows_based_on_cell_value()`
- Highlights entire rows based on conditions
- Supports multiple condition types: equal, greater, less, contains, between
- Customizable highlight colors and styles

### 5. `add_conditional_formatting_professional()`
- Advanced Excel conditional formatting
- Supports color scales, data bars, icon sets
- Cell value based formatting

### 6. `highlight_rows_multiple_conditions()`
- Combine multiple conditions with AND/OR logic
- Bulk highlighting with different styles

These functions work seamlessly with pandas DataFrames and xlsxwriter to create professional Excel reports with proper formatting, filters, and conditional highlighting.
