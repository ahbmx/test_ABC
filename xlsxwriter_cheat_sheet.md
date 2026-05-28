# Comprehensive Python xlsxWriter with pandas Cheat Sheet

## Table of Contents
1. [Installation & Setup](#installation)
2. [Basic Workbook Operations](#basic-workbook)
3. [Working with Worksheets](#worksheets)
4. [Cell Formatting](#formatting)
5. [Writing Data](#writing-data)
6. [Charts](#charts)
7. [Formulas & Functions](#formulas)
8. [Conditional Formatting](#conditional-formatting)
9. [pandas Integration](#pandas)
10. [Advanced Features](#advanced)

---

## 1. Installation & Setup {#installation}

```python
# Installation
pip install xlsxwriter pandas

# Import libraries
import xlsxwriter
import pandas as pd
import numpy as np
```

## 2. Basic Workbook Operations {#basic-workbook}

```python
# Create new workbook
workbook = xlsxwriter.Workbook('output.xlsx')
workbook = xlsxwriter.Workbook('output.xlsx', {'strings_to_numbers': True})
workbook = xlsxwriter.Workbook('output.xlsx', {'constant_memory': True})  # For large files

# Access worksheet
worksheet = workbook.add_worksheet()  # Default "Sheet1"
worksheet = workbook.add_worksheet("Sales Data")
worksheet = workbook.add_worksheet()  # Creates "Sheet2", "Sheet3", etc.

# Close workbook (MUST DO)
workbook.close()

# Better practice: use context manager
with xlsxwriter.Workbook('output.xlsx') as workbook:
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Hello World')
```

## 3. Working with Worksheets {#worksheets}

```python
# Activate worksheet
worksheet.activate()

# Hide worksheet
worksheet.hide()

# Set worksheet as first tab
worksheet.set_first_sheet()

# Set tab color
worksheet.set_tab_color('red')

# Protect worksheet
worksheet.protect()
worksheet.protect('password', {'select_locked_cells': False})

# Set zoom
worksheet.set_zoom(150)

# Gridlines visibility
worksheet.hide_gridlines(2)  # 0=default, 1=hide screen, 2=hide printed

# Set page layout
worksheet.set_landscape()
worksheet.set_portrait()
worksheet.set_paper(9)  # 9 = A4
worksheet.set_margins(left=0.7, right=0.7, top=0.75, bottom=0.75)
worksheet.repeat_rows(0)  # Repeat first row at top of each page
worksheet.repeat_columns(0)  # Repeat first column on each page
worksheet.print_area('A1:G20')
worksheet.center_horizontally()
worksheet.center_vertically()
```

## 4. Cell Formatting {#formatting}

```python
# Create format objects
# Numeric formats
format_number = workbook.add_format({'num_format': '#,##0.00'})
format_percent = workbook.add_format({'num_format': '0.00%'})
format_currency = workbook.add_format({'num_format': '$#,##0.00'})
format_date = workbook.add_format({'num_format': 'yyyy-mm-dd'})
format_datetime = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})

# Text alignment
format_center = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
format_left = workbook.add_format({'align': 'left'})
format_right = workbook.add_format({'align': 'right'})
format_top = workbook.add_format({'valign': 'top'})
format_bottom = workbook.add_format({'valign': 'bottom'})

# Font formatting
format_bold = workbook.add_format({'bold': True})
format_italic = workbook.add_format({'italic': True})
format_underline = workbook.add_format({'underline': True})
format_font_color = workbook.add_format({'font_color': 'red'})
format_font = workbook.add_format({'font_name': 'Arial', 'font_size': 12})

# Borders
format_border = workbook.add_format({'border': 1})
format_border_thick = workbook.add_format({'border': 2})
format_border_bottom = workbook.add_format({'bottom': 1})
format_border_top = workbook.add_format({'top': 1})
format_border_left = workbook.add_format({'left': 1})
format_border_right = workbook.add_format({'right': 1})

# Fill/Background
format_fill = workbook.add_format({'bg_color': '#FFFF00'})
format_pattern = workbook.add_format({'pattern': 1, 'bg_color': '#FFC7CE'})
format_gradient = workbook.add_format({
    'bg_color': '#00FF00',
    'pattern': 1,
    'gradient': {'colors': ['#FF0000', '#0000FF']}
})

# Text wrapping
format_wrap = workbook.add_format({'text_wrap': True})
format_shrink = workbook.add_format({'shrink': True})

# Rotation
format_rotate = workbook.add_format({'rotation': 45})

# Merge formatting
format_header = workbook.add_format({
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'border': 1,
    'bg_color': '#D9E1F2',
    'font_size': 11
})

# Apply formatting
worksheet.write('A1', 'Header', format_header)
worksheet.write(0, 0, 'Value', format_bold)
```

## 5. Writing Data {#writing-data}

```python
# Write single values
worksheet.write('A1', 'Text')  # String
worksheet.write(0, 1, 42)  # Number (row, column)
worksheet.write(0, 2, 3.14159)  # Float
worksheet.write('A4', '=SUM(A2:A3)')  # Formula
worksheet.write(0, 4, '=B2+C2')  # Formula with coordinates
worksheet.write(0, 5, None)  # Blank
worksheet.write(0, 6, True)  # Boolean
worksheet.write_datetime(0, 7, datetime.datetime(2024, 1, 15), format_date)
worksheet.write_url(0, 8, 'https://github.com', string='GitHub')
worksheet.write_comment(0, 9, 'This is a comment')

# Write row/column
row = ['Name', 'Age', 'City']
col = ['Name', 'Alice', 'Bob', 'Charlie']
worksheet.write_row('A1', row)
worksheet.write_column('A1', col)

# Write multiple rows
data = [
    ['Name', 'Age', 'Salary'],
    ['Alice', 25, 50000],
    ['Bob', 30, 60000],
    ['Charlie', 35, 70000]
]
for r, row_data in enumerate(data):
    worksheet.write_row(r, 0, row_data)

# Set column widths
worksheet.set_column('A:A', 20)  # Single column
worksheet.set_column('B:D', 15)  # Multiple columns
worksheet.set_column('A:F', 12, format_center)  # With format
worksheet.set_column(0, 5, 15)  # Using indices

# Set row height
worksheet.set_row(0, 30)  # Row 0 height 30
worksheet.set_row(0, 30, format_header)  # With format

# Merge cells
worksheet.merge_range('A1:D1', 'Title', format_header)
worksheet.merge_range(0, 0, 0, 3, 'Title', format_header)  # (row, col, row_end, col_end)

# Insert image
worksheet.insert_image('B5', 'logo.png')
worksheet.insert_image('B5', 'logo.png', {'x_offset': 10, 'y_offset': 10})
worksheet.insert_image('B5', 'logo.png', {'x_scale': 0.5, 'y_scale': 0.5})

# Insert chart (see charts section)
worksheet.insert_chart('E2', chart)

# Add filters
worksheet.autofilter('A1:D100')
worksheet.filter_column('A', 'x > 5')
worksheet.filter_column('B', 'x == "Apple"')

# Data validation
worksheet.data_validation('B2', {
    'validate': 'integer',
    'criteria': 'between',
    'minimum': 1,
    'maximum': 100
})
worksheet.data_validation('C2', {
    'validate': 'list',
    'source': ['Yes', 'No', 'Maybe']
})
worksheet.data_validation('D2', {
    'validate': 'date',
    'criteria': '>',
    'value': datetime.date(2024, 1, 1)
})
```

## 6. Charts {#charts}

```python
# Create chart objects
chart_column = workbook.add_chart({'type': 'column'})
chart_bar = workbook.add_chart({'type': 'bar'})
chart_line = workbook.add_chart({'type': 'line'})
chart_pie = workbook.add_chart({'type': 'pie'})
chart_area = workbook.add_chart({'type': 'area'})
chart_scatter = workbook.add_chart({'type': 'scatter'})
chart_stock = workbook.add_chart({'type': 'stock'})
chart_radar = workbook.add_chart({'type': 'radar'})

# Basic column chart
chart_column.add_series({
    'name': 'Sales 2023',
    'categories': 'Sheet1!$A$2:$A$6',
    'values': 'Sheet1!$B$2:$B$6',
    'fill': {'color': '#5B9BD5'},
    'line': {'color': '#2E75B6', 'width': 2}
})

# Multiple series
chart_column.add_series({
    'name': 'Sales 2024',
    'categories': 'Sheet1!$A$2:$A$6',
    'values': 'Sheet1!$C$2:$C$6',
    'fill': {'color': '#ED7D31'}
})

# Chart title and axes
chart_column.set_title({'name': 'Sales Report'})
chart_column.set_x_axis({'name': 'Months'})
chart_column.set_y_axis({'name': 'Sales (USD)', 'min': 0, 'max': 100000})
chart_column.set_legend({'position': 'bottom'})

# Pie chart configuration
chart_pie.add_series({
    'name': 'Market Share',
    'categories': 'Sheet1!$A$2:$A$6',
    'values': 'Sheet1!$B$2:$B$6',
    'data_labels': {'percentage': True, 'category': True}
})

# Line chart with markers
chart_line.add_series({
    'name': 'Trend',
    'categories': 'Sheet1!$A$2:$A$20',
    'values': 'Sheet1!$B$2:$B$20',
    'marker': {'type': 'circle', 'size': 5},
    'trendline': {'type': 'linear', 'name': 'Linear Trend', 'display_equation': True}
})

# Combined chart
chart_combined = workbook.add_chart({'type': 'column'})
chart_combined.add_series({
    'name': 'Revenue',
    'values': '=Sheet1!$B$2:$B$7',
    'y2_axis': 0
})
chart_combined.add_series({
    'name': 'Profit Margin %',
    'values': '=Sheet1!$C$2:$C$7',
    'y2_axis': 1,
    'type': 'line'
})
chart_combined.set_y2_axis({'name': 'Percentage (%)'})
```

## 7. Formulas & Functions {#formulas}

```python
# Basic formulas
worksheet.write_formula('E2', '=SUM(B2:D2)')
worksheet.write_formula('E3', '=AVERAGE(B3:D3)')
worksheet.write_formula('E4', '=COUNT(B4:D4)')
worksheet.write_formula('E5', '=MAX(B5:D5)')
worksheet.write_formula('E6', '=MIN(B6:D6)')
worksheet.write_formula('E7', '=IF(A7>10,"High","Low")')
worksheet.write_formula('E8', '=VLOOKUP(A8,$F$2:$G$100,2,FALSE)')
worksheet.write_formula('E9', '=INDEX(A1:C10,MATCH("Value",A1:A10,0),3)')

# Array formulas (Ctrl+Shift+Enter)
worksheet.write_array_formula('D2:D10', '=A2:A10*B2:B10')

# Dynamic formulas (Excel 365)
worksheet.write_dynamic_array_formula('E2', '=SORT(FILTER(A2:B10,B2:B10>100, "No results"))')

# Using cell references
worksheet.write_formula(4, 4, '=SUM(B2:D2)')  # Row 4, Col 4 = E5

# Named ranges
workbook.define_name('Sales', '=Sheet1!$B$2:$B$100')
worksheet.write_formula('F2', '=SUM(Sales)')
```

## 8. Conditional Formatting {#conditional-formatting}

```python
# Color scale (3 colors)
worksheet.conditional_format('B2:B10', {
    'type': '3_color_scale',
    'min_color': '#F8696B',
    'mid_color': '#FFEB84',
    'max_color': '#63BE7B'
})

# Color scale (2 colors)
worksheet.conditional_format('C2:C10', {
    'type': '2_color_scale',
    'min_color': '#F8696B',
    'max_color': '#63BE7B'
})

# Data bars
worksheet.conditional_format('D2:D10', {
    'type': 'data_bar',
    'bar_color': '#5B9BD5'
})

# Icon sets
worksheet.conditional_format('E2:E10', {
    'type': 'icon_set',
    'icon_style': '3_arrows'
})

# Cell value conditions
worksheet.conditional_format('F2:F10', {
    'type': 'cell',
    'criteria': 'greater than',
    'value': 100,
    'format': format_fill
})

worksheet.conditional_format('G2:G10', {
    'type': 'cell',
    'criteria': 'between',
    'minimum': 50,
    'maximum': 100,
    'format': format_bold
})

# Formula-based condition
worksheet.conditional_format('H2:H10', {
    'type': 'formula',
    'criteria': '=AND(H2>0,H2<50)',
    'format': format_bold
})

# Top/Bottom rules
worksheet.conditional_format('I2:I10', {
    'type': 'top',
    'value': 10,  # Top 10%
    'format': format_header
})

# Duplicate values
worksheet.conditional_format('J2:J10', {
    'type': 'duplicate',
    'format': format_fill
})

# Text contains
worksheet.conditional_format('K2:K10', {
    'type': 'text',
    'criteria': 'containing',
    'value': 'Error',
    'format': format_font_color
})
```

## 9. pandas Integration {#pandas}

```python
# Convert pandas DataFrame to Excel with xlsxwriter engine
df = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie', 'David'],
    'Age': [25, 30, 35, 28],
    'Salary': [50000, 60000, 70000, 55000],
    'Department': ['Sales', 'IT', 'Marketing', 'Sales']
})

# Basic pandas to Excel
df.to_excel('output.xlsx', sheet_name='Data', index=False)

# Using xlsxwriter engine with formatting
with pd.ExcelWriter('formatted_output.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Employees', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Employees']
    
    # Add formats
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'bg_color': '#D9E1F2',
        'border': 1
    })
    
    money_format = workbook.add_format({'num_format': '$#,##0'})
    
    # Apply formats
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
    worksheet.set_column('A:A', 15)  # Name
    worksheet.set_column('B:B', 8)   # Age
    worksheet.set_column('C:C', 12, money_format)  # Salary
    worksheet.set_column('D:D', 12)  # Department
    
    # Add total row
    total_row = len(df) + 1
    worksheet.write(total_row, 1, 'Total:')
    worksheet.write_formula(total_row, 2, f'=SUM(C2:C{total_row})', money_format)

# Multi-sheet with formatting
with pd.ExcelWriter('multi_sheet.xlsx', engine='xlsxwriter') as writer:
    # Write data to multiple sheets
    df.to_excel(writer, sheet_name='Raw_Data', index=False)
    
    # Create pivot table style
    pivot = df.groupby('Department')['Salary'].agg(['mean', 'sum', 'count'])
    pivot.to_excel(writer, sheet_name='Summary', index=True)
    
    # Add charts to summary sheet
    workbook = writer.book
    worksheet = writer.sheets['Summary']
    
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'Average Salary',
        'categories': '=Summary!$A$2:$A$5',
        'values': '=Summary!$B$2:$B$5'
    })
    worksheet.insert_chart('E2', chart)

# Apply conditional formatting to DataFrame export
with pd.ExcelWriter('conditional_format.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Data', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Data']
    
    # Highlight cells based on value
    worksheet.conditional_format('C2:C5', {
        'type': 'cell',
        'criteria': '>',
        'value': 60000,
        'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    })
    
    # Add data bars
    worksheet.conditional_format('C2:C5', {
        'type': 'data_bar',
        'bar_color': '#5B9BD5'
    })

# pandas operations for Excel output
# Sort and filter before export
df_sorted = df.sort_values('Salary', ascending=False)
df_filtered = df[df['Salary'] > 50000]

# Add calculated columns
df['Bonus'] = df['Salary'] * 0.1
df['Total_Comp'] = df['Salary'] + df['Bonus']

# Export with specific columns and order
df[['Name', 'Department', 'Salary']].to_excel('selected_columns.xlsx', index=False)

# Export multiple DataFrames to same sheet
with pd.ExcelWriter('multiple_dfs.xlsx', engine='xlsxwriter') as writer:
    df1.to_excel(writer, sheet_name='Sheet1', startrow=0, index=False)
    df2.to_excel(writer, sheet_name='Sheet1', startrow=len(df1)+2, index=False)
```

## 10. Advanced Features {#advanced}

```python
# Rich strings (multiple formats in one cell)
worksheet.write_rich_string('A1',
    workbook.add_format({'bold': True}), 'Bold ',
    workbook.add_format({'italic': True}), 'Italic ',
    workbook.add_format({'font_color': 'red'}), 'Red')

# Sparklines
worksheet.add_sparkline('F2', {
    'range': 'Sheet1!A2:E2',
    'type': 'line',
    'markers': True,
    'high_point': True,
    'low_point': True
})

# Add table
worksheet.add_table('A1:D5', {
    'columns': [
        {'header': 'Name'},
        {'header': 'Age'},
        {'header': 'Salary'},
        {'header': 'Bonus'}
    ],
    'autofilter': True,
    'total_row': True,
    'style': 'Table Style Medium 9'
})

# Custom properties
workbook.set_properties({
    'title': 'Sales Report',
    'subject': 'Q1 2024 Sales',
    'author': 'John Doe',
    'manager': 'Jane Smith',
    'company': 'Tech Corp',
    'category': 'Financial',
    'keywords': 'sales, Q1, 2024',
    'comments': 'Generated by Python script',
    'status': 'Final'
})

# Set specific worksheet properties
worksheet.set_default_row(15)
worksheet.set_row(0, 30, format_header)
worksheet.set_column('B:B', 15, None, {'hidden': True})

# Page breaks
worksheet.set_h_pagebreaks([20, 40])  # Page break after rows 20 and 40
worksheet.set_v_pagebreaks([8])  # Page break after column 8

# Header and footer
worksheet.set_header('&C&"Arial,Bold"Report Title')
worksheet.set_footer('&P of &N')
worksheet.set_header('&LPage &P &RPrinted: &D &T')

# Outline/Grouping levels
worksheet.set_row(5, None, None, {'level': 1, 'hidden': True})
worksheet.set_column('C:C', None, None, {'level': 1, 'hidden': True})

# Password protection
workbook = xlsxwriter.Workbook('protected.xlsx', {'options': {'password': 'secret'}})

# Add multiple charts to same worksheet
worksheet.insert_chart('A10', chart1)
worksheet.insert_chart('A30', chart2)

# Create a macro-enabled workbook
workbook = xlsxwriter.Workbook('macro.xlsm')

# Add comments with formatting
worksheet.write_comment('A1', 'This is a comment', {
    'author': 'Python Script',
    'width': 200,
    'height': 100,
    'x_scale': 1.5,
    'y_scale': 1.2,
    'visible': True
})

# Cell locking (must call protect() after)
format_locked = workbook.add_format({'locked': True})
format_unlocked = workbook.add_format({'locked': False})
worksheet.write('A1', 'Locked', format_locked)
worksheet.write('B1', 'Unlocked', format_unlocked)
worksheet.protect('password')
```

## Common Patterns & Best Practices

```python
# Template for professional reports
def create_professional_report(df, filename, title):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Data', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Data']
        
        # Define formats
        formats = {
            'header': workbook.add_format({
                'bold': True,
                'bg_color': '#2C3E50',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            }),
            'money': workbook.add_format({'num_format': '$#,##0.00'}),
            'percent': workbook.add_format({'num_format': '0.00%'}),
            'date': workbook.add_format({'num_format': 'yyyy-mm-dd'}),
            'center': workbook.add_format({'align': 'center'}),
            'total': workbook.add_format({
                'bold': True,
                'top': 2,
                'bg_color': '#ECF0F1'
            })
        }
        
        # Apply header format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, formats['header'])
        
        # Auto-fit columns
        for i, col in enumerate(df.columns):
            column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, min(column_width, 30))
        
        # Add title
        worksheet.merge_range('A1:Z1', title, formats['header'])
        
        # Freeze header row
        worksheet.freeze_panes(1, 0)
        
        # Add totals row
        total_row = len(df) + 1
        worksheet.write(total_row, 0, 'TOTAL', formats['total'])
        for i, col in enumerate(df.select_dtypes(include=['number']).columns):
            col_idx = df.columns.get_loc(col)
            worksheet.write_formula(total_row, col_idx, 
                                  f'=SUM({chr(65+col_idx)}2:{chr(65+col_idx)}{total_row})',
                                  formats['total'])

# Usage
df = pd.DataFrame({
    'Product': ['A', 'B', 'C'],
    'Sales': [1000, 2000, 1500],
    'Growth': [0.05, 0.08, 0.03],
    'Date': pd.date_range('2024-01-01', periods=3)
})
create_professional_report(df, 'report.xlsx', '2024 Sales Report')
```

This cheat sheet covers the most common and useful features of xlsxWriter with pandas integration. For more details, refer to the [xlsxWriter documentation](https://xlsxwriter.readthedocs.io/) and [pandas documentation](https://pandas.pydata.org/docs/).
