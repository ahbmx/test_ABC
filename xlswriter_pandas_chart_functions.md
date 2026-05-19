Here are comprehensive functions to build professional-looking charts with xlsxwriter and pandas:

```python
import pandas as pd
import xlsxwriter
from typing import List, Dict, Optional, Tuple, Any
from datetime import datetime

class ProfessionalChartBuilder:
    """Advanced chart builder for professional Excel reports"""
    
    def __init__(self, workbook: xlsxwriter.Workbook):
        self.workbook = workbook
        self.chart_styles = self._define_chart_styles()
    
    def _define_chart_styles(self) -> Dict:
        """Define professional chart style templates"""
        return {
            'corporate': {
                'colors': ['#2C3E50', '#34495E', '#5D6D7E', '#7F8C8D', '#95A5A6'],
                'title_font': {'name': 'Calibri', 'size': 14, 'bold': True, 'color': '#2C3E50'},
                'axis_font': {'name': 'Calibri', 'size': 10, 'color': '#7F8C8D'},
                'legend_font': {'name': 'Calibri', 'size': 9, 'color': '#2C3E50'},
                'gridlines': {'major': {'visible': True, 'line': {'color': '#E5E7E9'}}}
            },
            'modern': {
                'colors': ['#3498DB', '#E74C3C', '#2ECC71', '#F39C12', '#9B59B6'],
                'title_font': {'name': 'Segoe UI', 'size': 14, 'bold': True, 'color': '#2C3E50'},
                'axis_font': {'name': 'Segoe UI', 'size': 10, 'color': '#7F8C8D'},
                'legend_font': {'name': 'Segoe UI', 'size': 9, 'color': '#2C3E50'},
                'gridlines': {'major': {'visible': True, 'line': {'color': '#BDC3C7', 'dash_type': 'dash'}}}
            },
            'financial': {
                'colors': ['#1B4F72', '#117A65', '#B03A2E', '#D4AC0D', '#5B2C6F'],
                'title_font': {'name': 'Arial', 'size': 12, 'bold': True, 'color': '#1B4F72'},
                'axis_font': {'name': 'Arial', 'size': 9, 'color': '#566573'},
                'legend_font': {'name': 'Arial', 'size': 8, 'color': '#1B4F72'},
                'gridlines': {'major': {'visible': True, 'line': {'color': '#D5D8DC'}}}
            },
            'gradient': {
                'colors': ['#00C6FB', '#005BEA', '#00C6FB', '#005BEA', '#00C6FB'],
                'title_font': {'name': 'Calibri', 'size': 14, 'bold': True, 'color': '#2C3E50'},
                'axis_font': {'name': 'Calibri', 'size': 10, 'color': '#7F8C8D'},
                'legend_font': {'name': 'Calibri', 'size': 9, 'color': '#2C3E50'},
                'gridlines': {'major': {'visible': False}}
            }
        }
    
    def create_column_chart(
        self,
        df: pd.DataFrame,
        categories_col: str,
        values_cols: List[str],
        title: str,
        x_axis_name: str = '',
        y_axis_name: str = '',
        chart_style: str = 'corporate',
        stacked: bool = False,
        percentage: bool = False,
        data_labels: bool = True
    ) -> xlsxwriter.chart.Chart:
        """
        Create a professional column/bar chart
        
        Parameters:
        -----------
        df : pd.DataFrame
            Data source
        categories_col : str
            Column for x-axis categories
        values_cols : list
            Columns for y-axis values
        title : str
            Chart title
        x_axis_name : str
            X-axis label
        y_axis_name : str
            Y-axis label
        chart_style : str
            Style template ('corporate', 'modern', 'financial', 'gradient')
        stacked : bool
            Create stacked column chart
        percentage : bool
            Create percentage stacked chart
        data_labels : bool
            Show data labels
        """
        chart_type = 'column'
        if stacked:
            chart_type = 'column_stacked'
        if percentage:
            chart_type = 'column_stacked_percent'
        
        chart = self.workbook.add_chart({'type': chart_type})
        style = self.chart_styles[chart_style]
        
        # Add series
        for idx, col in enumerate(values_cols):
            color = style['colors'][idx % len(style['colors'])]
            chart.add_series({
                'name': col,
                'categories': [df.name, 1, df.columns.get_loc(categories_col), 
                              len(df), df.columns.get_loc(categories_col)],
                'values': [df.name, 1, df.columns.get_loc(col), 
                          len(df), df.columns.get_loc(col)],
                'fill': {'color': color},
                'border': {'color': color},
                'data_labels': {'value': data_labels, 'position': 'outside_end'} if data_labels else None,
                'line': {'color': color, 'width': 1.5} if not stacked else None
            })
        
        # Configure chart
        chart.set_title({
            'name': title,
            'name_font': style['title_font']
        })
        
        chart.set_x_axis({
            'name': x_axis_name,
            'name_font': style['axis_font'],
            'label_font': style['axis_font'],
            'major_gridlines': style['gridlines']['major'] if style['gridlines']['major']['visible'] else None
        })
        
        chart.set_y_axis({
            'name': y_axis_name,
            'name_font': style['axis_font'],
            'label_font': style['axis_font'],
            'major_gridlines': style['gridlines']['major'] if style['gridlines']['major']['visible'] else None
        })
        
        chart.set_legend({
            'font': style['legend_font'],
            'position': 'bottom',
            'delete_series': [] if len(values_cols) <= 4 else [4,5,6,7]  # Limit legend entries
        })
        
        chart.set_size({'width': 600, 'height': 400})
        
        return chart
    
    def create_line_chart(
        self,
        df: pd.DataFrame,
        categories_col: str,
        values_cols: List[str],
        title: str,
        x_axis_name: str = '',
        y_axis_name: str = '',
        chart_style: str = 'corporate',
        show_markers: bool = True,
        smooth_lines: bool = False,
        trendlines: bool = False,
        data_labels: bool = False
    ) -> xlsxwriter.chart.Chart:
        """
        Create professional line chart with trends
        """
        chart = self.workbook.add_chart({'type': 'line'})
        style = self.chart_styles[chart_style]
        
        for idx, col in enumerate(values_cols):
            color = style['colors'][idx % len(style['colors'])]
            series_config = {
                'name': col,
                'categories': [df.name, 1, df.columns.get_loc(categories_col), 
                              len(df), df.columns.get_loc(categories_col)],
                'values': [df.name, 1, df.columns.get_loc(col), 
                          len(df), df.columns.get_loc(col)],
                'line': {'color': color, 'width': 2.5},
                'marker': {'type': 'circle', 'size': 6, 'fill': {'color': color}} if show_markers else None,
                'smooth': smooth_lines,
                'data_labels': {'value': data_labels} if data_labels else None
            }
            
            chart.add_series(series_config)
            
            # Add trendline if requested
            if trendlines:
                chart.add_series({
                    'trendline': {
                        'type': 'linear',
                        'name': f'{col} Trend',
                        'line': {'color': color, 'dash_type': 'dash', 'width': 1.5}
                    }
                })
        
        chart.set_title({'name': title, 'name_font': style['title_font']})
        chart.set_x_axis({
            'name': x_axis_name,
            'name_font': style['axis_font'],
            'label_font': style['axis_font'],
            'major_gridlines': style['gridlines']['major'] if style['gridlines']['major']['visible'] else None
        })
        chart.set_y_axis({
            'name': y_axis_name,
            'name_font': style['axis_font'],
            'label_font': style['axis_font'],
            'major_gridlines': style['gridlines']['major'] if style['gridlines']['major']['visible'] else None
        })
        chart.set_legend({'font': style['legend_font'], 'position': 'bottom'})
        chart.set_size({'width': 650, 'height': 400})
        
        return chart
    
    def create_pie_chart(
        self,
        df: pd.DataFrame,
        categories_col: str,
        values_col: str,
        title: str,
        chart_style: str = 'corporate',
        show_percentage: bool = True,
        show_values: bool = True,
        explode_largest: bool = True,
        donut_hole_size: int = 0
    ) -> xlsxwriter.chart.Chart:
        """
        Create professional pie/donut chart
        """
        chart_type = 'pie' if donut_hole_size == 0 else 'doughnut'
        chart = self.workbook.add_chart({'type': chart_type})
        style = self.chart_styles[chart_style]
        
        # Prepare data labels
        data_labels_config = {
            'category': True,
            'percentage': show_percentage,
            'value': show_values,
            'separator': ' - ',
            'font': style['axis_font']
        }
        
        # Create series
        series = {
            'name': title,
            'categories': [df.name, 1, df.columns.get_loc(categories_col), 
                          len(df), df.columns.get_loc(categories_col)],
            'values': [df.name, 1, df.columns.get_loc(values_col), 
                      len(df), df.columns.get_loc(values_col)],
            'data_labels': data_labels_config,
            'points': []
        }
        
        # Explode largest slice if requested
        if explode_largest:
            max_idx = df[values_col].idxmax()
            for i in range(len(df)):
                if i == max_idx:
                    series['points'].append({'explosion': 10})
                else:
                    series['points'].append({'explosion': 0})
        
        chart.add_series(series)
        
        # Donut hole size
        if donut_hole_size > 0:
            chart.set_hole_size(donut_hole_size)
        
        chart.set_title({'name': title, 'name_font': style['title_font']})
        chart.set_legend({'font': style['legend_font'], 'position': 'right'})
        chart.set_size({'width': 500, 'height': 400})
        
        return chart
    
    def create_area_chart(
        self,
        df: pd.DataFrame,
        categories_col: str,
        values_cols: List[str],
        title: str,
        x_axis_name: str = '',
        y_axis_name: str = '',
        chart_style: str = 'gradient',
        stacked: bool = True,
        transparency: int = 50
    ) -> xlsxwriter.chart.Chart:
        """
        Create professional area chart with transparency
        """
        chart_type = 'area_stacked' if stacked else 'area'
        chart = self.workbook.add_chart({'type': chart_type})
        style = self.chart_styles[chart_style]
        
        for idx, col in enumerate(values_cols):
            color = style['colors'][idx % len(style['colors'])]
            chart.add_series({
                'name': col,
                'categories': [df.name, 1, df.columns.get_loc(categories_col), 
                              len(df), df.columns.get_loc(categories_col)],
                'values': [df.name, 1, df.columns.get_loc(col), 
                          len(df), df.columns.get_loc(col)],
                'fill': {'color': color, 'transparency': transparency},
                'line': {'color': color, 'width': 1.5},
                'border': {'color': color}
            })
        
        chart.set_title({'name': title, 'name_font': style['title_font']})
        chart.set_x_axis({
            'name': x_axis_name,
            'name_font': style['axis_font'],
            'label_font': style['axis_font'],
            'major_gridlines': style['gridlines']['major'] if style['gridlines']['major']['visible'] else None
        })
        chart.set_y_axis({
            'name': y_axis_name,
            'name_font': style['axis_font'],
            'label_font': style['axis_font'],
            'major_gridlines': style['gridlines']['major'] if style['gridlines']['major']['visible'] else None
        })
        chart.set_legend({'font': style['legend_font'], 'position': 'bottom'})
        chart.set_size({'width': 600, 'height': 400})
        
        return chart
    
    def create_scatter_chart(
        self,
        df: pd.DataFrame,
        x_col: str,
        y_cols: List[str],
        title: str,
        x_axis_name: str = '',
        y_axis_name: str = '',
        chart_style: str = 'modern',
        show_trendlines: bool = True,
        bubble_size_col: Optional[str] = None
    ) -> xlsxwriter.chart.Chart:
        """
        Create professional scatter/bubble chart
        """
        chart_type = 'bubble' if bubble_size_col else 'scatter'
        chart = self.workbook.add_chart({'type': chart_type})
        style = self.chart_styles[chart_style]
        
        for idx, col in enumerate(y_cols):
            color = style['colors'][idx % len(style['colors'])]
            
            if bubble_size_col:
                series_config = {
                    'name': col,
                    'categories': [df.name, 1, df.columns.get_loc(x_col), 
                                  len(df), df.columns.get_loc(x_col)],
                    'values': [df.name, 1, df.columns.get_loc(col), 
                              len(df), df.columns.get_loc(col)],
                    'bubble_sizes': [df.name, 1, df.columns.get_loc(bubble_size_col), 
                                    len(df), df.columns.get_loc(bubble_size_col)],
                    'fill': {'color': color, 'transparency': 30},
                    'border': {'color': color},
                    'marker': {'type': 'circle', 'size': 8}
                }
            else:
                series_config = {
                    'name': col,
                    'categories': [df.name, 1, df.columns.get_loc(x_col), 
                                  len(df), df.columns.get_loc(x_col)],
                    'values': [df.name, 1, df.columns.get_loc(col), 
                              len(df), df.columns.get_loc(col)],
                    'marker': {'type': 'circle', 'size': 7, 'fill': {'color': color}},
                    'line': {'none': True},
                    'trendline': {'type': 'linear'} if show_trendlines else None
                }
            
            chart.add_series(series_config)
        
        chart.set_title({'name': title, 'name_font': style['title_font']})
        chart.set_x_axis({
            'name': x_axis_name,
            'name_font': style['axis_font'],
            'label_font': style['axis_font'],
            'major_gridlines': style['gridlines']['major'] if style['gridlines']['major']['visible'] else None
        })
        chart.set_y_axis({
            'name': y_axis_name,
            'name_font': style['axis_font'],
            'label_font': style['axis_font'],
            'major_gridlines': style['gridlines']['major'] if style['gridlines']['major']['visible'] else None
        })
        chart.set_legend({'font': style['legend_font'], 'position': 'bottom'})
        chart.set_size({'width': 600, 'height': 400})
        
        return chart
    
    def create_combined_chart(
        self,
        df: pd.DataFrame,
        categories_col: str,
        column_cols: List[str],
        line_cols: List[str],
        title: str,
        x_axis_name: str = '',
        y1_axis_name: str = '',
        y2_axis_name: str = '',
        chart_style: str = 'corporate'
    ) -> xlsxwriter.chart.Chart:
        """
        Create combination chart (columns + line on secondary axis)
        """
        chart = self.workbook.add_chart({'type': 'column'})
        style = self.chart_styles[chart_style]
        
        # Add column series on primary axis
        for idx, col in enumerate(column_cols):
            color = style['colors'][idx % len(style['colors'])]
            chart.add_series({
                'name': col,
                'categories': [df.name, 1, df.columns.get_loc(categories_col), 
                              len(df), df.columns.get_loc(categories_col)],
                'values': [df.name, 1, df.columns.get_loc(col), 
                          len(df), df.columns.get_loc(col)],
                'fill': {'color': color},
                'y2_axis': False
            })
        
        # Add line series on secondary axis
        for idx, col in enumerate(line_cols):
            color = style['colors'][(len(column_cols) + idx) % len(style['colors'])]
            chart.add_series({
                'name': col,
                'categories': [df.name, 1, df.columns.get_loc(categories_col), 
                              len(df), df.columns.get_loc(categories_col)],
                'values': [df.name, 1, df.columns.get_loc(col), 
                          len(df), df.columns.get_loc(col)],
                'line': {'color': color, 'width': 2.5},
                'marker': {'type': 'circle', 'size': 6},
                'y2_axis': True
            })
        
        chart.set_title({'name': title, 'name_font': style['title_font']})
        chart.set_x_axis({'name': x_axis_name, 'name_font': style['axis_font']})
        chart.set_y_axis({'name': y1_axis_name, 'name_font': style['axis_font']})
        chart.set_y2_axis({'name': y2_axis_name, 'name_font': style['axis_font']})
        chart.set_legend({'font': style['legend_font'], 'position': 'bottom'})
        chart.set_size({'width': 700, 'height': 450})
        
        return chart


# Wrapper functions for easy use

def create_dashboard_report(
    df: pd.DataFrame,
    output_file: str,
    dashboard_title: str = "Executive Dashboard"
) -> None:
    """
    Create a complete Excel dashboard with multiple professional charts
    """
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write data to worksheet
        df.to_excel(writer, sheet_name='Data', index=False)
        
        # Create dashboard sheet
        dashboard = writer.book.add_worksheet('Dashboard')
        chart_builder = ProfessionalChartBuilder(writer.book)
        
        # Add title
        title_format = writer.book.add_format({
            'bold': True,
            'font_size': 20,
            'font_color': '#2C3E50',
            'align': 'center',
            'valign': 'vcenter'
        })
        dashboard.merge_range('A1:H2', dashboard_title, title_format)
        
        # Create various charts based on data columns
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        date_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        
        row_offset = 3
        col_offset = 0
        
        # 1. Column chart for top numeric metrics
        if len(numeric_cols) >= 2 and categorical_cols:
            chart1 = chart_builder.create_column_chart(
                df, categorical_cols[0], numeric_cols[:3],
                f"Top Metrics by {categorical_cols[0]}",
                x_axis_name=categorical_cols[0],
                y_axis_name="Values",
                chart_style='corporate'
            )
            dashboard.insert_chart(row_offset, col_offset, chart1)
            col_offset += 8
        
        # 2. Line chart for trends if date column exists
        if date_cols and numeric_cols:
            chart2 = chart_builder.create_line_chart(
                df, date_cols[0], numeric_cols[:2],
                "Trend Analysis",
                x_axis_name=date_cols[0],
                y_axis_name="Values",
                chart_style='modern',
                show_markers=True
            )
            dashboard.insert_chart(row_offset, col_offset, chart2)
            row_offset += 20
            col_offset = 0
        
        # 3. Pie chart for distribution
        if categorical_cols and numeric_cols:
            chart3 = chart_builder.create_pie_chart(
                df, categorical_cols[0], numeric_cols[0],
                f"Distribution by {categorical_cols[0]}",
                chart_style='financial',
                show_percentage=True
            )
            dashboard.insert_chart(row_offset, col_offset, chart3)
            col_offset += 6
        
        # 4. Area chart for cumulative data
        if date_cols and len(numeric_cols) >= 2:
            chart4 = chart_builder.create_area_chart(
                df, date_cols[0], numeric_cols[:2],
                "Cumulative Performance",
                x_axis_name=date_cols[0],
                y_axis_name="Values",
                chart_style='gradient',
                stacked=True
            )
            dashboard.insert_chart(row_offset, col_offset, chart4)


# Advanced chart with custom formatting

def create_advanced_financial_chart(
    df: pd.DataFrame,
    output_file: str
) -> None:
    """
    Create advanced financial charts with professional formatting
    """
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Financial Data', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Financial Data']
        chart_builder = ProfessionalChartBuilder(workbook)
        
        # Create multiple chart layouts
        charts = {}
        
        # Revenue vs Expenses waterfall style (using column chart)
        if 'Revenue' in df.columns and 'Expenses' in df.columns:
            charts['profit_analysis'] = chart_builder.create_column_chart(
                df, 'Month', ['Revenue', 'Expenses'],
                "Revenue vs Expenses Analysis",
                x_axis_name="Month",
                y_axis_name="Amount ($)",
                chart_style='financial',
                stacked=False,
                data_labels=True
            )
            worksheet.insert_chart('J2', charts['profit_analysis'])
        
        # Profit margin trend with secondary axis
        if 'Profit_Margin' in df.columns:
            charts['margin_trend'] = chart_builder.create_line_chart(
                df, 'Month', ['Profit_Margin'],
                "Profit Margin Trend",
                x_axis_name="Month",
                y_axis_name="Margin (%)",
                chart_style='modern',
                show_markers=True,
                smooth_lines=True
            )
            worksheet.insert_chart('J20', charts['margin_trend'])
        
        # Portfolio allocation pie chart
        if 'Investment' in df.columns and 'Asset_Class' in df.columns:
            # Aggregate data for pie chart
            pie_data = df.groupby('Asset_Class')['Investment'].sum().reset_index()
            pie_data.name = 'Pie_Data'
            
            charts['portfolio'] = chart_builder.create_pie_chart(
                pie_data, 'Asset_Class', 'Investment',
                "Portfolio Allocation",
                chart_style='modern',
                show_percentage=True,
                explode_largest=True,
                donut_hole_size=40  # Creates a donut chart
            )
            worksheet.insert_chart('A30', charts['portfolio'])
        
        # Set column widths
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:E', 12)


# Quick chart generation function

def quick_chart(
    df: pd.DataFrame,
    chart_type: str,
    categories: str,
    values: Union[str, List[str]],
    title: str,
    output_file: str,
    **kwargs
) -> None:
    """
    Quick one-line chart generation
    
    Parameters:
    -----------
    chart_type : str
        'column', 'line', 'pie', 'area', 'scatter', 'combined'
    categories : str
        Column for x-axis/categories
    values : str or list
        Column(s) for values
    title : str
        Chart title
    output_file : str
        Output Excel file path
    """
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Data', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Data']
        chart_builder = ProfessionalChartBuilder(workbook)
        
        if isinstance(values, str):
            values = [values]
        
        if chart_type == 'column':
            chart = chart_builder.create_column_chart(
                df, categories, values, title, **kwargs
            )
        elif chart_type == 'line':
            chart = chart_builder.create_line_chart(
                df, categories, values, title, **kwargs
            )
        elif chart_type == 'pie':
            chart = chart_builder.create_pie_chart(
                df, categories, values[0], title, **kwargs
            )
        elif chart_type == 'area':
            chart = chart_builder.create_area_chart(
                df, categories, values, title, **kwargs
            )
        elif chart_type == 'scatter':
            chart = chart_builder.create_scatter_chart(
                df, categories, values, title, **kwargs
            )
        else:
            raise ValueError(f"Unsupported chart type: {chart_type}")
        
        worksheet.insert_chart('J2', chart)


# Usage examples

if __name__ == "__main__":
    # Create sample data
    dates = pd.date_range('2024-01-01', periods=12, freq='M')
    sample_df = pd.DataFrame({
        'Month': dates,
        'Product_A': np.random.randint(100, 500, 12),
        'Product_B': np.random.randint(150, 600, 12),
        'Product_C': np.random.randint(80, 400, 12),
        'Revenue': np.random.randint(50000, 150000, 12),
        'Expenses': np.random.randint(30000, 100000, 12),
        'Profit_Margin': np.random.uniform(0.15, 0.35, 12),
        'Region': np.random.choice(['North', 'South', 'East', 'West'], 12),
        'Investment': np.random.randint(1000, 10000, 12),
        'Asset_Class': np.random.choice(['Stocks', 'Bonds', 'Real Estate', 'Cash'], 12)
    })
    
    # Method 1: Create complete dashboard
    create_dashboard_report(sample_df, 'professional_dashboard.xlsx')
    
    # Method 2: Create advanced financial charts
    create_advanced_financial_chart(sample_df, 'financial_analysis.xlsx')
    
    # Method 3: Quick chart generation
    quick_chart(
        sample_df,
        chart_type='line',
        categories='Month',
        values=['Product_A', 'Product_B', 'Product_C'],
        title='Monthly Product Sales Trends',
        output_file='quick_sales_chart.xlsx',
        chart_style='modern',
        show_markers=True,
        smooth_lines=True
    )
    
    # Method 4: Custom combination chart
    with pd.ExcelWriter('combination_chart.xlsx', engine='xlsxwriter') as writer:
        sample_df.to_excel(writer, sheet_name='Sales Data', index=False)
        chart_builder = ProfessionalChartBuilder(writer.book)
        
        combined_chart = chart_builder.create_combined_chart(
            sample_df,
            categories_col='Month',
            column_cols=['Revenue'],
            line_cols=['Profit_Margin'],
            title='Revenue vs Profit Margin Analysis',
            x_axis_name='Month',
            y1_axis_name='Revenue ($)',
            y2_axis_name='Profit Margin (%)',
            chart_style='corporate'
        )
        
        writer.sheets['Sales Data'].insert_chart('J2', combined_chart)
    
    print("All professional charts generated successfully!")
```

## Key Features of the Professional Chart Builder:

### 1. **Multiple Chart Types**
- Column/Bar charts (stacked, percentage)
- Line charts (with trends and markers)
- Pie/Donut charts (with explosion effects)
- Area charts (stacked with transparency)
- Scatter/Bubble charts
- Combination charts (dual axis)

### 2. **Professional Styling Templates**
- **Corporate**: Clean, business-appropriate colors
- **Modern**: Vibrant, contemporary design
- **Financial**: Conservative, finance-focused
- **Gradient**: Modern gradient effects

### 3. **Advanced Features**
- Automatic data labels
- Trendlines for analysis
- Secondary axes for dual metrics
- Customizable transparency
- Bubble charts for three-dimensional data
- Donut charts with adjustable hole size

### 4. **Dashboard Capabilities**
- Multiple charts in one sheet
- Automatic layout positioning
- Professional title formatting
- Gridlines and axis styling

### 5. **Quick Generation Functions**
- One-line chart creation
- Flexible parameter passing
- Automatic chart sizing

These functions create publication-quality charts suitable for executive dashboards, financial reports, and business presentations. The styling templates ensure consistency across all visualizations while maintaining professional appearance.
