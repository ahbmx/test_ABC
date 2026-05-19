Here's a comprehensive **Matplotlib + Pandas Integration Cheat Sheet** tailored for a storage admin who needs professional-looking charts and report integration.

---

# 📊 MATPLOTLIB + PANDAS INTEGRATION CHEAT SHEET
## For Storage Administrators – Professional Charts & Reporting

---

## 1. QUICK SETUP

```python
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.dates as mdates

# Set professional style globally
plt.style.use('seaborn-v0_8-darkgrid')
plt.rcParams['figure.figsize'] = (12, 6)
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 14
plt.rcParams['axes.labelsize'] = 12
```

---

## 2. LOADING STORAGE DATA

```python
# Typical storage metrics
df = pd.read_csv('storage_metrics.csv', parse_dates=['timestamp'])

# Sample storage dataframe
data = {
    'timestamp': pd.date_range('2024-01-01', periods=100, freq='D'),
    'capacity_tb': np.random.normal(500, 50, 100).cumsum(),
    'iops_read': np.random.poisson(1000, 100),
    'iops_write': np.random.poisson(800, 100),
    'latency_ms': np.random.gamma(2, 3, 100),
    'pool_utilization_pct': np.random.uniform(60, 95, 100)
}
df = pd.DataFrame(data)
```

---

## 3. CORE PANDAS + MATPLOTLIB PLOTTING

### Time Series (Storage Capacity Trend)
```python
fig, ax = plt.subplots(figsize=(14, 7))
df.plot(x='timestamp', y='capacity_tb', ax=ax, 
        color='#2E86AB', linewidth=2, marker='o', markersize=3)
ax.set_title('Storage Capacity Growth Trend', fontweight='bold', pad=15)
ax.set_xlabel('Date')
ax.set_ylabel('Capacity (TB)')
ax.grid(True, alpha=0.3)
ax.legend(['Used Capacity'], loc='upper left')
plt.tight_layout()
```

### Dual-Axis (IOPS + Latency)
```python
fig, ax1 = plt.subplots(figsize=(14, 7))

# IOPS on primary axis
ax1.plot(df['timestamp'], df['iops_read'], label='Read IOPS', 
         color='#1B98A0', linewidth=1.5)
ax1.plot(df['timestamp'], df['iops_write'], label='Write IOPS', 
         color='#E9C46A', linewidth=1.5)
ax1.set_xlabel('Date')
ax1.set_ylabel('IOPS', color='#1B98A0')
ax1.tick_params(axis='y', labelcolor='#1B98A0')
ax1.legend(loc='upper left')

# Latency on secondary axis
ax2 = ax1.twinx()
ax2.plot(df['timestamp'], df['latency_ms'], color='#E76F51', 
         linewidth=1.5, linestyle='--', label='Latency')
ax2.set_ylabel('Latency (ms)', color='#E76F51')
ax2.tick_params(axis='y', labelcolor='#E76F51')
ax2.legend(loc='upper right')

plt.title('Storage Performance: IOPS vs Latency', fontweight='bold')
plt.tight_layout()
```

### Utilization Distribution (Histogram)
```python
fig, ax = plt.subplots(figsize=(10, 6))
df['pool_utilization_pct'].plot(kind='hist', bins=20, 
                                 color='#2E86AB', edgecolor='black', 
                                 alpha=0.7, ax=ax)
ax.axvline(df['pool_utilization_pct'].mean(), color='red', 
           linestyle='--', linewidth=2, label=f'Mean: {df["pool_utilization_pct"].mean():.1f}%')
ax.axvline(80, color='orange', linestyle=':', linewidth=2, 
           label='Warning Threshold (80%)')
ax.axvline(90, color='red', linestyle=':', linewidth=2, 
           label='Critical Threshold (90%)')
ax.set_title('Storage Pool Utilization Distribution', fontweight='bold')
ax.set_xlabel('Utilization (%)')
ax.set_ylabel('Frequency')
ax.legend()
plt.tight_layout()
```

---

## 4. PROFESSIONAL FINE-TUNING TECHNIQUES

### Color Palettes for Storage Admin
```python
# Corporate/Professional palettes
storage_colors = {
    'primary': '#2E86AB',      # Steel blue
    'secondary': '#A23B72',    # Plum
    'accent': '#F18F01',       # Orange
    'warning': '#C73E1D',      # Red-orange
    'success': '#6A994E',      # Green
    'neutral': '#BCBDC0'       # Gray
}

# Custom color cycle
plt.rcParams['axes.prop_cycle'] = plt.cycler(color=[
    '#2E86AB', '#A23B72', '#F18F01', '#6A994E', '#C73E1D'
])
```

### Annotation and Thresholds
```python
fig, ax = plt.subplots(figsize=(14, 7))
df.plot(x='timestamp', y='capacity_tb', ax=ax, color='#2E86AB', linewidth=2)

# Add threshold lines
ax.axhline(y=df['capacity_tb'].quantile(0.95), color='orange', 
           linestyle='--', linewidth=1.5, alpha=0.7)
ax.axhline(y=df['capacity_tb'].max(), color='red', 
           linestyle=':', linewidth=1.5, alpha=0.7)

# Annotate peak point
max_idx = df['capacity_tb'].idxmax()
max_val = df.loc[max_idx, 'capacity_tb']
max_date = df.loc[max_idx, 'timestamp']
ax.annotate(f'Peak: {max_val:.0f} TB', 
            xy=(max_date, max_val), 
            xytext=(10, 10), 
            textcoords='offset points',
            arrowprops=dict(arrowstyle='->', color='red'),
            fontweight='bold', color='red')

# Add data labels on last point
last_val = df['capacity_tb'].iloc[-1]
last_date = df['timestamp'].iloc[-1]
ax.text(last_date, last_val, f'  {last_val:.0f} TB', 
        verticalalignment='bottom', fontweight='bold')

plt.title('Storage Capacity with Peak Annotation', fontweight='bold')
plt.tight_layout()
```

### Custom Tick Formatting
```python
fig, ax = plt.subplots(figsize=(14, 7))
df.plot(x='timestamp', y='capacity_tb', ax=ax, color='#2E86AB')

# Format x-axis dates
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=2))
plt.xticks(rotation=45, ha='right')

# Format y-axis with commas and TB suffix
def tb_formatter(x, p):
    return f'{x:,.0f} TB'
ax.yaxis.set_major_formatter(plt.FuncFormatter(tb_formatter))

plt.title('Formatted Axes for Reports', fontweight='bold')
plt.tight_layout()
```

---

## 5. ADVANCED STORAGE CHARTS

### Heatmap (Daily Utilization Pattern)
```python
# Create pivot table for heatmap
df['day_of_week'] = df['timestamp'].dt.day_name()
df['hour'] = df['timestamp'].dt.hour
heatmap_data = df.pivot_table(values='pool_utilization_pct', 
                              index='day_of_week', 
                              columns='hour', 
                              aggfunc='mean')

# Order days properly
days_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
heatmap_data = heatmap_data.reindex(days_order)

fig, ax = plt.subplots(figsize=(16, 8))
im = ax.imshow(heatmap_data.values, cmap='RdYlGn_r', aspect='auto', 
               interpolation='nearest')
plt.colorbar(im, ax=ax, label='Utilization (%)')

# Labels
ax.set_xticks(range(len(heatmap_data.columns)))
ax.set_xticklabels(heatmap_data.columns)
ax.set_yticks(range(len(heatmap_data.index)))
ax.set_yticklabels(heatmap_data.index)
ax.set_xlabel('Hour of Day')
ax.set_ylabel('Day of Week')
ax.set_title('Storage Pool Utilization Heatmap (by Day/Hour)', fontweight='bold')

# Add text annotations
for i in range(len(heatmap_data.index)):
    for j in range(len(heatmap_data.columns)):
        text = ax.text(j, i, f'{heatmap_data.values[i, j]:.0f}',
                      ha="center", va="center", color="black", fontsize=8)

plt.tight_layout()
```

### Box Plot (Latency by Day)
```python
fig, ax = plt.subplots(figsize=(12, 6))
df['weekday'] = df['timestamp'].dt.day_name()
df.boxplot(column='latency_ms', by='weekday', ax=ax, 
           color=storage_colors, patch_artist=True)

# Customize boxplot
for box in ax.artists:
    box.set_facecolor('#2E86AB')
    box.set_alpha(0.7)
    
ax.set_title('Latency Distribution by Day of Week', fontweight='bold')
ax.set_xlabel('Day of Week')
ax.set_ylabel('Latency (ms)')
plt.suptitle('')  # Remove automatic 'Boxplot grouped by...' title
plt.xticks(rotation=45)
plt.tight_layout()
```

### Area Chart (Capacity Components)
```python
# Example with stacked components
df_stack = pd.DataFrame({
    'timestamp': df['timestamp'],
    'active_data': df['capacity_tb'] * 0.6,
    'snapshots': df['capacity_tb'] * 0.25,
    'unused': df['capacity_tb'] * 0.15
})

fig, ax = plt.subplots(figsize=(14, 7))
df_stack.set_index('timestamp').plot(kind='area', ax=ax, 
                                      color=['#2E86AB', '#A23B72', '#BCBDC0'],
                                      alpha=0.7, stacked=True)
ax.set_title('Storage Capacity Breakdown', fontweight='bold')
ax.set_xlabel('Date')
ax.set_ylabel('Capacity (TB)')
ax.legend(loc='upper left', title='Components')
ax.grid(True, alpha=0.3)
plt.tight_layout()
```

---

## 6. SAVING IMAGES

```python
# High-resolution for reports
fig.savefig('storage_chart.png', dpi=300, bbox_inches='tight', 
            facecolor='white', edgecolor='none')

# Vector format (best for reports)
fig.savefig('storage_chart.pdf', format='pdf', bbox_inches='tight')

# Transparent background (for presentations)
fig.savefig('storage_chart_transparent.png', dpi=150, 
            transparent=True, bbox_inches='tight')

# Multiple formats in one go
for ext in ['png', 'pdf', 'svg']:
    fig.savefig(f'storage_chart.{ext}', dpi=300, bbox_inches='tight')
```

---

## 7. PDF REPORT INTEGRATION

### Method 1: Multi-Page PDF Report
```python
def create_storage_report_pdf(df, filename='storage_report.pdf'):
    """Generate a comprehensive storage report PDF"""
    
    with PdfPages(filename) as pdf:
        
        # Page 1: Capacity Trend
        fig1, ax = plt.subplots(figsize=(11, 8.5))  # Letter size
        df.plot(x='timestamp', y='capacity_tb', ax=ax, 
                color='#2E86AB', linewidth=2)
        ax.set_title('Storage Capacity Trend - Monthly Report', fontsize=14, fontweight='bold')
        ax.set_xlabel('Date')
        ax.set_ylabel('Capacity (TB)')
        ax.grid(True, alpha=0.3)
        pdf.savefig(fig1, bbox_inches='tight')
        plt.close(fig1)
        
        # Page 2: Performance Metrics
        fig2, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8.5))
        df.plot(x='timestamp', y='iops_read', ax=ax1, color='#1B98A0', label='Read IOPS')
        df.plot(x='timestamp', y='iops_write', ax=ax1, color='#E9C46A', label='Write IOPS')
        ax1.set_title('IOPS Performance', fontweight='bold')
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        
        df.plot(x='timestamp', y='latency_ms', ax=ax2, color='#E76F51')
        ax2.set_title('Latency Trend', fontweight='bold')
        ax2.set_ylabel('Latency (ms)')
        ax2.grid(True, alpha=0.3)
        
        plt.tight_layout()
        pdf.savefig(fig2, bbox_inches='tight')
        plt.close(fig2)
        
        # Page 3: Statistical Summary Table
        fig3, ax = plt.subplots(figsize=(11, 8.5))
        ax.axis('tight')
        ax.axis('off')
        
        summary_stats = df[['capacity_tb', 'iops_read', 'iops_write', 'latency_ms']].describe()
        table = ax.table(cellText=summary_stats.round(2).values,
                         rowLabels=summary_stats.index,
                         colLabels=summary_stats.columns,
                         cellLoc='center',
                         loc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1.2, 1.5)
        ax.set_title('Storage Performance Statistics', fontweight='bold', pad=20)
        
        pdf.savefig(fig3, bbox_inches='tight')
        plt.close(fig3)
        
    print(f"Report saved as {filename}")

# Generate report
create_storage_report_pdf(df, 'storage_monthly_report.pdf')
```

### Method 2: Custom Report with Cover Page
```python
def create_professional_report(df, filename='storage_report.pdf'):
    with PdfPages(filename) as pdf:
        
        # Cover Page
        fig_cover = plt.figure(figsize=(11, 8.5))
        fig_cover.patch.set_facecolor('#2E86AB')
        ax_cover = fig_cover.add_subplot(111)
        ax_cover.axis('off')
        ax_cover.text(0.5, 0.6, 'STORAGE INFRASTRUCTURE REPORT', 
                     transform=fig_cover.transFigure, 
                     fontsize=24, ha='center', color='white', fontweight='bold')
        ax_cover.text(0.5, 0.4, f'Report Date: {pd.Timestamp.now().strftime("%B %d, %Y")}',
                     transform=fig_cover.transFigure, 
                     fontsize=14, ha='center', color='white')
        ax_cover.text(0.5, 0.3, f'Data Period: {df["timestamp"].min().date()} to {df["timestamp"].max().date()}',
                     transform=fig_cover.transFigure,
                     fontsize=12, ha='center', color='white')
        pdf.savefig(fig_cover, bbox_inches='tight')
        plt.close(fig_cover)
        
        # Content pages (as before)
        # ... add your charts
        
        # Add metadata
        d = pdf.infodict()
        d['Title'] = 'Storage Performance Report'
        d['Author'] = 'Storage Admin'
        d['Subject'] = 'Monthly Storage Metrics'
        d['CreationDate'] = pd.Timestamp.now()

create_professional_report(df)
```

---

## 8. QUICK REFERENCE: COMMON ADJUSTMENTS

```python
# Figure size (letter, A4, etc.)
plt.figure(figsize=(11, 8.5))   # Letter
plt.figure(figsize=(8.27, 11.69))  # A4 portrait

# Improve readability
plt.rcParams['font.size'] = 11
plt.rcParams['lines.linewidth'] = 2
plt.rcParams['axes.linewidth'] = 1.2

# Remove spines (clean look)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Add watermark or footer
fig.text(0.99, 0.01, 'Confidential - Storage Team', 
         transform=fig.transFigure, ha='right', fontsize=8, alpha=0.5)
```

---

## 9. BONUS: Dashboard-Style Layout

```python
fig = plt.figure(figsize=(16, 10))
fig.suptitle('Storage Dashboard', fontsize=16, fontweight='bold', y=0.98)

# Grid layout (3 rows, 2 columns)
gs = fig.add_gridspec(3, 2, hspace=0.3, wspace=0.3)

# Capacity chart
ax1 = fig.add_subplot(gs[0, :])
df.plot(x='timestamp', y='capacity_tb', ax=ax1, color='#2E86AB')
ax1.set_title('Capacity Utilization', fontweight='bold')

# IOPS chart
ax2 = fig.add_subplot(gs[1, 0])
df.plot(x='timestamp', y='iops_read', ax=ax2, color='#1B98A0')
df.plot(x='timestamp', y='iops_write', ax=ax2, color='#E9C46A')
ax2.set_title('IOPS Trend')

# Latency chart
ax3 = fig.add_subplot(gs[1, 1])
df.plot(x='timestamp', y='latency_ms', ax=ax3, color='#E76F51')
ax3.set_title('Latency (ms)')

# Utilization histogram
ax4 = fig.add_subplot(gs[2, 0])
df['pool_utilization_pct'].hist(bins=20, ax=ax4, color='#2E86AB', alpha=0.7)
ax4.axvline(80, color='orange', linestyle='--')
ax4.set_title('Utilization Distribution')

# Summary metrics as text
ax5 = fig.add_subplot(gs[2, 1])
ax5.axis('off')
stats_text = f"""
Key Metrics Summary:
• Avg Capacity: {df['capacity_tb'].mean():.1f} TB
• Peak IOPS Read: {df['iops_read'].max():,.0f}
• Peak IOPS Write: {df['iops_write'].max():,.0f}
• Avg Latency: {df['latency_ms'].mean():.1f} ms
• 95th %ile Latency: {df['latency_ms'].quantile(0.95):.1f} ms
• Current Utilization: {df['pool_utilization_pct'].iloc[-1]:.1f}%
"""
ax5.text(0.1, 0.5, stats_text, transform=ax5.transAxes, 
         fontsize=11, verticalalignment='center', fontfamily='monospace')
ax5.set_title('System Summary', fontweight='bold')

plt.tight_layout()
```

---

This cheat sheet gives you **production-ready patterns** for storage analytics. Save the snippets as a script or Jupyter notebook for reuse. Adjust colors, thresholds, and layouts to match your organization's branding.
