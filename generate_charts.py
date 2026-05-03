"""
DMart Opportunities Analysis - Chart Generation
Generates all visualizations for the 10-slide presentation
"""
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import os
import numpy as np

# Setup
sns.set_theme(style="whitegrid")
COLORS = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#3B1F2B', '#44BBA4', '#E94F37', '#393E41', '#8D6A9F', '#5BC0EB']
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'charts')
os.makedirs(OUTPUT_DIR, exist_ok=True)

df = pd.read_excel(os.path.join(os.path.dirname(__file__), 'DMart Data Store.xlsx'), sheet_name='Order Data')
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.to_period('M')
df['Profit Margin'] = (df['Profit'] / df['Sales'] * 100).round(2)

def save(fig, name):
    fig.savefig(os.path.join(OUTPUT_DIR, name), dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Saved {name}")

# ---- CHART 1: Total Sales & Profit by Segment (Pie + Bar) ----
print("Chart 1: Segment overview...")
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
seg = df.groupby('Segment').agg({'Sales':'sum','Profit':'sum'}).sort_values('Sales', ascending=False)
axes[0].pie(seg['Sales'], labels=seg.index, autopct='%1.1f%%', colors=COLORS[:3], startangle=90, textprops={'fontsize':11})
axes[0].set_title('Sales Share by Segment', fontsize=13, fontweight='bold')
seg[['Sales','Profit']].plot(kind='bar', ax=axes[1], color=[COLORS[0], COLORS[1]])
axes[1].set_title('Sales & Profit by Segment', fontsize=13, fontweight='bold')
axes[1].set_ylabel('Amount')
axes[1].tick_params(axis='x', rotation=0)
axes[1].yaxis.set_major_formatter(mticker.FuncFormatter(lambda x,p: f'{x/1e6:.1f}M' if x>=1e6 else f'{x/1e3:.0f}K'))
fig.suptitle('Segment Performance Overview', fontsize=15, fontweight='bold', y=1.02)
fig.tight_layout()
save(fig, 'chart1_segment_overview.png')

# ---- CHART 2: Category spend share by Segment ----
print("Chart 2: Category by segment...")
fig, ax = plt.subplots(figsize=(10, 6))
ct = df.pivot_table(values='Sales', index='Segment', columns='Category', aggfunc='sum')
ct_pct = ct.div(ct.sum(axis=1), axis=0) * 100
ct_pct.plot(kind='barh', stacked=True, ax=ax, color=COLORS[:3])
ax.set_title('Category Spend Share by Segment', fontsize=14, fontweight='bold')
ax.set_xlabel('% of Total Sales')
ax.legend(title='Category', bbox_to_anchor=(1.01, 1))
for bars in ax.containers:
    ax.bar_label(bars, fmt='%.1f%%', label_type='center', fontsize=9)
fig.tight_layout()
save(fig, 'chart2_category_by_segment.png')

# ---- CHART 3: Geographic Spend - Top 10 Countries ----
print("Chart 3: Geographic spend...")
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
country = df.groupby('Country').agg({'Sales':'sum','Profit':'sum','Order ID':'nunique'}).sort_values('Sales', ascending=True)
country['Sales'].plot(kind='barh', ax=axes[0], color=COLORS[0])
axes[0].set_title('Total Sales by Country', fontsize=13, fontweight='bold')
axes[0].set_xlabel('Sales')
axes[0].xaxis.set_major_formatter(mticker.FuncFormatter(lambda x,p: f'{x/1e3:.0f}K'))
region = df.groupby('Region').agg({'Sales':'sum','Profit':'sum'})
axes[1].pie(region['Sales'], labels=region.index, autopct='%1.1f%%', colors=COLORS[:3], startangle=90, textprops={'fontsize':12})
axes[1].set_title('Sales by Region', fontsize=13, fontweight='bold')
fig.suptitle('Geographic Distribution of Sales', fontsize=15, fontweight='bold', y=1.02)
fig.tight_layout()
save(fig, 'chart3_geographic_sales.png')

# ---- CHART 4: Year-over-Year Growth Trend ----
print("Chart 4: YoY growth...")
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
yearly = df.groupby('Year').agg({'Sales':'sum','Profit':'sum','Order ID':'nunique'}).rename(columns={'Order ID':'Orders'})
yearly[['Sales','Profit']].plot(kind='bar', ax=axes[0], color=[COLORS[0], COLORS[1]])
axes[0].set_title('Sales & Profit by Year', fontsize=13, fontweight='bold')
axes[0].set_ylabel('Amount')
axes[0].tick_params(axis='x', rotation=0)
axes[0].yaxis.set_major_formatter(mticker.FuncFormatter(lambda x,p: f'{x/1e3:.0f}K'))
yoy_growth = yearly['Sales'].pct_change() * 100
yoy_growth.dropna().plot(kind='bar', ax=axes[1], color=COLORS[2])
axes[1].set_title('Year-over-Year Sales Growth %', fontsize=13, fontweight='bold')
axes[1].set_ylabel('Growth %')
axes[1].tick_params(axis='x', rotation=0)
for i, v in enumerate(yoy_growth.dropna()):
    axes[1].text(i, v + 0.5, f'{v:.1f}%', ha='center', fontweight='bold')
fig.suptitle('Sales Growth Trend (2011-2014)', fontsize=15, fontweight='bold', y=1.02)
fig.tight_layout()
save(fig, 'chart4_yoy_growth.png')

# ---- CHART 5: Sub-Category Opportunity Matrix (Sales vs Profit Margin) ----
print("Chart 5: Sub-category opportunity...")
fig, ax = plt.subplots(figsize=(12, 7))
sub = df.groupby('Sub-Category').agg({'Sales':'sum','Profit':'sum','Quantity':'sum'})
sub['Margin'] = (sub['Profit'] / sub['Sales'] * 100)
scatter = ax.scatter(sub['Sales'], sub['Margin'], s=sub['Quantity']/5, c=range(len(sub)), cmap='viridis', alpha=0.7, edgecolors='black')
for i, txt in enumerate(sub.index):
    ax.annotate(txt, (sub['Sales'].iloc[i], sub['Margin'].iloc[i]), fontsize=8, ha='center', va='bottom')
ax.axhline(y=sub['Margin'].median(), color='red', linestyle='--', alpha=0.5, label=f'Median Margin ({sub["Margin"].median():.1f}%)')
ax.axvline(x=sub['Sales'].median(), color='blue', linestyle='--', alpha=0.5, label=f'Median Sales ({sub["Sales"].median():,.0f})')
ax.set_xlabel('Total Sales', fontsize=12)
ax.set_ylabel('Profit Margin %', fontsize=12)
ax.set_title('Sub-Category Opportunity Matrix\n(Bubble size = Quantity sold)', fontsize=14, fontweight='bold')
ax.legend()
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x,p: f'{x/1e3:.0f}K'))
fig.tight_layout()
save(fig, 'chart5_opportunity_matrix.png')

# ---- CHART 6: Top 10 Cities by Sales with Profit Overlay ----
print("Chart 6: Top cities...")
fig, ax = plt.subplots(figsize=(11, 6))
city = df.groupby('City').agg({'Sales':'sum','Profit':'sum'}).sort_values('Sales', ascending=False).head(10)
x = range(len(city))
w = 0.35
ax.bar([i-w/2 for i in x], city['Sales'], w, label='Sales', color=COLORS[0])
ax.bar([i+w/2 for i in x], city['Profit'], w, label='Profit', color=COLORS[1])
ax.set_xticks(list(x))
ax.set_xticklabels(city.index, rotation=45, ha='right')
ax.set_title('Top 10 Cities: Sales vs Profit', fontsize=14, fontweight='bold')
ax.set_ylabel('Amount')
ax.legend()
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x,p: f'{x/1e3:.0f}K'))
fig.tight_layout()
save(fig, 'chart6_top_cities.png')

# ---- CHART 7: Discount Impact on Profitability ----
print("Chart 7: Discount impact...")
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
df['Discount_Bin'] = pd.cut(df['Discount'], bins=[-.01, 0, 0.1, 0.2, 0.3, 1.0], labels=['0%','1-10%','11-20%','21-30%','30%+'])
disc = df.groupby('Discount_Bin', observed=True).agg({'Profit':'mean','Sales':'sum'})
disc['Profit'].plot(kind='bar', ax=axes[0], color=[COLORS[3] if v<0 else COLORS[4] for v in disc['Profit']])
axes[0].set_title('Avg Profit by Discount Level', fontsize=13, fontweight='bold')
axes[0].set_ylabel('Avg Profit')
axes[0].tick_params(axis='x', rotation=0)
axes[0].axhline(y=0, color='black', linewidth=0.5)
seg_disc = df.groupby(['Segment','Discount_Bin'], observed=True)['Profit'].mean().unstack()
seg_disc.plot(kind='bar', ax=axes[1], color=COLORS[:5])
axes[1].set_title('Avg Profit by Segment & Discount', fontsize=13, fontweight='bold')
axes[1].set_ylabel('Avg Profit')
axes[1].tick_params(axis='x', rotation=0)
axes[1].axhline(y=0, color='black', linewidth=0.5)
axes[1].legend(title='Discount', fontsize=8)
fig.suptitle('Discount Impact Analysis', fontsize=15, fontweight='bold', y=1.02)
fig.tight_layout()
save(fig, 'chart7_discount_impact.png')

# ---- CHART 8: Segment Growth by Region (Heatmap) ----
print("Chart 8: Segment-region heatmap...")
fig, ax = plt.subplots(figsize=(10, 5))
sr = df.pivot_table(values='Sales', index='Region', columns='Segment', aggfunc='sum')
sns.heatmap(sr, annot=True, fmt=',.0f', cmap='YlOrRd', ax=ax, linewidths=1)
ax.set_title('Sales Heatmap: Region x Segment', fontsize=14, fontweight='bold')
fig.tight_layout()
save(fig, 'chart8_region_segment_heatmap.png')

# ---- CHART 9: Corporate & Home Office Penetration Opportunity ----
print("Chart 9: Penetration opportunity...")
fig, ax = plt.subplots(figsize=(11, 6))
seg_country = df.pivot_table(values='Sales', index='Country', columns='Segment', aggfunc='sum', fill_value=0)
seg_country['Total'] = seg_country.sum(axis=1)
seg_country = seg_country.sort_values('Total', ascending=True)
seg_country[['Consumer','Corporate','Home Office']].plot(kind='barh', stacked=True, ax=ax, color=COLORS[:3])
ax.set_title('Segment Penetration by Country\n(Adjacent Opportunity in Corporate & Home Office)', fontsize=13, fontweight='bold')
ax.set_xlabel('Total Sales')
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x,p: f'{x/1e3:.0f}K'))
ax.legend(title='Segment')
fig.tight_layout()
save(fig, 'chart9_penetration_opportunity.png')

# ---- CHART 10: Key Metrics Summary Dashboard ----
print("Chart 10: Summary dashboard...")
fig, axes = plt.subplots(2, 2, figsize=(12, 8))
# Top customers
top_cust = df.groupby('Customer Name')['Sales'].sum().sort_values(ascending=False).head(10)
top_cust.plot(kind='barh', ax=axes[0,0], color=COLORS[0])
axes[0,0].set_title('Top 10 Customers by Sales', fontsize=12, fontweight='bold')
axes[0,0].invert_yaxis()
# Ship mode
ship = df.groupby('Ship Mode')['Sales'].sum().sort_values()
ship.plot(kind='barh', ax=axes[0,1], color=COLORS[2])
axes[0,1].set_title('Sales by Ship Mode', fontsize=12, fontweight='bold')
# Monthly trend
monthly = df.groupby(df['Order Date'].dt.to_period('Q'))['Sales'].sum()
monthly.index = monthly.index.astype(str)
axes[1,0].plot(range(len(monthly)), monthly.values, color=COLORS[0], linewidth=2)
axes[1,0].fill_between(range(len(monthly)), monthly.values, alpha=0.2, color=COLORS[0])
axes[1,0].set_title('Quarterly Sales Trend', fontsize=12, fontweight='bold')
axes[1,0].set_xticks(range(0, len(monthly), 2))
axes[1,0].set_xticklabels([monthly.index[i] for i in range(0, len(monthly), 2)], rotation=45, fontsize=7)
# Feedback
fb = df['Feedback?'].value_counts()
axes[1,1].pie(fb, labels=['No Feedback','Has Feedback'], autopct='%1.1f%%', colors=[COLORS[3], COLORS[4]], startangle=90)
axes[1,1].set_title('Customer Feedback Rate', fontsize=12, fontweight='bold')
fig.suptitle('Key Business Metrics Dashboard', fontsize=15, fontweight='bold')
fig.tight_layout()
save(fig, 'chart10_dashboard.png')

print("\nAll charts generated successfully!")
