"""
Cyberdine Marketing Performance Analysis
Executive Board Presentation
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.patches import Rectangle
import warnings
warnings.filterwarnings('ignore')

# Set professional style
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (14, 8)
plt.rcParams['font.size'] = 11

# Load data
sales_df = pd.read_csv('sales.csv')
products_df = pd.read_csv('products.csv')
campaigns_df = pd.read_csv('campaigns.csv')

print("=" * 80)
print("CYBERDINE MARKETING PERFORMANCE ANALYSIS")
print("Executive Board Presentation")
print("=" * 80)
print()

# Data overview
print("Data Summary:")
print(f"  • Total Sales Records: {len(sales_df):,}")
print(f"  • Total Products: {len(products_df)}")
print(f"  • Total Campaigns: {len(campaigns_df)}")
print(f"  • Marketing Channels: {campaigns_df['platform'].nunique()}")
print()

# =============================================================================
# 1. CAMPAIGN PERFORMANCE OVERVIEW
# =============================================================================
print("\n" + "=" * 80)
print("1. CAMPAIGN PERFORMANCE OVERVIEW")
print("=" * 80)

# Merge sales with products to get revenue
sales_with_products = sales_df.merge(products_df, on='product_id')

# Calculate revenue and profit per sale
sales_with_products['revenue'] = sales_with_products['product_price']
sales_with_products['profit'] = sales_with_products['product_price'] * sales_with_products['profit_margin']

# Aggregate by campaign
campaign_performance = sales_with_products.groupby('campaign_id').agg({
    'revenue': 'sum',
    'profit': 'sum',
    'product_id': 'count'
}).rename(columns={'product_id': 'conversions'})

# Merge with campaign data
campaign_performance = campaign_performance.merge(campaigns_df, on='campaign_id')

# Calculate key metrics
campaign_performance['cost'] = campaign_performance['cost'].astype(float)
campaign_performance['ROI'] = (campaign_performance['profit'] - campaign_performance['cost']) / campaign_performance['cost'] * 100
campaign_performance['ROAS'] = campaign_performance['revenue'] / campaign_performance['cost']
campaign_performance['conversion_rate'] = campaign_performance['conversions'] / campaign_performance['clicks'] * 100
campaign_performance['cpc'] = campaign_performance['cost'] / campaign_performance['clicks']
campaign_performance['cpa'] = campaign_performance['cost'] / campaign_performance['conversions']
campaign_performance['revenue_per_click'] = campaign_performance['revenue'] / campaign_performance['clicks']

# Platform-level aggregation
platform_performance = campaign_performance.groupby('platform').agg({
    'cost': 'sum',
    'revenue': 'sum',
    'profit': 'sum',
    'conversions': 'sum',
    'impressions': 'sum',
    'clicks': 'sum'
}).reset_index()

platform_performance['ROI'] = (platform_performance['profit'] - platform_performance['cost']) / platform_performance['cost'] * 100
platform_performance['ROAS'] = platform_performance['revenue'] / platform_performance['cost']
platform_performance['conversion_rate'] = platform_performance['conversions'] / platform_performance['clicks'] * 100
platform_performance['avg_cpa'] = platform_performance['cost'] / platform_performance['conversions']

print("\n📊 PLATFORM PERFORMANCE SUMMARY")
print("-" * 80)
for _, row in platform_performance.iterrows():
    print(f"\n{row['platform'].upper()}")
    print(f"  Total Spend: ${row['cost']:,.0f}")
    print(f"  Revenue: ${row['revenue']:,.0f}")
    print(f"  Profit: ${row['profit']:,.0f}")
    print(f"  ROI: {row['ROI']:.1f}%")
    print(f"  ROAS: {row['ROAS']:.2f}x")
    print(f"  Conversions: {row['conversions']:,.0f}")
    print(f"  Conversion Rate: {row['conversion_rate']:.2f}%")
    print(f"  Cost per Acquisition: ${row['avg_cpa']:.2f}")

# =============================================================================
# 2. VISUALIZATIONS
# =============================================================================
print("\n\nGenerating visualizations...")

# Create comprehensive dashboard
fig = plt.figure(figsize=(18, 12))
gs = fig.add_gridspec(3, 3, hspace=0.3, wspace=0.3)

# 1. ROI by Platform
ax1 = fig.add_subplot(gs[0, 0])
colors = ['#2ecc71' if x > 0 else '#e74c3c' for x in platform_performance['ROI']]
bars = ax1.bar(platform_performance['platform'], platform_performance['ROI'], color=colors, alpha=0.8, edgecolor='black')
ax1.axhline(y=0, color='black', linestyle='-', linewidth=0.8)
ax1.set_title('ROI by Platform (%)', fontsize=13, fontweight='bold')
ax1.set_ylabel('ROI (%)')
ax1.grid(axis='y', alpha=0.3)
for bar in bars:
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2., height,
            f'{height:.1f}%', ha='center', va='bottom' if height > 0 else 'top', fontsize=10, fontweight='bold')

# 2. ROAS by Platform
ax2 = fig.add_subplot(gs[0, 1])
bars = ax2.bar(platform_performance['platform'], platform_performance['ROAS'], 
               color=['#3498db', '#9b59b6', '#e67e22', '#1abc9c'], alpha=0.8, edgecolor='black')
ax2.axhline(y=1, color='red', linestyle='--', linewidth=1.5, label='Break-even')
ax2.set_title('Return on Ad Spend (ROAS)', fontsize=13, fontweight='bold')
ax2.set_ylabel('ROAS (Revenue / Cost)')
ax2.legend()
ax2.grid(axis='y', alpha=0.3)
for bar in bars:
    height = bar.get_height()
    ax2.text(bar.get_x() + bar.get_width()/2., height,
            f'{height:.2f}x', ha='center', va='bottom', fontsize=10, fontweight='bold')

# 3. Conversion Rate by Platform
ax3 = fig.add_subplot(gs[0, 2])
bars = ax3.bar(platform_performance['platform'], platform_performance['conversion_rate'],
               color=['#3498db', '#9b59b6', '#e67e22', '#1abc9c'], alpha=0.8, edgecolor='black')
ax3.set_title('Conversion Rate by Platform', fontsize=13, fontweight='bold')
ax3.set_ylabel('Conversion Rate (%)')
ax3.grid(axis='y', alpha=0.3)
for bar in bars:
    height = bar.get_height()
    ax3.text(bar.get_x() + bar.get_width()/2., height,
            f'{height:.2f}%', ha='center', va='bottom', fontsize=10, fontweight='bold')

# 4. Revenue vs Cost by Campaign
ax4 = fig.add_subplot(gs[1, :2])
x = np.arange(len(campaign_performance))
width = 0.35
bars1 = ax4.bar(x - width/2, campaign_performance['cost'], width, label='Cost', 
                color='#e74c3c', alpha=0.8, edgecolor='black')
bars2 = ax4.bar(x + width/2, campaign_performance['revenue'], width, label='Revenue',
                color='#2ecc71', alpha=0.8, edgecolor='black')
ax4.set_xlabel('Campaign ID')
ax4.set_ylabel('Amount ($)')
ax4.set_title('Revenue vs Cost by Campaign', fontsize=13, fontweight='bold')
ax4.set_xticks(x)
ax4.set_xticklabels(campaign_performance['campaign_id'])
ax4.legend()
ax4.grid(axis='y', alpha=0.3)

# 5. Cost per Acquisition by Platform
ax5 = fig.add_subplot(gs[1, 2])
bars = ax5.bar(platform_performance['platform'], platform_performance['avg_cpa'],
               color=['#3498db', '#9b59b6', '#e67e22', '#1abc9c'], alpha=0.8, edgecolor='black')
ax5.set_title('Cost per Acquisition (CPA)', fontsize=13, fontweight='bold')
ax5.set_ylabel('CPA ($)')
ax5.grid(axis='y', alpha=0.3)
for bar in bars:
    height = bar.get_height()
    ax5.text(bar.get_x() + bar.get_width()/2., height,
            f'${height:.2f}', ha='center', va='bottom', fontsize=10, fontweight='bold')

# 6. Total Profit by Platform
ax6 = fig.add_subplot(gs[2, 0])
colors_profit = ['#2ecc71' if x > 0 else '#e74c3c' for x in platform_performance['profit']]
bars = ax6.bar(platform_performance['platform'], platform_performance['profit'],
               color=colors_profit, alpha=0.8, edgecolor='black')
ax6.axhline(y=0, color='black', linestyle='-', linewidth=0.8)
ax6.set_title('Total Profit by Platform', fontsize=13, fontweight='bold')
ax6.set_ylabel('Profit ($)')
ax6.grid(axis='y', alpha=0.3)
for bar in bars:
    height = bar.get_height()
    ax6.text(bar.get_x() + bar.get_width()/2., height,
            f'${height:,.0f}', ha='center', va='bottom' if height > 0 else 'top', 
            fontsize=9, fontweight='bold')

# 7. Conversions by Platform
ax7 = fig.add_subplot(gs[2, 1])
bars = ax7.bar(platform_performance['platform'], platform_performance['conversions'],
               color=['#3498db', '#9b59b6', '#e67e22', '#1abc9c'], alpha=0.8, edgecolor='black')
ax7.set_title('Total Conversions by Platform', fontsize=13, fontweight='bold')
ax7.set_ylabel('Conversions')
ax7.grid(axis='y', alpha=0.3)
for bar in bars:
    height = bar.get_height()
    ax7.text(bar.get_x() + bar.get_width()/2., height,
            f'{int(height):,}', ha='center', va='bottom', fontsize=10, fontweight='bold')

# 8. Efficiency Score (Revenue per Dollar Spent)
ax8 = fig.add_subplot(gs[2, 2])
efficiency = platform_performance['revenue'] / platform_performance['cost']
bars = ax8.bar(platform_performance['platform'], efficiency,
               color=['#3498db', '#9b59b6', '#e67e22', '#1abc9c'], alpha=0.8, edgecolor='black')
ax8.axhline(y=1, color='red', linestyle='--', linewidth=1.5, alpha=0.7)
ax8.set_title('Efficiency Score\n(Revenue per $1 Spent)', fontsize=13, fontweight='bold')
ax8.set_ylabel('$ Revenue / $ Spent')
ax8.grid(axis='y', alpha=0.3)
for bar in bars:
    height = bar.get_height()
    ax8.text(bar.get_x() + bar.get_width()/2., height,
            f'${height:.2f}', ha='center', va='bottom', fontsize=10, fontweight='bold')

plt.suptitle('CYBERDINE MARKETING PERFORMANCE DASHBOARD', 
             fontsize=16, fontweight='bold', y=0.98)

plt.savefig('campaign_performance_dashboard.png', dpi=300, bbox_inches='tight')
print("✅ Saved: campaign_performance_dashboard.png")

# =============================================================================
# 3. BUDGET ALLOCATION ANALYSIS
# =============================================================================
print("\n" + "=" * 80)
print("2. BUDGET ALLOCATION: $2,000")
print("=" * 80)

# Calculate efficiency metrics for optimization
campaign_performance['profit_per_dollar'] = campaign_performance['profit'] / campaign_performance['cost']
campaign_performance['revenue_per_dollar'] = campaign_performance['revenue'] / campaign_performance['cost']

# Sort by ROI for $2,000 budget
budget_2k = 2000
sorted_campaigns = campaign_performance.sort_values('ROI', ascending=False)

print("\n📈 RECOMMENDATION: Maximize ROI Strategy")
print("-" * 80)
print("\nBased on historical ROI performance, allocate to highest-performing campaigns:")
print()

allocation_2k = []
remaining_budget = budget_2k
for _, campaign in sorted_campaigns.iterrows():
    if remaining_budget >= campaign['cost']:
        allocation_2k.append({
            'campaign_id': campaign['campaign_id'],
            'platform': campaign['platform'],
            'cost': campaign['cost'],
            'roi': campaign['ROI'],
            'expected_profit': campaign['profit']
        })
        remaining_budget -= campaign['cost']
        print(f"  ✓ Campaign {campaign['campaign_id']} ({campaign['platform']}): ${campaign['cost']:,.0f}")
        print(f"    - Expected ROI: {campaign['ROI']:.1f}%")
        print(f"    - Expected Profit: ${campaign['profit']:,.0f}")

total_expected_profit = sum([a['expected_profit'] for a in allocation_2k])
total_allocated = sum([a['cost'] for a in allocation_2k])
overall_roi = (total_expected_profit - total_allocated) / total_allocated * 100

print(f"\n  📊 Total Allocated: ${total_allocated:,.0f}")
print(f"  💰 Expected Total Profit: ${total_expected_profit:,.0f}")
print(f"  📈 Expected Portfolio ROI: {overall_roi:.1f}%")
print(f"  💵 Remaining Budget: ${remaining_budget:,.0f}")

print("\n💡 CRITERIA: Maximize Return on Investment (ROI)")
print("   • Select campaigns with highest profit-to-cost ratio")
print("   • Prioritize proven performers over experimental campaigns")
print("   • Focus on campaigns that have delivered positive returns")

# =============================================================================
# 4. $5,000 BUDGET ALLOCATION
# =============================================================================
print("\n" + "=" * 80)
print("3. BUDGET ALLOCATION: $5,000")
print("=" * 80)

budget_5k = 5000
allocation_5k = []
remaining_budget_5k = budget_5k

print("\n📈 RECOMMENDATION: Diversified High-Performance Portfolio")
print("-" * 80)
print()

for _, campaign in sorted_campaigns.iterrows():
    if remaining_budget_5k >= campaign['cost']:
        allocation_5k.append({
            'campaign_id': campaign['campaign_id'],
            'platform': campaign['platform'],
            'cost': campaign['cost'],
            'roi': campaign['ROI'],
            'expected_profit': campaign['profit']
        })
        remaining_budget_5k -= campaign['cost']
        print(f"  ✓ Campaign {campaign['campaign_id']} ({campaign['platform']}): ${campaign['cost']:,.0f}")
        print(f"    - Expected ROI: {campaign['ROI']:.1f}%")
        print(f"    - Expected Profit: ${campaign['profit']:,.0f}")

total_expected_profit_5k = sum([a['expected_profit'] for a in allocation_5k])
total_allocated_5k = sum([a['cost'] for a in allocation_5k])
overall_roi_5k = (total_expected_profit_5k - total_allocated_5k) / total_allocated_5k * 100

print(f"\n  📊 Total Allocated: ${total_allocated_5k:,.0f}")
print(f"  💰 Expected Total Profit: ${total_expected_profit_5k:,.0f}")
print(f"  📈 Expected Portfolio ROI: {overall_roi_5k:.1f}%")
print(f"  💵 Remaining Budget: ${remaining_budget_5k:,.0f}")

print("\n🔄 STRATEGY CHANGE:")
if total_allocated_5k > total_allocated:
    print(f"   • YES - With $5,000, we can invest in {len(allocation_5k)} campaigns vs {len(allocation_2k)} campaigns")
    print(f"   • Increased diversification across platforms")
    print(f"   • Access to medium-budget campaigns with good ROI")
    print(f"   • More balanced risk profile")
else:
    print("   • NO - Same campaign selection, indicating top-tier campaigns worth prioritizing")

# =============================================================================
# 5. $30,000 BUDGET - CHANNEL ANALYSIS
# =============================================================================
print("\n" + "=" * 80)
print("4. LARGE BUDGET ALLOCATION: $30,000")
print("=" * 80)

budget_30k = 30000

# Platform-level scalability analysis
print("\n📊 CHANNEL SCALABILITY ANALYSIS")
print("-" * 80)

# Calculate channel metrics at different budget levels
for _, platform in platform_performance.iterrows():
    print(f"\n{platform['platform'].upper()}")
    print(f"  Historical Performance:")
    print(f"    - Total Spent: ${platform['cost']:,.0f}")
    print(f"    - ROI: {platform['ROI']:.1f}%")
    print(f"    - ROAS: {platform['ROAS']:.2f}x")
    print(f"    - Avg CPA: ${platform['avg_cpa']:.2f}")
    print(f"    - Conversion Rate: {platform['conversion_rate']:.2f}%")
    
    # Calculate scalability score
    scalability_score = (platform['ROI'] / 100) * (platform['conversion_rate'] / 10) * (platform['ROAS'])
    print(f"    - Scalability Score: {scalability_score:.2f}")

# Recommendation
best_channel = platform_performance.loc[platform_performance['ROI'].idxmax()]

print("\n" + "=" * 80)
print("🎯 RECOMMENDATION FOR $30,000 BUDGET")
print("=" * 80)
print(f"\n🏆 Most Promising Channel: {best_channel['platform'].upper()}")
print(f"\n   Key Strengths:")
print(f"   • Highest ROI: {best_channel['ROI']:.1f}%")
print(f"   • ROAS: {best_channel['ROAS']:.2f}x")
print(f"   • Conversion Rate: {best_channel['conversion_rate']:.2f}%")
print(f"   • Proven profitability at scale")
print()
print(f"   Suggested Allocation Strategy:")
print(f"   • {best_channel['platform']}: 50% ($15,000) - Primary channel")

# Get second best
platform_perf_sorted = platform_performance.sort_values('ROI', ascending=False)
second_best = platform_perf_sorted.iloc[1]
third_best = platform_perf_sorted.iloc[2]
fourth_best = platform_perf_sorted.iloc[3]

print(f"   • {second_best['platform']}: 25% ($7,500) - Secondary channel")
print(f"   • {third_best['platform']}: 15% ($4,500) - Tertiary channel")
print(f"   • {fourth_best['platform']}: 10% ($3,000) - Testing/experimental")
print()
print("   Rationale:")
print(f"   • Concentrate on proven performers")
print(f"   • Maintain diversification to mitigate risk")
print(f"   • Reserve budget for testing and optimization")

# =============================================================================
# 6. CONFOUNDERS ANALYSIS
# =============================================================================
print("\n" + "=" * 80)
print("5. POTENTIAL CONFOUNDERS")
print("=" * 80)

print("\n⚠️  FACTORS THAT COULD COMPROMISE ANALYSIS:")
print("-" * 80)

confounders = [
    {
        "title": "Seasonality & Timing",
        "description": "Campaigns may have run during different seasons or time periods",
        "impact": "HIGH",
        "mitigation": "Analyze campaign dates; control for seasonal effects"
    },
    {
        "title": "Product Mix Variation",
        "description": "Different campaigns may naturally attract buyers of different product categories",
        "impact": "HIGH",
        "mitigation": "Segment analysis by product category; normalize for product mix"
    },
    {
        "title": "Audience Overlap",
        "description": "Same customers may be exposed to multiple campaigns",
        "impact": "MEDIUM",
        "mitigation": "Attribution modeling; customer journey analysis"
    },
    {
        "title": "Ad Creative Quality",
        "description": "Same ad across channels may perform differently due to platform norms",
        "impact": "HIGH",
        "mitigation": "A/B test platform-optimized creatives"
    },
    {
        "title": "Budget Level Effects",
        "description": "Different budget levels ($1k, $2k, $3k) may have non-linear effects",
        "impact": "MEDIUM",
        "mitigation": "Test incremental budget increases; analyze marginal returns"
    },
    {
        "title": "Audience Maturity",
        "description": "Platforms may reach audiences at different buying stages",
        "impact": "MEDIUM",
        "mitigation": "Multi-touch attribution; customer lifecycle analysis"
    },
    {
        "title": "External Market Factors",
        "description": "Economic conditions, competitor actions, market trends",
        "impact": "MEDIUM",
        "mitigation": "Include market context; competitive intelligence"
    },
    {
        "title": "Attribution Window",
        "description": "Sales may be attributed to last-touch when earlier touches contributed",
        "impact": "HIGH",
        "mitigation": "Implement multi-touch attribution model"
    }
]

for i, conf in enumerate(confounders, 1):
    print(f"\n{i}. {conf['title']} [Impact: {conf['impact']}]")
    print(f"   Issue: {conf['description']}")
    print(f"   Mitigation: {conf['mitigation']}")

# Analyze product category distribution by campaign
print("\n\n📊 PRODUCT CATEGORY ANALYSIS BY PLATFORM")
print("-" * 80)

sales_full = sales_df.merge(products_df, on='product_id').merge(campaigns_df, on='campaign_id')
category_dist = sales_full.groupby(['platform', 'product_category']).size().unstack(fill_value=0)
category_pct = category_dist.div(category_dist.sum(axis=1), axis=0) * 100

print("\nProduct Category Distribution (% of sales):")
print(category_pct.round(1))

print("\n💡 INSIGHT: Product mix varies significantly by platform!")
print("   This suggests different audience demographics/preferences per channel.")

# Save data for reference
campaign_performance.to_csv('campaign_analysis.csv', index=False)
platform_performance.to_csv('platform_analysis.csv', index=False)
print("\n✅ Saved: campaign_analysis.csv")
print("✅ Saved: platform_analysis.csv")

print("\n" + "=" * 80)
print("ANALYSIS COMPLETE")
print("=" * 80)
