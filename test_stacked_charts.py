#!/usr/bin/env python3
"""
Test the new stacked bar and stacked area chart functionality
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
from ppt_skill_generator import PowerPointSkillGenerator

def create_stacked_chart_data():
    """Create sample data suitable for stacked charts."""
    np.random.seed(42)
    
    # Monthly financial data for stacked area chart (100%)
    months = pd.date_range(start='2023-01-01', end='2023-12-31', freq='ME')
    financial_data = pd.DataFrame({
        'Month': months,
        'Revenue': np.cumsum(np.random.uniform(50, 150, len(months))) + 1000,
        'Expenses': np.cumsum(np.random.uniform(30, 80, len(months))) + 600,
        'Operating_Income': None,
        'Net_Income': None
    })
    financial_data['Operating_Income'] = financial_data['Revenue'] - financial_data['Expenses']
    financial_data['Net_Income'] = financial_data['Operating_Income'] - financial_data['Operating_Income'] * 0.1
    
    # Product performance data for stacked bar chart
    products = ['Wealth Management', 'Personal Banking', 'Commercial Banking', 'Capital Markets']
    quarters = ['Q1', 'Q2', 'Q3', 'Q4']
    
    product_data = []
    for product in products:
        for quarter in quarters:
            product_data.append({
                'Product': product,
                'Quarter': quarter,
                'Revenue': np.random.uniform(100, 500),
                'Profit': np.random.uniform(20, 100)
            })
    
    product_df = pd.DataFrame(product_data)
    
    # Profile distribution data (good for 100% area chart)
    age_groups = ['18-25', '26-35', '36-45', '46-55', '56-65', '65+']
    profile_data = []
    for age_group in age_groups:
        profile_data.append({
            'Age_Group': age_group,
            'Conservative': np.random.uniform(10, 30),
            'Moderate': np.random.uniform(20, 40),
            'Aggressive': np.random.uniform(15, 35),
            'Very_Aggressive': np.random.uniform(5, 20)
        })
    
    profile_df = pd.DataFrame(profile_data)
    
    return {
        'financial': financial_data,
        'products': product_df,
        'profiles': profile_df
    }

def test_stacked_charts():
    """Test the new stacked chart functionality."""
    print("=== Testing Stacked Charts with RBC Branding ===\n")
    
    # Create sample data
    datasets = create_stacked_chart_data()
    
    # Initialize the skill generator
    generator = PowerPointSkillGenerator()
    
    # Test 1: Stacked Bar Chart (100%)
    print("1. Creating Stacked Bar Chart (100%)...")
    
    # Pivot product data for stacked bar chart
    pivot_data = datasets['products'].pivot_table(
        index='Quarter',
        columns='Product',
        values='Revenue',
        aggfunc='sum',
        fill_value=0
    ).reset_index()
    
    # Calculate percentages
    numeric_cols = pivot_data.columns[1:]  # Skip Quarter column
    pivot_data[numeric_cols] = pivot_data[numeric_cols].div(pivot_data[numeric_cols].sum(axis=1), axis=0) * 100
    
    output1 = generator.create_custom_chart(
        chart_type='stacked_bar',
        data_source=pivot_data,
        config={
            'x_column': 'Quarter',
            'y_columns': numeric_cols.tolist()
        },
        output_path="rbc_stacked_bar_chart.pptx"
    )
    print(f"   Generated: {output1}")
    
    # Test 2: 100% Stacked Area Chart
    print("\n2. Creating 100% Stacked Area Chart...")
    
    # Prepare financial data for stacked area chart
    financial_data = datasets['financial'][['Month', 'Revenue', 'Expenses', 'Operating_Income', 'Net_Income']]
    
    output2 = generator.create_custom_chart(
        chart_type='stacked_area',
        data_source=financial_data,
        config={
            'x_column': 'Month',
            'y_columns': ['Revenue', 'Expenses', 'Operating_Income', 'Net_Income']
        },
        output_path="rbc_stacked_area_chart.pptx"
    )
    print(f"   Generated: {output2}")
    
    # Test 3: Profile Distribution (100% Area Chart)
    print("\n3. Creating Profile Distribution (100% Area Chart)...")
    
    output3 = generator.create_custom_chart(
        chart_type='stacked_area',
        data_source=datasets['profiles'],
        config={
            'x_column': 'Age_Group',
            'y_columns': ['Conservative', 'Moderate', 'Aggressive', 'Very_Aggressive']
        },
        output_path="rbc_profile_distribution.pptx"
    )
    print(f"   Generated: {output3}")
    
    return [output1, output2, output3]

def display_chart_info():
    """Display information about the new chart types."""
    print("\n" + "="*60)
    print("NEW STACKED CHART TYPES")
    print("="*60)
    
    print("\n📊 STACKED BAR CHART:")
    print("   • Vertical bars showing composition of multiple data series")
    print("   • Each bar shows 100% of the total")
    print("   • Good for comparing categories with sub-components")
    print("   • Uses RBC categorical color palette")
    
    print("\n📈 100% STACKED AREA CHART:")
    print("   • Area chart showing composition over time")
    print("   • Y-axis shows percentage (0-100%)")
    print("   • Good for profile distribution analysis")
    print("   • Uses RBC sequential color palette")
    
    print("\n🎯 USE CASES:")
    print("   • Financial performance breakdown")
    print("   • Customer profile distributions")
    print("   • Market share evolution")
    print("   • Risk profile analysis")

if __name__ == "__main__":
    print("PowerPoint Skill Generator - Stacked Charts Test")
    print("=" * 60)
    
    # Display chart information
    display_chart_info()
    
    # Test stacked charts
    generated_files = test_stacked_charts()
    
    print("\n" + "=" * 60)
    print("STACKED CHARTS TEST SUMMARY")
    print("=" * 60)
    
    print(f"✓ Generated {len(generated_files)} stacked chart presentations:")
    for file in generated_files:
        if os.path.exists(file):
            size = os.path.getsize(file) / 1024  # Size in KB
            print(f"  📄 {file} ({size:.1f} KB)")
    
    print("\n🎯 All stacked charts use RBC branding and show 100% composition!")
    print("📋 Perfect for understanding profile distributions and component breakdowns.")
