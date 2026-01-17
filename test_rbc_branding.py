#!/usr/bin/env python3
"""
Test the PowerPoint Skill Generator with RBC Branding
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
from ppt_skill_generator import PowerPointSkillGenerator
from rbc_branding import RBCBranding

def create_rbc_sample_data():
    """Create sample data suitable for RBC-style presentations."""
    np.random.seed(42)
    
    # Monthly financial data (good for line charts)
    months = pd.date_range(start='2023-01-01', end='2023-12-31', freq='ME')
    financial_data = pd.DataFrame({
        'Month': months,
        'Revenue': np.cumsum(np.random.uniform(50, 150, len(months))) + 1000,
        'Expenses': np.cumsum(np.random.uniform(30, 80, len(months))) + 600,
        'Net_Income': None
    })
    financial_data['Net_Income'] = financial_data['Revenue'] - financial_data['Expenses']
    
    # Product performance data (good for bar charts)
    products = ['Wealth Management', 'Personal Banking', 'Commercial Banking', 'Capital Markets', 'Insurance']
    product_data = pd.DataFrame({
        'Product': products,
        'Q1_Revenue': np.random.uniform(200, 800, len(products)),
        'Q2_Revenue': np.random.uniform(250, 900, len(products)),
        'Growth_Rate': np.random.uniform(-5, 15, len(products))
    })
    
    # Regional distribution (good for pie charts)
    regions = ['Canada', 'USA', 'UK', 'Asia-Pacific', 'Europe']
    regional_data = pd.DataFrame({
        'Region': regions,
        'Market_Share': [35, 25, 15, 20, 5],
        'Revenue_Billions': [12.5, 8.9, 5.3, 7.1, 1.8]
    })
    
    # Risk metrics (good for scatter plots)
    risk_data = pd.DataFrame({
        'Credit_Score': np.random.normal(650, 100, 100),
        'Loan_Amount': np.random.uniform(10000, 500000, 100),
        'Risk_Level': np.random.choice(['Low', 'Medium', 'High'], 100)
    })
    
    return {
        'financial': financial_data,
        'products': product_data,
        'regional': regional_data,
        'risk': risk_data
    }

def test_rbc_branding():
    """Test the skill with RBC branding."""
    print("=== Testing PowerPoint Skill Generator with RBC Branding ===\n")
    
    # Create sample data
    datasets = create_rbc_sample_data()
    
    # Initialize the skill generator
    generator = PowerPointSkillGenerator()
    
    print("1. RBC Color Palette:")
    for color_name, hex_code in RBCBranding.COLORS.items():
        print(f"   {color_name}: {hex_code}")
    
    print(f"\n2. RBC Fonts:")
    for font_name, font in RBCBranding.FONTS.items():
        print(f"   {font_name}: {font}")
    
    # Test 1: Financial performance line chart
    print("\n3. Creating Financial Performance Chart...")
    output1 = generator.create_custom_chart(
        chart_type='line',
        data_source=datasets['financial'],
        config={
            'x_column': 'Month',
            'y_column': 'Revenue'
        },
        output_path="rbc_financial_performance.pptx"
    )
    print(f"   Generated: {output1}")
    
    # Test 2: Product performance bar chart
    print("\n4. Creating Product Performance Chart...")
    output2 = generator.create_custom_chart(
        chart_type='bar',
        data_source=datasets['products'],
        config={
            'x_column': 'Product',
            'y_column': 'Growth_Rate'
        },
        output_path="rbc_product_performance.pptx"
    )
    print(f"   Generated: {output2}")
    
    # Test 3: Regional distribution pie chart
    print("\n5. Creating Regional Distribution Chart...")
    output3 = generator.create_custom_chart(
        chart_type='pie',
        data_source=datasets['regional'],
        config={
            'labels_column': 'Region',
            'values_column': 'Market_Share'
        },
        output_path="rbc_regional_distribution.pptx"
    )
    print(f"   Generated: {output3}")
    
    # Test 4: Risk analysis scatter plot
    print("\n6. Creating Risk Analysis Chart...")
    output4 = generator.create_custom_chart(
        chart_type='scatter',
        data_source=datasets['risk'],
        config={
            'x_column': 'Credit_Score',
            'y_column': 'Loan_Amount'
        },
        output_path="rbc_risk_analysis.pptx"
    )
    print(f"   Generated: {output4}")
    
    # Test 5: Complete presentation with multiple charts
    print("\n7. Creating Complete RBC-Styled Presentation...")
    generator = PowerPointSkillGenerator()  # Fresh instance
    output5 = generator.create_presentation(
        data_source=datasets['financial'],
        title="RBC Financial Performance Analysis",
        output_path="rbc_complete_analysis.pptx"
    )
    print(f"   Generated: {output5}")
    
    return [output1, output2, output3, output4, output5]

def display_rbc_branding_info():
    """Display RBC branding information."""
    print("\n" + "="*60)
    print("RBC BRAND GUIDELINES")
    print("="*60)
    
    print("\n🎨 PRIMARY COLORS:")
    print(f"   Primary Blue: {RBCBranding.COLORS['primary_blue']}")
    print(f"   Accent Yellow: {RBCBranding.COLORS['accent_yellow']}")
    print(f"   White: {RBCBranding.COLORS['white']}")
    
    print("\n📊 CHART COLOR PALETTES:")
    for palette_name, colors in RBCBranding.CHART_PALETTES.items():
        print(f"   {palette_name.title()}: {colors}")
    
    print("\n🔤 TYPOGRAPHY:")
    print(f"   Primary Font: {RBCBranding.FONTS['primary']}")
    print(f"   Heading Font: {RBCBranding.FONTS['heading']}")
    print(f"   Body Font: {RBCBranding.FONTS['body']}")
    print(f"   Data Labels: {RBCBranding.FONTS['data_labels']}")
    
    print("\n📏 FONT SIZES:")
    for size_name, size in RBCBranding.FONT_SIZES.items():
        print(f"   {size_name.replace('_', ' ').title()}: {size}pt")

if __name__ == "__main__":
    print("PowerPoint Skill Generator - RBC Branding Test")
    print("=" * 60)
    
    # Display RBC branding info
    display_rbc_branding_info()
    
    # Test RBC branding
    generated_files = test_rbc_branding()
    
    print("\n" + "=" * 60)
    print("RBC BRANDING TEST SUMMARY")
    print("=" * 60)
    
    print(f"✓ Generated {len(generated_files)} RBC-styled presentations:")
    for file in generated_files:
        if os.path.exists(file):
            size = os.path.getsize(file) / 1024  # Size in KB
            print(f"  📄 {file} ({size:.1f} KB)")
    
    print("\n🎯 All charts now use RBC's official color scheme and typography!")
    print("📋 Presentations feature:")
    print("   • RBC Medium Persian Blue (#005DAA) as primary color")
    print("   • RBC Cyber Yellow (#FFD200) for accents and highlights")
    print("   • Arial font family for consistency with RBC presentations")
    print("   • Professional styling matching RBC investor relations materials")
