#!/usr/bin/env python3
"""
Test the PowerPoint Skill Generator with credit card origination data
"""

import pandas as pd
import os
from credit_card_data_generator import generate_credit_card_origination_data, create_credit_card_summary_stats
from ppt_skill_generator import PowerPointSkillGenerator

def test_credit_card_analysis():
    """Test the skill with credit card origination data."""
    print("=== Testing PowerPoint Skill Generator with Credit Card Data ===\n")
    
    # Generate synthetic credit card data
    print("1. Generating synthetic credit card origination data...")
    df = generate_credit_card_origination_data(num_records=1000)
    
    # Save the data
    data_file = "credit_card_originations.xlsx"
    df.to_excel(data_file, index=False)
    print(f"   Generated {len(df)} credit card applications")
    print(f"   Data saved to: {data_file}")
    
    # Display basic statistics
    summary = create_credit_card_summary_stats(df)
    print(f"\n   Approval Rate: {summary['approval_rate']:.1%}")
    print(f"   Average Credit Score: {summary['avg_credit_score']:.0f}")
    print(f"   Average Annual Income: ${summary['avg_income']:,.0f}")
    
    # Initialize the skill generator
    print("\n2. Initializing PowerPoint Skill Generator...")
    generator = PowerPointSkillGenerator()
    
    # Load data and get insights
    print("3. Analyzing data structure...")
    generator.load_data(data_file)
    insights = generator.get_data_insights()
    
    print(f"   Data Shape: {insights['shape']}")
    print(f"   Numeric Columns: {len(insights['numeric_columns'])}")
    print(f"   Categorical Columns: {len(insights['categorical_columns'])}")
    print(f"   Recommended Primary Chart: {insights['recommendations']['primary_chart']}")
    print(f"   Confidence: {insights['recommendations']['confidence']:.2f}")
    
    # Get all recommendations
    print("\n4. Chart Recommendations:")
    recommendations = generator.recommend_visualizations(max_charts=4)
    for i, rec in enumerate(recommendations, 1):
        print(f"   {i}. {rec['chart_type'].title()} Chart (Confidence: {rec.get('confidence', 0):.2f})")
        if 'x_column' in rec:
            print(f"      X: {rec['x_column']}, Y: {rec.get('y_column', 'N/A')}")
        elif 'column' in rec:
            print(f"      Column: {rec['column']}")
    
    # Create presentation
    print("\n5. Generating PowerPoint presentation...")
    output_path = generator.create_presentation(
        data_source=data_file,
        title="Credit Card Origination Analysis",
        output_path="credit_card_analysis.pptx"
    )
    
    print(f"   Presentation created: {output_path}")
    
    return output_path, insights, recommendations

def test_specific_analyses():
    """Test specific analyses on credit card data."""
    print("\n=== Testing Specific Analyses ===\n")
    
    # Generate data
    df = generate_credit_card_origination_data(num_records=1000)
    generator = PowerPointSkillGenerator()
    
    # Test 1: Approval analysis by product type
    print("1. Creating approval analysis by product type...")
    approval_by_product = df.groupby('Product_Type').agg({
        'Approved': ['count', 'sum', 'mean'],
        'Approved_Limit': 'mean'
    }).round(2)
    
    approval_by_product.columns = ['Total_Applications', 'Approved_Count', 'Approval_Rate', 'Avg_Approved_Limit']
    approval_by_product = approval_by_product.reset_index()
    
    # Save and analyze
    approval_file = "approval_by_product.xlsx"
    approval_by_product.to_excel(approval_file, index=False)
    
    output1 = generator.create_custom_chart(
        chart_type='bar',
        data_source=approval_file,
        config={
            'x_column': 'Product_Type',
            'y_column': 'Approval_Rate'
        },
        output_path="approval_rates_by_product.pptx"
    )
    print(f"   Created: {output1}")
    
    # Test 2: Credit score distribution
    print("2. Creating credit score distribution analysis...")
    credit_score_data = df[['Credit_Score', 'Approved']].copy()
    credit_score_file = "credit_score_distribution.xlsx"
    credit_score_data.to_excel(credit_score_file, index=False)
    
    output2 = generator.create_custom_chart(
        chart_type='histogram',
        data_source=credit_score_file,
        config={
            'column': 'Credit_Score',
            'bins': 20
        },
        output_path="credit_score_distribution.pptx"
    )
    print(f"   Created: {output2}")
    
    # Test 3: Monthly trends
    print("3. Creating monthly application trends...")
    monthly_data = df.groupby('Month').agg({
        'Application_ID': 'count',
        'Approved': 'sum'
    }).reset_index()
    monthly_data.columns = ['Month', 'Total_Applications', 'Approved_Applications']
    monthly_data['Approval_Rate'] = monthly_data['Approved_Applications'] / monthly_data['Total_Applications']
    
    monthly_file = "monthly_trends.xlsx"
    monthly_data.to_excel(monthly_file, index=False)
    
    output3 = generator.create_custom_chart(
        chart_type='line',
        data_source=monthly_file,
        config={
            'x_column': 'Month',
            'y_column': 'Total_Applications'
        },
        output_path="monthly_application_trends.pptx"
    )
    print(f"   Created: {output3}")
    
    return [output1, output2, output3]

def test_income_vs_credit_limit():
    """Test relationship between income and approved credit limits."""
    print("\n=== Testing Income vs Credit Limit Analysis ===\n")
    
    # Generate data
    df = generate_credit_card_origination_data(num_records=1000)
    
    # Filter for approved applications only
    approved_df = df[df['Approved']].copy()
    
    # Create income brackets
    approved_df['Income_Bracket'] = pd.cut(approved_df['Annual_Income'], 
                                          bins=[0, 40000, 75000, 125000, 200000, float('inf')],
                                          labels=['<40k', '40k-75k', '75k-125k', '125k-200k', '>200k'])
    
    # Aggregate by income bracket
    income_analysis = approved_df.groupby('Income_Bracket').agg({
        'Approved_Limit': ['mean', 'median', 'count'],
        'Credit_Score': 'mean'
    }).round(2)
    
    income_analysis.columns = ['Avg_Limit', 'Median_Limit', 'Count', 'Avg_Credit_Score']
    income_analysis = income_analysis.reset_index()
    
    # Save and analyze
    income_file = "income_vs_limit.xlsx"
    income_analysis.to_excel(income_file, index=False)
    
    generator = PowerPointSkillGenerator()
    output = generator.create_custom_chart(
        chart_type='bar',
        data_source=income_file,
        config={
            'x_column': 'Income_Bracket',
            'y_column': 'Avg_Limit'
        },
        output_path="income_vs_credit_limit.pptx"
    )
    
    print(f"Created income vs credit limit analysis: {output}")
    return output

def cleanup_test_files():
    """Clean up test files."""
    test_files = [
        "credit_card_originations.xlsx",
        "approval_by_product.xlsx",
        "credit_score_distribution.xlsx",
        "monthly_trends.xlsx",
        "income_vs_limit.xlsx"
    ]
    
    print("\n=== Cleaning Up Test Files ===")
    for file in test_files:
        if os.path.exists(file):
            os.remove(file)
            print(f"Removed: {file}")

def main():
    """Run all tests."""
    print("PowerPoint Skill Generator - Credit Card Data Test")
    print("=" * 60)
    
    try:
        # Main test
        main_output, insights, recommendations = test_credit_card_analysis()
        
        # Specific analyses
        specific_outputs = test_specific_analyses()
        
        # Income analysis
        income_output = test_income_vs_credit_limit()
        
        # Summary
        print("\n" + "=" * 60)
        print("TEST SUMMARY")
        print("=" * 60)
        print(f"✓ Main analysis presentation: {main_output}")
        print(f"✓ Specific analysis presentations: {len(specific_outputs)} created")
        print(f"✓ Income analysis presentation: {income_output}")
        print(f"✓ Total charts recommended: {len(recommendations)}")
        print(f"✓ Data quality score: {insights['recommendations']['data_quality']['completeness']:.1f}%")
        
        print("\nGenerated PowerPoint files:")
        all_outputs = [main_output] + specific_outputs + [income_output]
        for output in all_outputs:
            if os.path.exists(output):
                size = os.path.getsize(output) / 1024  # Size in KB
                print(f"  📄 {output} ({size:.1f} KB)")
        
        # Ask about cleanup
        print("\n" + "=" * 60)
        print("All tests completed successfully!")
        print("Check the generated .pptx files to see the visualizations.")
        
    except Exception as e:
        print(f"\n❌ Error during testing: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
