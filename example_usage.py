#!/usr/bin/env python3
"""
Example usage of the PowerPoint Skill Generator
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
from ppt_skill_generator import PowerPointSkillGenerator

def create_sample_data():
    """Create sample datasets for demonstration."""
    
    # Sales data (good for bar charts)
    np.random.seed(42)
    sales_data = pd.DataFrame({
        'Region': ['North', 'South', 'East', 'West', 'Central'] * 4,
        'Product': ['A', 'B', 'C', 'D'] * 5,
        'Sales': np.random.randint(1000, 10000, 20),
        'Quarter': ['Q1', 'Q2', 'Q3', 'Q4'] * 5
    })
    
    # Time series data (good for line charts)
    dates = pd.date_range(start='2023-01-01', end='2023-12-31', freq='M')
    time_series_data = pd.DataFrame({
        'Date': dates,
        'Revenue': np.cumsum(np.random.randint(5000, 15000, len(dates))),
        'Expenses': np.cumsum(np.random.randint(3000, 8000, len(dates))),
        'Profit': None
    })
    time_series_data['Profit'] = time_series_data['Revenue'] - time_series_data['Expenses']
    
    # Customer demographics (good for pie charts)
    demographics = pd.DataFrame({
        'Age_Group': ['18-25', '26-35', '36-45', '46-55', '56+'],
        'Customer_Count': [1200, 2800, 2400, 1800, 800]
    })
    
    # Performance metrics (good for scatter plots)
    performance = pd.DataFrame({
        'Experience_Years': np.random.uniform(1, 20, 50),
        'Performance_Score': 50 + np.random.uniform(0, 50, 50) + np.random.normal(0, 5, 50),
        'Department': np.random.choice(['Sales', 'Marketing', 'IT', 'HR'], 50)
    })
    
    return {
        'sales': sales_data,
        'time_series': time_series_data,
        'demographics': demographics,
        'performance': performance
    }

def example_basic_usage():
    """Basic usage example."""
    print("=== Basic Usage Example ===")
    
    # Create sample data
    datasets = create_sample_data()
    
    # Initialize the skill generator
    generator = PowerPointSkillGenerator()
    
    # Create presentation from sales data
    output_path = generator.create_presentation(
        data_source=datasets['sales'],
        title="Sales Analysis Report",
        output_path="sales_analysis.pptx"
    )
    
    print(f"Generated presentation: {output_path}")
    
    # Get data insights
    insights = generator.get_data_insights()
    print(f"Recommended chart: {insights['recommendations']['primary_chart']}")
    print(f"Confidence: {insights['recommendations']['confidence']:.2f}")

def example_custom_charts():
    """Example with custom chart configurations."""
    print("\n=== Custom Charts Example ===")
    
    datasets = create_sample_data()
    generator = PowerPointSkillGenerator()
    
    # Create custom bar chart
    output_path = generator.create_custom_chart(
        chart_type='bar',
        data_source=datasets['demographics'],
        config={
            'x_column': 'Age_Group',
            'y_column': 'Customer_Count'
        },
        output_path="customer_demographics.pptx"
    )
    
    print(f"Generated custom chart: {output_path}")

def example_time_series():
    """Example with time series data."""
    print("\n=== Time Series Example ===")
    
    datasets = create_sample_data()
    generator = PowerPointSkillGenerator()
    
    # Create presentation from time series data
    output_path = generator.create_presentation(
        data_source=datasets['time_series'],
        title="Financial Performance 2023",
        output_path="financial_performance.pptx",
        target_column='Revenue'  # Focus on Revenue column
    )
    
    print(f"Generated time series presentation: {output_path}")

def example_batch_processing():
    """Example of batch processing multiple files."""
    print("\n=== Batch Processing Example ===")
    
    # Create sample data files
    datasets = create_sample_data()
    
    # Save datasets to files
    file_paths = []
    for name, data in datasets.items():
        file_path = f"{name}_data.xlsx"
        data.to_excel(file_path, index=False)
        file_paths.append(file_path)
    
    # Process all files
    generator = PowerPointSkillGenerator()
    generated_files = generator.batch_process(
        data_files=file_paths,
        output_dir="batch_output"
    )
    
    print(f"Generated {len(generated_files)} presentations in batch_output/")
    
    # Clean up sample files
    for file_path in file_paths:
        if os.path.exists(file_path):
            os.remove(file_path)

def example_from_file():
    """Example loading data from file."""
    print("\n=== File Loading Example ===")
    
    # Create a sample CSV file
    sample_data = pd.DataFrame({
        'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        'Sales': [12000, 15000, 18000, 14000, 20000, 22000],
        'Expenses': [8000, 9000, 10000, 8500, 11000, 12000]
    })
    
    csv_file = "monthly_data.csv"
    sample_data.to_csv(csv_file, index=False)
    
    # Load from file and create presentation
    generator = PowerPointSkillGenerator()
    output_path = generator.create_presentation(
        data_source=csv_file,
        title="Monthly Performance",
        output_path="monthly_performance.pptx"
    )
    
    print(f"Generated presentation from CSV: {output_path}")
    
    # Clean up
    if os.path.exists(csv_file):
        os.remove(csv_file)

if __name__ == "__main__":
    print("PowerPoint Skill Generator - Example Usage")
    print("=" * 50)
    
    # Run all examples
    example_basic_usage()
    example_custom_charts()
    example_time_series()
    example_batch_processing()
    example_from_file()
    
    print("\n" + "=" * 50)
    print("All examples completed! Check the generated .pptx files.")
    print("\nTo install dependencies, run:")
    print("pip install -r requirements.txt")
