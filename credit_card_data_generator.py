import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

def generate_credit_card_origination_data(num_records=1000, start_date='2023-01-01', end_date='2023-12-31'):
    """
    Generate synthetic credit card origination data with realistic patterns.
    
    Args:
        num_records: Number of credit card applications to generate
        start_date: Start date for applications
        end_date: End date for applications
    
    Returns:
        pandas DataFrame with credit card origination data
    """
    np.random.seed(42)
    random.seed(42)
    
    # Date range
    start = datetime.strptime(start_date, '%Y-%m-%d')
    end = datetime.strptime(end_date, '%Y-%m-%d')
    date_range = (end - start).days
    
    # Generate data
    data = []
    
    # Credit card products
    products = ['Platinum', 'Gold', 'Classic', 'Secured', 'Student', 'Business']
    product_weights = [0.15, 0.25, 0.20, 0.10, 0.15, 0.15]
    
    # Credit score ranges with realistic distribution
    credit_score_ranges = [
        (300, 579, 0.15),   # Poor
        (580, 669, 0.25),   # Fair
        (670, 739, 0.35),   # Good
        (740, 799, 0.20),   # Very Good
        (800, 850, 0.05)    # Excellent
    ]
    
    # Income brackets
    income_brackets = [
        (20000, 40000, 0.25),
        (40000, 75000, 0.35),
        (75000, 125000, 0.25),
        (125000, 200000, 0.10),
        (200000, 500000, 0.05)
    ]
    
    # States with different approval rates
    states = ['CA', 'TX', 'FL', 'NY', 'IL', 'PA', 'OH', 'GA', 'NC', 'MI']
    state_approval_factors = {
        'CA': 1.1, 'TX': 1.05, 'FL': 0.95, 'NY': 0.98, 'IL': 1.02,
        'PA': 1.0, 'OH': 0.97, 'GA': 0.96, 'NC': 1.03, 'MI': 0.99
    }
    
    for i in range(num_records):
        # Application date (weighted towards recent months)
        days_offset = int(np.random.exponential(date_range/3))
        days_offset = min(days_offset, date_range)
        application_date = start + timedelta(days=days_offset)
        
        # Customer demographics
        age = np.random.normal(38, 12)
        age = max(18, min(80, int(age)))
        
        # Credit score
        credit_score = 0
        rand = np.random.random()
        cumulative = 0
        for min_score, max_score, weight in credit_score_ranges:
            cumulative += weight
            if rand <= cumulative:
                credit_score = np.random.randint(min_score, max_score + 1)
                break
        
        # Annual income
        rand = np.random.random()
        cumulative = 0
        min_income, max_income = 30000, 60000
        for min_inc, max_inc, weight in income_brackets:
            cumulative += weight
            if rand <= cumulative:
                min_income, max_income = min_inc, max_inc
                break
        annual_income = np.random.uniform(min_income, max_income)
        
        # State
        state = np.random.choice(states)
        
        # Product choice (influenced by credit score)
        if credit_score < 580:
            product_weights_adj = [0.05, 0.10, 0.15, 0.40, 0.20, 0.10]
        elif credit_score < 670:
            product_weights_adj = [0.10, 0.20, 0.30, 0.15, 0.15, 0.10]
        elif credit_score < 740:
            product_weights_adj = [0.15, 0.30, 0.25, 0.05, 0.10, 0.15]
        else:
            product_weights_adj = [0.25, 0.35, 0.15, 0.05, 0.05, 0.15]
        
        product = np.random.choice(products, p=product_weights_adj)
        
        # Credit limit requested
        if product == 'Secured':
            requested_limit = np.random.uniform(500, 5000)
        elif product == 'Student':
            requested_limit = np.random.uniform(1000, 8000)
        elif product == 'Classic':
            requested_limit = np.random.uniform(2000, 15000)
        elif product == 'Gold':
            requested_limit = np.random.uniform(5000, 25000)
        elif product == 'Platinum':
            requested_limit = np.random.uniform(10000, 50000)
        else:  # Business
            requested_limit = np.random.uniform(15000, 100000)
        
        # Approval decision (based on multiple factors)
        approval_probability = 0.5
        
        # Credit score impact
        if credit_score >= 740:
            approval_probability += 0.35
        elif credit_score >= 670:
            approval_probability += 0.25
        elif credit_score >= 580:
            approval_probability += 0.10
        else:
            approval_probability -= 0.20
        
        # Income impact
        if annual_income >= 100000:
            approval_probability += 0.15
        elif annual_income >= 50000:
            approval_probability += 0.10
        
        # Age impact
        if age >= 25 and age <= 65:
            approval_probability += 0.05
        
        # State factor
        approval_probability *= state_approval_factors[state]
        
        # Product type impact
        if product == 'Secured':
            approval_probability += 0.30
        elif product == 'Student':
            approval_probability -= 0.10
        
        approval_probability = max(0.1, min(0.95, approval_probability))
        approved = np.random.random() < approval_probability
        
        # Approved credit limit (if approved)
        if approved:
            if credit_score >= 740:
                approved_limit = requested_limit * np.random.uniform(0.9, 1.1)
            elif credit_score >= 670:
                approved_limit = requested_limit * np.random.uniform(0.7, 0.95)
            else:
                approved_limit = requested_limit * np.random.uniform(0.3, 0.7)
            
            approved_limit = min(approved_limit, 100000)  # Cap at 100k
        else:
            approved_limit = 0
        
        # Application channel
        channels = ['Online', 'Branch', 'Phone', 'Mobile App', 'Mail']
        channel_weights = [0.35, 0.25, 0.15, 0.20, 0.05]
        application_channel = np.random.choice(channels, p=channel_weights)
        
        # Processing time (days)
        if approved:
            processing_time = np.random.exponential(3) + 1
        else:
            processing_time = np.random.exponential(2) + 0.5
        processing_time = min(processing_time, 30)  # Cap at 30 days
        
        data.append({
            'Application_ID': f'CC{2023}{i:06d}',
            'Application_Date': application_date.strftime('%Y-%m-%d'),
            'Customer_Age': age,
            'Credit_Score': credit_score,
            'Annual_Income': annual_income,
            'State': state,
            'Product_Type': product,
            'Requested_Limit': requested_limit,
            'Approved_Limit': approved_limit,
            'Approved': approved,
            'Application_Channel': application_channel,
            'Processing_Time_Days': processing_time,
            'Month': application_date.strftime('%Y-%m'),
            'Quarter': f"Q{((application_date.month-1)//3)+1}"
        })
    
    return pd.DataFrame(data)

def create_credit_card_summary_stats(df):
    """Create summary statistics for the credit card data."""
    summary = {
        'total_applications': len(df),
        'approval_rate': df['Approved'].mean(),
        'avg_credit_score': df['Credit_Score'].mean(),
        'avg_income': df['Annual_Income'].mean(),
        'avg_approved_limit': df[df['Approved']]['Approved_Limit'].mean() if df['Approved'].any() else 0,
        'applications_by_product': df['Product_Type'].value_counts().to_dict(),
        'approval_by_product': df.groupby('Product_Type')['Approved'].mean().to_dict(),
        'applications_by_state': df['State'].value_counts().to_dict(),
        'applications_by_channel': df['Application_Channel'].value_counts().to_dict(),
        'monthly_trends': df.groupby('Month').size().to_dict()
    }
    return summary

if __name__ == "__main__":
    # Generate the dataset
    print("Generating credit card origination dataset...")
    df = generate_credit_card_origination_data(num_records=1000)
    
    # Save to Excel
    df.to_excel("credit_card_originations.xlsx", index=False)
    print(f"Generated {len(df)} credit card applications")
    print("Data saved to: credit_card_originations.xlsx")
    
    # Display summary statistics
    summary = create_credit_card_summary_stats(df)
    print("\n=== Summary Statistics ===")
    print(f"Total Applications: {summary['total_applications']}")
    print(f"Approval Rate: {summary['approval_rate']:.1%}")
    print(f"Average Credit Score: {summary['avg_credit_score']:.0f}")
    print(f"Average Annual Income: ${summary['avg_income']:,.0f}")
    print(f"Average Approved Limit: ${summary['avg_approved_limit']:,.0f}")
    
    print("\nApplications by Product:")
    for product, count in summary['applications_by_product'].items():
        print(f"  {product}: {count} ({summary['approval_by_product'][product]:.1%} approved)")
    
    print("\nTop 5 States by Applications:")
    for state, count in list(summary['applications_by_state'].items())[:5]:
        print(f"  {state}: {count}")
