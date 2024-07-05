import pandas as pd

# Define the path to the .xls file
file_path = 'task.xls'

# Load the .xls file with all sheets
sheets = pd.read_excel(file_path, sheet_name=None)

# Access individual sheets by their names
customer_acquisition = sheets['Customer Acqusition']
spend = sheets['Spend']
repayment = sheets['Repayment']

# Cleaning:
# Convert 'Month' to datetime in spend dataframe
spend['Month'] = pd.to_datetime(spend['Month'])

# Group spend by customer and month
monthly_spend = spend.groupby(['Costomer', spend['Month'].dt.to_period('M')])['Amount'].sum().reset_index()

# Merge monthly spend with customer acquisition data to get credit limits
merged_data = pd.merge(monthly_spend, customer_acquisition[['Customer', 'Limit']], 
                       left_on='Costomer', right_on='Customer', how='left')

# Calculate the difference between spend and limit
merged_data['Overspend'] = merged_data['Amount'] - merged_data['Limit']

# Filter for customers who have overspent
overspent = merged_data[merged_data['Overspend'] > 0].sort_values('Overspend', ascending=False)
print("Cleaning")
print(overspent)


# Task 1: Monthly spend of each customer
spend['Month'] = pd.to_datetime(spend['Month'])
spend['Year-Month'] = spend['Month'].dt.to_period('M')

# Pivot the spend data to have Year-Month as columns
monthly_spend = spend.pivot_table(values='Amount', 
                                  index='Costomer', 
                                  columns='Year-Month', 
                                  aggfunc='sum', 
                                  fill_value=0).reset_index()

# Display the first few rows of the result
print("Task 1: Monthly spend of each customer")
print(monthly_spend.head())

# Task 2: Monthly repayment of each customer
repayment['Month'] = pd.to_datetime(repayment['Month'])
repayment['Year-Month'] = repayment['Month'].dt.to_period('M')

# Pivot the repayment data to have Year-Month as columns
monthly_repayment = repayment.pivot_table(values='Amount', 
                                          index='Costomer', 
                                          columns='Year-Month', 
                                          aggfunc='sum', 
                                          fill_value=0).reset_index()

# Display the first few rows of the result
print("Task 2: Monthly repayment of each customer")
print(monthly_repayment.head())

# Task 3: Highest paying 10 customers
total_repayment = repayment.groupby('Costomer')['Amount'].sum().reset_index()
total_repayment_sorted = total_repayment.sort_values('Amount', ascending=False)
top_10_customers = total_repayment_sorted.head(10)
print("Task 3: Highest paying 10 customers")
print(top_10_customers.to_string())

# Task 4: People in which segment are spending more money
merged_data = pd.merge(spend, customer_acquisition[['Customer', 'Segment']], 
                       left_on='Costomer', right_on='Customer', how='left')
segment_spend = merged_data.groupby('Segment')['Amount'].sum().reset_index()
segment_spend_sorted = segment_spend.sort_values('Amount', ascending=False)
highest_spending_segment = segment_spend_sorted.iloc[0]
print("Task 4: People in which segment are spending more money")
print(f"\nThe segment spending the most money is '{highest_spending_segment['Segment']}' "
      f"with a total spend of {highest_spending_segment['Amount']:.2f}")

# Task 5: Which age group is spending more money
def age_group(age):
    if age < 20:
        return '< 20'
    elif 20 <= age < 30:
        return '20-29'
    elif 30 <= age < 40:
        return '30-39'
    elif 40 <= age < 50:
        return '40-49'
    elif 50 <= age < 60:
        return '50-59'
    else:
        return '60+'

customer_acquisition['AgeGroup'] = customer_acquisition['Age'].apply(age_group)
merged_data = pd.merge(spend, customer_acquisition[['Customer', 'AgeGroup']], 
                       left_on='Costomer', right_on='Customer', how='left')
age_group_spend = merged_data.groupby('AgeGroup')['Amount'].sum().reset_index()
age_group_spend_sorted = age_group_spend.sort_values('Amount', ascending=False)
highest_spending_age_group = age_group_spend_sorted.iloc[0]
print("Task 5: Which age group is spending more money")
print(f"\nThe age group spending the most money is '{highest_spending_age_group['AgeGroup']}' "
      f"with a total spend of {highest_spending_age_group['Amount']:.2f}")

# Task 6: Which is the most profitable segment
merged_spend = pd.merge(spend, customer_acquisition[['Customer', 'Segment']], 
                        left_on='Costomer', right_on='Customer', how='left')
merged_repayment = pd.merge(repayment, customer_acquisition[['Customer', 'Segment']], 
                            left_on='Costomer', right_on='Customer', how='left')
segment_spend = merged_spend.groupby('Segment')['Amount'].sum().reset_index()
segment_repayment = merged_repayment.groupby('Segment')['Amount'].sum().reset_index()
segment_profit = pd.merge(segment_spend, segment_repayment, on='Segment', suffixes=('_spend', '_repayment'))
segment_profit['Profit'] = segment_profit['Amount_repayment'] - segment_profit['Amount_spend']
segment_profit_sorted = segment_profit.sort_values('Profit', ascending=False)
most_profitable_segment = segment_profit_sorted.iloc[0]
print("Task 6: Which is the most profitable segment")
print(f"\nThe most profitable segment is '{most_profitable_segment['Segment']}' "
      f"with a profit of {most_profitable_segment['Profit']:.2f}")

# Task 7: In which category the customers are spending more money
category_spend = spend.groupby('Type')['Amount'].sum().reset_index()
category_spend_sorted = category_spend.sort_values('Amount', ascending=False)
highest_spending_category = category_spend_sorted.iloc[0]
print("Task 7: In which category the customers are spending more money")
print(f"\nThe category with the highest spending is '{highest_spending_category['Type']}' "
      f"with a total spend of {highest_spending_category['Amount']:.2f}")

# Task 8: Impose an interest rate of 2.9% for each customer for any due amount
customer_spend = spend.groupby('Costomer')['Amount'].sum().reset_index()
customer_repayment = repayment.groupby('Costomer')['Amount'].sum().reset_index()
customer_balance = pd.merge(customer_spend, customer_repayment, on='Costomer', suffixes=('_spend', '_repayment'))
customer_balance['Due_Amount'] = customer_balance['Amount_spend'] - customer_balance['Amount_repayment']
interest_rate = 0.029
customer_balance['Interest'] = customer_balance['Due_Amount'].apply(lambda x: max(x, 0) * interest_rate)
customer_balance['Total_Due'] = customer_balance['Due_Amount'] + customer_balance['Interest']
customer_balance_sorted = customer_balance.sort_values('Total_Due', ascending=False)
print("Task 8: Impose an interest rate of 2.9% for each customer for any due amount")
print(customer_balance_sorted.head())

# Task 9: Monthly profit for the bank
spend['Month'] = pd.to_datetime(spend['Month'])
repayment['Month'] = pd.to_datetime(repayment['Month'])
monthly_spend = spend.groupby(spend['Month'].dt.to_period('M'))['Amount'].sum().reset_index()
monthly_repayment = repayment.groupby(repayment['Month'].dt.to_period('M'))['Amount'].sum().reset_index()
monthly_profit = pd.merge(monthly_spend, monthly_repayment, on='Month', suffixes=('_spend', '_repayment'))
monthly_profit['Profit'] = monthly_profit['Amount_repayment'] - monthly_profit['Amount_spend']
monthly_profit['Cumulative_Spend'] = monthly_profit['Amount_spend'].cumsum()
monthly_profit['Cumulative_Repayment'] = monthly_profit['Amount_repayment'].cumsum()
monthly_profit['Due_Amount'] = monthly_profit['Cumulative_Spend'] - monthly_profit['Cumulative_Repayment']
annual_interest_rate = 0.029
monthly_interest_rate = annual_interest_rate / 12
monthly_profit['Interest'] = monthly_profit['Due_Amount'].apply(lambda x: max(x, 0) * monthly_interest_rate)
monthly_profit['Profit_with_Interest'] = monthly_profit['Profit'] + monthly_profit['Interest']
print("Task 9: Monthly profit for the bank")
print(monthly_profit.head())
