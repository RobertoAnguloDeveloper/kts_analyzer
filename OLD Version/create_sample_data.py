import pandas as pd
import numpy as np

# Create sample data matching the user's format
data = {
    'Metric': ['Ore Mined', 'Overburden', 'Ore Mined', 'Overburden', 
               'Active Fleet Count (Aprox)', 'Liter of Diesel Consumed'],
    'Category': ['RGM', 'RGM', 'Sar', 'Sar', '', ''],
    'Unit': ['kt', 'kt', 'kt', 'kt', '', ''],
    'ene-20': [406.8, 4273.2, 91.3, 445.5, 831, 5145492],
    'feb-20': [549.1, 4272.4, 3.1, 731.9, 848, 5190595],
    'mar-20': [805.9, 3235.0, 377.8, 669.0, 865, 5319575],
    'abr-20': [609.1, 2696.9, 257.6, 465.4, 795, 5109433],
    'may-20': [620.0, 3292.4, 233.0, 242.4, 875, 5022389],
    'jun-20': [62.3, 1384.4, 105.8, 89.6, 765, 1862203],
    'jul-20': [0.0, 0.0, 34.6, 70.7, 486, 481172],
    'ago-20': [58.8, 195.2, 255.8, 244.3, 692, 2311407],
    'sep-20': [341.1, 1432.9, 232.5, 486.3, 766, 3691536],
    'oct-20': [398.2, 3454.9, 419.4, 921.6, 848, 5339833],
    'nov-20': [86.2, 3110.2, 402.9, 439.5, 834, 5197323],
    'dic-20': [253.4, 2829.0, 426.9, 365.3, 821, 5554960]
}

# Add more months (sample data for demonstration)
months_2021 = ['ene-21', 'feb-21', 'mar-21', 'abr-21', 'may-21', 'jun-21', 
               'jul-21', 'ago-21', 'sep-21', 'oct-21', 'nov-21', 'dic-21']
months_2022 = ['ene-22', 'feb-22', 'mar-22', 'abr-22', 'may-22', 'jun-22',
               'jul-22', 'ago-22', 'sep-22', 'oct-22', 'nov-22', 'dic-22']
months_2023 = ['ene-23', 'feb-23', 'mar-23', 'abr-23', 'may-23', 'jun-23',
               'jul-23', 'ago-23', 'sep-23', 'oct-23', 'nov-23', 'dic-23']
months_2024 = ['ene-24', 'feb-24']

all_months = months_2021 + months_2022 + months_2023 + months_2024

# Generate random but realistic data for additional months
np.random.seed(42)
for month in all_months:
    data[month] = [
        round(np.random.uniform(100, 700), 1),      # Ore Mined RGM
        round(np.random.uniform(2000, 4000), 1),    # Overburden RGM
        round(np.random.uniform(50, 400), 1),       # Ore Mined Sar
        round(np.random.uniform(200, 1500), 1),     # Overburden Sar
        int(np.random.uniform(650, 900)),           # Fleet Count
        int(np.random.uniform(3000000, 6000000))    # Diesel
    ]

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
output_file = '/mnt/user-data/outputs/sample_mining_data.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Mining_Data', index=False)
    
    # Add a second sheet with metadata
    metadata = pd.DataFrame({
        'Information': ['Data Type', 'Period', 'Units', 'Last Updated'],
        'Value': ['Mining Operations Data', 'Jan 2020 - Feb 2024', 
                 'kt (kilotons), Liters', 'October 2025']
    })
    metadata.to_excel(writer, sheet_name='Metadata', index=False)

print(f"Sample Excel file created: {output_file}")
print("\nData structure preview:")
print(df.iloc[:, :6])  # Show first 6 columns
print(f"\nTotal columns: {len(df.columns)}")
print(f"Total rows: {len(df)}")
