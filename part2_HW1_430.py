import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Load Excel file
df = pd.read_excel('/Users/sharif/Desktop/UNDERGRAD_COLLEGE/Semesters/FALL/Fall_2024/ECEN430/homeworks/hw1/Valentine Wind Data copy.xlsx', sheet_name='Wind Data')

# Clean and convert wind speed columns to numeric (forcing any errors to NaN)
df['SPEED 10 M/SEC'] = pd.to_numeric(df['SPEED 10 M/SEC'], errors='coerce')
df['SPEED 25 M/SEC'] = pd.to_numeric(df['SPEED 25 M/SEC'], errors='coerce')
df['SPEED 40A M/SEC'] = pd.to_numeric(df['SPEED 40A M/SEC'], errors='coerce')

# Remove rows with missing or non-positive wind speed values
df_clean = df[(df['SPEED 10 M/SEC'] > 0) & (df['SPEED 25 M/SEC'] > 0) & (df['SPEED 40A M/SEC'] > 0)]

# Convert the timestamp columns (assuming you have 'YEAR', 'MONTH', 'DAY' columns)
df_clean['DATE'] = pd.to_datetime(df_clean[['YEAR', 'MONTH', 'DAY']])

# Step 1: Group by month and calculate the average wind speed for each height (10m, 25m, 40m)
df_clean['MONTH'] = df_clean['DATE'].dt.month
monthly_avg = df_clean.groupby('MONTH')[['SPEED 10 M/SEC', 'SPEED 25 M/SEC', 'SPEED 40A M/SEC']].mean()

# Step 2: Calculate the estimated wind speeds at 25m and 40m using the 1/7th power law based on 10m data
monthly_avg['Estimated SPEED 25'] = monthly_avg['SPEED 10 M/SEC'] * (25 / 10) ** (1/7)
monthly_avg['Estimated SPEED 40'] = monthly_avg['SPEED 10 M/SEC'] * (40 / 10) ** (1/7)

# Step 3: Plot the actual vs estimated wind speeds for 10m, 25m, and 40m
plt.figure(figsize=(10, 6))

# Plot actual wind speeds
plt.plot(monthly_avg.index, monthly_avg['SPEED 10 M/SEC'], label='Actual SPEED 10m', marker='o')
plt.plot(monthly_avg.index, monthly_avg['SPEED 25 M/SEC'], label='Actual SPEED 25m', marker='o')
plt.plot(monthly_avg.index, monthly_avg['SPEED 40A M/SEC'], label='Actual SPEED 40m', marker='o')

# Plot estimated wind speeds
plt.plot(monthly_avg.index, monthly_avg['Estimated SPEED 25'], label='Estimated SPEED 25m', linestyle='--', marker='x')
plt.plot(monthly_avg.index, monthly_avg['Estimated SPEED 40'], label='Estimated SPEED 40m', linestyle='--', marker='x')

# Customize plot
plt.title('Monthly Average Wind Speeds at 10m, 25m, and 40m')
plt.xlabel('Month')
plt.ylabel('Wind Speed (m/s)')
plt.xticks(monthly_avg.index, ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
plt.legend()
plt.grid(True)
plt.tight_layout()

# Display the plot
plt.show()

# Optional: Save the plot to a file
# plt.savefig('monthly_wind_speeds.png')

# Print the monthly average wind speed data for verification
print(monthly_avg)

# # Save the monthly average wind speed data to an Excel file
# monthly_avg.to_excel('/Users/sharif/Desktop/UNDERGRAD_COLLEGE/Semesters/FALL/Fall_2024/ECEN430/homeworks/monthly_avg_ws.xlsx', index=False)
# print("Monthly average wind speed saved to Excel.")