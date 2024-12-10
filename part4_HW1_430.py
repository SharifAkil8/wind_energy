import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Constants
AIR_DENSITY = 1.225  # kg/m^3, at sea level and 15°C

# Load Excel file
df = pd.read_excel('/Users/sharif/Desktop/UNDERGRAD_COLLEGE/Semesters/FALL/Fall_2024/ECEN430/homeworks/hw1/Valentine Wind Data copy.xlsx', sheet_name='Wind Data')

# Clean and convert wind speed column to numeric (forcing any errors to NaN)
df['SPEED 40A M/SEC'] = pd.to_numeric(df['SPEED 40A M/SEC'], errors='coerce')

# Remove rows with missing or non-positive wind speed values
df_clean = df[df['SPEED 40A M/SEC'] > 0]

# Convert the timestamp columns (assuming you have 'YEAR', 'MONTH', 'DAY' columns)
df_clean['DATE'] = pd.to_datetime(df_clean[['YEAR', 'MONTH', 'DAY']])

# Step 1: Compute wind power density for each hour at 40m
# Formula: P = 0.5 * rho * v^3
df_clean['WPD'] = 0.5 * AIR_DENSITY * df_clean['SPEED 40A M/SEC'] ** 3

# Step 2: Group by month and calculate monthly and annual averages
df_clean['MONTH'] = df_clean['DATE'].dt.month

# Compute monthly average wind power density
monthly_avg_wpd = df_clean.groupby('MONTH')['WPD'].mean()

# Compute annual average wind power density
annual_avg_wpd = df_clean['WPD'].mean()

# Display the results
print(f"Monthly Average Wind Power Density (W/m^2):\n{monthly_avg_wpd}")
print(f"\nAnnual Average Wind Power Density (W/m^2): {annual_avg_wpd:.2f} W/m^2")

# Step 3: Plot the monthly average wind power density on a bar graph
plt.figure(figsize=(10, 6))
monthly_avg_wpd.plot(kind='bar', color='skyblue')

# Customize plot
plt.title('Monthly Average Wind Power Density at 40m')
plt.xlabel('Month')
plt.ylabel('Wind Power Density (W/m^2)')
plt.xticks(ticks=range(12), labels=['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], rotation=0)
plt.grid(True)
plt.tight_layout()

# Display the plot
plt.show()

# # Save the Monthly Average Wind Power Density to an Excel file
# monthly_avg_wpd.to_excel('/Users/sharif/Desktop/UNDERGRAD_COLLEGE/Semesters/FALL/Fall_2024/ECEN430/homeworks/monthly_avg_wpd.xlsx', index=False)
# print("Monthly Average Wind Power Density saved to Excel.")