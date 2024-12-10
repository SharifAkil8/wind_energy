import pandas as pd
import numpy as np

# Load Excel file
df = pd.read_excel('/Users/sharif/Desktop/UNDERGRAD_COLLEGE/Semesters/FALL/Fall_2024/ECEN430/homeworks/hw1/Valentine Wind Data copy.xlsx', sheet_name='Wind Data')

# Constants
AIR_DENSITY = 1.225  # kg/m^3, at sea level and 15°C

# Step 1: Convert 'SPEED 40A M/SEC' to numeric, forcing errors to NaN
df['SPEED 40A M/SEC'] = pd.to_numeric(df['SPEED 40A M/SEC'], errors='coerce')

# Step 2: Clean the data - Remove rows with NaN values or non-positive wind speeds
df_clean = df[df['SPEED 40A M/SEC'] > 0]

# Step 3: Calculate the annual average wind speed at 40m
annual_avg_wind_speed = df_clean['SPEED 40A M/SEC'].mean()
print(f'Annual Average Wind Speed at 40m: {annual_avg_wind_speed:.2f} m/s')

# Step 4: Calculate the wind power density
# Formula: P = 0.5 * rho * v^3 (where rho is air density and v is wind speed in m/s)
wind_speeds = df_clean['SPEED 40A M/SEC']
wind_power_density = 0.5 * AIR_DENSITY * np.mean(wind_speeds ** 3)
print(f'Wind Power Density at 40m: {wind_power_density:.2f} W/m^2')