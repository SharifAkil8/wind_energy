import pandas as pd
import numpy as np

# Load Excel file
df = pd.read_excel('/Users/sharif/Desktop/UNDERGRAD_COLLEGE/Semesters/FALL/Fall_2024/ECEN430/homeworks/hw1/Valentine Wind Data copy.xlsx', sheet_name='Wind Data')

# Clean and convert wind speed column to numeric (forcing any errors to NaN)
df['SPEED 40A M/SEC'] = pd.to_numeric(df['SPEED 40A M/SEC'], errors='coerce')

# Remove rows with missing or non-positive wind speed values
df_clean = df[df['SPEED 40A M/SEC'] > 0]

# Step 1: Define the power curve from the provided graph
def get_power_output(wind_speed):
    if wind_speed < 3:
        return 0  # Below cut-in speed
    elif 3 <= wind_speed < 4:
        return 0.2 * 1500  # Roughly interpolating between points
    elif 4 <= wind_speed < 5:
        return 0.4 * 1500
    elif 5 <= wind_speed < 6:
        return 0.55 * 1500
    elif 6 <= wind_speed < 7:
        return 0.75 * 1500
    elif 7 <= wind_speed < 8:
        return 0.9 * 1500
    elif 8 <= wind_speed < 9:
        return 1100  # From the graph
    elif 9 <= wind_speed < 10:
        return 1300
    elif 10 <= wind_speed < 11:
        return 1400
    elif 11 <= wind_speed < 12:
        return 1450
    elif 12 <= wind_speed <= 25:
        return 1500  # Rated power
    else:
        return 0  # Above cut-out speed

# Step 2: Apply the power curve to the wind speed data to get power output for each hour
df_clean['Power Output (kW)'] = df_clean['SPEED 40A M/SEC'].apply(get_power_output)

# Step 3: Compute the hourly energy production (in kWh)
# Energy = Power * Time (1 hour per data point)
df_clean['Energy (kWh)'] = df_clean['Power Output (kW)']

# Step 4: Calculate the total annual energy production (sum of all hourly energy)
annual_energy_production = df_clean['Energy (kWh)'].sum()

# Constants for Capacity Factor
RATED_POWER_KW = 1500  # Rated power of the turbine in kW
HOURS_IN_YEAR = 365 * 24  # Total hours in a year

# Step 5: Compute the maximum possible energy production (in kWh)
max_energy_production = RATED_POWER_KW * HOURS_IN_YEAR

# Step 6: Compute the Capacity Factor
capacity_factor = annual_energy_production / max_energy_production

# Display the Capacity Factor
print(f"Capacity Factor: {capacity_factor:.4f}")

# Optional: Display as percentage
print(f"Capacity Factor (Percentage): {capacity_factor * 100:.2f}%")