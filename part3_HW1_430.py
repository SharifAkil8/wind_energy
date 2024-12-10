import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import rayleigh

# Load Excel file
df = pd.read_excel('/Users/sharif/Desktop/UNDERGRAD_COLLEGE/Semesters/FALL/Fall_2024/ECEN430/homeworks/Valentine Wind Data copy.xlsx', sheet_name='Wind Data')

# Clean and convert wind speed column to numeric (forcing any errors to NaN)
df['SPEED 40A M/SEC'] = pd.to_numeric(df['SPEED 40A M/SEC'], errors='coerce')

# Remove rows with missing or non-positive wind speed values
df_clean = df[df['SPEED 40A M/SEC'] > 0]

# Step 1: Extract the wind speeds at 40m and calculate the mean wind speed
wind_speeds_40m = df_clean['SPEED 40A M/SEC']
mean_wind_speed = wind_speeds_40m.mean()
print(f"Mean Wind Speed at 40m: {mean_wind_speed:.2f} m/s")

# Step 2: Create a histogram of wind speeds with 0.5 m/s bins
bin_width = 0.5
bins = np.arange(0, wind_speeds_40m.max() + bin_width, bin_width)

plt.figure(figsize=(10, 6))
plt.hist(wind_speeds_40m, bins=bins, density=True, alpha=0.6, color='blue', label='Wind Speed Data')

# Step 3: Calculate the Rayleigh distribution for comparison
sigma = mean_wind_speed / np.sqrt(np.pi / 2)  # Scale parameter for Rayleigh
x = np.linspace(0, wind_speeds_40m.max(), 500)  # Wind speed values for the Rayleigh plot
rayleigh_pdf = (x / sigma**2) * np.exp(-x**2 / (2 * sigma**2))

# Plot the Rayleigh distribution
plt.plot(x, rayleigh_pdf, color='red', label='Rayleigh Distribution', linewidth=2)

# Step 4: Customize the plot
plt.title('Wind Speed Distribution at 40m with Rayleigh Distribution Overlay')
plt.xlabel('Wind Speed (m/s)')
plt.ylabel('Probability Density')
plt.legend()
plt.grid(True)
plt.tight_layout()

# Step 5: Display the plot
plt.show()