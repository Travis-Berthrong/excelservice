import numpy as np
import pandas as pd

# Set a random seed for reproducibility
np.random.seed(42)

# Generate data for 5 different distributions
n_samples = 1000

# Normal distribution
normal = np.random.normal(loc=0, scale=1, size=n_samples)

# Skewed distribution (log-normal)
skewed = np.random.lognormal(mean=0, sigma=0.5, size=n_samples)

# High kurtosis distribution (t-distribution with low degrees of freedom)
high_kurtosis = np.random.standard_t(df=3, size=n_samples)

# Low kurtosis distribution (uniform)
low_kurtosis = np.random.uniform(low=-2, high=2, size=n_samples)

# Bimodal distribution
bimodal = np.concatenate([
    np.random.normal(loc=-2, scale=0.5, size=n_samples//2),
    np.random.normal(loc=2, scale=0.5, size=n_samples//2)
])

# Create a DataFrame
df = pd.DataFrame({
    'Normal': normal,
    'Skewed': skewed,
    'High_Kurtosis': high_kurtosis,
    'Low_Kurtosis': low_kurtosis,
    'Bimodal': bimodal
})

# Save to CSV
df.to_csv('data_analysis_test.csv', index=False)

print("CSV file 'data_analysis_test.csv' has been created.")
