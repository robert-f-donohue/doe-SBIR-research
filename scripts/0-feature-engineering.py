import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from sklearn.preprocessing import StandardScaler


def analyze_distribution(df, feature):
    """
    Analyzes the distribution of a feature with various transformations, including:
        1. Log transformation
        2. Box-Cox transformation
        3. Standardization transformation

    :param data: DataFrame containing the feature values
    :param feature: ML feature to analyze
    :return: none
    """

    # Drop NaNs and non-positive values for log/Box-Cox transformations
    df = df.loc[df[feature].notna() & (df[feature] > 0)]

    # Display original distribution
    plt.figure(figsize = (10,6))
    sns.histplot(data=df, x=feature, bins=30, kde=True, color='orange')
    plt.title(f'Distribution of {feature} (Original)')
    plt.xlabel(feature)
    plt.ylabel('Count')
    plt.show()

    # Calculate and display original skewness and kurtosis
    original_skewness = stats.skew(df[feature])
    original_kurtosis = stats.kurtosis(df[feature])
    print(f'Original skewness: {original_skewness}')
    print(f'Original kurtosis: {original_kurtosis}')

    # Apply transformations
    log_transformed = np.log1p(df[feature])
    boxcox_transformed, lambda_bc = stats.boxcox(df[feature])
    standardized = StandardScaler().fit_transform(df[[feature]]).flatten()

    # Calculate skewness and kurtosis for each transformation
    transformed_df = pd.DataFrame({
        "Log Transformed": log_transformed,
        "Box-Cox Transformed": boxcox_transformed,
        "Standardized": standardized
    })

    # Plot transformed distributions
    plt.figure(figsize = (30, 6))

    # Log Transformation
    plt.subplot(1, 3, 1)
    sns.histplot(data=transformed_df, x='Log Transformed', bins=30, kde=True, color='blue')
    plt.title(f'Log Transformation: {feature}')
    plt.xlabel(f'Log Transformed {feature}')
    plt.ylabel('Count')

    # Box-Cox Transformation
    plt.subplot(1, 3, 2)
    sns.histplot(data=transformed_df, x='Box-Cox Transformed', bins=30, kde=True, color='green')
    plt.title(f'Box-Cox Transformation (λ={lambda_bc}): {feature}')
    plt.xlabel(f'Box-Cox Transformed {feature}')
    plt.ylabel('Count')

    # Standardization
    plt.subplot(1, 3, 3)
    sns.histplot(data=transformed_df, x='Standardized', bins=30, kde=True, color='red')
    plt.title(f'Standardization: {feature}')
    plt.xlabel(f'Standardized {feature}')
    plt.ylabel('Count')

    plt.tight_layout()
    plt.show()

    # Calculate skewness and kurtosis for each transformation
    transformations = {
        "Log Transformation": log_transformed,
        "Box-Cox Transformation": boxcox_transformed,
        "Standardization": standardized
    }

    print(f'Skewness: and Kurtosis: {original_skewness}')
    for name, transformed_data in transformations.items():
        skewness = stats.skew(transformed_data)
        kurtosis = stats.kurtosis(transformed_data)
        print(f'{name} - Skewness: {skewness}')
        print(f'{name} - Kurtosis: {kurtosis}')

    # # Evaluate best fit distributions for original data
    # dist_names = ['norm', 'lognorm', 'gamma', 'beta', 'expon', 'uniform']
    # results = {}
    # for dist_name in dist_names:
    #     dist = getattr(stats, dist_name)
    #     param = dist.fit(df[feature])
    #     ks_stat, p_value = stats.kstest(df[feature], dist_name, args=param)
    #     results[dist_name] = p_value
    #
    # # Get sorted results
    # sorted_results = sorted(results.items(), key=lambda x: x[1], reverse=True)
    # print(f'Best fit distribution for {feature}:')
    # for dist_name, p_value in sorted_results:
    #     print(f'{dist_name}: p-value = {p_value}')



# Define the list of features to analyze
features = [
    'Reported Gross Floor Area (Sq Ft)',
    'Site EUI (Energy Use Intensity kBtu/ft2)',
    'Total Site Energy Usage (kBtu)',
    'Estimated Total GHG Emissions (kgCO2e)',
    'Percentage Electrification'
]

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Setting Up Analysis ------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# File path to 2024 BERDO data
file_path_berdo_2024 = '../data-files/berdo_data_files/BERDO_Data-2024-clean.csv'

# create DataFrames for 2021 and 2024 data
df_2024 = pd.read_csv(file_path_berdo_2024)
print(f"Size of Original Dataset: {df_2024.shape[0]}")

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Add Electrification Metric -----------------------------------------
# --------------------------------------------------------------------------------------------------------

# make sure metrics are numeric
df_2024['Total Site Energy Usage (kBtu)'] = pd.to_numeric(df_2024['Total Site Energy Usage (kBtu)'], errors='coerce')

# calculate percentage electricication for properties
df_2024['Percentage Electrification'] = (
    (3.41215 * df_2024['Electricity Usage (kWh)']) / (df_2024['Total Site Energy Usage (kBtu)'])
) * 100

# export to csv
df_2024.to_csv('../data-files/berdo_data_files/BERDO_Data-2024-features.csv', index=False)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Data Cleaning ------------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# keep only multifamily
multifamily_df = df_2024[df_2024['Largest Property Type'] == 'Multifamily Housing']

# convert columns to numeric
for feature in features:
    multifamily_df.loc[:, feature] = pd.to_numeric(multifamily_df[feature], errors='coerce')
print(f"Original Size of Multifamily Housing Dataset: {multifamily_df.shape[0]}")

# drop any rows that became NaN after conversion
multifamily_df = multifamily_df.dropna(subset=features)
print(f"Size of Multifamily Housing Dataset w/o NaN: {multifamily_df.shape[0]}")

# drop any rows where electricity usage is greater than 100%
multifamily_df = multifamily_df[multifamily_df['Percentage Electrification'] <= 100]
print(f"Size of Multifamily Housing Dataset w/o Electrification > 100%: {multifamily_df.shape[0]}")

# drop any rows with GSF below 20,000 SF
min_gsf = 20000
multifamily_df = multifamily_df[multifamily_df['Reported Gross Floor Area (Sq Ft)'] >= min_gsf]
print(f"Size of Multifamily Housing Dataset w/o GSF < 20,000 SF: {multifamily_df.shape[0]}")

# drop any rows with EUI below 15 kBtu/sf
min_eui = 15.0
multifamily_df = multifamily_df[multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'] >= min_eui]
print(f"Size of Multifamily Housing Dataset w/o EUI < 15 kBtu/sf: {multifamily_df.shape[0]}")

# ensure Site EUI remains a float value
multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'] = multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'].astype(float)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Percentile-based Filtering -----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Percentile-based Filtering
gsf_upper_threshold = np.percentile(multifamily_df['Reported Gross Floor Area (Sq Ft)'], 99)                  # Top 1% for GSF
eui_upper_threshold = np.percentile(multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'], 98)           # Top 2% for EUI
site_energy_upper_threshold = np.percentile(multifamily_df['Total Site Energy Usage (kBtu)'], 98)             # Top 2% for Site Energy Usage
ghg_emissions_upper_threshold = np.percentile(multifamily_df['Estimated Total GHG Emissions (kgCO2e)'], 98)   # Top 2% for GHG Emissions

# Filter out extreme values based on percentiles
filtered_df = multifamily_df[
    (multifamily_df['Reported Gross Floor Area (Sq Ft)'] <= gsf_upper_threshold) &
    (multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'] <= eui_upper_threshold) &
    (multifamily_df['Total Site Energy Usage (kBtu)'] <= site_energy_upper_threshold) &
    (multifamily_df['Estimated Total GHG Emissions (kgCO2e)'] <= ghg_emissions_upper_threshold)
].copy()

# export multifamily data to avoid splitting data twice
# export to csv
filtered_df.to_csv('../data-files/berdo_data_files/BERDO_Data-2024-multifamily.csv', index=False)


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Check Distribution -------------------------------------------------
# --------------------------------------------------------------------------------------------------------

for feature in features:
    print(f'\n--- Analysis for {feature} ---')
    analyze_distribution(multifamily_df, feature)