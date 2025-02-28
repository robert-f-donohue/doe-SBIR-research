import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.feature_selection import mutual_info_classif
from sklearn.metrics import silhouette_score, davies_bouldin_score, calinski_harabasz_score
from sklearn.manifold import TSNE
sns.set_theme(style="whitegrid", palette="pastel")


def analyze_distribution(df, feature):
    """
    Analyzes the distribution of a feature with Log transformation

    :param df: DataFrame containing the feature values
    :param feature: ML feature to analyze
    :return: none
    """

    # Drop NaNs and non-positive values for log/Box-Cox transformations
    df = df.loc[df[feature].notna() & (df[feature] > 0)]

    # Apply transformations
    log_feature = np.log1p(df[feature])

    # Create figure to display distributions
    plt.figure(figsize = (12, 10))
    # Display original distribution
    plt.subplot(1, 2, 1)
    sns.histplot(data=df, x=feature, bins=30, kde=True, alpha=0.7,)
    plt.title(f'Distribution of {feature}')
    plt.xlabel(feature)
    plt.ylabel('Count')

    # Log transformed distribution
    plt.subplot(1, 2, 2)
    sns.histplot(data=df, x=log_feature, bins=30, kde=True, alpha=0.7,)
    plt.title(f'Log Distribution of {feature}')
    plt.xlabel(f'Log Transformed {feature}')
    plt.ylabel('Count')

    plt.show()

    # Calculate and display original skewness and kurtosis
    original_skewness = stats.skew(df[feature])
    original_kurtosis = stats.kurtosis(df[feature])
    print(f'Original skewness: {original_skewness}')
    print(f'Original kurtosis: {original_kurtosis}')

    # Calculate and display log transformed skewness and kurtosis
    log_skewness = stats.skew(log_feature)
    log_kurtosis = stats.kurtosis(log_feature)
    print(f'Log Transformed skewness: {log_skewness}')
    print(f'Log Transformed kurtosis: {log_kurtosis}')


# Define the list of features to analyze
features = [
    'Reported Gross Floor Area (Sq Ft)',
    'Site EUI (Energy Use Intensity kBtu/ft2)',
    'Total Site Energy Usage (kBtu)',
    'Estimated Total GHG Emissions (kgCO2e)',
    'Estimated Carbon Emissions Intensity (kg CO2e/sf)'
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
# ----------------------------------- Add Electrification & CEI Metrics ----------------------------------
# --------------------------------------------------------------------------------------------------------

# Esure metrics are numeric
df_2024['Total Site Energy Usage (kBtu)'] = pd.to_numeric(df_2024['Total Site Energy Usage (kBtu)'], errors='coerce')

# Calculate percentage electrification for properties
df_2024['Percentage Electrification'] = (
    (3.41215 * df_2024['Electricity Usage (kWh)']) / (df_2024['Total Site Energy Usage (kBtu)'])
) * 100

# Calculate CEI for properties
df_2024['Estimated Carbon Emissions Intensity (kg CO2e/sf)'] = (
    (df_2024['Estimated Total GHG Emissions (kgCO2e)']) / (df_2024['Reported Gross Floor Area (Sq Ft)'])
)

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

# drop any rows with GSF below 20,000 SF
min_gsf = 20000
multifamily_df = multifamily_df[multifamily_df['Reported Gross Floor Area (Sq Ft)'] >= min_gsf]
print(f"Size of Multifamily Housing Dataset w/o GSF < 20,000 SF: {multifamily_df.shape[0]}")

# drop any rows with EUI below 15 kBtu/sf
min_eui = 15.0
multifamily_df = multifamily_df[multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'] >= min_eui]
print(f"Size of Multifamily Housing Dataset w/o EUI < 15 kBtu/sf: {multifamily_df.shape[0]}")

# drop any rows where electricity usage is greater than 100%
multifamily_df = multifamily_df[multifamily_df['Percentage Electrification'] <= 100]
print(f"Size of Multifamily Housing Dataset w/o Electrification > 100%: {multifamily_df.shape[0]}")

# ensure Site EUI remains a float value
multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'] = multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'].astype(float)


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Log Transformation -------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Log Transformation for GSF and EUI
multifamily_df['Log GSF'] = np.log(multifamily_df['Reported Gross Floor Area (Sq Ft)'])
multifamily_df['Log EUI'] = np.log(multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'])
multifamily_df['Log Total Site Energy'] = np.log(multifamily_df['Total Site Energy Usage (kBtu)'])
multifamily_df['Log GHG Emissions'] = np.log(multifamily_df['Estimated Total GHG Emissions (kgCO2e)'])
multifamily_df['Log CEI'] = np.log(multifamily_df['Estimated Carbon Emissions Intensity (kg CO2e/sf)'])
multifamily_df['Log % Electrification'] = np.log(multifamily_df['Percentage Electrification'])

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Percentile-based Filtering -----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Percentile-based Filtering
gsf_upper_threshold = np.percentile(multifamily_df['Log GSF'], 99)                       # Top 1% for GSF
eui_upper_threshold = np.percentile(multifamily_df['Log EUI'], 98)                       # Top 2% for EUI
site_energy_upper_threshold = np.percentile(multifamily_df['Log Total Site Energy'], 98) # Top 2% for Site Energy Usage
ghg_emissions_upper_threshold = np.percentile(multifamily_df['Log GHG Emissions'], 98)   # Top 2% for GHG Emissions
cei_upper_threshold = np.percentile(multifamily_df['Log CEI'], 98)                       # Top 2% for CEI

# Filter out extreme values based on percentiles
filtered_df = multifamily_df[
    (multifamily_df['Log GSF'] <= gsf_upper_threshold) &
    (multifamily_df['Log EUI'] <= eui_upper_threshold) &
    (multifamily_df['Log Total Site Energy'] <= site_energy_upper_threshold) &
    (multifamily_df['Log GHG Emissions'] <= ghg_emissions_upper_threshold) &
    (multifamily_df['Log CEI'] <= cei_upper_threshold)
].copy()

print(f"Dataset Size After Percentile-based Filtering: {filtered_df.shape[0]}")

# # export multifamily data to avoid splitting data twice
# # export to csv
filtered_df.to_csv('../data-files/berdo_data_files/BERDO_Data-2024-multifamily.csv', index=False)


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Check Distribution -------------------------------------------------
# --------------------------------------------------------------------------------------------------------
#
# for feature in features:
#     print(f'\n--- Analysis for {feature} ---')
#     analyze_distribution(multifamily_df, feature)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Correlation Matrix -------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# get a clean copy for analysis
corr_df = filtered_df.copy()

# remove irrelevant columns
columns_to_remove = [
    '_id', 'BERDO ID', 'Tax Parcel ID', 'Property Owner Name', 'Building Address', 'Building Address City',
    'Building Address Zip  Code', 'Geospatial Address', 'Geospatial Address.1', 'Parcel Address', 'Parcel Address City',
    'Parcel Address Zip Code', 'Reported Enclosed Parking Area (Sq Ft)', 'Largest Property Type', 'All Property Types and GFAs',
    'District Hot Water Usage (kBtu)', 'District Hot Water Emissions (kgCO2e)', 'District Chilled Water Usage (kBtu)',
    'District Chilled Water Emissions (kgCO2e)', 'Fuel Oil 5 and 6 Usage (kBtu)', 'Fuel Oil 5 and 6 Emissions (kgCO2e)',
    'Kerosene Usage (kBtu)', 'Kerosene Emissions (kgCO2e)', 'Compliance Status', 'Notes', 'Corresponding Campus ID',
    'Community Choice Electricity Participation', 'Renewable Energy Purchased through a Power Purchase Agreement (',
    'Renewable Energy Certificate (REC) Purchase', 'Backup Generator', 'Battery Storage', 'Electric Vehicle (EV) Charging',
    'Energy Star Score', 'Renewable System Electricity Usage Onsite (kBtu)', 'District Steam Usage (kBtu)',
    'District Steam Emissions (kgCO2e)', 'Fuel Oil 1 Usage (kBtu)', 'Fuel Oil 1 Emissions (kgCO2e)', 'Fuel Oil 2 Usage (kBtu)',
    'Fuel Oil 2 Emissions (kgCO2e)', 'Fuel Oil 4 Usage (kBtu)', 'Fuel Oil 4 Emissions (kgCO2e)', 'Propane Usage (kBtu)',
    'Propane Emissions (kgCO2e)', 'Diesel Usage (kBtu)', 'Diesel Emissions (kgCO2e)', 'First Emissions Compliance Year (Projected)'
]
corr_df.drop(columns=columns_to_remove, inplace=True)

# Convert all relevant columns to numeric (force coercion of errors)
corr_df = corr_df.apply(pd.to_numeric, errors='coerce').fillna(0)

# create correlation matrix
corr_matrix = corr_df.corr()

# visualize the correlation matrix using a heatmap
plt.figure(figsize=(24, 20))
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', fmt=".2f", cbar_kws={'shrink': .8})
plt.title('Correlation Heatmap of Building Performance Features')
plt.show()

# identify highly correlated pairs (above 0.85 or below -0.85)
high_corr_pairs = []
threshold = 0.85

# iterate over correlation matrix columns
for col in corr_matrix.columns:
    for idx in corr_matrix.index:
        if col != idx and abs(corr_matrix.loc[idx, col]) > threshold:
            high_corr_pairs.append((col, idx, corr_matrix.loc[idx, col]))

# display highly correlated pairs
print("\nHighly Correlated Feature Pairs (|correlation| > 0.85):")
for pair in high_corr_pairs:
    print(f"{pair[0]} and {pair[1]}: Correlation = {pair[2]}")

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Mutual Information Analysis ----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Features to test
mi_score_features = [
    'Log GSF', 'Log EUI', 'Log Total Site Energy', 'Log GHG Emissions',
    'Log CEI'
]

# Create mutual information DataFrame
mutual_df = multifamily_df[mi_score_features].copy()

# Drop NaN values after Box Cox transformation
mutual_df = mutual_df.dropna()

# Standardize features
scaler_mi = StandardScaler()
X_scaled = scaler_mi.fit_transform(mutual_df)

# Run K-Means cluster with the current K (8 clusters)
kmeans_mi = KMeans(n_clusters=8, random_state=42)
cluster_labels = kmeans_mi.fit_predict(X_scaled)

# Calculate mutual information
mi_scores = mutual_info_classif(X_scaled, cluster_labels)
mi_scores_series = pd.Series(mi_scores, index=mutual_df.columns).sort_values(ascending=False)

# Plot scores
plt.figure(figsize=(12, 12))
mi_scores_series.plot(kind='bar', color='blue')
plt.title(f'Mutual Information Score for Feature Selection')
plt.xlabel('Features')
plt.ylabel('Mutual Information Score')
plt.xticks(rotation=45)
plt.show()

print("\nMutual Information Scores:")
print(mi_scores_series)


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Feature Selection Analysis ----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Features to test
selected_features = [
    'Log GSF', 'Log EUI', 'Log Total Site Energy', 'Log GHG Emissions',
    'Log CEI'
]

# Create mutual information DataFrame
df_final = multifamily_df[selected_features].copy()

# Drop NaN values after Box Cox transformation
df_final = mutual_df.dropna()

# Standardize features
scaler_final = StandardScaler()
X_final = scaler_final.fit_transform(df_final[selected_features])

# Run K-Means for 7 and 8 Clusters
for k in [4, 5, 6, 7, 8]:
    kmeans_final = KMeans(n_clusters=k, random_state=42)
    cluster_labels_final = kmeans_final.fit_predict(X_final)

    # Calculate performance metrics
    silhouette = silhouette_score(X_final, cluster_labels_final)
    davies_bouldin = davies_bouldin_score(X_final, cluster_labels_final)
    calinski_harabasz = calinski_harabasz_score(X_final, cluster_labels_final)

    print(f"\nK-Means Clustering ({k} Clusters) Metrics:")
    print(f"Silhouette Score: {silhouette}")
    print(f"Davies-Bouldin Index: {davies_bouldin}")
    print(f"Calinski-Harabasz Index: {calinski_harabasz}")

    # Add cluster labels to DataFrame for analysis
    df_final['Cluster'] = cluster_labels_final

    # Visualize distributions for each feature by cluster
    for feature in selected_features:
        plt.figure(figsize=(10, 6))
        sns.boxplot(x='Cluster', y=feature, data=df_final)
        plt.title(f'Distribution of {feature} by Cluster (k={k})')
        # plt.show()

    # Visualize clusters using PCA
    pca = PCA(n_components=2)
    X_pca = pca.fit_transform(X_final)
    plt.figure(figsize=(10, 6))
    sns.scatterplot(x=X_pca[:, 0], y=X_pca[:, 1], hue=cluster_labels_final)
    plt.title(f'PCA Visualization of {k} Clusters')
    plt.xlabel('PCA Component 1')
    plt.ylabel('PCA Component 2')
    # plt.show()

# Add cluster labels to the DataFrame
df_core['Cluster'] = cluster_labels_core

# Cross-tabulate Cluster vs. Building Size Category
df_final['Size Category'] = pd.cut(df_final['Log GSF'], bins=3, labels=['Small', 'Medium', 'Large'])
print("\nCross-tabulation with Building Size:")
print(pd.crosstab(df_final['Size Category'], df_final['Cluster']))
