import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import silhouette_score
import matplotlib.pyplot as plt
import seaborn as sns
sns.set_theme(style="whitegrid", palette="pastel")


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Setting Up Analysis ------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# File path to BERDO data & read in CSV
file_path_berdo = '../data-files/berdo_data_files/BERDO_Data-2024.csv'
df = pd.read_csv(file_path_berdo)

# keep only multifamily
multifamily_df = df[df['Largest Property Type'] == 'Multifamily Housing']

# relevant features for clustering
features = [
    'Reported Gross Floor Area (Sq Ft)',
    'Site EUI (Energy Use Intensity kBtu/ft2)',
    'Total Site Energy Usage (kBtu)',
    'Estimated Total GHG Emissions (kgCO2e)'
]

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Data Cleaning ------------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# filter out bad data
multifamily_df = multifamily_df.dropna(subset=features)

# convert columns to numeric
for feature in features:
    multifamily_df[feature] = pd.to_numeric(multifamily_df[feature], errors='coerce')

# drop any rows that became NaN after conversion
multifamily_df = multifamily_df.dropna(subset=features)

print(f"Original Dataset Size: {multifamily_df.shape[0]}")

# drop any rows with GSF below 20,000 SF
min_gsf = 20000
multifamily_df = multifamily_df[multifamily_df['Reported Gross Floor Area (Sq Ft)'] >= min_gsf]

print(f"Dataset Size After Filtering by GSF (>= {min_gsf} sq ft): {multifamily_df.shape[0]}")

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Visualize Distributions --------------------------------------------
# --------------------------------------------------------------------------------------------------------

# log scale distributions
plt.figure(figsize=(20, 10))
# log gross floor area distribution
plt.subplot(1, 2, 1)
sns.histplot(np.log(multifamily_df['Reported Gross Floor Area (Sq Ft)']), bins=30, kde=True)
plt.title('Log Scale: Gross Floor Area Distribution')
plt.xlabel('Log Gross Floor Area')
plt.ylabel('Count')

# Log Site EUI Distribution
plt.subplot(1, 2, 2)
sns.histplot(np.log(multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)']), bins=30, kde=True)
plt.title('Log Scale: Site EUI Distribution')
plt.xlabel('Log Site EUI')
plt.ylabel('Count')

plt.tight_layout()
plt.show()

# log scale distributions
plt.figure(figsize=(20, 10))

# Log Gross Floor Area Distribution
plt.subplot(1, 2, 1)
sns.histplot(np.log(multifamily_df['Total Site Energy Usage (kBtu)']), bins=30, kde=True)
plt.title('Log Scale: Total Site Energy Usage Distribution')
plt.xlabel('Log Total Site Energy Usage')
plt.ylabel('Count')

# Log Site EUI Distribution
plt.subplot(1, 2, 2)
sns.histplot(np.log(multifamily_df['Estimated Total GHG Emissions (kgCO2e)']), bins=30, kde=True)
plt.title('Log Scale: Total GHG Emissions Distribution')
plt.xlabel('Log Total GHG Emissions')
plt.ylabel('Count')

plt.tight_layout()
plt.show()

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Log Transformation -------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Log Transformation for GSF and EUI
cleaned_df = multifamily_df.copy()  # Keep a clean copy for transformations
cleaned_df['Log GSF'] = np.log(cleaned_df['Reported Gross Floor Area (Sq Ft)'])
cleaned_df['Log EUI'] = np.log(cleaned_df['Site EUI (Energy Use Intensity kBtu/ft2)'])
cleaned_df['Log Total Site Energy'] = np.log(cleaned_df['Total Site Energy Usage (kBtu)'])
cleaned_df['Log GHG Emissions'] = np.log(cleaned_df['Estimated Total GHG Emissions (kgCO2e)'])

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Percentile-based Filtering -----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Percentile-based Filtering
gsf_upper_threshold = np.percentile(cleaned_df['Log GSF'], 99)                       # Top 1% for GSF
eui_upper_threshold = np.percentile(cleaned_df['Log EUI'], 98)                       # Top 2% for EUI
site_energy_upper_threshold = np.percentile(cleaned_df['Log Total Site Energy'], 98) # Top 2% for Site Energy Usage
ghg_emissions_upper_threshold = np.percentile(cleaned_df['Log GHG Emissions'], 98)   # Top 2% for GHG Emissions

# Filter out extreme values based on percentiles
filtered_df = cleaned_df[
    (cleaned_df['Log GSF'] <= gsf_upper_threshold) &
    (cleaned_df['Log EUI'] <= eui_upper_threshold) &
    (cleaned_df['Log Total Site Energy'] <= site_energy_upper_threshold) &
    (cleaned_df['Log GHG Emissions'] <= ghg_emissions_upper_threshold)
].copy()

print(f"Dataset Size After Percentile-based Filtering: {filtered_df.shape[0]}")

# --------------------------------------------------------------------------------------------------------
# ---------------------------------- Normalization for Log-Transformed Data (UPDATED) --------------------
# --------------------------------------------------------------------------------------------------------

# Normalize the Log-Transformed Data
log_features = ['Log GSF', 'Log EUI', 'Log Total Site Energy', 'Log GHG Emissions']

# Drop NaN values after log transformation
filtered_df = filtered_df.dropna(subset=log_features)
print(f"Dataset Size After Dropping NaN Values: {filtered_df.shape[0]}")

scaler = StandardScaler()
X = scaler.fit_transform(filtered_df[log_features])

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Elbow Curve Analysis -----------------------------------------------
# --------------------------------------------------------------------------------------------------------

# determine the optimal number of clusters using Elbow Method
sse = []
range_n_clusters = range(1, 11)
for k in range_n_clusters:
    km = KMeans(n_clusters=k, random_state=42)
    km.fit(X)
    sse.append(km.inertia_)

# plot elbow curve
# create figure
plt.figure(figsize=(12, 8))
plt.plot(range_n_clusters, sse, marker='o')
plt.title('Elbow Method for Optimal Clusters')
plt.xlabel('Number of Clusters')
plt.ylabel('SSE (Sum of Squared Errors')
plt.grid(True)
plt.show()

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Cluster Analysis ---------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# perform K-Means Clustering with 2 clusters (prototypes)
optimal_clusters_simple = 2
kmeans_simple = KMeans(n_clusters=optimal_clusters_simple, random_state=42)
labels_simple = kmeans_simple.fit_predict(X)
filtered_df.loc[:, 'Cluster_simple'] = labels_simple

# perform K-Means Clustering with 8 clusters (prototypes)
optimal_clusters_complex = 6
kmeans_complex = KMeans(n_clusters=optimal_clusters_complex, random_state=42)
labels_complex = kmeans_complex.fit_predict(X)
filtered_df.loc[:, 'Cluster_complex'] = labels_complex

# evaluate clustering with Silhouette Score
sil_score_simple = silhouette_score(X, labels_simple)
sil_score_complex = silhouette_score(X, labels_complex)
print(f'Silhouette Score for {optimal_clusters_simple} clusters: {sil_score_simple}')
print(f'Silhouette Score for {optimal_clusters_complex} clusters: {sil_score_complex}')

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- PCA Analysis -------------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# visualize clusters using PCA (2D)
pca_simple = PCA(n_components=4)
pca_complex = PCA(n_components=4)
pca_components_simple = pca_simple.fit_transform(X)
pca_components_complex = pca_complex.fit_transform(X)

# add columns
filtered_df.loc[:, 'PCA_simple_1'] = pca_components_simple[:, 0]
filtered_df.loc[:, 'PCA_simple_2'] = pca_components_simple[:, 1]
filtered_df.loc[:, 'PCA_complex_1'] = pca_components_complex[:, 0]
filtered_df.loc[:, 'PCA_complex_2'] = pca_components_complex[:, 1]

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- PCA Visualizations -------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Simple Clustering
plt.figure(figsize=(12, 8))
sns.scatterplot(
    data=filtered_df,
    x='PCA_simple_1',
    y='PCA_simple_2',
    hue='Cluster_simple',
    palette='Set1',
    alpha=0.7,
)
plt.title('Cluster Visualization for Simple Clustering using PCA')
plt.xlabel('PCA Component 1')
plt.ylabel('PCA Component 2')
plt.legend(title='Cluster')
plt.grid(True)
plt.show()

# Complex Clustering
plt.figure(figsize=(12, 8))
sns.scatterplot(
    data=filtered_df,
    x='PCA_complex_1',
    y='PCA_complex_2',
    hue='Cluster_complex',
    palette='Set1',
    alpha=0.7,
)
plt.title('Cluster Visualization for Complex Clustering using PCA')
plt.xlabel('PCA Component 1')
plt.ylabel('PCA Component 2')
plt.legend(title='Cluster')
plt.grid(True)
plt.show()

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Log Cluster Visualizations -----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Simple Clustering Scatter Plot
plt.figure(figsize=(12, 8))
sns.scatterplot(
    data=filtered_df,
    x='Log GSF',
    y='Log EUI',
    hue='Cluster_simple',
    palette='Set2',
    alpha=0.7
)
plt.title('Simple Clusters on Log GSF vs. Log EUI (Percentile-Filtered)')
plt.xlabel('Log Gross Floor Area (Sq Ft)')
plt.ylabel('Log Site EUI (kBtu/sf)')
plt.legend(title='Cluster')
plt.grid(True)
plt.show()

# Complex Clustering Scatter Plot
plt.figure(figsize=(12, 8))
sns.scatterplot(
    data=filtered_df,
    x='Log GSF',
    y='Log EUI',
    hue='Cluster_complex',
    palette='Set2',
    alpha=0.7
)
plt.title('Complex Clusters on Log GSF vs. Log EUI (Percentile-Filtered)')
plt.xlabel('Gross Floor Area (Sq Ft)')
plt.ylabel('Site EUI (kBtu/sf)')
plt.legend(title='Cluster')
plt.grid(True)
plt.show()

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Normal Cluster Visualizations --------------------------------------
# --------------------------------------------------------------------------------------------------------

# Simple Clustering Scatter Plot
plt.figure(figsize=(12, 8))
sns.scatterplot(
    data=filtered_df,
    x='Reported Gross Floor Area (Sq Ft)',
    y='Site EUI (Energy Use Intensity kBtu/ft2)',
    hue='Cluster_simple',
    palette='Set2',
    alpha=0.7
)
plt.title('Simple Clusters on GSF vs. Site EUI (Percentile-Filtered)')
plt.xlabel('Gross Floor Area (Sq Ft)')
plt.ylabel('Site EUI (kBtu/sf)')
plt.legend(title='Cluster')
plt.grid(True)
plt.show()

# Complex Clustering Scatter Plot
plt.figure(figsize=(12, 8))
sns.scatterplot(
    data=filtered_df,
    x='Reported Gross Floor Area (Sq Ft)',
    y='Site EUI (Energy Use Intensity kBtu/ft2)',
    hue='Cluster_complex',
    palette='Set2',
    alpha=0.7
)
plt.title('Complex Clusters on GSF vs. Site EUI (Percentile-Filtered)')
plt.xlabel('Gross Floor Area (Sq Ft)')
plt.ylabel('Site EUI (kBtu/sf)')
plt.legend(title='Cluster')
plt.grid(True)
plt.show()
# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Prototype Properties  ----------------------------------------------
# --------------------------------------------------------------------------------------------------------

# get cluster centriods of prototypes
cluster_centers_simple = scaler.inverse_transform(kmeans_simple.cluster_centers_)
cluster_centers_complex = scaler.inverse_transform(kmeans_complex.cluster_centers_)

# create dataframes of the prototype features
prototype_simple_df = np.exp(pd.DataFrame(cluster_centers_simple, columns=log_features))
prototype_complex_df = np.exp(pd.DataFrame(cluster_centers_complex, columns=log_features))

# display properties of prototypes (centriods)
print("Prototypes (Centriods of Clusters - Simple):")
print(prototype_simple_df)
print("Prototypes (Centriods of Clusters - Complex):")
print(prototype_complex_df)
