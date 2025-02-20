import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans, DBSCAN
from sklearn.mixture import GaussianMixture
from sklearn.decomposition import PCA
from sklearn.metrics import silhouette_score, davies_bouldin_score, calinski_harabasz_score
import matplotlib.pyplot as plt
import seaborn as sns
sns.set_theme(style="whitegrid", palette="pastel")


def histograms(df):
    # plot histograms to see distributions
    plt.figure(figsize=(20, 6))
    # Gross Floor Area Distribution
    plt.subplot(1, 4, 1)
    sns.histplot(df['Reported Gross Floor Area (Sq Ft)'], bins=30, kde=True)
    plt.title('Gross Floor Area (GSF) Distribution')
    plt.xlabel('Gross Floor Area (Sq Ft)')
    plt.ylabel('Count')
    # Site EUI Distribution
    plt.subplot(1, 4, 2)
    sns.histplot(df['Site EUI (Energy Use Intensity kBtu/ft2)'], bins=30, kde=True)
    plt.title('Site EUI Distribution')
    plt.xlabel('Site EUI (kBtu/ft2)')
    plt.ylabel('Count')
    # Total Energy Distribution
    plt.subplot(1, 4, 3)
    sns.histplot(df['Total Site Energy Usage (kBtu)'], bins=30, kde=True)
    plt.title('Total Energy Distribution')
    plt.xlabel('Site Energy Usage (kBtu)')
    plt.ylabel('Count')
    # GHG Emissions Distribution
    plt.subplot(1, 4, 4)
    sns.histplot(df['Estimated Total GHG Emissions (kgCO2e)'], bins=30, kde=True)
    plt.title('GHG Emissions Distribution')
    plt.xlabel('GHG Emissions (kg CO2e)')
    plt.ylabel('Count')

    plt.tight_layout()
    plt.show()

def histogram_log(df):
    # Log Scale Distributions
    plt.figure(figsize=(20, 6))
    # Log Gross Floor Area Distribution
    plt.subplot(1, 4, 1)
    sns.histplot(np.log(df['Reported Gross Floor Area (Sq Ft)']), bins=30, kde=True)
    plt.title('Log Scale: Gross Floor Area Distribution')
    plt.xlabel('Log Gross Floor Area')
    plt.ylabel('Count')
    # Log Site EUI Distribution
    plt.subplot(1, 4, 2)
    sns.histplot(np.log(df['Site EUI (Energy Use Intensity kBtu/ft2)']), bins=30, kde=True)
    plt.title('Log Scale: Site EUI Distribution')
    plt.xlabel('Log Site EUI')
    plt.ylabel('Count')
    # Log Total Energy Distribution
    plt.subplot(1, 4, 3)
    sns.histplot(np.log(df['Total Site Energy Usage (kBtu)']), bins=30, kde=True)
    plt.title('Log Scale: Total Energy Distribution')
    plt.xlabel('Log Site Energy Usage (kBtu)')
    plt.ylabel('Count')
    # Log GHG Emissions Distribution
    plt.subplot(1, 4, 4)
    sns.histplot(np.log(df['Estimated Total GHG Emissions (kgCO2e)']), bins=30, kde=True)
    plt.title('Log Scale: GHG Emissions Distribution')
    plt.xlabel('Log GHG Emissions (kg CO2e)')
    plt.ylabel('Count')

    plt.tight_layout()
    plt.show()



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

# print original dataset size
print(f"Original Dataset Size: {multifamily_df.shape[0]}")

# drop any rows with GSF below 20,000 SF
min_gsf = 20000
multifamily_df = multifamily_df[multifamily_df['Reported Gross Floor Area (Sq Ft)'] >= min_gsf]

# print new dataset size
print(f"Dataset Size After Filtering by GSF (>= {min_gsf} sq ft): {multifamily_df.shape[0]}")

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Visualize Distributions --------------------------------------------
# --------------------------------------------------------------------------------------------------------

# visualize feature distribution
histograms(multifamily_df)
histogram_log(multifamily_df)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Log Transformations ------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# log transformation for GSF and EUI
cleaned_df = multifamily_df.copy()
cleaned_df['Log GSF'] = np.log(cleaned_df['Reported Gross Floor Area (Sq Ft)'])
cleaned_df['Log EUI'] = np.log(cleaned_df['Site EUI (Energy Use Intensity kBtu/ft2)'])
cleaned_df['Log Total Energy'] = np.log(cleaned_df['Total Site Energy Usage (kBtu)'])
cleaned_df['Log GHG Emissions'] = np.log(cleaned_df['Estimated Total GHG Emissions (kgCO2e)'])

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Percentile-Based Filtering -----------------------------------------
# --------------------------------------------------------------------------------------------------------

# set percentiles for filtering
gsf_upper_bound = np.percentile(cleaned_df['Log GSF'], 99)              # top 1% for GSF
eui_upper_bound = np.percentile(cleaned_df['Log EUI'], 98)              # top 2% for EUI
energy_upper_bound = np.percentile(cleaned_df['Log Total Energy'], 98)  # top 2% for Total Energy
ghg_upper_bound = np.percentile(cleaned_df['Log GHG Emissions'], 98)    # top 2% for GHG

# filter out the extreme values based on desired percentiles for all features
filtered_df = cleaned_df[
    (cleaned_df['Log GSF'] <= gsf_upper_bound) &
    (cleaned_df['Log EUI'] <= eui_upper_bound) &
    (cleaned_df['Log Total Energy'] <= energy_upper_bound) &
    (cleaned_df['Log GHG Emissions'] <= ghg_upper_bound)
].copy()

# print new dataset size
print(f"Dataset Size After Percentile-based Filtering: {filtered_df.shape[0]}")

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Normalization (Log) ------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# normalize log-transformed data
log_features = ['Log GSF', 'Log EUI', 'Log Total Energy', 'Log GHG Emissions']

# initialize StandardScaler object
scaler = StandardScaler()

# fit scaler to log data
X = scaler.fit_transform(filtered_df[log_features])

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Elbow Analysis -----------------------------------------------------
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
# ----------------------------------- Cluster Analysis ----------------------------------------------
# --------------------------------------------------------------------------------------------------------

# DBSCAN Parameters (Start with these, then tune as needed)
eps = 0.5
min_samples = 3

# DBSCAN Clustering
dbscan = DBSCAN(eps=eps, min_samples=min_samples)
labels_dbscan = dbscan.fit_predict(X)
filtered_df.loc[:, 'Cluster_dbscan'] = labels_dbscan

# Handle noise points (-1 label)
num_clusters = len(set(labels_dbscan)) - (1 if -1 in labels_dbscan else 0)
print(f"\nNumber of Clusters found by DBSCAN: {num_clusters}")
print(f"Number of Noise Points: {list(labels_dbscan).count(-1)}")

# --------------------------------------------------------------------------------------------------------
# ----------------- Evaluate Clustering Performance --------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Silhouette Score (Only calculate if more than 1 cluster is found)
if num_clusters > 1:
    sil_score_dbscan = silhouette_score(X, labels_dbscan)
    db_index_dbscan = davies_bouldin_score(X, labels_dbscan)
    ch_index_dbscan = calinski_harabasz_score(X, labels_dbscan)
else:
    sil_score_dbscan = None
    db_index_dbscan = None
    ch_index_dbscan = None

print("\nDBSCAN Clustering Metrics:")
print(f"Silhouette Score: {sil_score_dbscan}")
print(f"Davies-Bouldin Index: {db_index_dbscan}")
print(f"Calinski-Harabasz Index: {ch_index_dbscan}")

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- PCA Analysis -------------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# visualize clusters using PCA (2D)
pca_simple = PCA(n_components=2)
pca_complex = PCA(n_components=2)
pca_components_simple = pca_simple.fit_transform(X)
pca_components_complex = pca_complex.fit_transform(X)

# add columns
filtered_df.loc[:, 'PCA_simple_1'] = pca_components_simple[:, 0]
filtered_df.loc[:, 'PCA_simple_2'] = pca_components_simple[:, 1]
filtered_df.loc[:, 'PCA_complex_1'] = pca_components_complex[:, 0]
filtered_df.loc[:, 'PCA_complex_2'] = pca_components_complex[:, 1]

# scatter plots of clusters
# Simple Clustering
plt.figure(figsize=(12, 8))
sns.scatterplot(
    data=filtered_df,
    x='PCA_simple_1',
    y='PCA_simple_2',
    hue='Cluster_dbscan',
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
    hue='Cluster_dbscan',
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
# ----------------------------- Pair Plot for Log Features -----------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Pair Plot for Log-Transformed Features (SIMPLE)
pair_plot_features = ['Log GSF', 'Log EUI', 'Log Total Energy', 'Log GHG Emissions', 'Cluster_dbscan']
sns.pairplot(
    filtered_df[pair_plot_features],
    hue='Cluster_dbscan',
    palette='Set2')
plt.suptitle("Pair Plot of Log Features (2 Clusters)", y=1.02)
plt.show()

# Pair Plot for Log-Transformed Features (COMPLEX)
pair_plot_features = ['Log GSF', 'Log EUI', 'Log Total Energy', 'Log GHG Emissions', 'Cluster_dbscan']
sns.pairplot(
    filtered_df[pair_plot_features],
    hue='Cluster_dbscan',
    palette='Set2')
plt.suptitle("Pair Plot of Log Features (Complex Clusters)", y=1.02)
plt.show()

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Prototype Properties -----------------------------------------------
# --------------------------------------------------------------------------------------------------------
#
# # get cluster centriods of prototypes
# cluster_centers_simple = scaler.inverse_transform(gmm_simple.cluster_centers_)
# cluster_centers_complex = scaler.inverse_transform(gmm_complex.cluster_centers_)
#
# # create dataframes of the prototype features
# prototype_simple_df = np.exp(pd.DataFrame(cluster_centers_simple, columns=features))
# prototype_complex_df = np.exp(pd.DataFrame(cluster_centers_complex, columns=features))
# # add cluster index
# prototype_simple_df['Cluster_simple'] = prototype_simple_df.index
# prototype_complex_df['Cluster_complex'] = prototype_complex_df.index
#
# # display properties of prototypes (centriods)
# print("Prototypes (Centriods of Clusters - Simple):")
# print(prototype_simple_df)
# print("Prototypes (Centriods of Clusters - Complex):")
# print(prototype_complex_df)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Visualize Clusters -------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# plot characteristics of prototypes
# simple
plt.figure(figsize=(12, 8))
sns.scatterplot(
    data=filtered_df,
    x='Reported Gross Floor Area (Sq Ft)',
    y='Site EUI (Energy Use Intensity kBtu/ft2)',
    hue='Cluster_dbscan',
    palette='Set2'
)
plt.title('Simple Clusters on GSF vs. EUI (Percentile-Filtered)')
plt.xlabel('Gross Floor Area (Sq Ft)')
plt.ylabel('Site EUI (kBtu/sf)')
plt.legend(title='Cluster')
plt.grid(True)
plt.show()

# complex
plt.figure(figsize=(12, 8))
sns.scatterplot(
    data=filtered_df,
    x='Reported Gross Floor Area (Sq Ft)',
    y='Site EUI (Energy Use Intensity kBtu/ft2)',
    hue='Cluster_dbscan',
    palette='Set2'
)
plt.title('Complex Clusters on GSF vs. EUI (Percentile-Filtered)')
plt.xlabel('Gross Floor Area (Sq Ft)')
plt.ylabel('Site EUI (kBtu/sf)')
plt.legend(title='Cluster')
plt.grid(True)
plt.show()
