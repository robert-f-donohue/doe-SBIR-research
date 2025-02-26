import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from yellowbrick.cluster import SilhouetteVisualizer
from sklearn.decomposition import PCA
from sklearn.metrics import silhouette_score, davies_bouldin_score, calinski_harabasz_score
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns
sns.set_theme(style="whitegrid", palette="pastel")


# --------------------------------------------------------------------------------------------------------
# ---------------------------- Log Transformed Features ---------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# List of features that were log-transformed before clustering
features = [
    'Reported Gross Floor Area (Sq Ft)',
    'Site EUI (Energy Use Intensity kBtu/ft2)',
    'Total Site Energy Usage (kBtu)',
    'Estimated Total GHG Emissions (kgCO2e)',
    'Percentage Electrification'
]


def get_prototypes_and_export(centers, cluster_count, filename_prefix):
    """
    Inverse transform and export cluster centroids as prototypes.

    Parameters:
    - centers: Cluster centroids from KMeans or KMedoids
    - cluster_count: Number of clusters (e.g., 2 or 8)
    - filename_prefix: Prefix for the output file (e.g., 'simple' or 'complex')
    """
    # Inverse transform to return to original scale
    inverse_transformed = scaler.inverse_transform(centers)

    # Apply inverse log transform (exp) BEFORE creating the DataFrame
    for i, feature in enumerate(features):
        if feature in features:
            inverse_transformed[:, i] = np.exp(inverse_transformed[:, i])

    # Create DataFrame for centroids
    prototype_df = pd.DataFrame(inverse_transformed, columns=features)

    # Add cluster labels for easy reference
    prototype_df['Cluster'] = range(cluster_count)

    # Display the prototypes
    print(f"\nPrototypes (Centroids of Clusters - {filename_prefix.capitalize()}):")
    print(prototype_df)

    # Export to CSV
    output_filename = f'../data-files/0-final-analysis/{optimal_clusters_complex}_clusters/{filename_prefix}_centroids_{cluster_count}_clusters.xlsx'
    # prototype_df.to_excel(output_filename, index=False)
    print(f"Exported prototypes to {output_filename}.xlsx")

    return prototype_df


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Setting Up Analysis ------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# File path to 2024 BERDO data
file_path_multifamily_2024 = '../data-files/berdo_data_files/BERDO_Data-2024-multifamily.csv'

# create DataFrames for 2021 and 2024 data
df_2024 = pd.read_csv(file_path_multifamily_2024)

# Keep clean copy for transformation
cleaned_df = df_2024.copy()

# Create dictionary to store Box-Cox lambda values for each feature
boxcox_lambdas = {}

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Box-Cox Transformation -------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Apply Box-Cox Transformation for each feature
for feature in features:
    # ensure all are positive
    cleaned_df = cleaned_df[cleaned_df[feature] > 0]

    # apply transformation and store lambda values
    transformed_data, lambda_value = stats.boxcox(cleaned_df[feature])
    cleaned_df[f'BoxCox {feature}'] = transformed_data
    boxcox_lambdas[feature] = lambda_value


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
range_n_clusters = range(2, 9)
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
# plt.show()

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Silhouette Analysis -----------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Create subplots: 2 rows, 3 columns
fig, axs = plt.subplots(2, 3, figsize=(24, 16), sharey=True)
fig.suptitle('Silhouette Analysis for KMeans Clustering', fontsize=24)

# Flatten the axs array for easier iteration
axs = axs.flatten()

# Loop through the number of clusters and create silhouette plots
for ax, n_clusters in zip(axs, range_n_clusters):
    # Initialize the KMeans model
    kmeans = KMeans(n_clusters=n_clusters, random_state=42)

    # Create SilhouetteVisualizer
    visualizer = SilhouetteVisualizer(kmeans, colors='coolwarm', ax=ax)
    visualizer.fit(X)
    visualizer.finalize()
    ax.set_title(f'{n_clusters} Clusters')

# Remove any empty subplots (if the number of clusters is not a multiple of the grid)
for i in range(len(range_n_clusters), len(axs)):
    fig.delaxes(axs[i])

plt.tight_layout(rect=[0, 0.03, 1, 0.95])  # Adjust layout to fit the title
# save figure to correct file path
plt.savefig(
    f'../data-files/0-final-analysis/0-summary/silhouette_scores.jpg',
    format='jpg',
    dpi=300,
)
# plt.show()

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Cluster Analysis - KMeans ------------------------------------------
# --------------------------------------------------------------------------------------------------------

# perform K-Means Clustering with 2 clusters (prototypes)
optimal_clusters_simple = 2
kmeans_simple = KMeans(n_clusters=optimal_clusters_simple, random_state=42)
labels_simple = kmeans_simple.fit_predict(X)
filtered_df.loc[:, 'Cluster_simple'] = labels_simple

# perform K-Means Clustering with 8 clusters (prototypes)
optimal_clusters_complex = 8
kmeans_complex = KMeans(n_clusters=optimal_clusters_complex, random_state=42)
labels_complex = kmeans_complex.fit_predict(X)
filtered_df.loc[:, 'Cluster_complex'] = labels_complex

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Cluster Performance ------------------------------------------------
# --------------------------------------------------------------------------------------------------------

cluster_type = 'K-Means'

# Metrics for Simple Clustering (2 Clusters)
sil_score_simple = silhouette_score(X, labels_simple)
db_index_simple = davies_bouldin_score(X, labels_simple)
ch_index_simple = calinski_harabasz_score(X, labels_simple)

print(f"\nSimple {cluster_type} Clustering (2 Clusters) Metrics:")
print(f"Silhouette Score: {sil_score_simple}")
print(f"Davies-Bouldin Index: {db_index_simple}")
print(f"Calinski-Harabasz Index: {ch_index_simple}")

# Metrics for Complex Clustering (6 Clusters)
sil_score_complex = silhouette_score(X, labels_complex)
db_index_complex = davies_bouldin_score(X, labels_complex)
ch_index_complex = calinski_harabasz_score(X, labels_complex)

print(f"\nComplex {cluster_type} Clustering ({optimal_clusters_complex} Clusters) Metrics:")
print(f"Silhouette Score: {sil_score_complex}")
print(f"Davies-Bouldin Index: {db_index_complex}")
print(f"Calinski-Harabasz Index: {ch_index_complex}")

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

# # save figure to correct file path
# plt.savefig(
#     f'../data-files/0-final-analysis/{optimal_clusters_simple}_clusters/scatter_pca_{optimal_clusters_simple}_clusters.jpg',
#     format='jpg',
#     dpi=300,
# )
# plt.show()

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

# # save figure to correct file path
# plt.savefig(
#     f'../data-files/0-final-analysis/{optimal_clusters_complex}_clusters/scatter_pca_{optimal_clusters_complex}_clusters.jpg',
#     format='jpg',
#     dpi=300,
# )
# plt.show()

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

# # save figure to correct file path
# plt.savefig(
#     f'../data-files/0-final-analysis/{optimal_clusters_simple}_clusters/scatter_gsf_vs_eui_{optimal_clusters_simple}_clusters_log.jpg',
#     format='jpg',
#     dpi=300,
# )
# plt.show()

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

# # save figure to correct file path
# plt.savefig(
#     f'../data-files/0-final-analysis/{optimal_clusters_complex}_clusters/scatter_gsf_vs_eui_{optimal_clusters_complex}_clusters_log.jpg',
#     format='jpg',
#     dpi=300,
# )
# plt.show()

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

# # save figure to correct file path
# plt.savefig(
#     f'../data-files/0-final-analysis/{optimal_clusters_simple}_clusters/scatter_gsf_vs_eui_{optimal_clusters_simple}_clusters.jpg',
#     format='jpg',
#     dpi=300,
# )
# plt.show()

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

# # save figure to correct file path
# plt.savefig(
#     f'../data-files/0-final-analysis/{optimal_clusters_complex}_clusters/scatter_gsf_vs_eui_{optimal_clusters_complex}_clusters.jpg',
#     format='jpg',
#     dpi=300,
# )
# plt.show()

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Violin Plot Visualizations -----------------------------------------
# --------------------------------------------------------------------------------------------------------

# initialize list to store statistics
statistics_list = []

# Choose the cluster column
cluster_column = 'Cluster_complex'

# Create subplots
fig, axs = plt.subplots(2, 2, figsize=(24, 16))
fig.suptitle('Distribution of Metrics within Each Cluster', fontsize=24)
axs = axs.flatten()

# Loop through metrics and create a violin plot for each
for ax, metric in zip(axs, features):
    sns.violinplot(
        data=filtered_df,
        x=cluster_column,
        y=metric,
        hue='Cluster_complex',
        palette='Set2',
        ax=ax
    )
    # calculate and annotate statistics
    for cluster in sorted(filtered_df[cluster_column].unique()):
        # filter data for the cluster
        cluster_data = filtered_df[filtered_df[cluster_column] == cluster][metric]

        # calculate statistics
        min_val = cluster_data.min()
        max_val = cluster_data.max()
        median_val = cluster_data.median()
        q1 = cluster_data.quantile(0.25)
        q3 = cluster_data.quantile(0.75)
        mean_val = cluster_data.mean()
        std_val = cluster_data.std()

        # annotate statistics on plot
        ax.text(cluster, median_val, 'Median Value: ' + str(median_val),
                color='black', ha='center', va='center', fontsize=10)
        ax.text(cluster, q1, 'Q1: ' + str(q1),
                color='blue', ha='center', va='center', fontsize=9)
        ax.text(cluster, q3, 'Q3: ' + str(q3),
                color='blue', ha='center', va='center', fontsize=9)

        # store statistics for export
        statistics_list.append({
            'Metric': metric,
            'Cluster': cluster,
            'Min': min_val,
            'Max': max_val,
            'Median': median_val,
            'Q1': q1,
            'Q3': q3,
            'Mean': mean_val,
            'Std Dev': std_val,
        })

    # set titles and labels
    ax.set_title(metric)
    ax.set_xlabel('Cluster')
    ax.set_ylabel(metric)
    ax.grid(True)

plt.tight_layout(rect=[0, 0.03, 1, 0.95])  # Adjust layout for title
plt.show()

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Export Prototype Properties  ---------------------------------------
# --------------------------------------------------------------------------------------------------------

# Get and export prototypes for 2 clusters (Simple Clustering)
prototype_simple_df = get_prototypes_and_export(kmeans_simple.cluster_centers_, 2, 'simple')

# Get and export prototypes for 8 clusters (Complex Clustering)
prototype_complex_df = get_prototypes_and_export(
    kmeans_complex.cluster_centers_,
    optimal_clusters_complex,
    'complex')

# display properties of prototypes (centriods)
print("Prototypes (Centriods of Clusters - Simple):")
print(prototype_simple_df)
print(f"Prototypes (Centriods of Clusters - {optimal_clusters_complex}):")
print(prototype_complex_df)

# --------------------------------------------------------------------------------------------------------
# ---------------------------- Cross-Reference with Original Dataset -------------------------------------
# --------------------------------------------------------------------------------------------------------

# Merge the original columns back to the cleaned DataFrame before clustering
# This ensures that we keep all relevant features for cross-referencing
original_df = multifamily_df.copy()

# Keep original index to merge back later
original_df.reset_index(drop=True, inplace=True)
filtered_df.reset_index(drop=True, inplace=True)

# Add cluster labels for both 2 and 8 clusters
filtered_df['Cluster_2'] = labels_simple
filtered_df[f'Cluster_{optimal_clusters_complex}'] = labels_complex

# Merge cluster labels back to the original dataset
# This keeps all columns, including the ones not used in clustering
original_df['Cluster_2'] = filtered_df['Cluster_2']
original_df[f'Cluster_{optimal_clusters_complex}'] = filtered_df[f'Cluster_{optimal_clusters_complex}']

# --------------------------------------------------------------------------------------------------------
# ---------------------------- Export Cross-Referenced Data ----------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Export the full dataset with cluster labels and all original features
# original_df.to_excel(
#     f'../data-files/0-final-analysis/{optimal_clusters_complex}_clusters/context_data_{optimal_clusters_complex}.xlsx',
#     index=False
# )
print("Exported cross-referenced clustered data to cross_referenced_clustered_data.xlsx")

