import os
os.makedirs("results/figures", exist_ok=True)
os.makedirs("results/tables", exist_ok=True)
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.decomposition import PCA
from scipy.cluster.hierarchy import dendrogram, linkage

# ==============================
# Global plotting style
# ==============================
plt.rcParams['font.family'] = 'Times New Roman'
plt.rcParams['font.size'] = 12
plt.rcParams['axes.titlesize'] = 14
plt.rcParams['axes.labelsize'] = 12
plt.rcParams['xtick.labelsize'] = 11
plt.rcParams['ytick.labelsize'] = 11

# ==============================
# 1. Read standardized input data
# ==============================
file_path = "results/tables/something.xlsx"
sheet_name = "Sheet1"

df = pd.read_excel(file_path, sheet_name=sheet_name)
df = df.select_dtypes(include=[np.number])

# ==============================
# 2. Descriptive statistics
# ==============================
desc = pd.DataFrame({
    "Mean": df.mean(),
    "Standard Deviation": df.std(ddof=1),
    "N": df.count()
})

# ==============================
# 3. Correlation matrix
# ==============================
corr_matrix = df.corr()

# ==============================
# 4. PCA
# ==============================
pca = PCA()
scores = pca.fit_transform(df)

eigenvalues = pca.explained_variance_
explained_ratio = pca.explained_variance_ratio_
cumulative = np.cumsum(explained_ratio)

difference = np.empty_like(eigenvalues)
difference[:-1] = eigenvalues[:-1] - eigenvalues[1:]
difference[-1] = np.nan

pc_names = [f"PC{i+1}" for i in range(len(df.columns))]

# ==============================
# 5. Eigenvectors
# ==============================
eigenvectors = pd.DataFrame(
    pca.components_.T,
    index=df.columns,
    columns=pc_names
)

# Flip PC1 sign to match Excel convention
if "Weekly Mean Air Temperature (°C)" in eigenvectors.index:
    if eigenvectors.loc["Weekly Mean Air Temperature (°C)", "PC1"] < 0:
        eigenvectors["PC1"] = -eigenvectors["PC1"]
        scores[:, 0] = -scores[:, 0]

# ==============================
# 6. Component loadings
# ==============================
loadings = eigenvectors.copy()
for i, ev in enumerate(eigenvalues):
    loadings.iloc[:, i] = eigenvectors.iloc[:, i] * np.sqrt(ev)

# ==============================
# 7. Squared loadings and communalities
# ==============================
cos2 = loadings ** 2
communalities = cos2.cumsum(axis=1)

# ==============================
# 8. Variance explained table
# ==============================
variance_table = pd.DataFrame({
    "#": range(1, len(eigenvalues) + 1),
    "Eigenvalue": eigenvalues,
    "Difference": difference,
    "Proportion": explained_ratio,
    "Cumulative": cumulative
})

# ==============================
# 9. Variable contributions
# ==============================
contrib = loadings.pow(2).div(loadings.pow(2).sum(axis=0), axis=1) * 100

# ==============================
# 10. Save results to Excel
# ==============================
output_file = r"D:\DERIVE\PCA\PCA_results_python.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    desc.to_excel(writer, sheet_name="Descriptive_Statistics")
    corr_matrix.to_excel(writer, sheet_name="Correlation_Matrix")
    variance_table.to_excel(writer, sheet_name="Variance_Explained", index=False)
    eigenvectors.to_excel(writer, sheet_name="Eigenvectors")
    loadings.to_excel(writer, sheet_name="Component_Loadings")
    cos2.to_excel(writer, sheet_name="Squared_Loadings_cos2")
    communalities.to_excel(writer, sheet_name="Communalities")
    contrib.to_excel(writer, sheet_name="Variable_Contributions")
    pd.DataFrame(scores, columns=pc_names).to_excel(writer, sheet_name="Scores", index=False)

print("PCA tables saved to Excel.")

# ============================================================
# FIGURE 1. Scree Plot
# ============================================================
plt.figure(figsize=(8,5))
plt.plot(range(1,len(eigenvalues)+1), eigenvalues, marker='o')
plt.xlabel("Component")
plt.ylabel("Eigenvalue")
plt.title("Scree Plot")
plt.xticks(range(1,len(eigenvalues)+1))
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig("results/figures/pca_loading_plot.png", dpi=300)
plt.close()

# ============================================================
# FIGURE 2. Variance Explained
# ============================================================
plt.figure(figsize=(8,5))
plt.plot(range(1,len(explained_ratio)+1), explained_ratio, marker='o', label="Proportion")
plt.plot(range(1,len(cumulative)+1), cumulative, marker='o', linestyle='--', label="Cumulative")
plt.xlabel("Component")
plt.ylabel("Variance Explained")
plt.title("Variance Explained by PCA Components")
plt.legend()
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig("results/figures/pca_loading_plot.png", dpi=300)
plt.close()

# ============================================================
# FIGURE 3. Loading Plot
# ============================================================
plt.figure(figsize=(9,9))

for var in loadings.index:
    x = loadings.loc[var,"PC1"]
    y = loadings.loc[var,"PC2"]
    plt.arrow(0,0,x,y, head_width=0.03)
    plt.text(x*1.07, y*1.07, var, fontsize=9)

theta = np.linspace(0,2*np.pi,400)
plt.plot(np.cos(theta), np.sin(theta))

plt.axhline(0)
plt.axvline(0)
plt.xlabel(f"PC1 ({explained_ratio[0]*100:.2f}%)")
plt.ylabel(f"PC2 ({explained_ratio[1]*100:.2f}%)")
plt.title("Variable Loadings Plot (PC1 vs PC2)")
plt.axis("equal")
plt.grid(alpha=0.3)

plt.tight_layout()
plt.savefig("results/figures/pca_loading_plot.png", dpi=300)
plt.close()

# ============================================================
# FIGURE 4. Score Plot
# ============================================================
plt.figure(figsize=(8,6))
plt.scatter(scores[:,0], scores[:,1], alpha=0.7)

plt.axhline(0)
plt.axvline(0)

plt.xlabel(f"PC1 ({explained_ratio[0]*100:.2f}%)")
plt.ylabel(f"PC2 ({explained_ratio[1]*100:.2f}%)")
plt.title("Sample Scores Plot (PC1 vs PC2)")
plt.grid(alpha=0.3)

plt.tight_layout()
plt.savefig("results/figures/pca_loading_plot.png", dpi=300)
plt.close()

# ============================================================
# FIGURE 5. Loadings Heatmap
# ============================================================
fig, ax = plt.subplots(figsize=(10,8))

heat_data = loadings.iloc[:, :5].values
im = ax.imshow(heat_data, aspect='auto')

ax.set_xticks(np.arange(5))
ax.set_xticklabels(loadings.columns[:5])
ax.set_yticks(np.arange(len(loadings.index)))
ax.set_yticklabels(loadings.index)

plt.title("Component Loadings Heatmap (PC1-PC5)")
plt.colorbar(im)

plt.tight_layout()
plt.savefig("results/figures/pca_loading_plot.png", dpi=300)
plt.close()

# ============================================================
# FIGURE 6. cos2 Heatmap
# ============================================================
fig, ax = plt.subplots(figsize=(10,8))

heat_data = cos2.iloc[:, :5].values
im = ax.imshow(heat_data, aspect='auto')

ax.set_xticks(np.arange(5))
ax.set_xticklabels(cos2.columns[:5])
ax.set_yticks(np.arange(len(cos2.index)))
ax.set_yticklabels(cos2.index)

plt.title("Squared Loadings (cos2) Heatmap (PC1-PC5)")
plt.colorbar(im)

plt.tight_layout()
plt.savefig("results/figures/pca_loading_plot.png", dpi=300)
plt.close()

# ============================================================
# FIGURE 7. Contribution PC1
# ============================================================
pc1_sorted = contrib["PC1"].sort_values(ascending=False)

plt.figure(figsize=(10,6))
plt.bar(pc1_sorted.index, pc1_sorted.values)
plt.xticks(rotation=90)
plt.ylabel("Contribution (%)")
plt.title("Variable Contributions to PC1")

plt.tight_layout()
plt.savefig("results/figures/pca_loading_plot.png", dpi=300)
plt.close()

# ============================================================
# FIGURE 8. Contribution PC2
# ============================================================
pc2_sorted = contrib["PC2"].sort_values(ascending=False)

plt.figure(figsize=(10,6))
plt.bar(pc2_sorted.index, pc2_sorted.values)
plt.xticks(rotation=90)
plt.ylabel("Contribution (%)")
plt.title("Variable Contributions to PC2")

plt.tight_layout()
 plt.savefig("results/figures/pca_loading_plot.png", dpi=300)
plt.close()

# ============================================================
# FIGURE 9. Hierarchical clustering of variables 
# ============================================================

data_vars = df.T
Z = linkage(data_vars, method='ward', metric='euclidean')
plt.rcParams['lines.linewidth'] = 0.8
plt.figure(figsize=(10,6))
dendrogram(
    Z,
    labels=data_vars.index,
    leaf_rotation=90,
    leaf_font_size=10,
    color_threshold=0,
    above_threshold_color='black',
    link_color_func=lambda k: 'black'
)
plt.title("Hierarchical Clustering of Variables", color='black')
plt.ylabel("Distance", color='black')
plt.xticks(color='black')
plt.yticks(color='black')
# make axes black
ax = plt.gca()
for spine in ax.spines.values():
    spine.set_color('black')
plt.tight_layout()
plt.savefig(
    r"D:\DERIVE\PCA\fig9_variable_dendrogram.png",
    dpi=600,
    bbox_inches="tight"
)
plt.close()

from sklearn.cluster import KMeans

# ============================================================
# K-means clustering of variables
# ============================================================

# transpose so variables become observations
data_vars = df.T

# number of clusters (adjust if needed)
k = 3

kmeans_vars = KMeans(n_clusters=k, random_state=42, n_init=20)
var_clusters = kmeans_vars.fit_predict(data_vars)

# save cluster assignments
var_cluster_table = pd.DataFrame({
    "Variable": data_vars.index,
    "Cluster": var_clusters
})

var_cluster_table.to_excel(
    r"D:\DERIVE\PCA\kmeans_variable_clusters.xlsx",
    index=False
)

print("Variable clusters saved.")

# ============================================================
# FIGURE: Variable clusters in PCA loading space
# ============================================================

plt.figure(figsize=(9,9))

colors = ["black","red","blue","green","purple","orange"]

# plot variables
for i, var in enumerate(loadings.index):

    cluster = var_clusters[i]

    x = loadings.loc[var,"PC1"]
    y = loadings.loc[var,"PC2"]

    plt.scatter(
        x,
        y,
        color=colors[cluster],
        s=80
    )

    plt.text(
        x*1.05,
        y*1.05,
        var,
        fontsize=9
    )

# -------- LEGEND (ADDED HERE) --------
for c in range(k):
    plt.scatter(
        [],
        [],
        color=colors[c],
        label=f"Cluster {c+1}"
    )

plt.legend(frameon=False)
# ------------------------------------

plt.axhline(0,color='black',linewidth=0.8)
plt.axvline(0,color='black',linewidth=0.8)

plt.xlabel(f"PC1 ({explained_ratio[0]*100:.2f}%)")
plt.ylabel(f"PC2 ({explained_ratio[1]*100:.2f}%)")

plt.title("K-means Clustering of Variables in PCA Space")

plt.grid(alpha=0.3)

plt.tight_layout()

plt.savefig(
    r"D:\DERIVE\PCA\fig10_kmeans_variable_clusters.png",
    dpi=600,
    bbox_inches="tight"
)

plt.close()
