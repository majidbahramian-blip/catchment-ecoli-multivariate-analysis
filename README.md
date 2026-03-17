
# E. coli Catchment Multivariate Analysis

This repository contains Python workflows for multivariate analysis of E. coli concentrations and environmental variables in catchments using:

- Principal Component Analysis (PCA)
- Hierarchical clustering
- K-means clustering
- Visualization of PCA loadings and cluster structures

---

## Research Context

This work supports environmental modelling of microbial contamination in catchments, particularly focusing on:

- Transport and persistence of E. coli
- Relationships with hydrological and meteorological variables
- Identification of dominant drivers using dimensionality reduction

---

## Methods

### 1. Data preprocessing
- Standardization (z-score normalization)
- Handling missing values
- Variable selection

### 2. PCA
- Correlation matrix-based PCA
- Extraction of principal components
- Interpretation of loadings and explained variance

### 3. Clustering
- Hierarchical clustering (linkage-based)
- K-means clustering (variable grouping)
- Cluster validation via PCA space

### 4. Visualization
- PCA loading plots
- Cluster-colored variable maps
- Scatter plots in PC space

---

## 📂 Repository Structure

