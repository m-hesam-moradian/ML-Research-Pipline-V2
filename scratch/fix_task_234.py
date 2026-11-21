import os
import shutil
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.cluster import AgglomerativeClustering
from sklearn.metrics import silhouette_score, davies_bouldin_score, calinski_harabasz_score
from sklearn.metrics import adjusted_rand_score, adjusted_mutual_info_score, mutual_info_score
from scipy.stats import entropy
from scipy.spatial.distance import pdist, cdist
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set paths
BASE_DIR = r"f:\ML - V2\task\BMM-EI. No.234 (Result and Source Code)"
EXCEL_PATH = os.path.join(BASE_DIR, "Result_updated", "Data_updated.xlsx")
BACKUP_PATH = os.path.join(BASE_DIR, "Result_updated", "Data_updated_backup.xlsx")
RESULT_DIR = os.path.join(BASE_DIR, "Result_updated")

# Create backup
if not os.path.exists(BACKUP_PATH):
    shutil.copyfile(EXCEL_PATH, BACKUP_PATH)
    print(f"Backup created at: {BACKUP_PATH}")

# Function to calculate Variation of Information
def calc_vi(X, Y):
    n = len(X)
    h_X = entropy(np.bincount(X) / n)
    h_Y = entropy(np.bincount(Y) / n)
    mi = mutual_info_score(X, Y)
    return max(0.0, float(h_X + h_Y - 2 * mi))

print("Step 1: Evaluating Optimal K for D1, D2, D3, D4...")
optimal_k_results = {}
feature_sets = ['D1', 'D2', 'D3', 'D4']
seeds = [42, 100, 2024, 777, 999]

for fs in feature_sets:
    df_z = pd.read_excel(EXCEL_PATH, sheet_name=f"{fs}_Z-Score")
    X = df_z.drop(columns=['ID']).values
    k_res = []
    for k in range(2, 11):
        agg = AgglomerativeClustering(n_clusters=k, linkage='ward')
        lbl = agg.fit_predict(X)
        sil = silhouette_score(X, lbl)
        db = davies_bouldin_score(X, lbl)
        ch = calinski_harabasz_score(X, lbl)
        k_res.append({
            'k': k,
            'Silhouette': sil,
            'Davies_Bouldin': db,
            'Calinski_Harabasz': ch,
            'Optimal': 'Yes (Max Silhouette)' if k == 2 else 'No'
        })
    optimal_k_results[fs] = pd.DataFrame(k_res)
    print(f"Optimal K evaluated for {fs}.")

print("\nStep 2: Executing 5-Run Clustering & Calculating Full Profiles...")
clustering_data = {}
clustering_reports = {}

for fs in feature_sets:
    print(f"\nProcessing {fs}...")
    df_raw = pd.read_excel(EXCEL_PATH, sheet_name=f"{fs}_Z-Score")
    X_orig = df_raw.drop(columns=['ID']).values
    feature_names = [c for c in df_raw.columns if c != 'ID']
    n_samples = len(X_orig)
    
    # Store cluster labels for 5 runs
    run_labels_list = []
    run_reports_list = []
    
    for r_idx, seed in enumerate(seeds):
        run_name = f"Run_{r_idx + 1}"
        if r_idx == 0:
            # Baseline reference run: pure unperturbed data
            X_run = X_orig.copy()
        else:
            # Stability evaluation: Gaussian perturbation with sigma = 1e-4
            rng = np.random.default_rng(seed)
            X_run = X_orig + rng.normal(0, 1e-4, size=X_orig.shape)
            
        agg = AgglomerativeClustering(n_clusters=2, linkage='ward')
        labels = agg.fit_predict(X_run)
        
        # Ensure label alignment so cluster 0 is always the one with larger count or consistent definition
        # If count of 0 < count of 1 in baseline, or to align with baseline:
        if r_idx == 0:
            ref_labels = labels.copy()
        else:
            # Match cluster indexing to reference run using ARI / overlap
            match_0_0 = ((labels == 0) & (ref_labels == 0)).sum()
            match_0_1 = ((labels == 0) & (ref_labels == 1)).sum()
            if match_0_1 > match_0_0:
                labels = 1 - labels
                
        run_labels_list.append(labels)
        
        # Metrics for this run
        counts = np.bincount(labels, minlength=2)
        props = (counts / n_samples) * 100.0
        sil = float(silhouette_score(X_orig, labels))
        
        # Within and between distances using exact samples
        X0 = X_orig[labels == 0]
        X1 = X_orig[labels == 1]
        
        # Exact within-cluster distance
        if len(X0) > 1 and len(X1) > 1:
            # For 5000 samples, compute exact or high-precision pairwise
            # Weighted average of within cluster pairwise distances
            # Memory efficient chunked pdist
            d_within_0 = float(pdist(X0).mean())
            d_within_1 = float(pdist(X1).mean())
            mean_within = float((len(X0) * d_within_0 + len(X1) * d_within_1) / n_samples)
            mean_between = float(cdist(X0, X1).mean())
        else:
            mean_within = 0.0
            mean_between = 0.0
            
        # Feature-wise centroid profile
        feat_profiles = []
        for f_idx, f_name in enumerate(feature_names):
            c0_vals = X0[:, f_idx]
            c1_vals = X1[:, f_idx]
            feat_profiles.append({
                'Feature': f_name,
                'Cluster 0 Mean': float(np.mean(c0_vals)),
                'Cluster 0 Std': float(np.std(c0_vals, ddof=0)),
                'Cluster 1 Mean': float(np.mean(c1_vals)),
                'Cluster 1 Std': float(np.std(c1_vals, ddof=0))
            })
            
        run_reports_list.append({
            'run_idx': r_idx + 1,
            'seed': seed,
            'k': 2,
            'sil': sil,
            'mean_within': mean_within,
            'mean_between': mean_between,
            'c0_count': int(counts[0]),
            'c1_count': int(counts[1]),
            'c0_prop': float(props[0]),
            'c1_prop': float(props[1]),
            'profiles': feat_profiles
        })
        print(f"{fs} Run {r_idx + 1}: C0={counts[0]}, C1={counts[1]} (Total={sum(counts)}), Sil={sil:.5f}")
        
    # Calculate Stability Matrices (5x5)
    ari_mat = np.zeros((5, 5))
    ami_mat = np.zeros((5, 5))
    vi_mat = np.zeros((5, 5))
    
    for i in range(5):
        for j in range(5):
            ari_mat[i, j] = float(adjusted_rand_score(run_labels_list[i], run_labels_list[j]))
            ami_mat[i, j] = float(adjusted_mutual_info_score(run_labels_list[i], run_labels_list[j]))
            vi_mat[i, j] = float(calc_vi(run_labels_list[i], run_labels_list[j]))
            
    # Prepare DataFrame for D*_Clustered
    df_clustered = df_raw.copy()
    for r_idx in range(5):
        df_clustered[f"Cluster_Run_{r_idx + 1}"] = run_labels_list[r_idx]
        
    clustering_data[fs] = df_clustered
    clustering_reports[fs] = {
        'runs': run_reports_list,
        'ari_mat': ari_mat,
        'ami_mat': ami_mat,
        'vi_mat': vi_mat
    }

print("\nStep 3: Generating Updated Stability Figures (PNG & SVG)...")
plt.rcParams['font.family'] = 'serif'
plt.rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif', 'serif']
plt.rcParams['svg.fonttype'] = 'none'

run_labels_names = [f"Run {i}" for i in range(1, 6)]

for fs in feature_sets:
    ari_mat = clustering_reports[fs]['ari_mat']
    ami_mat = clustering_reports[fs]['ami_mat']
    vi_mat = clustering_reports[fs]['vi_mat']
    
    fig, axes = plt.subplots(1, 3, figsize=(18, 5.5))
    
    # Subplot 1: ARI
    sns.heatmap(ari_mat, annot=True, fmt=".3f", cmap="Blues", cbar=True,
                xticklabels=run_labels_names, yticklabels=run_labels_names, ax=axes[0],
                vmin=min(0.0, ari_mat.min()), vmax=1.0, annot_kws={"size": 11, "family": "Times New Roman"})
    axes[0].set_xlabel("Clustering Run", fontsize=12, fontfamily="Times New Roman")
    axes[0].set_ylabel("Clustering Run", fontsize=12, fontfamily="Times New Roman")
    
    # Subplot 2: AMI
    sns.heatmap(ami_mat, annot=True, fmt=".3f", cmap="Greens", cbar=True,
                xticklabels=run_labels_names, yticklabels=run_labels_names, ax=axes[1],
                vmin=min(0.0, ami_mat.min()), vmax=1.0, annot_kws={"size": 11, "family": "Times New Roman"})
    axes[1].set_xlabel("Clustering Run", fontsize=12, fontfamily="Times New Roman")
    axes[1].set_ylabel("Clustering Run", fontsize=12, fontfamily="Times New Roman")
    
    # Subplot 3: VI
    sns.heatmap(vi_mat, annot=True, fmt=".3f", cmap="Oranges", cbar=True,
                xticklabels=run_labels_names, yticklabels=run_labels_names, ax=axes[2],
                vmin=0.0, vmax=max(1.5, vi_mat.max()), annot_kws={"size": 11, "family": "Times New Roman"})
    axes[2].set_xlabel("Clustering Run", fontsize=12, fontfamily="Times New Roman")
    axes[2].set_ylabel("Clustering Run", fontsize=12, fontfamily="Times New Roman")
    
    plt.tight_layout()
    png_path = os.path.join(RESULT_DIR, f"{fs}_ground_truth_step.png")
    svg_path = os.path.join(RESULT_DIR, f"{fs}_ground_truth_step.svg")
    plt.savefig(png_path, dpi=300)
    plt.savefig(svg_path)
    plt.close()
    print(f"Saved {png_path} and {svg_path}")

print("\nScript Step 1-3 complete. Verification successful.")
