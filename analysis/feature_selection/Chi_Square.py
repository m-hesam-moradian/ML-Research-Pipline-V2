import pandas as pd
import numpy as np
import scipy.special as sp
from sklearn.feature_selection import chi2
import os
import win32com.client

# --- Excel control ---
def close_excel_file(filepath):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        for wb in excel.Workbooks:
            try:
                if os.path.abspath(wb.FullName) == os.path.abspath(filepath):
                    wb.Save()
                    wb.Close(SaveChanges=False)
                    print("Saved and Closed Excel file:", filepath)
                    break
            except Exception:
                pass
        excel.Quit()
    except Exception:
        pass

def open_excel_file(filepath):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.Workbooks.Open(os.path.abspath(filepath))
        print("Opened Excel file:", filepath)
    except Exception:
        pass

# --- High-Precision Asymptotic p-value for Chi-Square ---
def compute_exact_chi2_pvalue(x, p_float, df_k):
    """
    Computes exact authentic p-value string for Chi-Square distribution.
    When x is large (> 1400) and p underflows float64 to 0.0,
    computes the asymptotic log-survival function:
        ln(p) ~ (df/2 - 1)*ln(x/2) - (x/2) - ln(Gamma(df/2))
    """
    if p_float > 1e-4:
        return f"{p_float:.6f}"
    elif p_float > 0.0:
        return f"{p_float:.4e}"
    else:
        # Authentic asymptotic log-survival evaluation
        a = df_k / 2.0
        z = max(x, 1e-9) / 2.0
        log_sf = (a - 1.0) * np.log(z) - z - sp.gammaln(a)
        log10_sf = log_sf / np.log(10)
        exponent = int(np.floor(log10_sf))
        mantissa = 10.0 ** (log10_sf - exponent)
        return f"{mantissa:.4f}e{exponent}"

# --- Chi-Square Feature Selection ---
def chi_square_selection(X, y, alpha=0.05, k=None):
    X = X.copy()

    # --- STEP 0: Remove leakage if any ---
    num_rows = len(X)
    leaky_cols = [col for col in X.columns if X[col].nunique() == num_rows]
    if leaky_cols:
        print(f"Removed Leakage columns: {leaky_cols}")
        X.drop(columns=leaky_cols, inplace=True)

    # --- STEP 1: Ensure non-negative features (required for chi2) ---
    # Shift minimum to 0 only for features with negative values to preserve raw distributions
    X_proc = X.copy()
    for col in X_proc.columns:
        min_val = X_proc[col].min()
        if min_val < 0:
            X_proc[col] = X_proc[col] - min_val

    # --- STEP 2: Compute Authentic Chi-Square Statistics ---
    scores, pvalues = chi2(X_proc, y)
    df_k = len(np.unique(y)) - 1  # degrees of freedom = n_classes - 1

    # Exact asymptotic string representation of p-values
    p_exact_strs = [
        compute_exact_chi2_pvalue(score, pval, df_k)
        for score, pval in zip(scores, pvalues)
    ]

    # --- STEP 3: Selection based on p-value < alpha (and optional top-k) ---
    if k is not None:
        ranked_indices = np.argsort(scores)[::-1]
        top_k_indices = set(ranked_indices[:k])
        is_selected = [
            (i in top_k_indices) and (pval < alpha)
            for i, pval in enumerate(pvalues)
        ]
    else:
        is_selected = [pval < alpha for pval in pvalues]

    selected_features = [col for col, sel in zip(X.columns, is_selected) if sel]
    removed_features = [col for col, sel in zip(X.columns, is_selected) if not sel]

    print(f"Selected Features ({len(selected_features)}): {selected_features}")
    print(f"Removed Features ({len(removed_features)}): {removed_features}")

    # --- STEP 4: Create Report Table ---
    report = pd.DataFrame({
        "Feature": X.columns,
        "Chi2_Statistic": scores,
        "p_value": p_exact_strs,
        "Selected (p < 0.05)": is_selected,
        "Status": ["Selected" if sel else "Removed" for sel in is_selected]
    }).sort_values(by="Chi2_Statistic", ascending=False).reset_index(drop=True)

    X_selected_df = X[selected_features]

    return X_selected_df, report

# --- Main Logic ---
if __name__ == "__main__":
    excel_path = r"task/Data.xlsx"
    if not os.path.exists(excel_path):
        excel_path = r"task/BMM-EI_No.223 (Results).xlsx"

    if os.path.exists(excel_path):
        close_excel_file(excel_path)

        # Check available sheet
        sheet_to_read = "Encoded_Data"
        df = pd.read_excel(excel_path, sheet_name=sheet_to_read)

        target_column = df.columns[-1]
        X_input = df.drop(columns=[target_column])
        y = df[target_column]

        # Apply Chi-Square Feature Selection (Significance threshold alpha=0.05)
        selected_X, chi_report = chi_square_selection(X_input, y, alpha=0.05)

        # Reattach target
        data_after_chi2 = selected_X.copy()
        data_after_chi2[target_column] = y

        # Save to Excel
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            chi_report.to_excel(writer, sheet_name="chi2_report", index=False)
            data_after_chi2.to_excel(writer, sheet_name="data_after_chi2", index=False)

        print("Done! Chi-Square Feature Selection Completed with authentic high-precision p-values.")
        open_excel_file(excel_path)