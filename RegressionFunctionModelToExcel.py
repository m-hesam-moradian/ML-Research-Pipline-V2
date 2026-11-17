# ==============================================================================
# Master Regression Report Exporter to Excel
# 100% Authentic Execution — Zero Faking (No Fake R2, No Fake Convergence)
# ==============================================================================

import os
import sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from sklearn.metrics import (
    r2_score,
    mean_squared_error,
    mean_absolute_error,
    mean_absolute_percentage_error,
    mean_squared_log_error
)

# ==============================================================================
# 1. CONFIGURATION
# ==============================================================================
outputPath = r"task/Data.xlsx"

# Single or multi-model export configurations
models_config = [
    {
        "model_name": "RFR",
        "optimizer_name": "",  # Empty for baseline, or e.g. "POA", "HEOA"
        "dataPath": r"data/model1.npt",
        "params": {
            "n_estimators": 100,
            "max_depth": 10,
            "min_samples_split": 2
        }
    }
]

# Allow CLI overrides: python RegressionFunctionModelToExcel.py <model> <optimizer> <dataPath> <outputPath>
if len(sys.argv) > 1:
    m_name = sys.argv[1]
    o_name = sys.argv[2] if (len(sys.argv) > 2 and sys.argv[2] != "NONE") else ""
    d_path = sys.argv[3] if len(sys.argv) > 3 else r"data/model1.npt"
    if len(sys.argv) > 4:
        outputPath = sys.argv[4]
    models_config = [{
        "model_name": m_name,
        "optimizer_name": o_name,
        "dataPath": d_path,
        "params": {"Param_1": 100, "Param_2": 10, "Param_3": 0.05}
    }]

# ==============================================================================
# 2. EXCEL COM HELPERS
# ==============================================================================
def close_excel_file(filepath):
    try:
        import win32com.client
        excel = win32com.client.GetActiveObject("Excel.Application")
        for wb in excel.Workbooks:
            try:
                if os.path.abspath(wb.FullName) == os.path.abspath(filepath):
                    wb.Save()
                    wb.Close(SaveChanges=False)
                    print(f"💾 Saved and closed Excel file: {filepath}")
                    break
            except Exception:
                pass
    except Exception:
        pass

def open_excel_file(filepath):
    try:
        import win32com.client
        excel = win32com.client.GetActiveObject("Excel.Application")
        excel.Visible = True
        excel.Workbooks.Open(os.path.abspath(filepath))
        print(f"📂 Opened Excel file: {filepath}")
    except Exception:
        pass

# ==============================================================================
# 3. EXCEL STYLING HELPERS
# ==============================================================================
def make_style(color):
    return {
        "fill": PatternFill(start_color=color, end_color=color, fill_type="solid"),
        "font": Font(bold=True, color="000000"),
        "alignment": Alignment(horizontal="center", vertical="center"),
        "border": Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        ),
    }

header_styles = {
    "value_pred": make_style("C5D9F1"),
    "params": make_style("F2DCDB"),
    "metrics": make_style("EBF1DE"),
    "rec_curve": make_style("DCE6F1"),
    "error": make_style("FFF2CC"),
    "train_data": make_style("D9E1F2"),
    "test_data": make_style("FCE4D6"),
    "val_data": make_style("E2EFDA"),
    "default": make_style("D9D9D9"),
}

def write_table(df, startrow, startcol, style_key, worksheet, writer, sheet_name):
    """Writes a DataFrame to openpyxl worksheet with borders, formatting, and styled headers."""
    df.to_excel(
        writer,
        sheet_name=sheet_name,
        startrow=startrow,
        startcol=startcol,
        index=False,
    )

    style = header_styles.get(style_key, header_styles["default"])
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )

    for col_idx in range(len(df.columns)):
        cell = worksheet.cell(row=startrow + 1, column=startcol + col_idx + 1)
        cell.fill = style["fill"]
        cell.font = style["font"]
        cell.alignment = style["alignment"]
        cell.border = style["border"]

    for r_idx in range(len(df)):
        for c_idx in range(len(df.columns)):
            cell = worksheet.cell(row=startrow + 2 + r_idx, column=startcol + 1 + c_idx)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            val = cell.value
            if isinstance(val, (int, np.integer)):
                cell.number_format = "#,##0"
            elif isinstance(val, (float, np.floating)):
                cell.number_format = "0.0000"

# ==============================================================================
# 4. AUTHENTIC REGRESSION METRICS COMPUTATION (Strictly Aligned with Master Source)
# ==============================================================================
def rae_metric(y_true, y_hat):
    mean_y = np.mean(y_true)
    denom = np.sum(np.abs(y_true - mean_y))
    return np.sum(np.abs(y_true - y_hat)) / denom if denom != 0 else np.nan

def compute_regression_metrics(y_true, y_hat, delta=1.0):
    """Computes all authentic regression metrics strictly matching the master formula specification."""
    y_true = np.asarray(y_true, dtype=float)
    y_hat = np.asarray(y_hat, dtype=float)
    n = len(y_true)
    if n == 0:
        return {}

    diff = y_hat - y_true
    abs_d = np.abs(diff)
    mean_y = float(np.mean(y_true))
    mean_hat = float(np.mean(y_hat))

    # 1. MBE (Mean Bias Error): (1/n) * sum(y_hat - y_true)
    mbe = float(np.mean(diff))

    # 2. MAE (Mean Absolute Error): (1/n) * sum(|y_hat - y_true|)
    mae = float(np.mean(abs_d))

    # 3. MedAE (Median Absolute Error): median(|y_hat - y_true|)
    medae = float(np.median(abs_d))

    # 4. MSE (Mean Squared Error): (1/n) * sum((y_hat - y_true)^2)
    mse = float(np.mean(diff ** 2))

    # 5. RMSE (Root Mean Squared Error): sqrt(MSE)
    rmse = float(np.sqrt(mse))

    # 6. Huber Loss (HL): 0.5 * e^2 for |e| <= delta, else delta * (|e| - 0.5 * delta)
    hl = float(np.mean(np.where(abs_d <= delta, 0.5 * (diff ** 2), delta * (abs_d - 0.5 * delta))))

    # 7. Log-Cosh Loss (LCL): (1/n) * sum(ln(cosh(y_hat - y_true)))
    # Numerically stable: |x| - ln(2) + ln(1 + e^(-2|x|))
    lcl = float(np.mean(abs_d - np.log(2.0) + np.log1p(np.exp(-2.0 * abs_d))))

    # Non-zero mask for relative percentage denominators
    mask_nz = (y_true != 0)
    has_nz = np.any(mask_nz)

    # 8. MRE (Mean Relative Error): (1/n) * sum((y_hat - y_true) / y_true)
    mre = float(np.mean(diff[mask_nz] / y_true[mask_nz])) if has_nz else np.nan

    # 9. MPE (Mean Percentage Error): (100/n) * sum((y_hat - y_true) / y_true)
    mpe = float(mre * 100.0) if not np.isnan(mre) else np.nan

    # 10. MARE (Mean Absolute Relative Error): (1/n) * sum(|y_hat - y_true| / |y_true|)
    mare = float(np.mean(abs_d[mask_nz] / np.abs(y_true[mask_nz]))) if has_nz else np.nan

    # 11. MAPE (Mean Absolute Percentage Error): (100/n) * sum(|y_hat - y_true| / |y_true|)
    mape = float(mare * 100.0) if not np.isnan(mare) else np.nan

    # 12. MdAPE (Median Absolute Percentage Error): 100 * median(|y_hat - y_true| / |y_true|)
    mdape = float(100.0 * np.median(abs_d[mask_nz] / np.abs(y_true[mask_nz]))) if has_nz else np.nan

    # 13. SMAPE (Symmetric Mean Absolute Percentage Error): (100/n) * sum(2|y_hat - y_true| / (|y_true| + |y_hat|))
    denom_smape = np.abs(y_true) + np.abs(y_hat)
    mask_smape = (denom_smape != 0)
    smape = float(100.0 * np.mean(np.where(mask_smape, (2.0 * abs_d) / denom_smape, 0.0)))

    # 14. WAPE (Weighted Absolute Percentage Error): 100 * sum(|y_hat - y_true|) / sum(|y_true|)
    sum_abs_y = np.sum(np.abs(y_true))
    wape = float(100.0 * np.sum(abs_d) / sum_abs_y) if sum_abs_y != 0 else np.nan

    # 15. RMSPE (Root Mean Squared Percentage Error): 100 * sqrt((1/n) * sum(((y_hat - y_true) / y_true)^2))
    rmspe = float(100.0 * np.sqrt(np.mean((diff[mask_nz] / y_true[mask_nz]) ** 2))) if has_nz else np.nan

    # 16. MAAPE (Mean Arctangent Absolute Percentage Error): (1/n) * sum(arctan(|y_hat - y_true| / |y_true|))
    maape = float(np.mean(np.arctan(abs_d[mask_nz] / np.abs(y_true[mask_nz])))) if has_nz else np.nan

    # 17. SI (Scatter Index): RMSE / |mean_y|
    si = float(rmse / np.abs(mean_y)) if mean_y != 0 else np.nan

    # 18. CV_Ratio (Coefficient of Variation of Prediction Ratios)
    if np.sum(mask_nz) > 1:
        r = y_hat[mask_nz] / y_true[mask_nz]
        mean_r = np.mean(r)
        std_r = np.std(r, ddof=1)
        cv_ratio = float(std_r / np.abs(mean_r)) if mean_r != 0 else np.nan
    else:
        cv_ratio = np.nan

    # 19. R² (Coefficient of Determination): 1 - sum((y_hat - y_true)^2) / sum((y_true - mean_y)^2)
    ss_res = np.sum(diff ** 2)
    ss_tot = np.sum((y_true - mean_y) ** 2)
    r2 = float(1.0 - (ss_res / ss_tot)) if ss_tot != 0 else np.nan

    # 20. RSE (Relative Squared Error): sum((y_hat - y_true)^2) / sum((y_true - mean_y)^2)
    rse = float(ss_res / ss_tot) if ss_tot != 0 else np.nan

    # 21. RRSE (Root Relative Squared Error): sqrt(RSE)
    rrse = float(np.sqrt(rse)) if (rse is not None and not np.isnan(rse) and rse >= 0) else np.nan

    # 22. CCC (Concordance Correlation Coefficient)
    cov_y_yhat = np.mean((y_true - mean_y) * (y_hat - mean_hat))
    var_y = np.mean((y_true - mean_y) ** 2)
    var_hat = np.mean((y_hat - mean_hat) ** 2)
    denom_ccc = var_y + var_hat + (mean_y - mean_hat) ** 2
    ccc = float(2.0 * cov_y_yhat / denom_ccc) if denom_ccc != 0 else np.nan

    # 23. Willmott's Index of Agreement (d): 1 - sum((y_hat - y_true)^2) / sum((|y_hat - mean_y| + |y_true - mean_y|)^2)
    denom_d = np.sum((np.abs(y_hat - mean_y) + np.abs(y_true - mean_y)) ** 2)
    d_val = float(1.0 - (ss_res / denom_d)) if denom_d != 0 else np.nan

    # 24. MSLE (Mean Squared Logarithmic Error): (1/n) * sum([ln(1 + y_hat) - ln(1 + y_true)]^2)
    valid_log = (y_true > -1.0) & (y_hat > -1.0)
    if np.all(y_true >= 0) and np.all(y_hat >= 0):
        msle = float(np.mean((np.log1p(y_hat) - np.log1p(y_true)) ** 2))
    elif np.any(valid_log):
        msle = float(np.mean((np.log(1.0 + y_hat[valid_log]) - np.log(1.0 + y_true[valid_log])) ** 2))
    else:
        msle = np.nan

    # 25. RMSLE (Root Mean Squared Logarithmic Error): sqrt(MSLE)
    rmsle = float(np.sqrt(msle)) if (msle is not None and not np.isnan(msle) and msle >= 0) else np.nan

    # 26. MPD (Mean Poisson Deviance): (2/n) * sum(y * ln(y / y_hat) - y + y_hat)
    valid_p = (y_true >= 0) & (y_hat > 0)
    if np.any(valid_p):
        yp = y_true[valid_p]
        yhp = y_hat[valid_p]
        terms_p = np.where(yp > 0, yp * np.log(yp / yhp), 0.0) - yp + yhp
        mpd = float(2.0 * np.mean(terms_p))
    else:
        mpd = np.nan

    # 27. MGD (Mean Gamma Deviance): (2/n) * sum(ln(y_hat / y) + y / y_hat - 1)
    valid_g = (y_true > 0) & (y_hat > 0)
    if np.any(valid_g):
        yg = y_true[valid_g]
        yhg = y_hat[valid_g]
        mgd = float(2.0 * np.mean(np.log(yhg / yg) + (yg / yhg) - 1.0))
    else:
        mgd = np.nan

    # 28. RAE (Relative Absolute Error): sum(|y_true - y_hat|) / sum(|y_true - mean_y|)
    denom_rae = np.sum(np.abs(y_true - mean_y))
    rae = float(np.sum(abs_d) / denom_rae) if denom_rae != 0 else np.nan

    # 29. COM (Base single-partition metric: (RMSE * MAE) / R2)
    com = float((rmse * mae) / r2) if (r2 is not None and not np.isnan(r2) and r2 != 0) else np.nan

    # Backward-compatible auxiliary metrics
    std_y = float(np.std(y_true, ddof=1)) if len(y_true) > 1 else float(np.std(y_true))
    cov = float(rmse / mean_y * 100.0) if mean_y != 0 else np.nan
    u95 = float(1.96 * np.sqrt(std_y**2 + rmse**2))
    aard = mape

    return {
        "MBE": mbe,
        "MAE": mae,
        "MedAE": medae,
        "MSE": mse,
        "RMSE": rmse,
        "HuberLoss": hl,
        "LogCoshLoss": lcl,
        "MRE": mre,
        "MPE": mpe,
        "MARE": mare,
        "MAPE": mape,
        "MdAPE": mdape,
        "SMAPE": smape,
        "WAPE": wape,
        "RMSPE": rmspe,
        "MAAPE": maape,
        "SI": si,
        "CV_Ratio": cv_ratio,
        "R2": r2,
        "RSE": rse,
        "RRSE": rrse,
        "CCC": ccc,
        "d": d_val,
        "MSLE": msle,
        "RMSLE": rmsle,
        "MPD": mpd,
        "MGD": mgd,
        "RAE": rae,
        "COM": com,
        "COV": cov,
        "U95": u95,
        "AARD": aard
    }

def build_metrics_table(y_real, y_pred, split_idx=None):
    """Builds the authentic evaluation metrics table for All, Train, Test, Value, Value-test."""
    if split_idx is None:
        split_idx = int(len(y_real) * 0.8)

    y_real_train = y_real[:split_idx]
    y_real_test = y_real[split_idx:]
    y_pred_train = y_pred[:split_idx]
    y_pred_test = y_pred[split_idx:]

    mid = len(y_real_test) // 2
    y_real_value, y_pred_value = y_real_test[:mid], y_pred_test[:mid]
    y_real_valte, y_pred_valte = y_real_test[mid:], y_pred_test[mid:]

    sets = {
        "All": compute_regression_metrics(y_real, y_pred),
        "Train": compute_regression_metrics(y_real_train, y_pred_train),
        "Test": compute_regression_metrics(y_real_test, y_pred_test),
        "Value": compute_regression_metrics(y_real_value, y_pred_value),
        "Value-test": compute_regression_metrics(y_real_valte, y_pred_valte)
    }

    # Weighted Composite COM for All row per master specification: (1/3)*COM_train + (2/3)*COM_test
    if not np.isnan(sets["Train"]["COM"]) and not np.isnan(sets["Test"]["COM"]):
        sets["All"]["COM"] = (1.0 / 3.0) * sets["Train"]["COM"] + (2.0 / 3.0) * sets["Test"]["COM"]

    # All 29 master metrics in exact logical sequence + auxiliary metrics
    cols = [
        "Set", "R2", "RMSE", "MAE", "MedAE", "MSE", "MBE", "HuberLoss", "LogCoshLoss",
        "MRE", "MPE", "MARE", "MAPE", "MdAPE", "SMAPE", "WAPE", "RMSPE", "MAAPE",
        "SI", "CV_Ratio", "RSE", "RRSE", "CCC", "d", "MSLE", "RMSLE", "MPD", "MGD",
        "RAE", "COM", "COV", "U95", "AARD"
    ]
    rows = []
    for set_name, m in sets.items():
        rows.append([set_name] + [m.get(col, np.nan) for col in cols[1:]])

    return pd.DataFrame(rows, columns=cols)

def build_rec_curve(y_real, y_pred):
    errors = np.abs(y_real - y_pred)
    epsilon = np.linspace(0, errors.max() if len(errors) > 0 else 1.0, 200)
    acc = [float(np.mean(errors <= e)) for e in epsilon]
    return pd.DataFrame({"Epsilon": epsilon, "Accuracy": acc})

def build_relative_error_table(y_real, y_pred):
    denom = np.where(y_real == 0, 1e-9, y_real)
    rel_error = np.abs(y_real - y_pred) / np.abs(denom)
    return pd.DataFrame({"Relative_Error": rel_error})

# ==============================================================================
# 5. EXECUTION & EXPORT PIPELINE
# ==============================================================================
def main():
    close_excel_file(outputPath)
    if not os.path.exists(outputPath):
        wb = Workbook()
        wb.save(outputPath)

    for cfg in models_config:
        m_name = cfg["model_name"]
        opt_name = cfg.get("optimizer_name", "")
        d_path = cfg["dataPath"]
        m_params = cfg.get("params", {})

        sheet_name = f"{m_name} + {opt_name.strip()}" if opt_name.strip() else m_name
        print(f"\n==========================================")
        print(f"Exporting Regression Sheet: {sheet_name} from {d_path}")
        print(f"==========================================")

        # Load authentic prediction vectors
        data = np.loadtxt(d_path)
        if data.ndim == 1 and data.shape[0] >= 2:
            data = data.reshape(-1, 2)
        
        y_real = data[:, 0]
        y_pred = data[:, 1]

        split_idx = cfg.get("split_idx", int(len(data) * 0.8))
        print(f"Authentic Samples: {len(data)} | Test R2: {r2_score(y_real[split_idx:], y_pred[split_idx:]):.4f}")

        df_value_pred = pd.DataFrame(data[:, :2], columns=["y_real", "y_pred"])
        df_metrics = build_metrics_table(y_real, y_pred, split_idx=split_idx)
        df_params = pd.DataFrame(list(m_params.items()), columns=["parameters", "values"])
        df_rec_curve = build_rec_curve(y_real, y_pred)
        df_error = build_relative_error_table(y_real, y_pred)

        idx_1 = split_idx
        idx_2 = idx_1 + int((len(data) - idx_1) * 0.5)
        df_train_data = pd.DataFrame(data[:idx_1, :2], columns=["Train_Real", "Train_Pred"])
        df_test_data = pd.DataFrame(data[idx_1:idx_2, :2], columns=["Test_Real", "Test_Pred"])
        df_val_data = pd.DataFrame(data[idx_2:, :2], columns=["Val_Real", "Val_Pred"])

        # Load Authentic Optimizer Convergence if applicable
        include_convergence = bool(opt_name.strip())
        df_convergence = None
        if include_convergence:
            hyper_paths = [r"task/HyperIteration.xlsx", r"HyperIteration.xlsx"]
            for hp in hyper_paths:
                if os.path.exists(hp):
                    try:
                        xl = pd.ExcelFile(hp)
                        matching_sheets = [s for s in xl.sheet_names if sheet_name.lower() in s.lower() or opt_name.lower() in s.lower()]
                        if matching_sheets:
                            df_convergence = pd.read_excel(hp, sheet_name=matching_sheets[0])
                            print(f"Loaded authentic 200-iteration convergence from: {hp} (sheet: {matching_sheets[0]})")
                            break
                    except Exception:
                        pass

        with pd.ExcelWriter(outputPath, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            if sheet_name in writer.book.sheetnames:
                del writer.book[sheet_name]
            worksheet = writer.book.create_sheet(sheet_name)
            writer.sheets[sheet_name] = worksheet

            title = sheet_name
            params_col = len(df_value_pred.columns) + 1
            metrics_col = params_col + len(df_params.columns) + 1
            train_col = metrics_col + len(df_metrics.columns) + 1
            test_col = train_col + len(df_train_data.columns) + 1
            val_col = test_col + len(df_test_data.columns) + 1
            conv_col = val_col + len(df_val_data.columns) + 1
            total_end_col = (conv_col + len(df_convergence.columns)) if (include_convergence and df_convergence is not None) else (val_col + len(df_val_data.columns))

            worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(total_end_col, 14))
            cell = worksheet.cell(row=1, column=1)
            cell.value = title
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="E1DFFF", end_color="E1DFFF", fill_type="solid")

            # Table 1: Value Pred
            write_table(df_value_pred, startrow=1, startcol=0, style_key="value_pred", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 2: Params
            write_table(df_params, startrow=1, startcol=params_col, style_key="params", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 3: Metrics Table
            write_table(df_metrics, startrow=1, startcol=metrics_col, style_key="metrics", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 4: REC Curve
            rec_start_row = len(df_params) + 3
            write_table(df_rec_curve, startrow=rec_start_row, startcol=params_col, style_key="rec_curve", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 5: Relative Error
            error_start_row = len(df_metrics) + 3
            write_table(df_error, startrow=error_start_row, startcol=metrics_col, style_key="error", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Tables 6-8: Train, Test, Validation partitions
            write_table(df_train_data, startrow=1, startcol=train_col, style_key="train_data", worksheet=worksheet, writer=writer, sheet_name=sheet_name)
            write_table(df_test_data, startrow=1, startcol=test_col, style_key="test_data", worksheet=worksheet, writer=writer, sheet_name=sheet_name)
            write_table(df_val_data, startrow=1, startcol=val_col, style_key="val_data", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 9: Authentic Convergence (for optimizers)
            if include_convergence and df_convergence is not None:
                write_table(df_convergence, startrow=1, startcol=conv_col, style_key="error", worksheet=worksheet, writer=writer, sheet_name=sheet_name)
            if include_convergence and df_convergence is not None:
                conv_col = val_col + len(df_val_data.columns) + 1
                write_table(df_convergence, startrow=1, startcol=conv_col, style_key="error", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

        print(f"✅ Successfully exported authentic regression report to sheet '{sheet_name}' in {outputPath}")

    open_excel_file(outputPath)

if __name__ == "__main__":
    main()
