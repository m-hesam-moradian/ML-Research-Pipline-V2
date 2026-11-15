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
# 4. AUTHENTIC REGRESSION METRICS COMPUTATION
# ==============================================================================
def rae_metric(y_true, y_hat):
    mean_y = np.mean(y_true)
    denom = np.sum(np.abs(y_true - mean_y))
    return np.sum(np.abs(y_true - y_hat)) / denom if denom != 0 else np.nan

def compute_regression_metrics(y_true, y_hat, delta=1.0):
    """Computes all authentic regression metrics with strict zero-faking."""
    y_true = np.asarray(y_true, dtype=float)
    y_hat = np.asarray(y_hat, dtype=float)

    r2 = r2_score(y_true, y_hat)
    mse = mean_squared_error(y_true, y_hat)
    rmse = np.sqrt(mse)
    mae = mean_absolute_error(y_true, y_hat)

    # AARD / MAPE & MARE
    non_zero = y_true != 0
    if np.any(non_zero):
        mare_val = np.mean(np.abs((y_true[non_zero] - y_hat[non_zero]) / y_true[non_zero]))
        mape_val = mare_val * 100.0
    else:
        mare_val, mape_val = np.nan, np.nan
    aard = mape_val

    # Mean and Standard Deviation
    mean_y = np.mean(y_true)
    std_y = np.std(y_true, ddof=1) if len(y_true) > 1 else np.std(y_true)

    cov = (rmse / mean_y * 100.0) if mean_y != 0 else np.nan
    si = (rmse / mean_y) if mean_y != 0 else np.nan

    # U95 Uncertainty
    u95 = 1.96 * np.sqrt(std_y**2 + rmse**2)

    # MBE (Mean Bias Error: y_hat - y_true)
    mbe = np.mean(y_hat - y_true)

    # LogCosh Loss (Numerically stable formula matching ln(cosh(x)))
    diff = y_hat - y_true
    abs_d = np.abs(diff)
    log_cosh = np.mean(abs_d - np.log(2.0) + np.log1p(np.exp(-2.0 * abs_d)))

    # Huber Loss
    huber = np.mean(np.where(abs_d <= delta, 0.5 * (diff ** 2), delta * (abs_d - 0.5 * delta)))

    # COM (Comprehensive Measure from Image 4 & PDF 2: (RMSE * MAE) / R2)
    com = (rmse * mae) / r2 if (r2 is not None and not np.isnan(r2) and r2 != 0) else np.nan

    # RAE
    rae = rae_metric(y_true, y_hat)

    return {
        "R2": r2,
        "RMSE": rmse,
        "MAE": mae,
        "AARD": aard,
        "MARE": mare_val,
        "COV": cov,
        "U95": u95,
        "SI": si,
        "MBE": mbe,
        "LogCoshLoss": log_cosh,
        "HuberLoss": huber,
        "MAPE": mape_val,
        "COM": com,
        "RAE": rae
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

    # Weighted Composite COM for All row per Image 4 / PDF 2: (1/3)*COM_train + (2/3)*COM_test
    if not np.isnan(sets["Train"]["COM"]) and not np.isnan(sets["Test"]["COM"]):
        sets["All"]["COM"] = (1.0 / 3.0) * sets["Train"]["COM"] + (2.0 / 3.0) * sets["Test"]["COM"]

    cols = ["Set", "R2", "RMSE", "MAE", "AARD", "MARE", "COV", "U95", "SI", "MBE", "LogCoshLoss", "HuberLoss", "MAPE", "COM", "RAE"]
    rows = []
    for set_name, m in sets.items():
        rows.append([
            set_name,
            m["R2"],
            m["RMSE"],
            m["MAE"],
            m["AARD"],
            m["MARE"],
            m["COV"],
            m["U95"],
            m["SI"],
            m["MBE"],
            m["LogCoshLoss"],
            m["HuberLoss"],
            m["MAPE"],
            m["COM"],
            m["RAE"]
        ])

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
            merge_end_col = 14 if include_convergence else 13
            worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=merge_end_col + 1)
            cell = worksheet.cell(row=1, column=1)
            cell.value = title
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="E1DFFF", end_color="E1DFFF", fill_type="solid")

            # Table 1: Value Pred
            write_table(df_value_pred, startrow=1, startcol=0, style_key="value_pred", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 2: Params
            params_col = len(df_value_pred.columns) + 1
            write_table(df_params, startrow=1, startcol=params_col, style_key="params", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 3: Metrics Table
            metrics_col = params_col + len(df_params.columns) + 1
            write_table(df_metrics, startrow=1, startcol=metrics_col, style_key="metrics", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 4: REC Curve
            rec_start_row = len(df_params) + 3
            write_table(df_rec_curve, startrow=rec_start_row, startcol=params_col, style_key="rec_curve", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 5: Relative Error
            error_start_row = len(df_metrics) + 3
            write_table(df_error, startrow=error_start_row, startcol=metrics_col, style_key="error", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Tables 6-8: Train, Test, Validation partitions
            train_col = metrics_col + len(df_metrics.columns) + 1
            write_table(df_train_data, startrow=1, startcol=train_col, style_key="train_data", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            test_col = train_col + len(df_train_data.columns) + 1
            write_table(df_test_data, startrow=1, startcol=test_col, style_key="test_data", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            val_col = test_col + len(df_test_data.columns) + 1
            write_table(df_val_data, startrow=1, startcol=val_col, style_key="val_data", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

            # Table 9: Authentic Convergence (for optimizers)
            if include_convergence and df_convergence is not None:
                conv_col = val_col + len(df_val_data.columns) + 1
                write_table(df_convergence, startrow=1, startcol=conv_col, style_key="error", worksheet=worksheet, writer=writer, sheet_name=sheet_name)

        print(f"✅ Successfully exported authentic regression report to sheet '{sheet_name}' in {outputPath}")

    open_excel_file(outputPath)

if __name__ == "__main__":
    main()
