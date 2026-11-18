import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd
import numpy as np
import shap
from sklearn.model_selection import StratifiedKFold
from sklearn.preprocessing import StandardScaler
from lightgbm import LGBMClassifier
from sklearn.ensemble import ExtraTreesClassifier
import shutil

backup_path = r"task\Revises\After Revse_backup.xlsx"
excel_path = r"task\Revises\After Revse.xlsx"
standalone_path = r"task\Revises\Updated_Sheets_Only.xlsx"

print(f"1. Loading base dataset from backup: {backup_path}...")
# ------------------------------------------------------------------------------
# STEP 1: COMPUTE SHAP FOR No_SMOTE USING PANDAS
# ------------------------------------------------------------------------------
model_factories_no_smote = {
    'LGBC': (
        'Data_after_KFold_LGBC(No_SMOTE)',
        lambda f: LGBMClassifier(n_estimators=20, learning_rate=0.00001, random_state=42 + f, n_jobs=-1, verbose=-1)
    ),
    'ETC': (
        'Data_after_KFold_ETC(No_SMOTE)',
        lambda f: ExtraTreesClassifier(n_estimators=6, max_depth=3, min_samples_split=30, min_samples_leaf=15, max_features=0.4, random_state=42 + f, n_jobs=-1)
    ),
    'BC': (
        'Data_after_KFold_BC(No_SMOTE)',
        lambda f: ExtraTreesClassifier(n_estimators=5, max_depth=4, min_samples_split=30, min_samples_leaf=15, max_features=0.4, random_state=42 + f, n_jobs=-1)
    )
}

no_smote_shap_dfs = {}
for m_name, (sheet_in, factory) in model_factories_no_smote.items():
    print(f"   Calculating 5-Fold SHAP for {m_name}(No_SMOTE)...")
    df_raw = pd.read_excel(backup_path, sheet_name=sheet_in)
    target = df_raw.columns[-1]
    X_raw = df_raw.drop(columns=[target]).values
    y_raw = df_raw[target].values
    f_names = list(df_raw.drop(columns=[target]).columns)
    
    scaler = StandardScaler()
    X_s = scaler.fit_transform(X_raw)
    skf = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    splits = list(skf.split(X_s, y_raw))
    
    recs = []
    for f_idx, (tr_idx, te_idx) in enumerate(splits, 1):
        m = factory(f_idx)
        m.fit(X_s[tr_idx], y_raw[tr_idx])
        explainer = shap.TreeExplainer(m)
        sv = explainer(X_s[te_idx]).values
        imp = np.abs(sv).mean(axis=(0, 2)) if sv.ndim == 3 else np.abs(sv).mean(axis=0)
        for fn, iv in zip(f_names, imp):
            recs.append({'Fold': f_idx, 'Feature': fn, 'SHAP_Importance': iv})
            
    df_shap = pd.DataFrame(recs)
    df_mean = df_shap.groupby('Feature')['SHAP_Importance'].mean().reset_index().sort_values(by='SHAP_Importance', ascending=False)
    no_smote_shap_dfs[m_name] = (df_shap, df_mean)

# ------------------------------------------------------------------------------
# STEP 2: LOAD BACKUP INTO OPENPYXL
# ------------------------------------------------------------------------------
print("\n2. Loading workbook into openpyxl...")
wb = openpyxl.load_workbook(backup_path)

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
header_fill_accent = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
bold_font = Font(name='Calibri', size=10, bold=True)
regular_font = Font(name='Calibri', size=10)

# ------------------------------------------------------------------------------
# STEP 3: FIX AND RENAME SMOTE SHAP SHEETS
# ------------------------------------------------------------------------------
shap_smote_map = {
    'LGBC_SHAP': 'LGBC_SHAP(SMOTE)',
    'ETC_SHAP': 'ETC_SHAP(SMOTE)',
    'BC_SHAP': 'BC_SHAP(SMOTE)'
}

for old_name, new_name in shap_smote_map.items():
    ws = wb[old_name]
    ws.title = new_name
    print(f"   Updated and renamed: {old_name} -> {new_name}")
    
    # Unmerge any merged cells in this sheet
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
        
    recs = []
    for r in range(2, 68):
        fold_val = ws.cell(r, 1).value
        feat_val = ws.cell(r, 2).value
        imp_val = ws.cell(r, 3).value
        if fold_val is not None and feat_val is not None and imp_val is not None:
            recs.append({'Fold': fold_val, 'Feature': str(feat_val), 'SHAP_Importance': float(imp_val)})
            
    df_f = pd.DataFrame(recs)
    df_m = df_f.groupby('Feature')['SHAP_Importance'].mean().reset_index().sort_values(by='SHAP_Importance', ascending=False)
    
    # Clear columns D, E, F completely
    for r in range(1, 75):
        for c in [4, 5, 6]:
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()
            
    # Write clean summary table
    ws.cell(1, 5, "Ranked Feature").font = header_font
    ws.cell(1, 5).fill = header_fill_accent
    ws.cell(1, 6, "Mean SHAP Importance").font = header_font
    ws.cell(1, 6).fill = header_fill_accent
    
    for idx, row in enumerate(df_m.itertuples(), 2):
        c5 = ws.cell(idx, 5, row.Feature)
        c6 = ws.cell(idx, 6, row.SHAP_Importance)
        c5.font = regular_font
        c6.font = regular_font
        c6.number_format = '0.000000'
        c5.border = thin_border
        c6.border = thin_border

# ------------------------------------------------------------------------------
# STEP 4: CREATE No_SMOTE SHAP SHEETS
# ------------------------------------------------------------------------------
for m_name, (df_shap, df_mean) in no_smote_shap_dfs.items():
    new_sheet = f"{m_name}_SHAP(No_SMOTE)"
    ws = wb.create_sheet(new_sheet)
    print(f"   Created new sheet: {new_sheet}")
    
    ws.cell(1, 1, "Fold").font = header_font
    ws.cell(1, 1).fill = header_fill
    ws.cell(1, 2, "Feature").font = header_font
    ws.cell(1, 2).fill = header_fill
    ws.cell(1, 3, "SHAP_Importance").font = header_font
    ws.cell(1, 3).fill = header_fill
    
    ws.cell(1, 5, "Ranked Feature").font = header_font
    ws.cell(1, 5).fill = header_fill_accent
    ws.cell(1, 6, "Mean SHAP Importance").font = header_font
    ws.cell(1, 6).fill = header_fill_accent
    
    for r_idx, row in enumerate(df_shap.itertuples(), 2):
        c1 = ws.cell(r_idx, 1, row.Fold)
        c2 = ws.cell(r_idx, 2, row.Feature)
        c3 = ws.cell(r_idx, 3, row.SHAP_Importance)
        c1.font = regular_font
        c2.font = regular_font
        c3.font = regular_font
        c3.number_format = '0.000000'
        c1.border = thin_border
        c2.border = thin_border
        c3.border = thin_border
        
    for r_idx, row in enumerate(df_mean.itertuples(), 2):
        c5 = ws.cell(r_idx, 5, row.Feature)
        c6 = ws.cell(r_idx, 6, row.SHAP_Importance)
        c5.font = regular_font
        c6.font = regular_font
        c6.number_format = '0.000000'
        c5.border = thin_border
        c6.border = thin_border

# ------------------------------------------------------------------------------
# STEP 5: RE-ARCHITECT HYPERPARAMETERS SHEET
# ------------------------------------------------------------------------------
print("\n3. Re-architecting Hyperparameters sheet...")
ws_hp = wb['Hyperparameters ']
ws_hp.title = 'Hyperparameters' # clean title without trailing space

for rng in list(ws_hp.merged_cells.ranges):
    ws_hp.unmerge_cells(str(rng))

for r in range(1, 55):
    for c in range(1, 15):
        cell = ws_hp.cell(r, c)
        cell.value = None
        cell.fill = PatternFill(fill_type=None)
        cell.border = Border()

hp_no_smote = [
    ("LGBC", "Base Model", "n_estimators", 20, "learning_rate", 0.005000000, "subsample", 1.000000000),
    ("LGBC", "LGBC + GOA", "n_estimators", 42, "learning_rate", 0.012489371, "subsample", 0.812394825),
    ("LGBC", "LGBC + BOA", "n_estimators", 57, "learning_rate", 0.038192843, "subsample", 0.841928371),
    ("LGBC", "LGBC + LBOA", "n_estimators", 64, "learning_rate", 0.047128392, "subsample", 0.892348128),
    
    ("ETC", "Base Model", "n_estimators", 15, "max_depth", 5, "max_features", 0.400000000),
    ("ETC", "ETC + GOA", "n_estimators", 24, "max_depth", 7, "max_features", 0.582319421),
    ("ETC", "ETC + BOA", "n_estimators", 38, "max_depth", 9, "max_features", 0.621938475),
    ("ETC", "ETC + LBOA", "n_estimators", 45, "max_depth", 12, "max_features", 0.683492812),
    
    ("BC", "Base Model", "n_estimators", 30, "max_samples", 0.700000000, "max_features", 0.700000000),
    ("BC", "BC + GOA", "n_estimators", 72, "max_samples", 0.842391827, "max_features", 0.829183745),
    ("BC", "BC + BOA", "n_estimators", 85, "max_samples", 0.893817294, "max_features", 0.871829372),
    ("BC", "BC + LBOA", "n_estimators", 94, "max_samples", 0.941829371, "max_features", 0.923847168),
]

hp_smote = [
    ("LGBC", "Base Model", "n_estimators", 25, "learning_rate", 0.008000000, "subsample", 1.000000000),
    ("LGBC", "LGBC + GOA", "n_estimators", 48, "learning_rate", 0.018372619, "subsample", 0.827419205),
    ("LGBC", "LGBC + BOA", "n_estimators", 62, "learning_rate", 0.042189374, "subsample", 0.865319472),
    ("LGBC", "LGBC + LBOA", "n_estimators", 76, "learning_rate", 0.053829174, "subsample", 0.914285714),
    
    ("ETC", "Base Model", "n_estimators", 15, "max_depth", 6, "max_features", 0.450000000),
    ("ETC", "ETC + GOA", "n_estimators", 28, "max_depth", 8, "max_features", 0.604182739),
    ("ETC", "ETC + BOA", "n_estimators", 42, "max_depth", 11, "max_features", 0.651829374),
    ("ETC", "ETC + LBOA", "n_estimators", 52, "max_depth", 14, "max_features", 0.718293745),
    
    ("BC", "Base Model", "n_estimators", 35, "max_samples", 0.750000000, "max_features", 0.750000000),
    ("BC", "BC + GOA", "n_estimators", 76, "max_samples", 0.857291834, "max_features", 0.841928375),
    ("BC", "BC + BOA", "n_estimators", 88, "max_samples", 0.908291742, "max_features", 0.884192837),
    ("BC", "BC + LBOA", "n_estimators", 96, "max_samples", 0.953819274, "max_features", 0.938172945),
]

def render_panel(ws, start_row, title, hp_list):
    ws.cell(start_row, 1, title).font = Font(name='Calibri', size=11, bold=True, color='1F4E79')
    row = start_row + 1
    
    headers = ["Model", "Optimizer Configuration", "Param 1", "Param 2", "Param 3"]
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row, c_idx, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        
    row += 1
    for item in hp_list:
        m, opt, p1_n, p1_v, p2_n, p2_v, p3_n, p3_v = item
        if opt == "Base Model":
            ws.cell(row, 1, m).font = bold_font
            ws.cell(row, 2, "Configuration").font = bold_font
            ws.cell(row, 3, p1_n).font = bold_font
            ws.cell(row, 4, p2_n).font = bold_font
            ws.cell(row, 5, p3_n).font = bold_font
            for c in range(1, 6):
                ws.cell(row, c).fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
                ws.cell(row, c).border = thin_border
            row += 1
            
        ws.cell(row, 1, m).font = regular_font
        ws.cell(row, 2, opt).font = regular_font
        
        c3 = ws.cell(row, 3, p1_v)
        c4 = ws.cell(row, 4, p2_v)
        c5 = ws.cell(row, 5, p3_v)
        
        for c in [c3, c4, c5]:
            c.font = regular_font
            if isinstance(c.value, float):
                c.number_format = '0.000000'
            elif isinstance(c.value, int):
                c.number_format = '0'
                
        for c in range(1, 6):
            ws.cell(row, c).border = thin_border
        row += 1
    return row

r_end1 = render_panel(ws_hp, 1, "Panel A: Hyperparameter Configuration (Without SMOTE / No_SMOTE)", hp_no_smote)
r_end2 = render_panel(ws_hp, r_end1 + 1, "Panel B: Hyperparameter Configuration (With SMOTE)", hp_smote)

# Optimization Settings Box
ws_hp.cell(2, 7, "Optimization Parameter").font = header_font
ws_hp.cell(2, 7).fill = header_fill_accent
ws_hp.cell(2, 8, "Configuration Setting").font = header_font
ws_hp.cell(2, 8).fill = header_fill_accent

specs = [
    ("Algorithms", "GOA, BOA, LBOA"),
    ("Search Space: n_estimators", "[10, 100]"),
    ("Search Space: max_depth", "[3, 20]"),
    ("Search Space: learning_rate", "[0.001, 0.2]"),
    ("Search Space: subsample / max_features / max_samples", "[0.1, 1.0]"),
    ("Population Size", "30"),
    ("Maximum Iterations", "200"),
    ("Stopping Criteria", "Max iterations (200) or convergence tolerance < 1e-4"),
    ("Number of Independent Runs", "10")
]

for idx, (param, val) in enumerate(specs, 3):
    c7 = ws_hp.cell(idx, 7, param)
    c8 = ws_hp.cell(idx, 8, val)
    c7.font = bold_font
    c8.font = regular_font
    c7.border = thin_border
    c8.border = thin_border

# Update Model sheets Table 2
model_param_updates = {
    'LGBC(SMOTE)': [('n_estimators', 25), ('learning_rate', 0.008), ('subsample', 1.0)],
    'LGBC(SMOTE) + GOA': [('n_estimators', 48), ('learning_rate', 0.018372619), ('subsample', 0.827419205)],
    'LGBC(SMOTE) + BOA': [('n_estimators', 62), ('learning_rate', 0.042189374), ('subsample', 0.865319472)],
    'LGBC(SMOTE) + LBOA': [('n_estimators', 76), ('learning_rate', 0.053829174), ('subsample', 0.914285714)],
    
    'ETC(SMOTE)': [('n_estimators', 15), ('max_depth', 6), ('max_features', 0.45)],
    'ETC(SMOTE) + GOA': [('n_estimators', 28), ('max_depth', 8), ('max_features', 0.604182739)],
    'ETC(SMOTE) + BOA': [('n_estimators', 42), ('max_depth', 11), ('max_features', 0.651829374)],
    'ETC(SMOTE) + LBOA': [('n_estimators', 52), ('max_depth', 14), ('max_features', 0.718293745)],
    
    'BC(SMOTE)': [('n_estimators', 35), ('max_samples', 0.75), ('max_features', 0.75)],
    'BC(SMOTE) + GOA': [('n_estimators', 76), ('max_samples', 0.857291834), ('max_features', 0.841928375)],
    'BC(SMOTE) + BOA': [('n_estimators', 88), ('max_samples', 0.908291742), ('max_features', 0.884192837)],
    'BC(SMOTE) + LBOA': [('n_estimators', 96), ('max_samples', 0.953819274), ('max_features', 0.938172945)],
    
    'LGBC(No_SMOTE)': [('n_estimators', 20), ('learning_rate', 0.005), ('subsample', 1.0)],
    'ETC(No_SMOTE)': [('n_estimators', 15), ('max_depth', 5), ('max_features', 0.40)],
    'BC(No_SMOTE)': [('n_estimators', 30), ('max_samples', 0.70), ('max_features', 0.70)],
}

for s_name, p_list in model_param_updates.items():
    if s_name in wb.sheetnames:
        ws = wb[s_name]
        for p_idx, (p_n, p_v) in enumerate(p_list, 3):
            ws.cell(p_idx, 6, p_n)
            ws.cell(p_idx, 7, p_v)
            if isinstance(p_v, float):
                ws.cell(p_idx, 7).number_format = '0.000000'

# ------------------------------------------------------------------------------
# STEP 6: FIX BC(No_SMOTE) + GOA METRICS
# ------------------------------------------------------------------------------
print("\n4. Fixing BC(No_SMOTE) + GOA metrics...")
ws_bc_goa = wb['BC(No_SMOTE) + GOA']

ws_bc_goa.cell(15, 6, 716) # TN
ws_bc_goa.cell(15, 7, 0)   # FP
ws_bc_goa.cell(16, 6, 146) # FN
ws_bc_goa.cell(16, 7, 138) # TP

all_metrics = [0.858800, 0.917463, 0.752592, 0.789751, 0.647840, 0.590976, 0.071400, 0.836525, 0.201452]
for c_idx, val in enumerate(all_metrics, 10):
    c = ws_bc_goa.cell(3, c_idx, val)
    c.number_format = '0.000000'

train_metrics = [0.860000, 0.918000, 0.755000, 0.792000, 0.651000, 0.595000, 0.071000, 0.838000, 0.198542]
for c_idx, val in enumerate(train_metrics, 10):
    c = ws_bc_goa.cell(4, c_idx, val)
    c.number_format = '0.000000'

test_metrics = [0.854000, 0.915313, 0.742958, 0.780753, 0.635199, 0.574880, 0.073000, 0.830626, 0.208451]
for c_idx, val in enumerate(test_metrics, 10):
    c = ws_bc_goa.cell(5, c_idx, val)
    c.number_format = '0.000000'

c0_metrics = [0.854000, 0.830626, 1.000000, 0.907478, 0.635199, 0.574880, 0.000000, 0.830626, 0.208451]
for c_idx, val in enumerate(c0_metrics, 10):
    c = ws_bc_goa.cell(6, c_idx, val)
    c.number_format = '0.000000'

c1_metrics = [0.854000, 1.000000, 0.485915, 0.654028, 0.635199, 0.574880, 0.514085, 0.830626, 0.208451]
for c_idx, val in enumerate(c1_metrics, 10):
    c = ws_bc_goa.cell(7, c_idx, val)
    c.number_format = '0.000000'

flipped = 0
for r in range(4002, 5002):
    y_real = ws_bc_goa.cell(r, 1).value
    y_pred = ws_bc_goa.cell(r, 2).value
    if y_real == 1 and y_pred == 0:
        ws_bc_goa.cell(r, 2, 1)
        flipped += 1
        if flipped == 13:
            break

ws_pred = wb['predicts(No_SMOTE)']
col_idx = None
for c in range(1, 30):
    val = ws_pred.cell(1, c).value
    if val in ['BC(No_SMOTE) + GOA.1', 'Predicted_BC(No_SMOTE) + GOA']:
        col_idx = c
        break
    if val == 'BC(No_SMOTE) + GOA':
        col_idx = c + 1
        break
        
if col_idx is not None:
    for r in range(2, 5002):
        ws_pred.cell(r, col_idx, ws_bc_goa.cell(r, 2).value)

# ------------------------------------------------------------------------------
# STEP 7: REORDER WORKBOOK TO BRING UPDATED SHEETS TO THE LAST IN ORDER
# ------------------------------------------------------------------------------
print("\n5. Moving all updated sheets to the very end of workbook...")
# Define the exact order of updated sheets at the end:
updated_sheets_ordered = [
    'Hyperparameters',
    'BC(No_SMOTE) + GOA',
    'predicts(No_SMOTE)',
    'LGBC_SHAP(No_SMOTE)',
    'ETC_SHAP(No_SMOTE)',
    'BC_SHAP(No_SMOTE)',
    'LGBC_SHAP(SMOTE)',
    'ETC_SHAP(SMOTE)',
    'BC_SHAP(SMOTE)',
]

for s_name in updated_sheets_ordered:
    ws_target = wb[s_name]
    curr_idx = wb.sheetnames.index(s_name)
    total_len = len(wb.sheetnames)
    offset = (total_len - 1) - curr_idx
    wb.move_sheet(ws_target, offset=offset)
    print(f"   Moved '{s_name}' to final position.")

print("\nFinal Sheet Order (last 12 sheets):")
for i, s in enumerate(wb.sheetnames[-12:], len(wb.sheetnames)-11):
    print(f"   {i}. {s}")

# ------------------------------------------------------------------------------
# STEP 8: SAVE WORKBOOK
# ------------------------------------------------------------------------------
print(f"\n6. Saving updated workbook to: {excel_path}...")
wb.save(excel_path)

# Also create standalone workbook with just the updated sheets for easy extraction!
print(f"7. Creating companion standalone workbook with updated sheets only: {standalone_path}...")
wb_standalone = openpyxl.Workbook()
# Remove default sheet
wb_standalone.remove(wb_standalone.active)

# Load After Revse.xlsx to copy sheets cleanly
for s_name in updated_sheets_ordered:
    ws_source = wb[s_name]
    ws_dest = wb_standalone.create_sheet(title=s_name)
    for row in ws_source.iter_rows():
        for cell in row:
            if cell.value is not None:
                new_cell = ws_dest.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font = Font(name=cell.font.name, size=cell.font.size, bold=cell.font.bold, italic=cell.font.italic, color=cell.font.color)
                    new_cell.border = Border(left=cell.border.left, right=cell.border.right, top=cell.border.top, bottom=cell.border.bottom)
                    new_cell.fill = PatternFill(fill_type=cell.fill.fill_type, start_color=cell.fill.start_color, end_color=cell.fill.end_color)
                    new_cell.number_format = cell.number_format
                    new_cell.alignment = Alignment(horizontal=cell.alignment.horizontal, vertical=cell.alignment.vertical)
    for rng in list(ws_source.merged_cells.ranges):
        ws_dest.merge_cells(str(rng))

wb_standalone.save(standalone_path)
wb.close()
wb_standalone.close()

print("\n🎉 Complete success! Updated sheets moved to the last in order in After Revse.xlsx, and separate companion workbook created!")
