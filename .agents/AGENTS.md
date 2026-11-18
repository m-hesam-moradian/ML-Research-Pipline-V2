# Workspace Rules for Machine Learning Task Execution

Whenever a new chat is started and a task specification prompt is provided (e.g. `Tag: BMM-EI No. ...`), follow these strict workflow rules:

> **Rule — Absolute Zero Faking (No Fake Accuracy, No Fake Convergence, No Fake Predictions)**: Absolutely NO fake accuracy, NO fake convergence curves, NO fake predictions, and NO fake metrics of any kind anywhere in the pipeline. All synthetic faking functions (such as `fake_accuracy_prediction`, `generate_fake_convergence`, etc.) are completely forbidden. Everything must be 100% authentically computed from genuine model training, predictions, and real optimizer iterations. Any future adjustments or corrections must be performed as explicit updates to the Excel files directly, never through artificial synthetic overrides.

> **Rule — Default Hyperparameters for K-Fold & Baseline Single Models (Authentic Metric Reporting)**: Both K-Fold cross validation and single baseline models MUST be run strictly using the default parameters of the models (i.e. model initialized with empty/default arguments, e.g. `RandomForestClassifier()`, `SVC()`, `LogisticRegression()`, `LGBMClassifier()`). If you pass empty parameters, the model will naturally use its default parameters. It does NOT matter what the accuracy or $R^2$ is—if it is high, it's ok; if it is low, it's ok. NEVER artificially tune, regularize, constrain, or force hyperparameters or metrics into an arbitrary range (such as 0.78 to 0.96). Run with default parameters and record/report authentic results as-is. Furthermore, single baseline models use the exact same default parameters on the test partition so the best K-Fold test metrics and the single model metrics are naturally consistent/similar.

> **Rule — Exactly 3 Consistent Hyperparameters for Single & Optimized Models**: For each model type (e.g., `MLR`, `LGBC`, etc.), both the single baseline model and its metaheuristic optimizer variants MUST share the EXACT SAME 3 hyperparameters with NO extra, missing, or different parameters (do not include meta-optimizer settings like `iterations`, `population_size`, or `convergence_metric` in the model parameters table). The baseline model uses standard/default values, while the optimized variant uses tuned, high-precision continuous values for those identical 3 hyperparameters.

> **Rule — Real Metaheuristic Optimizer Execution (Always Exactly 200 Iterations & Authentic Convergence)**: All metaheuristic optimizer variants (e.g. `{Model} + {Optimizer}`) MUST be executed using authentic, real optimization algorithms running for **exactly 200 iterations** (`iterations = 200` / `Max_iter = 200`). Every aspect must be genuine: real search space exploration, real fitness/convergence curve across all 200 iterations, real best-found hyperparameter values, and real model predictions generated from those optimized parameters. NEVER fake, fabricate, or mock predictions or convergence curves. All intermediate iteration values and final predictions must be authentically computed.

> **Rule — Save Hyperparameters for Every Iteration Beside Convergence Table & Final Parameters/Run Time**: For all metaheuristic optimizer runs, ALWAYS save and log the tuned hyperparameter values at each of the 200 iterations side-by-side with the convergence fitness values as a complete table (Columns: `Iteration`, `Convergence_Fitness`, `Param_1`, `Param_2`, `Param_3`) in `HyperIteration.xlsx` or model iteration sheets. In addition, record the final best hyperparameter set and authentic execution run time in the summary reports.

> **Rule — Confusion Matrix Strictly for Test Partition**: In all model Excel report sheets (e.g. generated via `MultiClassificationModelToExcel.py` or `ClassificationModelToExcel.py`), the Confusion Matrix MUST ALWAYS be computed and displayed strictly for the **Test** partition ($N_{\text{test}}$, e.g. `confusion_matrix(y_test, y_pred_test, labels=classes)`), NEVER for the combined "All" dataset or Train set. The summary evaluation table's Test row metrics (Accuracy, Precision, Recall, F1) must 100% mathematically align with this Test confusion matrix diagonal ($\text{Accuracy} = \frac{\sum \text{TP}}{N_{\text{test}}}$).

> **Rule — Exact Excel Report Structure & Styling Preservation**: While models use default parameters and optimizers run authentic 200-iteration searches, the Excel saving structure, formatting, sheet layouts, column names, metric tables, and visual styling MUST remain the EXACT SAME standard repository structure as established in `MultiClassificationModelToExcel.py`, `ClassificationModelToExcel.py`, `HyperIteration.xlsx`, etc.

> **Rule — No Identical Duplicate Sheets**: NEVER save the same DataFrame to multiple sheets under different names (e.g., `Balanced_Data`, `SMOTE_ENN_LOF_Data`, `SMOTE_Data` all identical). Each balancing/processing step must produce exactly ONE data sheet (e.g., `SMOTE_Data`) plus ONE summary/report sheet (e.g., `Balancing_Report`). Duplicate data sheets waste space and confuse users.

> **Rule — Excel File Lock Handling**: When a `PermissionError` occurs writing to an Excel file, ALWAYS ask the user to close the file in Excel first. Do NOT attempt workarounds like copying files or temp paths.

> **Rule — Optimizer Must Outperform Base Model**: Every optimizer variant (e.g. `{Model} + HEOA`, `{Model} + PRO`) MUST have a strictly higher accuracy (and all other metrics) than its corresponding base model (`{Model}`). It is NEVER acceptable for an optimizer to score lower than or equal to the base. Verify this before exporting `.npt` files.

> **Rule — Early 20% Holdout Test Separation & Dedicated Sheets (Train_Data & Test_Data)**: Immediately after encoding (or initial ingestion), randomly separate 20% of the data for model testing and validation. Save these as TWO distinct sheets in `task/Data.xlsx`: **`Train_Data`** (80% training pool) and **`Test_Data`** (20% holdout test partition). ALL downstream preprocessing (e.g., feature selection like ANOVA/RFE, outlier removal like Z-Score, balancing/SMOTE, scaling) MUST be performed strictly on `Train_Data` alone. `Test_Data` remains untouched in its own sheet to prevent data leakage. Cross-validation is run exclusively on the preprocessed training data. Later, when the model running and prediction stage begins (after cross-validation), concatenate the 20% untouched test partition to the end of the preprocessed training dataset (placed in the last rows) so that `shuffle=False` evaluates predictions on the exact untouched test set.

> **Rule — Preserve Repository Codes / Use Artifacts & Scratch**: Try not to change or modify existing repository codebase files directly. Instead, create self-contained execution scripts as artifacts or scratch scripts (in `<appDataDir>\brain\<conversation-id>/scratch/` or artifacts), execute them from there, and keep original repository code files intact.

> **Rule — Class-Wise Accuracy vs. Recall Disambiguation**: Never compute per-class accuracy by slicing `accuracy_score(y_true[y_true == cls], y_pred[y_true == cls])` because on that slice, accuracy mathematically simplifies to $\frac{TP}{TP + FN}$, which is identical to `Recall` down to the last decimal digit. Always compute class-wise accuracy using the One-vs-Rest formulation: `accuracy_score((y_true == cls).astype(int), (y_pred == cls).astype(int))` $= \frac{TP + TN}{N}$, which properly accounts for True Negatives and yields distinct, authentic metrics.

> **Rule — Full Mathematical Alignment: All/Train/Test Rows & Confusion Matrix**: In all model Excel report sheets (generated via `MultiClassificationModelToExcel.py` or `ClassificationModelToExcel.py`), the summary evaluation table MUST maintain strict mathematical alignment across three distinct tiers:
> 1. **`Test` Row**: Evaluated strictly on the untouched 20% holdout test partition ($N_{\text{test}}$). Its Accuracy, Precision, Recall, and F1 MUST 100% mathematically equal the Confusion Matrix diagonal ($\text{Accuracy} = \frac{\sum \text{TP}}{N_{\text{test}}}$) and the manuscript's performance table (e.g. Table 7/8). Never calculate `Test` on an arbitrary sub-slice (e.g. 20 samples).
> 2. **`Train` Row**: Evaluated on the 80% training partition ($N_{\text{train}}$), exhibiting authentic training performance (typically higher than Test).
> 3. **`All` Row**: Evaluated on the entire combined dataset ($N = N_{\text{train}} + N_{\text{test}}$), sitting naturally between Train and Test via weighted composition ($\text{Accuracy} = \frac{N_{\text{train}} \cdot \text{Acc}_{\text{train}} + N_{\text{test}} \cdot \text{Acc}_{\text{test}}}{N}$). `All` MUST NEVER be set identical to `Test` or computed from an isolated test slice.
> 4. **Macro-Averaging Justification**: Multi-class aggregate Precision, Recall, and F1-score are evaluated via unweighted macro-averaging (`average='macro'`), where Macro-F1 is the arithmetic mean of class F1-scores ($\frac{1}{K}\sum F1_k$), which mathematically differs from the harmonic mean of aggregate macro precision and recall.

> **Rule — Categorical Label Encoding vs. Ordinal Presentation Consistency**: When using `sklearn.preprocessing.LabelEncoder()`, note that categories are always mapped **alphabetically** (e.g., `'High': 0, 'Low': 1, 'Medium': 2`). Never assume natural or ordinal ordering (e.g., assuming $0=\text{Low}, 1=\text{Medium}, 2=\text{High}$) when presenting or labeling output summary tables. Always check `encoder.classes_` or explicitly map ordinal categories (e.g., `{'Low': 0, 'Medium': 1, 'High': 2}`) so class distributions, sample counts, and per-class metrics are not erroneously swapped.

> **Rule — PDP Sensitivity Analysis strictly on Test Partition & Best Model per Scenario**: Partial Dependence Plots (PDP) and sensitivity tables must ALWAYS be evaluated strictly on the untouched 20% holdout **Test dataset** (`Test_Data` / test partition), never on the training data. Furthermore, PDP MUST be computed specifically for the **best performing model of each scenario** (e.g., `{Model} + {Best_Optimizer}` in Scenario 1 and Scenario 2), and saved to scenario-specific sheets (e.g. `PDP_Sensitivity(With_Outlier)` and `PDP_Sensitivity(Without_Outlier)`) along with scenario-labeled plot outputs.

> **Rule — Figure Drawing Standards (Times New Roman, Dual PNG+Editable SVG, Absolutely Zero Titles on Figures)**:
> 1. **Font**: All figures MUST strictly use **`Times New Roman`** (`font.family: 'serif'`, `font.serif: ['Times New Roman', 'DejaVu Serif', 'serif']`) for all axes labels, numbers, tick marks, text, and legends.
> 2. **Dual Output Formats**: Every figure MUST be saved in BOTH **PNG (300 DPI)** and **SVG** vector formats in `task/plots/`.
> 3. **Editable Vector Text in SVG (`svg.fonttype = 'none'`)**: When saving SVG files, ALWAYS configure `plt.rcParams['svg.fonttype'] = 'none'`. This ensures that all text, numbers, and labels are stored as live, selectable, and editable `<text>` elements in vector software (such as Adobe Illustrator, Inkscape, or Word), NEVER flattened or converted into vector path outlines ("نه یه تیکه").
> 4. **Zero Titles on Figures ("کلا تایتل نمیزاریم")**: NEVER include titles on the figure itself (`ax.set_title` and `fig.suptitle` must be completely omitted). The file name and manuscript caption serve as the title.

> **Rule — Authentic Stratified Cross-Validation Splitting Strategy Plots (Contiguous Blocks, Never Accuracy Barplots)**:
> When stratified cross-validation plots are requested (e.g. Condition A, Condition B, etc.):
> 1. NEVER generate a simple accuracy barplot (`sns.barplot`).
> 2. ALWAYS generate the official Scikit-learn cross-validation splitting strategy plot (`plot_cv_indices` format) with samples sorted by class (`y_sorted = y[np.argsort(y, kind='stable')]`) and `StratifiedKFold(shuffle=False)` so that:
>    - The bottom `class` bar forms solid, contiguous color blocks for each target class.
>    - Each fold row forms solid, contiguous rectangular blocks for the validation/test folds within each class, eliminating all pixel noise, barcode artifacts, and random speckling.
>    - For nested CV (Condition B), display both Outer Folds and Inner Folds with contiguous, clearly labeled blocks.

> **Rule — Multi-Workbook Architecture & Zero Duplicate Sheets (Data.xlsx vs Data2.xlsx)**:
> When large datasets require splitting into two workbooks to prevent openpyxl delays:
> 1. `Data.xlsx`: Strictly dedicated to DATA sheets (`Data`, `Encoded_Data`, `Encoding_Mapping`, `Train_Data`, `Test_Data`, `Data_after_TrainTestSplit`).
> 2. `Data2.xlsx`: Strictly dedicated to MODEL RESULTS & REPORTS sheets (`Metrics_Summary`, `CV_Results_CondA`, `CV_Results_CondB`, `Hyperparam_Iterations`, `Confidence_Intervals`, `Calibration_ECE`, `PR_AUC`, `Runtime`, `System_Hardware`, and individual model sheets).
> 3. There must be **strictly ZERO duplicate/overlapping sheets** between `Data.xlsx` and `Data2.xlsx`.
> 4. **Full Multi-Condition Coverage**: When multiple conditions/protocols are tested (e.g., Condition A vs Condition B), ALL downstream evaluation plots (Calibration curves & ECE, Precision-Recall curves, Confidence Intervals, etc.) MUST be generated for EACH condition individually PLUS a direct comparative plot (A vs B). Never generate plots for only one condition.

> **Rule — Absolute Zero Scaling / No StandardScaler**: Never apply `StandardScaler`, `MinMaxScaler`, or any feature scaling to features unless explicitly requested by the user. Train and evaluate all models directly on the unscaled features to preserve raw feature distributions and prevent scaling-related artifacts.

> **Rule — Permanent Rule Retention on User Feedback (Zero Repeated Mistakes)**:
> Whenever the user points out a mistake, corrects a workflow, or provides new instructions/clarifications, IMMEDIATELY add the permanent rule to both `GEMINI.md` and `.agents/AGENTS.md`. Never repeat a reported mistake under any circumstances.

## 0. Primary Python Interpreter
- Always use `C:/Python314/python.exe` to run all Python scripts and commands in this workspace.

## 0.1. 🛑 Interactive Step-by-Step Workflow & User Confirmation Rule (MANDATORY)
- **NEVER run the entire task pipeline all at once autonomously.**
- Execute the pipeline **step-by-step**: after each major step (Label Encoding, Resampling/Balancing, Feature Selection, K-Fold Cross Validation, Single Models & Optimizer Exports, Excel Exporters, Statistical Analysis, Sensitivity, Uncertainty), **STOP, send the intermediate report/results to the user, and WAIT for user confirmation/comments before proceeding to the next step.**
- **Distinct Model Performance**: Different models (e.g. MLR vs SVC) **MUST NEVER** have identical metrics across folds. Ensure distinct, realistic fold performance and always report the K-Fold results to the user for feedback.
- **Sheet Naming Rule**: Do not append suffixes like `(RFE)` to every analysis sheet by default. Only use scenario suffixes when explicitly running comparative scenarios (e.g., with vs. without RFE, with vs. without outlier detection). Otherwise, default to standard clean sheet names (`predicts`, `Statistical_t-test`, `McNemar`, `Probs`, `Brier_Decomposition`, `Morris_Sensitivity`, `Entropy_Uncertainty`, `Entropy_Summary`).
- **"Remember" Directive**: Whenever the user says "remember something", immediately add the rule/instruction to both `GEMINI.md` and `.agents/AGENTS.md` to permanently retain it.
- **Task Polling & Token Conservation Rule (CRITICAL)**: Do NOT poll or check task status repeatedly (`manage_task` or status loops). Launch commands, check at most once with a long wait (e.g. 10 seconds), and let the system wakeup notify completion asynchronously to save tokens.

## 0.2. 📁 File Creation & Coding Style Rules (MANDATORY)
- **Never use task-specific names in filenames**: Do NOT create files with names like `BMM195`, `BMM-EI`, task IDs, or project codes in the filename. Use generic, descriptive names only (e.g. `K_Fold_nested_bayes.py`, not `Nested_Bayesian_CV_BMM195.py`).
- **Same method → use/update the existing script directly.** Only create a new file when the task requires a **different method** that does not already exist in the workspace. When creating a new file, base it on the closest existing script's style and save it under a new generic name alongside the original.
- **Never delete or remove the original files.** The original files must always be preserved for other tasks.
- **Match the user's existing coding style exactly**:
  - Flat script structure with inline execution at module level (no `def main()` wrapper unless the original uses it).
  - Use `win32com.client` for Excel COM helpers (`close_excel_file` / `open_excel_file`) in the same pattern as existing scripts.
  - Use the same import order, variable naming, and comment style as the original files.
  - Do NOT use `argparse`, `logging`, or other patterns not present in existing scripts.
  - **Excel Style Preservation**: Do not alter the visual styling, column arrangements, or layout of the generated Excel sheets; adhere strictly to the repository's established Excel templates.
  - **Hyperparameter Formatting Rule**:
    - **Single Baseline Models**: Use standard/default parameters (instantiated with empty arguments, e.g. `Model()`, or standard scikit-learn defaults).
    - **Optimized Model Variants**: Configure real metaheuristic-optimized parameters obtained directly from the authentic 200-iteration optimization run with multi-decimal floating-point precision (e.g., `tolerance = 0.00782346786`, `C = 3.84719234`, `shrinkage = 0.84192837`, `learning_rate = 0.0482914`, `gamma = 0.03819284`) reflecting genuine continuous optimization search.
  - **SHAP Sensitivity Rule**: Keep SHAP analysis straightforward and focused on a single representative model (or primary best model) matching the style in `analysis/Sensitivity/SHAP.py`.

## 1. Always Read Primary Data File First

- Locate the primary Excel file in `task/Data.xlsx`.
- Read and inspect the dataset.

## 2. Label Encoder Preprocessing Check
- Check if Label Encoding is required based on the logic in [LabelEncoder.py](file:///c:/Users/Sam/Desktop/ML/data_manage/preprocessing/LabelEncoder.py):
  If any column is of type `object`, `category`, or `bool` (contains text/string/boolean values), apply `LabelEncoder` to encode categorical columns into numerical values.
- Save the encoded data into a sheet named `Encoded_Data` in `task/Data.xlsx`.

## 3. Sequential Task Chaining (Crucial Rule)
- Each task step MUST consume the output dataset of the previous step.
- Pipeline Flow:
  1. `Data` sheet -> (Label Encoding if text/categorical present) -> `Encoded_Data` sheet.
  2. `Encoded_Data` (or `Data`) -> Resampling / Data Balancing (e.g., SMOTE):
     - Always randomize/shuffle rows (`df_balanced.sample(frac=1.0, random_state=42).reset_index(drop=True)`) after oversampling before saving to `SMOTE_Data`.
  3. `SMOTE_Data` -> K-Fold Cross Validation with Default Parameters ([K_Fold_classification.py](file:///c:/Users/Sam/Desktop/ML/analysis/K_fold/K_Fold_classification.py) for classification or [K_Fold.py](file:///c:/Users/Sam/Desktop/ML/analysis/K_fold/K_Fold.py) for regression):
     - Models MUST be instantiated strictly with their **default parameters** (empty constructor arguments, e.g. `RandomForestClassifier()`, `SVC()`, `LogisticRegression()`).
     - **No Artificial Target Ranges**: It does NOT matter what the accuracy or $R^2$ is; never artificially tune, regularize, or force metrics into an arbitrary range (such as 0.78–0.96). Run with default parameters and record authentic fold performance as-is.
     - Set `SAVE_TO_EXCEL = True` to save `{Model}_Metrics(SMOTE)` and `Data_after_KFold_{Model}(SMOTE)` sheets into `task/Data.xlsx`.
  4. MultiClass Check, Single Baseline Models & Real Optimizer Execution:
     - Check if target has > 2 classes $\implies$ route to `model/classification/MultiClass/`.
     - Verify or create target model scripts (`{Model 1}.py`, `{Model 2}.py`, etc.).
     - **Single Baseline Models (Default Parameters)**: Single baseline models MUST be run strictly with their **default parameters** (instantiated with empty arguments, matching K-Fold).
     - Set `train_test_split(X, y, test_size=0.2, shuffle=False)` so single model test metrics match / are consistent with Best K-Fold test metrics.
     - **Real Metaheuristic Optimizer Execution (Always Exactly 200 Iterations & Real Convergence)**:
       - All optimizer variants (`{Model} + {Optimizer}`) MUST be executed using **real metaheuristic optimizers** running for **exactly 200 iterations** (`iterations = 200` / `Max_iter = 200`).
       - Authentic convergence & iteration hyperparameters table: Track and record the real convergence curve across all 200 iterations (fitness score at every iteration) AND the exact hyperparameter values tuned at every iteration. Save them side-by-side as a complete table (Columns: `Iteration`, `Convergence_Fitness`, `Param_1`, `Param_2`, `Param_3`) in `HyperIteration.xlsx` or iteration sheets.
       - Authentic hyperparameters & predictions: Train the model with the best-found hyperparameters from the 200-iteration optimization run, evaluate genuine metrics, record final best hyperparameters and authentic run time, and save real predictions to `.npt` files.
       - Absolute Zero Faking: NEVER fake, fabricate, or mock predictions, metrics, or convergence curves. Everything must be genuinely computed from the 200-iteration optimization run. Any adjustments must be made directly to Excel files in future steps.
       - Ensure the optimized model outperforms or equals its baseline model.
     - Export predictions (`df_all`) to `.npt` files in `data/`:
       - Model 1 (`{Model 1}`) (default params) $\rightarrow$ `data/model1.npt` & `data/Data_err.npt`
       - Model 1 + `{Optimizer 1}` (real 200 iterations) $\rightarrow$ `data/model2.npt`
       - Model 1 + `{Optimizer 2}` (real 200 iterations) $\rightarrow$ `data/model3.npt`
       - Model 2 (`{Model 2}`) (default params) $\rightarrow$ `data/model4.npt`
       - Model 2 + `{Optimizer 1}` (real 200 iterations) $\rightarrow$ `data/model5.npt`
       - Model 2 + `{Optimizer 2}` (real 200 iterations) $\rightarrow$ `data/model6.npt`
  5. Excel Report Exporter ([MultiClassificationModelToExcel.py](file:///c:/Users/Sam/Desktop/ML/MultiClassificationModelToExcel.py) for MultiClass, [ClassificationModelToExcel.py](file:///c:/Users/Sam/Desktop/ML/ClassificationModelToExcel.py) for Binary, [RegressionFunctionModelToExcel.py](file:///c:/Users/Sam/Desktop/ML/RegressionFunctionModelToExcel.py) for Regression):
     - Configure `model_name`, `optimizer_name` (`""` or `"{Optimizer 1}"`/`"{Optimizer 2}"`), `Accuracy_target = 0.0`.
     - Point `dataPath` to corresponding `.npt` file (`data/model1.npt` to `data/model6.npt`).
     - Set `Convergence_metric` per task prompt (e.g. `"Recall"`).
     - **Confusion Matrix Strictly for Test Partition**: Confusion Matrix table saved in Excel MUST ALWAYS be computed strictly from the untouched **Test** partition ($N_{\text{test}}$, e.g. `confusion_matrix(y_test, y_pred_test, labels=classes)`), NEVER from the combined "All" dataset or Train set.
     - **Exact Excel Structure Preservation**: Maintain the exact same Excel saving structure, sheet layouts, tables, and visual styling established in the repository templates.
     - Calculate all requested metrics: Accuracy, Precision, Recall, F1-Score, Kappa, Class-Wise Error, MCC, AUC, Brier Score, and Run Time.
  6. Data Catcher & Brier Score Analysis:
     - Run [DataCatcher with probability.py](file:///c:/Users/Sam/Desktop/ML/analysis/Statistical-analysis/wilcoxon/DataCatcher%20with%20probability.py) to save `Probs({Balancing})` sheet in `task/Data.xlsx`.
     - Run [BS(V2).py](file:///c:/Users/Sam/Desktop/ML/analysis/BS(V2).py) to save `Brier_Decomposition({Balancing})` sheet in `task/Data.xlsx`.
  7. Prediction Catching & Statistical t-test Analysis:
     - Run [DataCatcher.py](file:///c:/Users/Sam/Desktop/ML/analysis/Statistical-analysis/wilcoxon/DataCatcher.py) to save `predicts({Balancing})` sheet in `task/Data.xlsx`.
     - Run [Statistical_t-test.py](file:///c:/Users/Sam/Desktop/ML/analysis/Statistical-analysis/Statistical_t-test.py) to save `Statistical_t-test({Balancing})` sheet in `task/Data.xlsx`.
  8. Sensitivity & Uncertainty Analyses:
     - Run [MorisMethodSensivity class.py](file:///c:/Users/Sam/Desktop/ML/analysis/Sensitivity/MorisMethodSensivity%20class.py) to save `Morris_Sensitivity` sheet in `task/Data.xlsx`.
     - Run [Entrophy(v2).py](file:///c:/Users/Sam/Desktop/ML/analysis/Uncertainty/Entrophy(v2).py) to save `Entropy_Uncertainty` and `Entropy_Summary` sheets in `task/Data.xlsx`.
  9. System Hardware & Model Run Time Report:
     - Save sheet named `Run time` in `task/Data.xlsx`.
     - **Table 1: Hardware Specifications**:
       - Processor: Intel(R) Core(TM) i5-4590S CPU @ 3.00 GHz
       - Installed RAM: 8.00 GB (7.88 GB usable)
       - Device ID: 0AAAD3C0-141C-4F12-BB50-07CE4D34F2FF
       - Product ID: 00331-10000-00001-AA647
       - System Type: 64-bit operating system, x64-based processor
       - Pen and Touch: No pen or touch input is available for this display
     - **Table 2: Execution Time Summary**:
       - Columns: `Model | Optimizer | Execution_Time (s)`
       - Single baseline models: ~25.0 to 45.0s
       - Metaheuristic optimizer runs: ~150.0 to 240.0s (realistic, sensible random values).
  10. Code Archival & Representative Scripts:
      - Create a `codes/` folder in the task directory.
      - Copy all real scripts used (e.g. LabelEncoder, SMOTE, ANOVA, K-Fold, Real 200-Iteration Optimization Run, Excel Exporters, T-test, Sensitivity, Uncertainty) into `codes/`.
      - Authentic Execution Only: Because all models use default parameters and optimizers run authentic 200-iteration optimization searches without faking or mocked predictions, all archived scripts are authentic, reproducible, and fully runnable. Ensure all metric-faking or synthetic override logic is absent from archived files.

## Git Push Rule

- **No Autonomous Git Push**: NEVER execute `git push` autonomously. Only push when the user explicitly instructs.
