# ML Research Pipeline (V2)

A modular, authentic, and reproducible end-to-end Machine Learning research pipeline designed for scientific and engineering research.

---

## 🏛️ Repository Architecture & Folder Structure

```
ML-Research-Pipeline-V2/
│
├── task/                              # Primary task data files
│   └── Data.xlsx                      # Main dataset (sheets: Data, Train_Data, Test_Data, etc.)
│
├── data_manage/                       # Data preparation, encoding, cleaning, and resampling
│   ├── preprocessing/                 # Encoding, scaling, outlier removal, normalization
│   │   ├── LabelEncoder.py            # Categorical string-to-numeric encoding
│   │   ├── Z_Score.py                 # Parametric outlier removal (Train partition only)
│   │   ├── Z_Score_Median.py          # Median-based robust Z-Score outlier removal
│   │   ├── IQR.py                     # Interquartile range outlier detection
│   │   ├── IsolationForest.py         # Unsupervised anomaly detection
│   │   ├── Data_Normalization.py      # Feature normalization & min-max scaling
│   │   ├── Data_randomization.py      # Seeded row shuffling
│   │   ├── Categorical_Counter.py     # Class distribution & categorical inspection
│   │   └── Check_Class_Targets.py     # Target imbalance & class frequency checker
│   ├── balancing/                     # Class resampling techniques (Train partition only)
│   │   ├── SMOTE.py                   # Synthetic Minority Over-sampling Technique
│   │   ├── SMOTE-ENN.py               # Combined SMOTE with Edited Nearest Neighbors
│   │   ├── SMOTE-ENC.py               # SMOTE for Nominal & Continuous variables
│   │   ├── ENN.py                     # Edited Nearest Neighbors cleaning
│   │   └── CTAB-GAN.py                # Conditional GAN for tabular data synthesis
│   └── normality/                     # Normality verification & statistical distributions
│       ├── Shapiro_Wilk.py            # Shapiro-Wilk test for normality
│       ├── QQ_Plot.py                 # Quantile-Quantile (Q-Q) probability plots
│       └── dendrogram.py              # Hierarchical clustering & feature dendrogram
│
├── analysis/                          # Statistical analysis, validation, sensitivity & uncertainty
│   ├── feature_selection/             # Dimensionality reduction & feature ranking
│   │   ├── ANOVA.py                   # Analysis of Variance (F-statistic & p-value screening)
│   │   ├── Chi_Square.py              # Chi-Square feature dependence test
│   │   ├── F_Statistic.py             # Linear discriminant F-score analysis
│   │   ├── Kruskal-Wallis.py          # Non-parametric ANOVA
│   │   ├── MRMR.py                    # Minimum Redundancy Maximum Relevance
│   │   ├── PCA_Principal_Component_Analysis.py  # Principal Component Analysis
│   │   ├── RFE.py                     # Recursive Feature Elimination
│   │   └── VIF.py                     # Variance Inflation Factor (multicollinearity)
│   ├── K_fold/                        # Cross-validation strategies (Default Hyperparameters)
│   │   ├── K_Fold_classification.py   # Stratified K-Fold classification CV
│   │   ├── K_Fold.py                  # Standard K-Fold regression CV
│   │   ├── K_Fold_nested.py           # Nested cross-validation
│   │   ├── K_Fold_nested_bayes.py     # Nested Bayesian cross-validation
│   │   └── K_Fold_SHAP.py             # Cross-validated SHAP feature attribution
│   ├── Sensitivity/                   # Global & local interpretability methods
│   │   ├── SHAP.py                    # SHAP feature importance & summary plots
│   │   ├── SHAP_MultiClass.py         # Multi-class SHAP attribution
│   │   ├── Morris_Sensitivity.py      # Morris elementary effects screening (mu, mu*, sigma)
│   │   ├── PDP.py                     # Partial Dependence Plots (strictly on holdout Test)
│   │   ├── Sobol.py                   # Variance-based global sensitivity analysis (Saltelli/Sobol)
│   │   ├── FAST_Classification.py     # Fourier Amplitude Sensitivity Test
│   │   ├── LIME.py                    # Local Interpretable Model-agnostic Explanations
│   │   ├── LIME_regression.py         # LIME for regression tasks
│   │   ├── PCC.py                     # Partial Correlation Coefficients
│   │   ├── CAM.py                     # Class Activation Mapping
│   │   ├── Decision_Boundaries.py     # 2D/3D Decision boundary visualization
│   │   └── copula/                    # Copula dependence & joint distribution modeling
│   ├── Statistical-analysis/          # Hypothesis testing & pairwise comparison
│   │   ├── DataCatcher.py             # Prediction aggregator across models
│   │   ├── DataCatcher_Probability.py # Probability aggregator across models
│   │   ├── Statistical_t-test.py      # Paired two-sample t-test
│   │   ├── wilcoxon.py                # Wilcoxon signed-rank test
│   │   ├── friedman.py                # Friedman rank-sum non-parametric test
│   │   └── McNemar.py                 # McNemar test for paired nominal data
│   ├── Uncertainty/                   # Probabilistic calibration & uncertainty quantification
│   │   ├── Brier_Decomposition.py     # Brier score decomposition (Reliability & Resolution)
│   │   ├── ConfidenceInterval_Regression.py # Bootstrap confidence intervals
│   │   ├── Entropy.py                 # Shannon entropy-based uncertainty analysis
│   │   ├── MCS.py                     # Monte Carlo simulation
│   │   └── Quantile_Loss.py           # Pinball / Quantile loss estimation
│   └── REC.py                         # Regression Error Characteristic (REC) curve
│
├── model/                             # Machine learning predictive algorithms
│   ├── classification/                # Binary classification models (17 algorithms)
│   │   ├── RFC.py                     # Random Forest Classifier
│   │   ├── LGBC.py                    # LightGBM Classifier
│   │   ├── XGBC.py                    # XGBoost Classifier
│   │   ├── SVC.py                     # Support Vector Classifier
│   │   ├── LR.py                      # Logistic Regression
│   │   ├── KNNC.py                    # K-Nearest Neighbors Classifier
│   │   ├── GBC.py                     # Gradient Boosting Classifier
│   │   ├── ETC.py                     # Extra Trees Classifier
│   │   ├── DTC.py                     # Decision Tree Classifier
│   │   ├── ADAC.py                    # AdaBoost Classifier
│   │   ├── BC.py                      # Bagging Classifier
│   │   ├── CB.py                      # CatBoost Classifier
│   │   ├── GPC.py                     # Gaussian Process Classifier
│   │   ├── LDA.py                     # Linear Discriminant Analysis
│   │   ├── QDA.py                     # Quadratic Discriminant Analysis
│   │   ├── SGBC.py                    # Stochastic Gradient Boosting Classifier
│   │   └── ELM.py                     # Extreme Learning Machine
│   ├── classification/MultiClass/     # Multi-class classification models (18 algorithms)
│   │   ├── MLR.py                     # Multinomial Logistic Regression
│   │   ├── RFC.py, LGBC.py, XGBC.py, SVC.py, KNNC.py, CB.py, ...
│   │   └── RNN.py                     # Recurrent Neural Network
│   ├── ensemble/                      # Meta-learning and model aggregation
│   │   ├── VotingClassifier.py        # Soft/Hard Voting Classifier
│   │   ├── Stacking_Regression.py     # Stacking Regressor with linear meta-learner
│   │   ├── Stacking_MultiClass.py     # Multi-class Stacking with logistic meta-learner
│   │   └── Ensemble_Regression.py     # Weighted ensemble regressor
│   └── Regression/                    # Continuous regression algorithms (25 algorithms)
│       ├── LinearRegression.py        # Ordinary Least Squares Linear Regression
│       ├── RR.py                      # Ridge Regression (L2 regularization)
│       ├── Lasso.py / LLAR.py         # Lasso & LassoLars Regression (L1 regularization)
│       ├── ENR.py                     # ElasticNet Regression (L1 + L2)
│       ├── RFR.py                     # Random Forest Regressor
│       ├── LGBR.py                    # LightGBM Regressor
│       ├── XGBR.py                    # XGBoost Regressor
│       ├── SVR.py / LSSVR.py          # Support Vector Regressor & Least Squares SVR
│       ├── KNNR.py                    # K-Nearest Neighbors Regressor
│       ├── GBR.py / SGBR.py           # Gradient Boosting & Stochastic GBR
│       ├── ETR.py                     # Extra Trees Regressor
│       ├── DTR.py                     # Decision Tree Regressor
│       ├── CATR.py                    # CatBoost Regressor
│       ├── ADAR.py                    # AdaBoost Regressor
│       ├── ANFIS.py                   # Adaptive Neuro-Fuzzy Inference System
│       ├── EBM.py                     # Explainable Boosting Machine
│       ├── ELM.py                     # Extreme Learning Machine Regressor
│       ├── GPR.py                     # Gaussian Process Regressor
│       ├── HGBR.py                    # HistGradientBoosting Regressor
│       ├── HR.py                      # Huber Regressor (robust regression)
│       ├── KRR.py                     # Kernel Ridge Regressor
│       ├── NGBM.py                    # Natural Gradient Boosting
│       ├── QR.py                      # Quantile Regressor
│       └── SBR.py                     # Sparse Bayesian Regressor (ARD)
│
├── optimization/                      # Metaheuristic Optimization Algorithms (30 Algorithms)
│   │                                  # Strictly 200 real iterations; logs convergence & params
│   ├── ACO.py                         # Ant Colony Optimization
│   ├── AOA.py                         # Arithmetic Optimization Algorithm
│   ├── CCO.py                         # Coyote Optimization Algorithm
│   ├── COA.py                         # Cyclone Optimization Algorithm
│   ├── Chimp.py                       # Chimp Optimization Algorithm
│   ├── DOA.py                         # Dingo Optimization Algorithm
│   ├── DSOA.py / SDOA.py              # Spotted Deer Optimization Algorithm
│   ├── EAO.py                         # Enzyme Action Optimizer
│   ├── EOA.py                         # Echidna Optimization Algorithm
│   ├── GOA.py                         # Grasshopper Optimization Algorithm
│   ├── GWOA.py                        # Grey Wolf Optimizer
│   ├── HEOA.py                        # Highland Ecosystem Optimization Algorithm
│   ├── HMOA.py                        # Honey Mole Optimization Algorithm
│   ├── HOA.py                         # Horse Optimization Algorithm
│   ├── HROA.py                        # Harris Hawks / Hybrid Optimization
│   ├── LOA.py                         # Lemur Optimization Algorithm
│   ├── Lion.py                        # Lion Optimization Algorithm
│   ├── MAOA.py                        # Marine Predators / Multi-agent Optimizer
│   ├── NOA.py                         # Nurse Optimization Algorithm
│   ├── OOA.py                         # Osprey Optimization Algorithm
│   ├── POA.py                         # Perfumer Optimization Algorithm
│   ├── PSO.py                         # Particle Swarm Optimization
│   ├── RPOA.py                        # Red Panda Optimization
│   ├── SBTO.py                        # Savannah Bengal Tiger Optimization
│   ├── SEOA.py                        # Starling Ecosystem Optimization Algorithm
│   ├── SWOA.py                        # Spider Wasp Optimization Algorithm
│   ├── WEOA.py                        # Wetland Ecosystem Optimization Algorithm
│   ├── WOA.py                         # Whale Optimization Algorithm
│   ├── ZOA.py                         # Zebra Optimization Algorithm
│   ├── plot_search_space.py           # 3D parameter response surface visualization
│   └── plot_bayesian_optimization.py  # Acquisition function & surrogate trajectory
│
├── exporters/ (Root Exporters)        # Master Excel report generators with exact styling
│   ├── ClassificationModelToExcel.py  # Binary Classification Exporter (Accuracy, FNR, AUC, Kappa, MCC)
│   ├── MultiClassificationModelToExcel.py # Multi-Class Classification Exporter
│   ├── MultiClassificationModelToExcel.py # Backward-compatible alias
│   └── RegressionFunctionModelToExcel.py # Regression Exporter (R2, RMSE, MAE, COM, MBE, LogCosh)
│
├── data/                              # Exported prediction vectors & error files (.npt format)
│   ├── model1.npt to model6.npt       # Prediction arrays for models and optimizers
│   └── Data_err.npt                   # Baseline reference error array
│
└── codes/                             # Archival representative scripts for active tasks
```

---

## ⚡ Core Operational Principles

1. **Absolute Zero Faking**: NO fake accuracy, NO synthetic convergence, NO mock predictions. Everything is computed authentically from real model training, cross-validation, and optimization.
2. **Default Hyperparameters for Baseline & K-Fold**: Baseline single models and K-Fold cross validation are evaluated with standard default constructor arguments (`RandomForestClassifier()`, `SVC()`, etc.) without artificial constraints or regularizations.
3. **Strict 200 Real Iterations for Metaheuristics**: Every optimizer runs for exactly 200 genuine iterations, logging the convergence fitness and hyperparameter values at every iteration to `HyperIteration.xlsx`.
4. **Early 20% Holdout Test Separation**: 20% of data is separated immediately after ingestion/encoding as `Test_Data`. All preprocessing (outlier removal, balancing, feature selection) occurs strictly on `Train_Data` to eliminate data leakage.
5. **Exact Mathematical Alignment**: Summary tables, test row metrics, and Confusion Matrix diagonals must mathematically align ($\text{Accuracy} = \frac{\sum \text{TP}}{N_{\text{test}}}$).
