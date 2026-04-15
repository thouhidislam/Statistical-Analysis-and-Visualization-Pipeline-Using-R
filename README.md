# Statistical Analysis and Visualization Pipeline Using R

This project provides a comprehensive and reproducible R workflow for statistical analysis, visualization, and reporting of biomedical and survey datasets. It integrates data preprocessing, categorical analysis, regression modeling, and graphical representation in a single structured pipeline.

The workflow is designed to handle multiple types of datasets, including health survey data, pharmacological experiments, and biological assay results.

---

## 📊 Features

### Data Preprocessing
- Data cleaning using `janitor`
- Conversion of categorical and numerical variables
- Handling missing values and formatting inconsistencies

### Statistical Analysis
- Multinomial logistic regression (`nnet::multinom`)
- Chi-square tests for association analysis
- Frequency table generation
- Summary statistics (mean, SEM)

### Visualization
- Line plots and bar plots using `ggplot2`
- Multi-sample comparison graphs
- Publication-quality figures

### Reporting
- Automated table generation using `flextable`
- Export to Word (.docx) and Excel (.xlsx)
- Model summaries with odds ratios, confidence intervals, and p-values

---

## 📦 Packages Used

- tidyverse  
- dplyr  
- nnet  
- flextable  
- officer  
- rio  
- ggplot2  

---

## 📁 Outputs

The pipeline generates:

- Cleaned datasets (.csv / .xlsx)
- Regression results tables
- Chi-square test reports
- Publication-ready Word documents
- Graphical visualizations (PDF/PNG)

---

## 🔬 Applications

This pipeline can be used for:

- Medical and health survey analysis  
- Breast cancer awareness studies  
- Pharmacological and antimicrobial experiments  
- Biochemical assay data analysis  
- Academic research and publication preparation  

---

## ⚙️ Workflow Summary

1. Load raw dataset  
2. Clean and standardize variables  
3. Perform statistical modeling  
4. Run association tests (Chi-square)  
5. Generate publication-ready tables  
6. Create visualization plots  
7. Export results to Word and Excel  

---

## 📜 License

This project is released under the MIT License, allowing free use, modification, and distribution with proper attribution.

---

## 👤 Author

Thouhid Islam  
Department of Pharmacy  
University of Chittagong
