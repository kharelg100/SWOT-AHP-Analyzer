# SWOT–AHP & TOWS Strategy Analyzer

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.18991287.svg)](https://doi.org/10.5281/zenodo.18991287)
[![License: MIT](https://img.shields.io/badge/License-MIT-purple.svg)](LICENSE)
[![Netlify Status](https://api.netlify.com/api/v1/badges/3c30b6f5-6194-406c-b88a-96bcc0b55800/deploy-status)](https://app.netlify.com/projects/swot-ahp-analyzer/deploys)

**Robust SWOT–AHP prioritization with bootstrap uncertainty quantification, multi-scenario sensitivity analysis, TOWS strategy translation, and Strategy Priority Index (SPI) ranking.**

> **Live tool (pick any):**
> - **Netlify:** [https://swot-ahp-analyzer.netlify.app](https://swot-ahp-analyzer.netlify.app)
> - **GitHub Pages:** [https://kharelg100.github.io/SWOT-AHP-Analyzer](https://kharelg100.github.io/SWOT-AHP-Analyzer/)
> - **Zenodo archive:** [https://doi.org/10.5281/zenodo.18991287](https://doi.org/10.5281/zenodo.18991287)

## Overview

This tool implements a complete SWOT–AHP (A'WOT) analytical pipeline for strategic planning in conservation, natural resource management, and related fields. It accepts Qualtrics-format pairwise comparison survey data and produces:

1. **AHP Consistency Diagnostics** — λmax, CI, RI, CR for all matrices (Saaty, 1977)
2. **Within-Category Priorities** — Eigenvector-derived local weights per SWOT category
3. **SWOT Category Weights** — Survey II–derived quadrant importance
4. **Global Factor Priorities** — Multiplicative synthesis of category × local weights
5. **Scenario Sensitivity** — Rank robustness across four quadrant-weight postures
6. **Bootstrap Uncertainty** — Respondent-level nonparametric resampling with rank acceptability
7. **TOWS Strategy Portfolio & SPI** — Strategy translation with uncertainty-quantified rankings

## Two Interfaces

| Interface | File | Requirements | Use case |
|-----------|------|-------------|----------|
| **Browser tool** | `index.html` | None (open in any browser) | Interactive analysis, no installation |
| **Python script** | `swot_ahp_analyzer.py` | numpy, pandas, matplotlib, openpyxl | Spyder, Jupyter, Google Colab |

## Quick Start

### Browser Tool (No Installation)

1. Open `index.html` in any modern browser, or visit the [live Netlify deployment](https://swot-ahp-analyzer.netlify.app)
2. Upload your Qualtrics CSVs (or click **▶ Run with demo data**)
3. Explore results across 7 analysis tabs
4. Download multi-sheet Excel workbook

All computation runs locally in your browser — **no data is uploaded anywhere**.

### Python Script

```bash
# Install dependencies
pip install numpy pandas matplotlib openpyxl

# Run with demo data
python swot_ahp_analyzer.py
```

For your own data, edit Section 1 of the script:
```python
USE_DEMO = False
SURVEY1_PATH = "your_survey1.csv"
SURVEY2_PATH = "your_survey2.csv"  # or None
```

**Google Colab:** Uncomment the Colab upload block in Section 1.

## Data Format

### Survey I — Within-Category Pairwise Comparisons

CSV with one row per respondent. Column naming convention:

```
S_S1_vs_S2, S_S1_vs_S3, S_S1_vs_S4, S_S1_vs_S5, S_S2_vs_S3, ...
W_W1_vs_W2, W_W1_vs_W3, ...
O_O1_vs_O2, ...
T_T1_vs_T2, ...
```

Values are integers 1–9 on the directional Saaty scale:

| Value | Saaty Ratio | Interpretation |
|-------|------------|----------------|
| 1 | 9 | Strong preference for left factor |
| 2 | 7 | Moderate-to-strong left |
| 3 | 5 | Moderate left |
| 4 | 3 | Slight left |
| 5 | 1 | Equal importance |
| 6 | 1/3 | Slight right |
| 7 | 1/5 | Moderate right |
| 8 | 1/7 | Moderate-to-strong right |
| 9 | 1/9 | Strong preference for right factor |

### Survey II — SWOT Category Comparisons

CSV with one row per respondent, 6 columns:

```
CAT_S_vs_W, CAT_S_vs_O, CAT_S_vs_T, CAT_W_vs_O, CAT_W_vs_T, CAT_O_vs_T
```

Same 1–9 scale. If omitted, equal category weights (25% each) are used.

### Sample Data

- `sample_survey1.csv` — 13 respondents, 40 within-category comparisons
- `sample_survey2.csv` — 12 respondents, 6 category-level comparisons

## Methodological References

### Core Methods

- **AHP:** Saaty, T. L. (1977). A scaling method for priorities in hierarchical structures. *Journal of Mathematical Psychology*, 15(3), 234–281. https://doi.org/10.1016/0022-2496(77)90033-5

- **AHP Group Decision Making:** Saaty, T. L. (1989). Group decision making and the AHP. In *The Analytic Hierarchy Process: Applications and Studies* (pp. 59–67). Springer.

- **Geometric Mean Aggregation:** Aczél, J., & Saaty, T. L. (1983). Procedures for synthesizing ratio judgements. *Journal of Mathematical Psychology*, 27(1), 93–102. https://doi.org/10.1016/0022-2496(83)90028-7

- **SWOT–AHP Hybrid (A'WOT):** Kurttila, M., Pesonen, M., Kangas, J., & Kajanus, M. (2000). Utilizing the analytic hierarchy process (AHP) in SWOT analysis — a hybrid method and its application to a forest-certification case. *Forest Policy and Economics*, 1(1), 41–52. https://doi.org/10.1016/S1389-9341(99)00004-0

- **MCDS in SWOT:** Kajanus, M., Leskinen, P., Kurttila, M., & Kangas, J. (2012). Making use of MCDS methods in SWOT analysis. *Forest Policy and Economics*, 20, 1–9. https://doi.org/10.1016/j.forpol.2012.03.005

- **TOWS Matrix:** Weihrich, H. (1982). The TOWS matrix — A tool for situational analysis. *Long Range Planning*, 15(2), 54–66. https://doi.org/10.1016/0024-6301(82)90120-0

### Uncertainty & Robustness

- **Bootstrap:** Efron, B., & Tibshirani, R. J. (1993). *An Introduction to the Bootstrap*. Chapman & Hall/CRC.

- **AHP Sensitivity:** Tóth, W., Vacik, H., Panagopoulos, T., & Varga, A. (2018). Sensitivity analysis and evaluation of forest management strategies with the AHP. *International Journal of the Analytic Hierarchy Process*, 10(2), 160–178.

- **Rank Acceptability (SMAA):** Lahdelma, R., Hokkanen, J., & Salminen, P. (1998). SMAA — Stochastic multiobjective acceptability analysis. *European Journal of Operational Research*, 106(1), 137–143. https://doi.org/10.1016/S0377-2217(97)00163-X

- **Consistency Review:** Ishizaka, A., & Labib, A. (2011). Review of the main developments in the analytic hierarchy process. *Expert Systems with Applications*, 38(11), 14336–14345. https://doi.org/10.1016/j.eswa.2011.04.143

## Citation

If you use this tool in published research, please cite:

```bibtex
@software{kharel2025swotahp,
  author       = {Kharel, Gehendra},
  title        = {{SWOT–AHP \& TOWS Strategy Analyzer: Browser-based 
                   tool for robust strategic prioritization with 
                   bootstrap uncertainty quantification}},
  year         = {2026},
  publisher    = {Zenodo},
  doi          = {10.5281/zenodo.18991287},
  url          = {https://doi.org/10.5281/zenodo.18991287}
}
```

## Repository Structure

```
swot-ahp-analyzer/
├── index.html                 # Browser-based tool (self-contained)
├── swot_ahp_analyzer.py       # Python script (Spyder/Jupyter/Colab)
├── sample_survey1.csv          # Sample Survey I data (13 respondents)
├── sample_survey2.csv          # Sample Survey II data (12 respondents)
├── README.md                   # This file
├── LICENSE                     # MIT License
├── CITATION.cff                # Citation metadata
├── .zenodo.json                # Zenodo metadata
├── netlify.toml                # Netlify deployment config
└── .gitignore                  # Git ignore rules
```

## License

MIT License — see [LICENSE](LICENSE) for details.

## Author

**Dr. Gehendra Kharel**
Texas Christian University
[g.kharel@tcu.edu](mailto:g.kharel@tcu.edu)

© 2025–2026 Dr. Gehendra Kharel. All rights reserved.
