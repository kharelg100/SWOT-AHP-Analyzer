"""
SWOT–AHP & TOWS Strategy Analyzer
===================================
Robust SWOT–AHP prioritization with bootstrap uncertainty,
scenario sensitivity, and SPI strategy ranking.

Compatible with: Spyder (Anaconda), Jupyter Notebook, Google Colab.

Author: Dr. Gehendra Kharel, Texas Christian University
Contact: g.kharel@tcu.edu
© 2025–2026 Dr. Gehendra Kharel. All rights reserved.

Methodological References:
--------------------------
AHP (Analytic Hierarchy Process):
  Saaty, T. L. (1977). A scaling method for priorities in hierarchical
    structures. Journal of Mathematical Psychology, 15(3), 234–281.
    https://doi.org/10.1016/0022-2496(77)90033-5
  Saaty, T. L. (1989). Group decision making and the AHP. In The Analytic
    Hierarchy Process: Applications and Studies (pp. 59–67). Springer.
  Saaty, T. L. (2008). Decision making with the analytic hierarchy process.
    International Journal of Services Sciences, 1(1), 83–98.
    https://doi.org/10.1504/IJSSCI.2008.017590

Geometric Mean Aggregation:
  Aczél, J., & Saaty, T. L. (1983). Procedures for synthesizing ratio
    judgements. Journal of Mathematical Psychology, 27(1), 93–102.
    https://doi.org/10.1016/0022-2496(83)90028-7

SWOT–AHP Hybrid (A'WOT):
  Kurttila, M., Pesonen, M., Kangas, J., & Kajanus, M. (2000). Utilizing
    the analytic hierarchy process (AHP) in SWOT analysis — a hybrid method
    and its application to a forest-certification case. Forest Policy and
    Economics, 1(1), 41–52. https://doi.org/10.1016/S1389-9341(99)00004-0
  Kajanus, M., Leskinen, P., Kurttila, M., & Kangas, J. (2012). Making use
    of MCDS methods in SWOT analysis — Lessons learnt in strategic natural
    resources management. Forest Policy and Economics, 20, 1–9.
    https://doi.org/10.1016/j.forpol.2012.03.005

Consistency Diagnostics:
  Ishizaka, A., & Labib, A. (2011). Review of the main developments in the
    analytic hierarchy process. Expert Systems with Applications, 38(11),
    14336–14345. https://doi.org/10.1016/j.eswa.2011.04.143

TOWS Strategy Matrix:
  Weihrich, H. (1982). The TOWS matrix — A tool for situational analysis.
    Long Range Planning, 15(2), 54–66.
    https://doi.org/10.1016/0024-6301(82)90120-0

Bootstrap Resampling:
  Efron, B., & Tibshirani, R. J. (1993). An Introduction to the Bootstrap.
    Chapman & Hall/CRC.
  Tóth, W., Vacik, H., Panagopoulos, T., & Varga, A. (2018). Sensitivity
    analysis and evaluation of forest management strategies with the AHP.
    International Journal of the Analytic Hierarchy Process, 10(2), 160–178.

Rank Acceptability:
  Lahdelma, R., Hokkanen, J., & Salminen, P. (1998). SMAA — Stochastic
    multiobjective acceptability analysis. European Journal of Operational
    Research, 106(1), 137–143.
    https://doi.org/10.1016/S0377-2217(97)00163-X

SWOT Analysis (general):
  Gürel, E., & Tat, M. (2017). SWOT analysis: A theoretical review. Journal
    of International Social Research, 10(51), 994–1006.
  Helms, M. M., & Nixon, J. (2010). Exploring SWOT analysis — where are we
    now? Journal of Strategy and Management, 3(3), 215–251.

Usage:
------
1. Spyder / Anaconda:
   - Open this file in Spyder and click Run (F5)
   - Modify the DATA SOURCE section (Section 1) to point to your CSVs
   - Plots display in the Plots pane; tables print to Console

2. Jupyter Notebook:
   - Copy cells into a notebook or run: %run swot_ahp_analyzer.py
   - Plots render inline with %matplotlib inline

3. Google Colab:
   - Upload this file or paste into cells
   - Use files.upload() to load CSVs (see Section 1)

Requirements:
    pip install numpy pandas matplotlib seaborn openpyxl
"""

# ═══════════════════════════════════════════════════════════════════════
# IMPORTS
# ═══════════════════════════════════════════════════════════════════════
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import LinearSegmentedColormap
import warnings
import os
from itertools import combinations
from io import BytesIO

warnings.filterwarnings("ignore")

# Detect environment
IN_COLAB = False
try:
    import google.colab
    IN_COLAB = True
except ImportError:
    pass

IN_JUPYTER = False
try:
    get_ipython()
    IN_JUPYTER = True
except NameError:
    pass

if IN_JUPYTER or IN_COLAB:
    try:
        get_ipython().run_line_magic('matplotlib', 'inline')
    except:
        pass

print("=" * 65)
print("  SWOT–AHP & TOWS Strategy Analyzer")
print("  Dr. Gehendra Kharel · Texas Christian University")
print("  g.kharel@tcu.edu")
print("=" * 65)
print(f"  Environment: {'Google Colab' if IN_COLAB else 'Jupyter' if IN_JUPYTER else 'Spyder/Script'}")
print("=" * 65)


# ═══════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════

# --- Color palette (TCU theme) ---
TCU = "#4d1979"
TCU2 = "#6b2fa0"
TCUL = "#f3edf9"
CAT_COLORS = {"S": "#2ecc71", "W": "#e74c3c", "O": "#3498db", "T": "#e67e22"}
CAT_NAMES = {"S": "Strengths", "W": "Weaknesses", "O": "Opportunities", "T": "Threats"}
CATEGORIES = ["S", "W", "O", "T"]

# --- Saaty scale mapping (Qualtrics 1-9 directional → ratio) ---
# 9-point ratio scale per Saaty (1977): {9, 7, 5, 3, 1, 1/3, 1/5, 1/7, 1/9}
SAATY_MAP = {1: 9, 2: 7, 3: 5, 4: 3, 5: 1, 6: 1/3, 7: 1/5, 8: 1/7, 9: 1/9}

# --- Random indices for consistency check (Saaty, 1977) ---
RI_TABLE = {1: 0, 2: 0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24,
            7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49}

# --- Default factor labels ---
FACTOR_LABELS = {
    "S1": "Content development", "S2": "Research programs",
    "S3": "Hands-on engagement", "S4": "Emotional connection",
    "S5": "Personnel expertise",
    "W1": "Resource scarcity", "W2": "Evaluation difficulty",
    "W3": "Personnel gaps", "W4": "Adult education gaps",
    "W5": "Oversimplification",
    "O1": "Virtual/social media", "O2": "School partnerships",
    "O3": "Local community programs", "O4": "Tourist engagement",
    "O5": "Policy influence",
    "T1": "Short-term experiences", "T2": "Funding instability",
    "T3": "Climate change", "T4": "Education effectiveness doubts",
    "T5": "Political opposition",
}

# --- Default TOWS strategies ---
TOWS_STRATEGIES = {
    "SO-1 Local stewardship pathway": ["S3", "S4", "O3", "O2"],
    "SO-2 Scalable program kits": ["S1", "S5", "O3", "O4"],
    "SO-3 Policy-relevant education": ["S3", "S4", "O5", "O3"],
    "ST-1 Localize climate action": ["S3", "S4", "T3", "O3"],
    "ST-2 Evidence pipeline": ["S2", "S1", "T4", "W2"],
    "ST-3 Navigate opposition": ["S5", "S3", "S4", "T5"],
    "WO-1 Co-delivery partnerships": ["W1", "W3", "O2", "O3"],
    "WO-2 Low-cost post-visit engagement": ["W1", "W2", "O1", "T1"],
    "WO-3 Evaluation partnerships": ["W2", "O2", "O3", "T4"],
    "WT-1 Financial resilience": ["W1", "T2", "T3"],
    "WT-2 Adult learning redesign": ["W4", "T1", "S4"],
    "WT-3 Accurate messaging": ["W5", "T5", "T4"],
}

# --- Bootstrap settings ---
N_BOOTSTRAP = 5000  # Change as needed: 1000–20000
RANDOM_SEED = 42

# --- Output ---
OUTPUT_EXCEL = "SWOT_AHP_Results.xlsx"  # Set to None to skip Excel export


# ═══════════════════════════════════════════════════════════════════════
# SECTION 1: DATA SOURCE
# ═══════════════════════════════════════════════════════════════════════
# Option A: Load from CSV files (uncomment and set paths)
# SURVEY1_PATH = "survey1.csv"
# SURVEY2_PATH = "survey2.csv"  # Set to None if not available

# Option B: Google Colab upload (uncomment for Colab)
# if IN_COLAB:
#     from google.colab import files
#     print("Upload Survey I CSV:")
#     uploaded = files.upload()
#     SURVEY1_PATH = list(uploaded.keys())[0]
#     print("Upload Survey II CSV (or press Cancel to skip):")
#     try:
#         uploaded2 = files.upload()
#         SURVEY2_PATH = list(uploaded2.keys())[0]
#     except:
#         SURVEY2_PATH = None

# Option C: Use generated demo data (DEFAULT)
USE_DEMO = True


# ═══════════════════════════════════════════════════════════════════════
# AHP ENGINE
# ═══════════════════════════════════════════════════════════════════════

def get_pairs(n):
    """Generate all pairwise combinations for n items."""
    return list(combinations(range(n), 2))


def qualtrics_to_saaty(val):
    """Convert Qualtrics ordinal value to Saaty ratio."""
    try:
        v = int(round(float(val)))
        return SAATY_MAP.get(v, np.nan)
    except (ValueError, TypeError):
        return np.nan


def geometric_mean_matrix(respondent_ratios, n):
    """Build reciprocal PCM from respondent-level ratio lists.
    Group judgments aggregated via geometric mean, the unique synthesis
    procedure preserving reciprocity (Aczél & Saaty, 1983)."""
    mat = np.ones((n, n))
    for i, j in get_pairs(n):
        vals = [r[(i, j)] for r in respondent_ratios
                if (i, j) in r and np.isfinite(r[(i, j)]) and r[(i, j)] > 0]
        if vals:
            gm = np.exp(np.mean(np.log(vals)))
            mat[i, j] = gm
            mat[j, i] = 1.0 / gm
    return mat


def ahp_eigenvector_weights(matrix):
    """Compute AHP priority vector via principal right eigenvector (Saaty, 1977)."""
    eigvals, eigvecs = np.linalg.eig(matrix)
    max_idx = np.argmax(eigvals.real)
    principal = eigvecs[:, max_idx].real
    weights = np.abs(principal / principal.sum())
    return weights


def ahp_consistency(matrix):
    """Compute λmax, CI, RI, CR (Saaty, 1977; Ishizaka & Labib, 2011).
    CR < 0.10 indicates acceptable judgmental consistency."""
    n = matrix.shape[0]
    w = ahp_eigenvector_weights(matrix)
    Aw = matrix @ w
    lambda_max = np.mean(Aw / w)
    CI = (lambda_max - n) / (n - 1) if n > 1 else 0
    RI = RI_TABLE.get(n, 1.49)
    CR = CI / RI if RI > 0 else 0
    return {"lambda_max": lambda_max, "CI": CI, "RI": RI, "CR": CR}


def parse_survey1(df, categories, factor_ids):
    """Parse Survey I CSV into respondent-level ratio dictionaries."""
    resp_by_cat = {}
    for cat in categories:
        fids = factor_ids[cat]
        ps = get_pairs(len(fids))
        respondents = []
        for _, row in df.iterrows():
            ratios = {}
            valid = True
            for i, j in ps:
                col = f"{cat}_{fids[i]}_vs_{fids[j]}"
                if col not in row:
                    valid = False
                    break
                s = qualtrics_to_saaty(row[col])
                if np.isnan(s):
                    valid = False
                    break
                ratios[(i, j)] = s
            if valid:
                respondents.append(ratios)
        resp_by_cat[cat] = respondents
    return resp_by_cat


def parse_survey2(df, categories):
    """Parse Survey II CSV into respondent-level ratio dictionaries."""
    ps = get_pairs(len(categories))
    respondents = []
    for _, row in df.iterrows():
        ratios = {}
        valid = True
        for i, j in ps:
            col = f"CAT_{categories[i]}_vs_{categories[j]}"
            if col not in row:
                valid = False
                break
            s = qualtrics_to_saaty(row[col])
            if np.isnan(s):
                valid = False
                break
            ratios[(i, j)] = s
        if valid:
            respondents.append(ratios)
    return respondents


def compute_global_weights(cat_weights, local_weights, categories):
    """Global weight = category weight × local factor weight.
    Standard AHP multiplicative synthesis (Saaty, 1977; Kurttila et al., 2000)."""
    gw = {}
    for cat in categories:
        cw = cat_weights.get(cat, 0.25)
        for fid, lw in local_weights.get(cat, {}).items():
            gw[fid] = cw * lw
    return gw


def scenario_weights(scenario, survey_cat_weights, categories):
    """Get category weights for a given scenario.
    Multi-scenario sensitivity analysis follows SWOT–AHP best practice
    (Kajanus et al., 2012; Kurttila et al., 2000)."""
    if scenario == "A":
        return survey_cat_weights
    elif scenario == "B":
        return {c: 0.25 for c in categories}
    elif scenario == "C":  # Threat-forward
        w = {c: 0.1 for c in categories}
        w["T"] = 0.7
        total = sum(w.values())
        return {c: v / total for c, v in w.items()}
    elif scenario == "D":  # Opportunity-forward
        w = {c: 0.1 for c in categories}
        w["O"] = 0.7
        total = sum(w.values())
        return {c: v / total for c, v in w.items()}
    return survey_cat_weights


# ═══════════════════════════════════════════════════════════════════════
# BOOTSTRAP ENGINE
# ═══════════════════════════════════════════════════════════════════════

def bootstrap_ahp(resp_by_cat, factor_ids, cat_resp, cat_weights,
                   categories, B=5000, seed=42):
    """Respondent-level nonparametric bootstrap for global weights.
    Resamples respondents with replacement within each survey
    (Efron & Tibshirani, 1993; Tóth et al., 2018)."""
    rng = np.random.RandomState(seed)
    results = []

    for b in range(B):
        if b % 1000 == 0 and b > 0:
            print(f"    Bootstrap: {b:,}/{B:,} replicates...", end="\r")

        # Resample within-category
        local_w = {}
        for cat in categories:
            rl = resp_by_cat[cat]
            n = len(factor_ids[cat])
            if not rl:
                continue
            idx = rng.choice(len(rl), size=len(rl), replace=True)
            sampled = [rl[i] for i in idx]
            mat = geometric_mean_matrix(sampled, n)
            w = ahp_eigenvector_weights(mat)
            local_w[cat] = {factor_ids[cat][k]: w[k] for k in range(n)}

        # Resample category-level
        if cat_resp and len(cat_resp) > 0:
            idx2 = rng.choice(len(cat_resp), size=len(cat_resp), replace=True)
            sampled2 = [cat_resp[i] for i in idx2]
            mat2 = geometric_mean_matrix(sampled2, len(categories))
            cw_vec = ahp_eigenvector_weights(mat2)
            cw = {categories[i]: cw_vec[i] for i in range(len(categories))}
        else:
            cw = cat_weights

        results.append(compute_global_weights(cw, local_w, categories))

    print(f"    Bootstrap: {B:,}/{B:,} replicates — done.     ")
    return results


def rank_acceptability(boot_results, factor_ids):
    """Compute rank acceptability from bootstrap, following the SMAA
    framework (Lahdelma et al., 1998)."""
    B = len(boot_results)
    n = len(factor_ids)
    weights = np.array([[gw.get(f, 0) for f in factor_ids] for gw in boot_results])
    ranks = np.zeros_like(weights)
    for b in range(B):
        ranks[b] = n - np.argsort(np.argsort(weights[b]))

    summary = {}
    for i, fid in enumerate(factor_ids):
        w_col = weights[:, i]
        r_col = ranks[:, i]
        summary[fid] = {
            "median": np.median(w_col),
            "ci_lo": np.percentile(w_col, 2.5),
            "ci_hi": np.percentile(w_col, 97.5),
            "p_rank1": np.mean(r_col == 1),
            "p_top3": np.mean(r_col <= 3),
            "p_top5": np.mean(r_col <= 5),
            "expected_rank": np.mean(r_col),
        }
    return summary


def compute_spi(strategies, global_weights):
    """Strategy Priority Index = sum of global weights of mapped factors.
    Strategies derived from TOWS matrix (Weihrich, 1982); SPI provides
    a transparent, additive ranking metric for the strategy portfolio
    (Kajanus et al., 2012)."""
    return {name: sum(global_weights.get(f, 0) for f in facs)
            for name, facs in strategies.items()}


def bootstrap_spi(boot_results, strategies, factor_ids):
    """Bootstrap SPI distributions."""
    B = len(boot_results)
    strat_names = list(strategies.keys())
    spi_mat = np.array([[sum(gw.get(f, 0) for f in strategies[s])
                         for s in strat_names] for gw in boot_results])
    n_s = len(strat_names)
    ranks = np.zeros_like(spi_mat)
    for b in range(B):
        ranks[b] = n_s - np.argsort(np.argsort(spi_mat[b]))

    summary = {}
    for si, name in enumerate(strat_names):
        col = spi_mat[:, si]
        r_col = ranks[:, si]
        summary[name] = {
            "median": np.median(col),
            "ci_lo": np.percentile(col, 2.5),
            "ci_hi": np.percentile(col, 97.5),
            "p_top1": np.mean(r_col == 1),
            "p_top3": np.mean(r_col <= 3),
            "p_top5": np.mean(r_col <= 5),
        }
    return summary


# ═══════════════════════════════════════════════════════════════════════
# DEMO DATA GENERATOR
# ═══════════════════════════════════════════════════════════════════════

def generate_demo_data():
    """Generate sample Qualtrics-style survey data."""
    rng = np.random.RandomState(2024)
    factor_ids = {cat: [f"{cat}{i}" for i in range(1, 6)] for cat in CATEGORIES}

    # Survey I: 13 respondents
    rows = []
    for r in range(13):
        row = {"respondent_id": r + 1}
        for cat in CATEGORIES:
            fids = factor_ids[cat]
            for i, j in get_pairs(len(fids)):
                col = f"{cat}_{fids[i]}_vs_{fids[j]}"
                row[col] = rng.choice([1, 2, 3, 4, 5, 6, 7, 8, 9],
                                       p=[0.05, 0.08, 0.12, 0.15, 0.20,
                                          0.15, 0.12, 0.08, 0.05])
        rows.append(row)
    df_s1 = pd.DataFrame(rows)

    # Survey II: 12 respondents
    rows2 = []
    for r in range(12):
        row = {}
        for i, j in get_pairs(4):
            col = f"CAT_{CATEGORIES[i]}_vs_{CATEGORIES[j]}"
            row[col] = rng.choice([1, 2, 3, 4, 5, 6, 7, 8, 9],
                                   p=[0.08, 0.10, 0.12, 0.15, 0.15,
                                      0.12, 0.10, 0.10, 0.08])
        rows2.append(row)
    df_s2 = pd.DataFrame(rows2)

    return df_s1, df_s2, factor_ids


# ═══════════════════════════════════════════════════════════════════════
# VISUALIZATION
# ═══════════════════════════════════════════════════════════════════════

def set_style():
    """Set publication-quality matplotlib style."""
    plt.rcParams.update({
        "figure.facecolor": "white",
        "axes.facecolor": "white",
        "font.family": "sans-serif",
        "font.sans-serif": ["Segoe UI", "Helvetica Neue", "Arial"],
        "font.size": 10,
        "axes.titlesize": 13,
        "axes.titleweight": "bold",
        "axes.labelsize": 11,
        "axes.edgecolor": "#c8cdd5",
        "axes.grid": True,
        "grid.alpha": 0.3,
        "grid.color": "#c8cdd5",
    })

set_style()


def plot_local_weights(local_weights, factor_ids, n_resp):
    """Plot within-category AHP weights (2×2 grid)."""
    fig, axes = plt.subplots(2, 2, figsize=(12, 9))
    fig.suptitle("Within-Category AHP Priorities", fontsize=16,
                 fontweight="bold", color=TCU, y=0.98)

    for idx, cat in enumerate(CATEGORIES):
        ax = axes[idx // 2][idx % 2]
        fids = factor_ids[cat]
        items = sorted([(f, local_weights[cat][f]) for f in fids],
                       key=lambda x: x[1])
        labels = [FACTOR_LABELS.get(f, f) for f, _ in items]
        vals = [w * 100 for _, w in items]
        color = CAT_COLORS[cat]

        bars = ax.barh(labels, vals, color=color, alpha=0.85, height=0.6,
                       edgecolor="white", linewidth=0.5)
        for bar, v in zip(bars, vals):
            ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height() / 2,
                    f"{v:.1f}%", va="center", fontsize=9, fontweight="bold")

        ax.set_title(f"{CAT_NAMES[cat]} (n={n_resp[cat]})", color=color,
                     fontweight="bold")
        ax.set_xlim(0, max(vals) * 1.25)
        ax.set_xlabel("Weight (%)")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig("01_local_weights.png", dpi=200, bbox_inches="tight")
    plt.show()


def plot_global_weights(global_w, top_n=20):
    """Plot global factor priorities."""
    sorted_items = sorted(global_w.items(), key=lambda x: x[1])[-top_n:]
    fids = [x[0] for x in sorted_items]
    vals = [x[1] * 100 for x in sorted_items]
    colors = [CAT_COLORS.get(f[0], "#8c95a5") for f in fids]
    labels = [FACTOR_LABELS.get(f, f) for f in fids]

    fig, ax = plt.subplots(figsize=(10, 8))
    bars = ax.barh(labels, vals, color=colors, alpha=0.85, height=0.6,
                   edgecolor="white", linewidth=0.5)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height() / 2,
                f"{v:.1f}%", va="center", fontsize=9, fontweight="bold")

    ax.set_title("Global Factor Priorities (Scenario A)", color=TCU,
                 fontsize=14, fontweight="bold")
    ax.set_xlabel("Global Weight (%)")
    ax.set_xlim(0, max(vals) * 1.2)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # Legend
    patches = [mpatches.Patch(color=CAT_COLORS[c], label=CAT_NAMES[c])
               for c in CATEGORIES]
    ax.legend(handles=patches, loc="lower right", framealpha=0.9)

    plt.tight_layout()
    plt.savefig("02_global_weights.png", dpi=200, bbox_inches="tight")
    plt.show()


def plot_scenario_heatmap(scenario_ranks, all_fids):
    """Plot rank robustness heatmap across scenarios."""
    scenarios = ["A", "B", "C", "D"]
    sc_labels = ["A: Stakeholder", "B: Equal", "C: Threat-fwd", "D: Opportunity-fwd"]

    # Sort by Scenario A rank
    sorted_fids = sorted(all_fids, key=lambda f: scenario_ranks[f].get("A", 20))
    n_f = len(sorted_fids)

    data = np.array([[scenario_ranks[f][sc] for sc in scenarios]
                     for f in sorted_fids])
    labels = [FACTOR_LABELS.get(f, f) for f in sorted_fids]

    cmap = LinearSegmentedColormap.from_list("rank",
        ["#2ecc71", "#f1c40f", "#e67e22", "#e74c3c"], N=n_f)

    fig, ax = plt.subplots(figsize=(8, 10))
    im = ax.imshow(data, cmap=cmap, aspect="auto", vmin=1, vmax=n_f)

    ax.set_xticks(range(len(scenarios)))
    ax.set_xticklabels(sc_labels, fontweight="bold")
    ax.set_yticks(range(n_f))
    ax.set_yticklabels(labels, fontsize=9)

    for i in range(n_f):
        for j in range(len(scenarios)):
            ax.text(j, i, str(data[i, j]), ha="center", va="center",
                    fontsize=9, fontweight="bold", color="white")

    ax.set_title("Rank Robustness Across Scenarios\n(lower rank = higher priority)",
                 color=TCU, fontsize=13, fontweight="bold")
    plt.colorbar(im, ax=ax, label="Rank", shrink=0.6)
    plt.tight_layout()
    plt.savefig("03_scenario_heatmap.png", dpi=200, bbox_inches="tight")
    plt.show()


def plot_bootstrap_forest(boot_summary, top_n=10):
    """Forest plot of bootstrap median + 95% CI."""
    sorted_items = sorted(boot_summary.items(),
                          key=lambda x: x[1]["median"], reverse=True)[:top_n]

    fig, ax = plt.subplots(figsize=(10, 6))
    y_pos = list(range(len(sorted_items)))[::-1]

    for yi, (fid, s) in zip(y_pos, sorted_items):
        color = CAT_COLORS.get(fid[0], "#8c95a5")
        label = FACTOR_LABELS.get(fid, fid)
        ax.plot([s["ci_lo"] * 100, s["ci_hi"] * 100], [yi, yi],
                color=color, linewidth=3, solid_capstyle="round")
        ax.plot(s["median"] * 100, yi, "o", color=color, markersize=8,
                markeredgecolor="white", markeredgewidth=1.5, zorder=5)
        ax.text(-0.5, yi, label, ha="right", va="center", fontsize=9,
                fontweight="bold", transform=ax.get_yaxis_transform())

    ax.set_yticks(y_pos)
    ax.set_yticklabels([""] * len(y_pos))
    ax.set_xlabel("Global Weight (%)")
    ax.set_title(f"Bootstrap Global Weights — Top {top_n} (median + 95% CI)",
                 color=TCU, fontsize=13, fontweight="bold")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)

    plt.tight_layout()
    plt.savefig("04_bootstrap_forest.png", dpi=200, bbox_inches="tight")
    plt.show()


def plot_rank_acceptability(boot_summary, top_n=10):
    """Stacked bar of P(rank=1), P(top-3), P(top-5)."""
    sorted_items = sorted(boot_summary.items(),
                          key=lambda x: x[1]["p_top5"], reverse=True)[:top_n]

    fids = [x[0] for x in sorted_items]
    labels = [FACTOR_LABELS.get(f, f) for f in fids]
    p1 = [x[1]["p_rank1"] * 100 for x in sorted_items]
    p3 = [(x[1]["p_top3"] - x[1]["p_rank1"]) * 100 for x in sorted_items]
    p5 = [(x[1]["p_top5"] - x[1]["p_top3"]) * 100 for x in sorted_items]

    fig, ax = plt.subplots(figsize=(10, 5))
    x = range(len(fids))
    ax.bar(x, p1, color="#27ae60", label="P(rank=1)")
    ax.bar(x, p3, bottom=p1, color="#2980b9", label="P(top-3) − P(1)")
    ax.bar(x, p5, bottom=[a + b for a, b in zip(p1, p3)],
           color="#f39c12", label="P(top-5) − P(top-3)")

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Probability (%)")
    ax.set_title("Rank Acceptability", color=TCU, fontsize=13, fontweight="bold")
    ax.legend(framealpha=0.9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout()
    plt.savefig("05_rank_acceptability.png", dpi=200, bbox_inches="tight")
    plt.show()


def plot_spi_forest(spi_summary):
    """Forest plot of SPI with bootstrap CIs."""
    sorted_items = sorted(spi_summary.items(),
                          key=lambda x: x[1]["median"], reverse=True)

    fig, ax = plt.subplots(figsize=(10, 7))
    y_pos = list(range(len(sorted_items)))[::-1]

    for yi, (name, s) in zip(y_pos, sorted_items):
        color = "#e74c3c" if s["median"] > 0.3 else "#e67e22" if s["median"] > 0.2 else "#8c95a5"
        ax.plot([s["ci_lo"], s["ci_hi"]], [yi, yi],
                color="#2d3748", linewidth=3, solid_capstyle="round")
        ax.plot(s["median"], yi, "o", color=color, markersize=9,
                markeredgecolor="white", markeredgewidth=1.5, zorder=5)

    ax.set_yticks(y_pos)
    ax.set_yticklabels([x[0] for x in sorted_items][::-1], fontsize=9)
    ax.set_xlabel("Strategy Priority Index")
    ax.set_title("TOWS Strategy Rankings — SPI (median + 95% CI)",
                 color=TCU, fontsize=13, fontweight="bold")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout()
    plt.savefig("06_spi_forest.png", dpi=200, bbox_inches="tight")
    plt.show()


# ═══════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════

def export_excel(local_w, consis, cat_w, global_w, sc_ranks,
                 boot_sum, spi_sum, strategies, filename):
    """Export all results to multi-sheet Excel workbook."""
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        # Local weights
        rows = []
        for cat in CATEGORIES:
            for fid in sorted(local_w[cat], key=lambda f: local_w[cat][f], reverse=True):
                rows.append({"Category": cat, "Factor": fid,
                             "Description": FACTOR_LABELS.get(fid, fid),
                             "Local Weight (%)": round(local_w[cat][fid] * 100, 2)})
        pd.DataFrame(rows).to_excel(writer, "Local Weights", index=False)

        # Consistency
        rows_c = [{"Matrix": CAT_NAMES.get(k, k), **{kk: round(vv, 4) for kk, vv in v.items()},
                    "Status": "Acceptable" if v["CR"] < 0.10 else "Exceeds 0.10"}
                   for k, v in consis.items()]
        pd.DataFrame(rows_c).to_excel(writer, "Consistency", index=False)

        # Category weights
        pd.DataFrame([{"Category": CAT_NAMES[c], "Weight (%)": round(cat_w[c] * 100, 2)}
                      for c in CATEGORIES]).to_excel(writer, "Category Weights", index=False)

        # Global weights
        gw_sorted = sorted(global_w.items(), key=lambda x: x[1], reverse=True)
        pd.DataFrame([{"Rank": i + 1, "Factor": f, "Category": f[0],
                        "Description": FACTOR_LABELS.get(f, f),
                        "Global Weight (%)": round(w * 100, 2)}
                       for i, (f, w) in enumerate(gw_sorted)]
                      ).to_excel(writer, "Global Weights", index=False)

        # Scenario ranks
        rows_s = []
        for fid in sorted(sc_ranks, key=lambda f: sc_ranks[f].get("A", 20)):
            r = sc_ranks[fid]
            rows_s.append({"Factor": fid, "Description": FACTOR_LABELS.get(fid, fid),
                           **{f"Rank_{s}": r[s] for s in ["A", "B", "C", "D"]},
                           "Range": max(r.values()) - min(r.values())})
        pd.DataFrame(rows_s).to_excel(writer, "Scenario Ranks", index=False)

        # Bootstrap
        rows_b = []
        for fid, s in sorted(boot_sum.items(), key=lambda x: x[1]["median"], reverse=True):
            rows_b.append({"Factor": fid, "Description": FACTOR_LABELS.get(fid, fid),
                           "Median (%)": round(s["median"] * 100, 2),
                           "CI_Low (%)": round(s["ci_lo"] * 100, 2),
                           "CI_High (%)": round(s["ci_hi"] * 100, 2),
                           "P(rank=1)": round(s["p_rank1"], 3),
                           "P(top-3)": round(s["p_top3"], 3),
                           "P(top-5)": round(s["p_top5"], 3),
                           "E[rank]": round(s["expected_rank"], 2)})
        pd.DataFrame(rows_b).to_excel(writer, "Bootstrap Factors", index=False)

        # SPI
        rows_t = []
        for name, s in sorted(spi_sum.items(), key=lambda x: x[1]["median"], reverse=True):
            rows_t.append({"Strategy": name,
                           "Factors": ", ".join(strategies.get(name, [])),
                           "SPI Median": round(s["median"], 3),
                           "CI_Low": round(s["ci_lo"], 3),
                           "CI_High": round(s["ci_hi"], 3),
                           "P(top-1)": round(s["p_top1"], 3),
                           "P(top-3)": round(s["p_top3"], 3),
                           "P(top-5)": round(s["p_top5"], 3)})
        pd.DataFrame(rows_t).to_excel(writer, "SPI Rankings", index=False)

    print(f"\n  ✓ Results exported to: {filename}")


# ═══════════════════════════════════════════════════════════════════════
# MAIN ANALYSIS
# ═══════════════════════════════════════════════════════════════════════

def main():
    print("\n" + "─" * 65)
    print("  STEP 1: Loading data")
    print("─" * 65)

    factor_ids = {cat: [f"{cat}{i}" for i in range(1, 6)] for cat in CATEGORIES}
    all_fids = [f for cat in CATEGORIES for f in factor_ids[cat]]

    if USE_DEMO:
        df_s1, df_s2, factor_ids = generate_demo_data()
        print("  Using demo data (13 Survey I + 12 Survey II respondents)")
    else:
        df_s1 = pd.read_csv(SURVEY1_PATH)
        print(f"  Survey I: {SURVEY1_PATH} ({len(df_s1)} rows)")
        if SURVEY2_PATH and os.path.exists(SURVEY2_PATH):
            df_s2 = pd.read_csv(SURVEY2_PATH)
            print(f"  Survey II: {SURVEY2_PATH} ({len(df_s2)} rows)")
        else:
            df_s2 = None
            print("  Survey II: not provided (using equal category weights)")

    all_fids = [f for cat in CATEGORIES for f in factor_ids[cat]]

    # ── Parse respondents ──
    print("\n" + "─" * 65)
    print("  STEP 2: AHP computation & consistency")
    print("─" * 65)

    resp_by_cat = parse_survey1(df_s1, CATEGORIES, factor_ids)
    cat_resp = parse_survey2(df_s2, CATEGORIES) if df_s2 is not None else []
    n_resp = {cat: len(resp_by_cat[cat]) for cat in CATEGORIES}

    for cat in CATEGORIES:
        print(f"    {CAT_NAMES[cat]}: {n_resp[cat]} valid respondents")
    if cat_resp:
        print(f"    Categories: {len(cat_resp)} valid respondents")

    # ── Within-category weights + consistency ──
    local_weights = {}
    consistency_results = {}
    for cat in CATEGORIES:
        fids = factor_ids[cat]
        mat = geometric_mean_matrix(resp_by_cat[cat], len(fids))
        w = ahp_eigenvector_weights(mat)
        c = ahp_consistency(mat)
        local_weights[cat] = {fids[i]: w[i] for i in range(len(fids))}
        consistency_results[cat] = c
        status = "✓" if c["CR"] < 0.10 else "⚠"
        print(f"    {CAT_NAMES[cat]:15s}  CR = {c['CR']:.4f}  {status}")

    # ── Category weights ──
    if cat_resp:
        cat_mat = geometric_mean_matrix(cat_resp, 4)
        cat_w_vec = ahp_eigenvector_weights(cat_mat)
        cat_weights_dict = {CATEGORIES[i]: cat_w_vec[i] for i in range(4)}
        consistency_results["Categories"] = ahp_consistency(cat_mat)
        print(f"    {'Categories':15s}  CR = {consistency_results['Categories']['CR']:.4f}")
    else:
        cat_weights_dict = {c: 0.25 for c in CATEGORIES}
        print("    Categories: equal weights (0.25 each)")

    print("\n  Category weights:")
    for c in CATEGORIES:
        print(f"    {CAT_NAMES[c]:15s}  {cat_weights_dict[c]*100:.1f}%")

    # ── Global weights ──
    print("\n" + "─" * 65)
    print("  STEP 3: Global priorities & scenario sensitivity")
    print("─" * 65)

    global_w = compute_global_weights(cat_weights_dict, local_weights, CATEGORIES)

    print("\n  Top 10 global factors (Scenario A):")
    for i, (f, w) in enumerate(sorted(global_w.items(),
                                       key=lambda x: x[1], reverse=True)[:10], 1):
        print(f"    {i:2d}. {f}  {FACTOR_LABELS.get(f,''):30s}  {w*100:.2f}%")

    # ── Scenario sensitivity ──
    sc_names = ["A", "B", "C", "D"]
    sc_ranks = {f: {} for f in all_fids}
    for sc in sc_names:
        sw = scenario_weights(sc, cat_weights_dict, CATEGORIES)
        gw = compute_global_weights(sw, local_weights, CATEGORIES)
        ranked = sorted(all_fids, key=lambda f: gw.get(f, 0), reverse=True)
        for rank, fid in enumerate(ranked, 1):
            sc_ranks[fid][sc] = rank

    print("\n  Scenario rank ranges (top 5):")
    top5_a = sorted(all_fids, key=lambda f: sc_ranks[f]["A"])[:5]
    for f in top5_a:
        r = sc_ranks[f]
        rng = max(r.values()) - min(r.values())
        print(f"    {f}  A={r['A']:2d}  B={r['B']:2d}  C={r['C']:2d}  D={r['D']:2d}  range={rng}")

    # ── Bootstrap ──
    print("\n" + "─" * 65)
    print(f"  STEP 4: Bootstrap uncertainty ({N_BOOTSTRAP:,} replicates)")
    print("─" * 65)

    boot_results = bootstrap_ahp(resp_by_cat, factor_ids, cat_resp,
                                  cat_weights_dict, CATEGORIES,
                                  B=N_BOOTSTRAP, seed=RANDOM_SEED)
    boot_summary = rank_acceptability(boot_results, all_fids)

    print("\n  Top 10 bootstrap summary:")
    print(f"  {'Factor':<6s} {'Median%':>8s} {'95% CI':>16s} {'P(r=1)':>7s} {'P(t3)':>6s} {'P(t5)':>6s} {'E[r]':>6s}")
    print("  " + "─" * 58)
    for fid, s in sorted(boot_summary.items(),
                         key=lambda x: x[1]["median"], reverse=True)[:10]:
        print(f"  {fid:<6s} {s['median']*100:8.2f} [{s['ci_lo']*100:6.2f}–{s['ci_hi']*100:6.2f}]"
              f" {s['p_rank1']:7.3f} {s['p_top3']:6.3f} {s['p_top5']:6.3f} {s['expected_rank']:6.2f}")

    # ── SPI ──
    print("\n" + "─" * 65)
    print("  STEP 5: TOWS strategies & SPI")
    print("─" * 65)

    spi_summary = bootstrap_spi(boot_results, TOWS_STRATEGIES, all_fids)

    print(f"\n  {'Strategy':<38s} {'SPI':>6s} {'95% CI':>16s} {'P(t1)':>6s} {'P(t3)':>6s}")
    print("  " + "─" * 72)
    for name, s in sorted(spi_summary.items(),
                          key=lambda x: x[1]["median"], reverse=True):
        print(f"  {name:<38s} {s['median']:6.3f} [{s['ci_lo']:.3f}–{s['ci_hi']:.3f}]"
              f" {s['p_top1']:6.3f} {s['p_top3']:6.3f}")

    # ── Plots ──
    print("\n" + "─" * 65)
    print("  STEP 6: Generating plots")
    print("─" * 65)

    plot_local_weights(local_weights, factor_ids, n_resp)
    plot_global_weights(global_w)
    plot_scenario_heatmap(sc_ranks, all_fids)
    plot_bootstrap_forest(boot_summary)
    plot_rank_acceptability(boot_summary)
    plot_spi_forest(spi_summary)

    print("  ✓ 6 plots saved as PNG files")

    # ── Excel export ──
    if OUTPUT_EXCEL:
        print("\n" + "─" * 65)
        print("  STEP 7: Excel export")
        print("─" * 65)
        export_excel(local_weights, consistency_results, cat_weights_dict,
                     global_w, sc_ranks, boot_summary, spi_summary,
                     TOWS_STRATEGIES, OUTPUT_EXCEL)

    print("\n" + "═" * 65)
    print("  ✓ Analysis complete!")
    print("═" * 65)

    # Return all results for interactive use
    return {
        "factor_ids": factor_ids,
        "local_weights": local_weights,
        "consistency": consistency_results,
        "category_weights": cat_weights_dict,
        "global_weights": global_w,
        "scenario_ranks": sc_ranks,
        "bootstrap_summary": boot_summary,
        "spi_summary": spi_summary,
        "boot_results": boot_results,
    }


# ═══════════════════════════════════════════════════════════════════════
# RUN
# ═══════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    results = main()
