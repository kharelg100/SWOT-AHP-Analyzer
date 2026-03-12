# Deployment Guide: GitHub + Zenodo + Netlify

## Step-by-Step Setup

---

## 1. GitHub Repository

### 1a. Create the Repository

1. Go to [github.com/new](https://github.com/new)
2. Settings:
   - **Repository name:** `swot-ahp-analyzer`
   - **Description:** `SWOT–AHP & TOWS Strategy Analyzer with bootstrap uncertainty quantification`
   - **Visibility:** Public
   - **Initialize:** Do NOT add README (we have one)
3. Click **Create repository**

### 1b. Push Files

Open terminal and run:

```bash
cd swot-ahp-repo

# Initialize and push
git init
git add .
git commit -m "Initial release: SWOT-AHP & TOWS Strategy Analyzer v1.0.0"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/swot-ahp-analyzer.git
git push -u origin main
```

Replace `YOUR_USERNAME` with your GitHub username (e.g., `gkharel`).

### 1c. Create a Release

1. Go to your repo → **Releases** → **Create a new release**
2. **Tag:** `v1.0.0`
3. **Title:** `SWOT–AHP & TOWS Strategy Analyzer v1.0.0`
4. **Description:**
   ```
   Initial release of the SWOT–AHP & TOWS Strategy Analyzer.

   Features:
   - Browser-based tool (single HTML file, no installation)
   - Python script (Spyder, Jupyter, Google Colab)
   - AHP eigenvector prioritization with Saaty consistency diagnostics
   - Geometric mean group judgment aggregation
   - Multi-scenario sensitivity analysis (4 postures)
   - Respondent-level bootstrap uncertainty (1,000–20,000 replicates)
   - Rank acceptability probabilities and percentile CIs
   - TOWS strategy matrix with Strategy Priority Index (SPI)
   - Multi-sheet Excel export
   - Sample data included
   ```
5. Click **Publish release**

### 1d. Enable GitHub Pages (Optional Alternative to Netlify)

If you prefer GitHub Pages over Netlify:

1. Go to repo → **Settings** → **Pages**
2. Source: **Deploy from a branch**
3. Branch: `main`, folder: `/ (root)`
4. Click **Save**
5. Your tool will be live at: `https://YOUR_USERNAME.github.io/swot-ahp-analyzer/`

---

## 2. Zenodo (DOI Assignment)

### 2a. Link GitHub to Zenodo

1. Go to [zenodo.org](https://zenodo.org) and log in with your GitHub account
2. Go to **Settings** → **GitHub** ([zenodo.org/account/settings/github/](https://zenodo.org/account/settings/github/))
3. Find `swot-ahp-analyzer` in your repository list
4. Toggle the switch to **ON** (enable Zenodo integration)

### 2b. Trigger DOI Creation

1. Go back to GitHub and create the release (Step 1c above, if not already done)
2. Zenodo will automatically:
   - Archive the release
   - Assign a DOI
   - Create a landing page
3. The `.zenodo.json` file in the repo provides metadata (title, description, keywords, references) that Zenodo reads automatically

### 2c. Update DOI in Repository Files

Once Zenodo assigns your DOI (e.g., `10.5281/zenodo.1234567`):

1. Update `README.md` — replace `XXXXXXX` with your Zenodo record number:
   ```
   [![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.1234567.svg)](https://doi.org/10.5281/zenodo.1234567)
   ```

2. Update `CITATION.cff`:
   ```yaml
   doi: "10.5281/zenodo.1234567"
   ```

3. Update the BibTeX block in `README.md`:
   ```bibtex
   doi = {10.5281/zenodo.1234567},
   url = {https://doi.org/10.5281/zenodo.1234567}
   ```

4. Add your ORCID to `CITATION.cff` and `.zenodo.json` if available:
   ```yaml
   orcid: "https://orcid.org/0000-0000-0000-0000"
   ```

5. Commit and push:
   ```bash
   git add -A
   git commit -m "Add Zenodo DOI"
   git push
   ```

### 2d. Verify

- Visit your Zenodo page: `https://zenodo.org/records/1234567`
- Confirm metadata, keywords, references, and license are correct
- The DOI badge in your README should now link to the Zenodo record

---

## 3. Netlify (Live Web Deployment)

### 3a. Connect Repository

1. Go to [app.netlify.com](https://app.netlify.com) and log in with GitHub
2. Click **Add new site** → **Import an existing project**
3. Select **GitHub** → authorize Netlify if prompted
4. Choose your `swot-ahp-analyzer` repository

### 3b. Configure Build

Netlify will auto-detect `netlify.toml`. Verify these settings:

| Setting | Value |
|---------|-------|
| Branch to deploy | `main` |
| Build command | *(leave blank — no build needed)* |
| Publish directory | `.` |

5. Click **Deploy site**

### 3c. Custom Domain (Optional)

1. After deployment, go to **Site configuration** → **Domain management**
2. Click **Add custom domain**
3. Options:
   - Free Netlify subdomain: `swot-ahp-analyzer.netlify.app` (rename via Site configuration → Site name)
   - Custom domain: Add a CNAME record pointing to your Netlify site

### 3d. Update README

Replace the Netlify badge placeholder in `README.md`:

1. Go to **Site configuration** → **Status badges**
2. Copy the Markdown badge
3. Replace the placeholder in README:
   ```markdown
   [![Netlify Status](https://api.netlify.com/api/v1/badges/YOUR-SITE-ID/deploy-status)](https://app.netlify.com/sites/swot-ahp-analyzer/deploys)
   ```

### 3e. Verify

- Visit your Netlify URL (e.g., `https://swot-ahp-analyzer.netlify.app`)
- Test: click "Run with demo data" to confirm everything works
- Test: upload sample CSVs to confirm file handling works
- Test: download Excel workbook

---

## 4. Post-Setup Checklist

After completing all three platforms:

- [ ] GitHub repo is public with all files
- [ ] GitHub release `v1.0.0` is published
- [ ] Zenodo DOI is assigned and metadata is correct
- [ ] DOI badge in README links to Zenodo
- [ ] Netlify site is live and functional
- [ ] Netlify badge in README shows deploy status
- [ ] `CITATION.cff` has correct DOI and ORCID
- [ ] BibTeX citation block in README has correct DOI
- [ ] Live URL works: demo data runs, file upload works, Excel downloads

## 5. Future Releases

When you update the tool:

```bash
# Make changes, then:
git add -A
git commit -m "Description of changes"
git push

# Create new release on GitHub (e.g., v1.1.0)
# Zenodo will automatically archive and update the DOI
# Netlify will automatically redeploy
```

---

## File Summary

| File | Purpose | Platform |
|------|---------|----------|
| `index.html` | Browser tool (self-contained) | Netlify serves this |
| `swot_ahp_analyzer.py` | Python script | Download from GitHub/Zenodo |
| `sample_survey1.csv` | Sample Survey I | Included in all |
| `sample_survey2.csv` | Sample Survey II | Included in all |
| `README.md` | Documentation + badges | GitHub |
| `LICENSE` | MIT License | GitHub, Zenodo |
| `CITATION.cff` | Citation metadata | GitHub "Cite this repository" |
| `.zenodo.json` | Zenodo metadata | Zenodo auto-reads |
| `netlify.toml` | Deployment config + headers | Netlify auto-reads |
| `.gitignore` | Git ignore rules | GitHub |
| `DEPLOYMENT.md` | This guide | Reference |
