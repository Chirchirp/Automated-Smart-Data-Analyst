# SokoData — Smart Data Explorer (Streamlit App)

A fully self-contained, API-free data analysis app.  
Upload a CSV or Excel file and get instant intelligent analytics.

## Features

- **Data quality report** — missing values, duplicates, outlier detection, completeness score
- **Descriptive statistics** — numeric & categorical summaries with skew, kurtosis, IQR
- **Auto-charts** — histograms, bar charts, pie/donut charts, scatter plots, box plots, violin plots
- **Time-series detection** — auto-detects date columns, plots trends with rolling averages
- **Correlation matrix** — interactive heatmap + strongest-pair table
- **Custom chart builder** — pick any column combination and chart type interactively
- **Plain-English insights** — generated entirely from statistics, zero API required
- **Downloadable reports** — missing values CSV, numeric summary CSV, de-duplicated CSV, full HTML report

---

## Run locally

```bash
# 1. Clone or copy this folder
cd sokodata_analytics

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch
streamlit run app.py
```

App opens at http://localhost:8501

---

## Deploy to Streamlit Community Cloud (free)

1. Push this folder to a **public GitHub repository**.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **New app**.
4. Select your repo, branch (`main`), and set main file path to `app.py`.
5. Click **Deploy**. Your app URL will be:  
   `https://<your-slug>.streamlit.app`

### Link it to your website

In `index.html` and `analytics_section.html`, replace every occurrence of:
```
https://YOUR-APP-NAME.streamlit.app
```
with your real Streamlit URL.

Also update `js/main.js` bottom section with the same URL.

---

## File structure

```
sokodata_analytics/
├── app.py                         ← Main Streamlit application
├── requirements.txt               ← Python dependencies
├── .streamlit/
│   └── config.toml                ← SokoData brand theme for Streamlit
├── analytics_section.html         ← HTML block to paste into index.html
├── analytics_styles_additions.css ← CSS to append to css/styles.css
└── README.md
```

---

## Integrating with the website

### Step 1 — Add the nav link
In `index.html`, add to both the desktop and mobile nav menus:
```html
<a href="#analytics" class="nav-link hover:text-primary transition-colors">
  <i class="fas fa-chart-pie mr-2"></i>Analytics
</a>
```

### Step 2 — Paste the section
Copy the entire contents of `analytics_section.html` and paste it into `index.html`
just before the closing `</main>` tag (after the `#testimonials` section).

### Step 3 — Add the styles
Append all contents of `analytics_styles_additions.css` to `css/styles.css`.

### Step 4 — Update the URL
Find and replace `https://YOUR-APP-NAME.streamlit.app` in `index.html` with your real URL.

---

## No API key needed

All insights are generated purely from pandas statistics — skewness, IQR, correlation,
cardinality, missing ratios, etc. Nothing is sent to any external AI service.
