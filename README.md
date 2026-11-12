# Tideline: sea-level rise and coastal protection

21CEO310T Global Warming and Climate Change, LLT / CLA-2 software project.

**Live website:** <https://tideline-sea-level.netlify.app>

**Notebook:** [![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Akinc5/tideline-sea-level/blob/main/Tideline_Colab.ipynb)

**Project report:** [`report/Tideline_Project_Report.docx`](report/Tideline_Project_Report.docx)

**Topic:** Sea-level rise and coastal protection. Model a rising water level and compare protective approaches such as
mangroves, raised structures and embankments.

Tideline forecasts global sea level to 2100 with explainable machine learning trained on real tide-gauge and satellite
data. It then uses the forecast to flood a model coastal town and compares four defences: mangroves, an embankment,
raised houses and a green-grey hybrid.

## Folder layout
| Path | What it is |
|---|---|
| `Tideline_Colab.ipynb` | The main notebook. It has already been run, so every table and figure is visible |
| `web/index.html` | The 3D parallax website, self-contained, opens in any browser |
| `web/template.html`, `web/build_web.py` | Website source and build script |
| `data/` | Snapshot of the source files (used if a server is offline) |
| `outputs/` | Figures `fig01–fig08`, metric CSVs, projections, defence comparison, `web_data.json` |

## Data (real, official sources)
| Source | Variable | Coverage |
|---|---|---|
| CSIRO (Church & White) | Global mean sea level from tide gauges, mm ± 1σ | 1880–2019 |
| NOAA Laboratory for Satellite Altimetry | Global mean sea level, TOPEX → Sentinel-6 | 1993–2025 |
| PSMSL station 43 | Mumbai (Apollo Bandar) relative sea level | 1878–2024 |
| NASA GISS GISTEMP v4 | Global temperature anomaly | 1880–2026 |

## ML results (test period 2000–2024, never seen in training)
| Model | MAE (mm) | MSE | RMSE (mm) | R² |
|---|---|---|---|---|
| A · Linear Regression (year, T) | 23.04 | 628.9 | 25.08 | −0.058 |
| A · SVR | 23.33 | 636.6 | 25.23 | −0.071 |
| A · Random Forest | 41.37 | 2307.3 | 48.04 | −2.882 |
| **B · Linear Regression (semi-empirical)** | **2.07** | **7.4** | **2.72** | **0.988** |
| B · SVR | 2.33 | 10.5 | 3.24 | 0.982 |
| B · Random Forest | 41.38 | 2307.5 | 48.04 | −2.883 |
| **C · Hybrid + SVR** | **1.96** | **6.4** | **2.53** | **0.989** |
| **C · Hybrid + Random Forest** | **1.97** | **6.3** | **2.51** | **0.989** |

* **Explainable:** Linear Regression on the semi-empirical features is the Vermeer–Rahmstorf law. It learns
  a = 2.1 mm/yr per °C, and its present-day rate of 4.0 mm/yr matches the satellites' 3.7 mm/yr.
* **2100 projections:** SSP1-2.6 +0.43 m, SSP2-4.5 +0.50 m, SSP5-8.5 +0.62 m (IPCC AR6 medians 0.44 / 0.56 / 0.77 m).

## Defence comparison (SSP2-4.5, 1-in-100 storm)
| Defence | Homes flooded 2050 | Homes flooded 2100 | Benefit-cost ratio | Outcome |
|---|---|---|---|---|
| No defence | 398 | 467 | — | floods every decade |
| Mangrove belt (400 m) | 256 | 321 | 29.5 | survives; drowns only under SSP5-8.5 |
| Embankment (4.5 m) | 0 | 467 | 1.2 | holds, then fails all at once |
| Raised houses (1 m) | 398 | 467 | 2.0 | still wet, far less damage |
| Hybrid (300 m mangroves + 3.8 m wall) | 0 | 341 | 1.4 | most robust |

Costs, surge statistics and the town are illustrative; the assumptions table is in the notebook and on the website.

## How to run
**Colab:**
1. Click the **Open in Colab** badge above.
2. Choose **Runtime → Run all**. The data downloads live.
3. If a server is down, upload `data/` to `/content/data`.

**Website:** open `web/index.html` in a browser. After re-running the notebook, refresh the site's numbers:

```bash
python web/build_web.py
```

**Redeploy to Netlify** (after rebuilding): copy `web/index.html` into `site/`, then run from the `tideline` folder:

```bash
netlify deploy --prod --dir site
```
