# The Southern Frontier Review

*Economic intelligence on Thailand's southern border provinces, 2021–2025*

**Live dashboard:** https://aifan22404-crypto.github.io/sbp-dashboard/

A descriptive, open-data study of the economy of Pattani, Yala, Narathiwat and four districts of Songkhla (Chana, Thepha, Saba Yoi, Na Thawi), built from official statistics of 19 public sources plus district population registers: 2,147 values across 113 indicators and 12 domains.

It asks four linked questions:

| Code | Question |
| --- | --- |
| P1 Opportunity Mapping | Which sectors are strengths, and which show room to grow? |
| P2 Development Priority | Which development gaps are largest relative to the country? |
| P3 Risk Factor Analysis | Which risk factors are associated with economic activity? |
| P4 Vulnerable Sectors | Which groups are most exposed to shocks? |

## Key findings

- **Raw materials, little processing.** Location quotients for agriculture, forestry and fishing are 2.69–4.14; for manufacturing 0.29–0.36. Pattani lands 7.35% of Thailand's marine catch (2024), yet manufacturing is 8.7% of its GPP against 24.2% nationally.
- **Poverty gap.** Pattani's poverty rate was 21–30% in 2021–2025, against 3–6% nationally.
- **Tourism keeps growing after COVID.** Narathiwat visitors rose from 431,003 (2022) to 1.24 million (2025), above the 2019 figure of 703,810; Yala hotel occupancy reached 80.9%. The 2019 comparison is tentative: the three border provinces now report 1.8–2.9× their 2019 visitors while Songkhla reports slightly fewer, which may reflect a change in estimation method.
- **Violence and new firms move in opposite directions once population is taken into account.** Across 37 districts, raw counts show no link (Spearman ρ = −0.06), but lethal events per 100,000 people and new firms per 10,000 people have ρ = −0.47 (approx. 95% CI −0.69 to −0.18). Controlling for population with a partial rank correlation gives −0.31 (CI −0.58 to 0.02). The association holds after controlling for population density (−0.52) but weakens to about −0.3 when districts are compared only within the same province, so it is an association, not an effect.
- **Rubber price swings.** RSS3 annual average fell from 127.96 baht/kg (2011) to 41.48 (2018); the monthly price has not reached 100 baht since May 2012.

The author, who is from the area, wrote eight prior expectations before seeing results; three were consistent with the data, four partly consistent, and one not supported (updated 20 Nov 2026 after adding district population and a budget-by-agency breakdown). See the dashboard's *ความเชื่อ vs ข้อมูล* section.

All results are descriptive associations, not causal estimates.

## Repository layout

| Folder | Contents |
| --- | --- |
| `docs/` | The website (GitHub Pages): `index.html` + `data.json`, no build step |
| `data/clean/` | 17 cleaned datasets (Excel): 14 core sources + district population, 2019 tourism baseline, budget by agency |
| `data/master_panel/` | Integrated long-format panel (`SBP_master_panel_long.csv`, 2,147 rows) |
| `data/reference/` | Source log (file names, URLs, download dates), data dictionary, red-flag register (26 data anomalies) |
| `scripts/` | Python scripts used for cleaning, panel building and analysis |
| `powerbi/` | Power BI report (`.pbix`), star-schema CSVs, DAX measures, theme, build guide (Thai) |
| `notes/` | Scope document, P1–P4 findings, red-flag notes (Thai) |

Raw downloads are not included. Every cleaned figure traces back to its original publisher through `data/reference/source_log.xlsx`.

## Data sources

Office of the National Economic and Social Development Council (NESDC) · National Statistical Office · Department of Business Development · Office of Agricultural Economics · Department of Fisheries · Fish Marketing Organization · Ministry of Tourism and Sports · Board of Investment · Provincial Commerce Offices · Provincial Administration Offices · Budget Bureau · Bank of Thailand Southern Office · Uppsala Conflict Data Program (UCDP GED) · Department of Provincial Administration (district population, via provincial GD Catalogs) · Rocket Media Lab (budget by agency, secondary) · OpenGISData-Thailand by chingchai (district land area, CC BY 4.0).

## Limitations

District analysis covers only 37 districts and population figures come from registration data (Su-ngai Kolok undercounted before 2024, see RF25); the 2019 tourism baseline may not be comparable (RF26); the budget-by-agency figures are a secondary compilation; the two incident series (official counts and UCDP) use different definitions and do not identify perpetrators; 2024 GPP is preliminary; BOI figures are approvals; juristic-person registrations miss informal firms; 2022 fruit output is excluded as a source anomaly. No fieldwork was done, and the work has not yet been peer- or faculty-reviewed.

## Author and AI use

Irfan Aroma — economics student, Al Qasimia University (UAE), from southern Thailand. The research questions, scope and prior expectations are the author's. Claude (Anthropic) was used to help collect and clean data, compute statistics, search the literature, and draft code and text.

## License

Code (`scripts/`, `docs/index.html`): MIT. Data remain the property of their original publishers and are redistributed here in cleaned, derived form for research and education, with attribution. Chart.js (`docs/vendor/`) is MIT-licensed.
