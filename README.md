# The Southern Frontier Review

*Economic intelligence on Thailand's southern border provinces, 2021–2025*

**Live dashboard:** https://aifan22404-crypto.github.io/sbp-dashboard/

A descriptive, open-data study of the economy of Pattani, Yala, Narathiwat and four districts of Songkhla (Chana, Thepha, Saba Yoi, Na Thawi), built from official statistics of 19 public sources: 2,147 values across 113 indicators and 12 domains.

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
- **Tourism keeps growing after COVID.** Narathiwat visitors rose from 431,003 (2022) to 1.24 million (2025); Yala hotel occupancy reached 80.9%. No pre-COVID baseline yet.
- **No clear district-level link between violence and new firms.** Spearman ρ = −0.06 across 37 districts (approx. 95% CI −0.37 to 0.27), not population-adjusted.
- **Rubber price swings.** RSS3 annual average fell from 127.96 baht/kg (2011) to 41.48 (2018); the monthly price has not reached 100 baht since May 2012.

The author, who is from the area, wrote eight prior expectations before seeing results; three were consistent with the data, two partly, two not supported, one untestable. See the dashboard's *ความเชื่อ vs ข้อมูล* section.

All results are descriptive associations, not causal estimates.

## Repository layout

| Folder | Contents |
| --- | --- |
| `docs/` | The website (GitHub Pages): `index.html` + `data.json`, no build step |
| `data/clean/` | 14 cleaned datasets (Excel), one per source/topic |
| `data/master_panel/` | Integrated long-format panel (`SBP_master_panel_long.csv`, 2,147 rows) |
| `data/reference/` | Source log (file names, URLs, download dates), data dictionary, red-flag register (24 data anomalies) |
| `scripts/` | Python scripts used for cleaning, panel building and analysis |
| `powerbi/` | Power BI report (`.pbix`), star-schema CSVs, DAX measures, theme, build guide (Thai) |
| `notes/` | Scope document, P1–P4 findings, red-flag notes (Thai) |

Raw downloads are not included. Every cleaned figure traces back to its original publisher through `data/reference/source_log.xlsx`.

## Data sources

Office of the National Economic and Social Development Council (NESDC) · National Statistical Office · Department of Business Development · Office of Agricultural Economics · Department of Fisheries · Fish Marketing Organization · Ministry of Tourism and Sports · Board of Investment · Provincial Commerce Offices · Provincial Administration Offices · Budget Bureau · Bank of Thailand Southern Office · Uppsala Conflict Data Program (UCDP GED).

## Limitations

District counts are not population-adjusted; tourism data start in 2021; the two incident series (official counts and UCDP) use different definitions and do not identify perpetrators; 2024 GPP is preliminary; BOI figures are approvals; juristic-person registrations miss informal firms; 2022 fruit output is excluded as a source anomaly. No fieldwork was done, and the work has not yet been peer- or faculty-reviewed.

## Author and AI use

Irfan Aroma — economics student, Al Qasimia University (UAE), from southern Thailand. The research questions, scope and prior expectations are the author's. Claude (Anthropic) was used to help collect and clean data, compute statistics, search the literature, and draft code and text.

## License

Code (`scripts/`, `docs/index.html`): MIT. Data remain the property of their original publishers and are redistributed here in cleaned, derived form for research and education, with attribution. Chart.js (`docs/vendor/`) is MIT-licensed.
