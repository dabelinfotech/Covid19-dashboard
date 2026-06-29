# COVID-19 Mortality Analysis Dashboard

An image-based data dashboard analysing COVID-19 death cases globally from **January 2020 to April 2021**, built with vanilla HTML & CSS — charts exported directly from Microsoft Excel.

## Live Demo

Open `index.html` in any browser or view the hosted version on GitHub Pages.

## Features

- **12 Excel charts** across 3 analysis sheets, embedded as high-quality PNG images
- KPI cards: total global deaths, CFR, worst-affected continent, worst month
- Monthly trend, continent vs country breakdowns, CFR league table, death-share analysis
- Responsive dark-navy layout — no external libraries, no CDN

## Dashboard Sections

### Chart 1 — Full Overview (6 charts)
| # | Chart |
|---|---|
| 1 | Monthly Trend — New Deaths per Month |
| 2 | Continent: Total Cases vs Total Deaths |
| 3 | Top 10 Countries: Total Cases vs Total Deaths |
| 4 | Total Deaths by Continent |
| 5 | Continent Share of Global Deaths |
| 6 | Case Fatality Rate (CFR) — Top 10 Countries |

### Chart 2 — Trend & Country Analysis (3 charts)
| # | Chart |
|---|---|
| 1 | Monthly Trend — New Deaths per Month |
| 2 | Top 10 Countries: Total Cases vs Total Deaths |
| 3 | Total Deaths by Continent |

### Chart 3 — CFR & Continent Deep Dive (3 charts)
| # | Chart |
|---|---|
| 1 | Case Fatality Rate (CFR) — Top 10 Countries |
| 2 | Continent: Total Cases vs Total Deaths |
| 3 | Continent Share of Global Deaths (Donut) |

## Dataset

Source: [Our World in Data — COVID-19 Deaths](https://ourworldindata.org/covid-deaths)  
Rows: 85,172 | Columns: 59 | Period: Jan 2020 – Apr 2021

## Key Findings

| Metric | Value |
|---|---|
| Global Deaths | 3,180,223 |
| Global Cases | 151,376,060 |
| Global CFR | 2.10% |
| Deadliest Continent | Europe (1,016,750 deaths, 31.97%) |
| Deadliest Country | United States (576,232 deaths) |
| Highest CFR | Mexico (9.25%) |
| Worst Month | January 2021 (411,021 deaths) |

## Tech Stack

| Layer | Technology |
|---|---|
| Markup | HTML5 |
| Styling | CSS3 (custom properties, flexbox, grid) |
| Charts | Microsoft Excel (exported as PNG) |
| Data source | Microsoft Excel (.xlsx) via Power Query |

## Power Query (M Language) — Key Transformations

```m
let
    Source        = Excel.Workbook(File.Contents("CovidDeaths.xlsx"), true, true),
    Sheet         = Source{[Name="Sheet1"]}[Data],
    Promoted      = Table.PromoteHeaders(Sheet),
    Typed         = Table.TransformColumnTypes(Promoted, {
                        {"total_cases",  type number},
                        {"total_deaths", type number},
                        {"date",         type date}
                    }),
    NoNulls       = Table.SelectRows(Typed, each [continent] <> null),
    LatestPerLoc  = Table.Group(NoNulls, {"location","continent"}, {
                        {"total_cases",  each List.Max([total_cases]),  type number},
                        {"total_deaths", each List.Max([total_deaths]), type number}
                    }),
    ByContinent   = Table.Group(LatestPerLoc, {"continent"}, {
                        {"total_cases",  each List.Sum([total_cases]),  type number},
                        {"total_deaths", each List.Sum([total_deaths]), type number}
                    }),
    AddCFR        = Table.AddColumn(ByContinent, "cfr_pct",
                        each [total_deaths] / [total_cases] * 100, type number),
    Sorted        = Table.Sort(AddCFR, {{"total_deaths", Order.Descending}})
in
    Sorted
```

## Author

**dabeltech** — [github.com/dabelinfotech](https://github.com/dabelinfotech)
