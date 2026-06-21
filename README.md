# COVID-19 Mortality Analysis Dashboard

An interactive data dashboard analysing COVID-19 death cases globally from **January 2020 to April 2021**, built with vanilla HTML, CSS, and Canvas 2D — no external libraries or CDN dependencies.

## Live Demo

Open `index.html` in any browser or view the hosted version on GitHub Pages.

## Features

- **7 interactive charts** rendered with the native Canvas 2D API
- Deaths by continent (bar chart)
- Top 10 countries by total deaths (horizontal bar)
- Continent share of global deaths (donut chart)
- Case Fatality Rate (CFR) by country (horizontal bar)
- Total cases vs total deaths by continent (dual-axis grouped bar)
- Top 10 countries — total cases vs total deaths (dual-axis grouped bar)
- Monthly global death toll Jan 2020 – Apr 2021 (area + line chart)
- KPI cards: total deaths, CFR, worst-affected country
- Case Fatality Rate league table
- Key insights panel
- Downloadable Excel workbook (`COVID19_Dashboard_Data.xlsx`) with 8 sheets and embedded charts

## Dataset

Source: [Our World in Data — COVID-19 Deaths](https://ourworldindata.org/covid-deaths)  
Rows: 85,172 | Columns: 59 | Period: 24 Feb 2020 – 30 Apr 2021

## Tech Stack

| Layer | Technology |
|---|---|
| Markup | HTML5 |
| Styling | CSS3 (custom properties, flexbox, grid) |
| Charts | Canvas 2D API (no Chart.js, no D3) |
| Data export | Blob + createObjectURL → .xlsx |
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

## Excel Workbook Sheets

| Sheet | Contents |
|---|---|
| Summary | Global totals and KPIs |
| By Continent | Deaths per continent |
| Top 10 Countries | Highest-death countries |
| Monthly Trend | Month-by-month global deaths + chart |
| Continent Analysis | Cases + Deaths per continent + dual-axis chart |
| Country Analysis | Cases + Deaths per country + dual-axis chart |
| Global Death Share | Continent % share + pie chart + bar chart |
| Country CFR | CFR per country + horizontal bar chart |

## Author

**dabeltech** — [github.com/dabelinfotech](https://github.com/dabelinfotech)
