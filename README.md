# GCE 2026 Examination Results Analysis

A Power BI dashboard that analyses the 2026 GCE and TVEE examination results across four examinations, ten regions, 277 honor roll centres and 805 honor roll entries.

**Project file:** `GCE_analysis.pbix`
**Data year:** 2026 only

**Tool:**
- Python to scrape the data from the GCE Board Website
- Excel to combine the data into a common workbook
- Microsoft Power BI Desktop (data loaded and cleaned with Power Query, measures written in DAX)

---

## 1. Project Goal

Give stakeholders a clear view of:

1. How candidates performed nationally (registered, sat, absent, passed, failed).
2. How performance differs between regions.
3. Who the top performers (honor roll) are.
4. Which examination centres produced the most top performers.

## 2. Examinations Covered

| Examination | Registered | Sat | Passed | Pass rate |
|---|---|---|---|---|
| GCE Ordinary Level | 98,040 | 88,053 | 65,539 | 74.43% |
| GCE Advanced Level | 70,739 | 69,947 | 43,211 | 61.78% |
| TVEE Intermediate Level | 29,054 | 27,775 | 16,804 | 60.50% |
| TVEE Advanced Level | 10,908 | 10,804 | 7,875 | 72.89% |
| **Total** | **208,741** | **196,579** | **133,429** | 67.88% |



## 3. Dashboard Pages

| Page | What it shows | Filters |
|---|---|---|
| **Overview** | KPI cards (registered, sat, passed, failed, absent, % passed, % failed), "National Outcome Flow" funnel, "National outcome by Examination" chart, "Performance Summary" table | None |
| **Region** | Best region, highest and lowest pass rate, total candidates and passes, pass rate by region, successful candidates by region, absenteeism by region, "Regional Performance Trend" | Education, Exam type |
| **Honor Roll** | Total, male and female honor roll counts, 25 / 33 / 40-point counts, points distribution, honors by sex, ranked candidate table | Exam, Education, Points |
| **Centres** | Total centres, top centres by honors candidates, top centres by points | None |



## 4. Data Source

The data was scraped from the GCE board website
- Combined into an Excel workbook: `GCE_Results_combined.xlsx` (8 sheets).
- Loaded with Power Query using `Excel.Workbook(File.Contents(...))`.
- The path currently points to a folder on the author's computer (see "How to Refresh the Data").

## 5. Data Model

| Table | Content | Rows | Used in a visual |
|---|---|---|---|
| `Overall perfromance` | National totals for the 4 examinations | 4 | Yes |
| `Results_by_regions` | Results by region, exam type and education | 37 | Yes |
| `Honor Roll candidates` | Honor roll entries | 805 | Yes |
| `results-by-divisions` | Results by division | 174 | No |
| `Performance_by_schooltype` | Results by centre type | 20 | No |
| `Performance_by_Specalty` | TVEE results by specialty | 43 | No |
| `OL results_by_subjects` | Ordinary Level results by subject | 101 | No |
| `AL results_by_subjects` | Advanced Level results by subject | 126 | No |

### Power Query steps worth knowing

- **Overall perfromance:** adds `Absent` (Reg 2026 − Sat 2026) and `% Fail 2026`; removes the 2025 columns and the last row.
- **Honor Roll candidates:** renames `Reg`, `Sat`, `Pass` to `Papers Registered`, `Papers Sat`, `Papers Passed`; replaces "Advance Level" with "Advanced Level"; adds `CIN`, an automatic row number.
- **Results_by_regions:** removes `S/N`; adds `Absent.1` (Reg − Sat).


## 6. Key Definitions

- **Registered**: candidates entered for the examination.
- **Sat**: candidates who sat the examination.
- **Absent**: Registered − Sat.
- **Pass rate**: Passed ÷ Sat (absent candidates are not counted as fails).
- **Absenteeism rate**: Absent ÷ Registered × 100.
- **Group naming in the region and division tables**: Ordinary Level General = GCE Ordinary Level; Advance Level General = GCE Advanced Level; Ordinary Level Technical = TVEE Intermediate; Advance Level Technical = TVEE Advanced.

## 7. Main Measures (DAX)

| Measure | Table | Logic |
|---|---|---|
| Total Registered / Sat / Passed / Failed / Absent | Overall perfromance | `SUM` of the matching column |
| Percentage Passed / Failed | Overall perfromance | `AVERAGE` of the pass / fail % column |
| Absenteeism Rate | Overall perfromance | `Total Absent / Total Registered * 100` |
| Percentage Sat | Overall perfromance | `Total Sat / Total Registered * 100` |
| Total Regional Candidates / Passes | Results_by_regions | `SUM` of `Reg` / `Pass` |
| Regional Pass Rate | Results_by_regions | `DIVIDE(Total Regional Passes, SUM(Sat), 0) * 100` |
| Best-Performing Region | Results_by_regions | `TOPN` 1 region by Regional Pass Rate |
| Highest / Lowest Pass Rate | Results_by_regions | `MAXX` / `MINX` over regions |
| Total Honors Candidates | Honor Roll candidates | `COUNT` of `Candidate Name` |
| Male / Female Candidates | Honor Roll candidates | `DISTINCTCOUNT` of `CIN` by sex |
| 25 / 33 / 40 Points Honors | Honor Roll candidates | `DISTINCTCOUNT` of `CIN` where `Points` equals the value |
| Highest Points, Rank by Points | Honor Roll candidates | `MAX` of `Points`; `RANKX` (dense) |
| Total Centres | Honor Roll candidates | `DISTINCTCOUNT` of `Exam Centre Name` |
| Number of Specalties | Performance_by_Specalty | `DISTINCTCOUNT` of `SpecialtyCode` (not used in a visual) |

## 8. Headline Results Shown on the Dashboard

- 208,741 registered; 196,579 sat; 12,162 absent (5.8% of registered).
- 133,429 passed; 63,150 failed.
- Best region: Centre (70.40%). Lowest region: North (48.77%).
- Honor roll: 805 entries (492 female, 313 male) from 277 centres.
- Most honor roll candidates: Du Vaal Bilingual Comprehensive College Beedi–Douala (24). Most points: Saint Stephen's International College Yaoundé (759).

## 9. How to Open the Project

1. Install the latest [Power BI Desktop](https://powerbi.microsoft.com/desktop/).
2. Open `GCE_analysis.pbix`.
3. Use the page tabs (Overview, Region, Honor Roll, Centres, Key Findings) to move around.

## 10. How to Refresh the Data

1. Save `GCE_Results_combined.xlsx` in a folder you can access.
2. In Power BI Desktop, go to **Home → Transform data → Data source settings**.
3. Select the Excel source, click **Change Source**, and browse to the new file location.
4. Click **Close & Apply**, then **Refresh**.

If the sheet names or column names in the workbook change, the Power Query steps will fail and need updating.

## 11. Known Issues and Suggested Improvements

**High priority**


1. **Five tables feed no visual** (divisions, centre type, specialties, two subject tables). Add pages for them.
2. **Some tables do not match the national totals.** TVEE Intermediate in the centre-type table (23,818 registered vs 29,054) and in the specialty table (22,913 registered, 8,624 passed vs 29,054 and 16,804); TVEE Advanced in the specialty table (10,871 registered, 7,183 passed vs 10,908 and 7,875). Regions and divisions match exactly.

**Medium priority**

5. Remove the meaningless `S/N` relationships and the inactive `Papers Passed → Pass` relationship.
6. `CIN` is a row number, not a candidate ID, so "candidates" are entries: 805 entries but 800 individuals.
7. The 25, 33 and 40-point cards mix scores from different examinations.
8. "Regional Performance Trend" is a line chart across regions with no natural order; use bars.
9. Region cards combine all four examinations unless a filter is applied.
10. "Fail" has two meanings: in the subject tables it equals Registered − Pass (includes absent); elsewhere it equals Sat − Pass.
11. Only 2026 data is loaded; the 2025 columns were removed in Power Query, so no year-on-year view is possible.
12. The data source is a local file path.

**Housekeeping**

- Fix spelling in table names (`perfromance`, `Specalty`) and values (`Ordinar level`, `Litoral`).
- Standardise labels (`Advance` / `Advanced`, `Northwest` / `North West`, `west`).
- Remove the duplicate `Absent.1` column and name `Column11` in the divisions table.
- Add a centre ID and a centre-to-region link.
- The Honor Roll lists candidate names; confirm sharing is permitted before wider distribution.

## 12. Limitations

- One year of data, so no trend analysis.
- Candidate-level data exists only for the honor roll.
- Associations in the data are not proof of cause.
- Small groups give unstable rates and should be read with care.

## 13. Suggested Repository Structure

```
GCE-results-analysis/
├── GCE_analysis.pbix
├── data/
│   └── GCE_Results_combined.xlsx
__ notebooks
     ____Scraper
├── docs/
│   └── GCE_2026_Stakeholder_Report
└── README.md
```

## 14. Author and License

- **Author:** _Desmond Yembi_
- **Contact:** _www.linkedin.com/in/desmond-yembi-4a897621b_
- **License:** _a"all rights reserved"_
# Cameroon-GCE-2026-Result-analysis
