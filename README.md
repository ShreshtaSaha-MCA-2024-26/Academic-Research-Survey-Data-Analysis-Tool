<img width="1017" height="590" alt="image" src="https://github.com/user-attachments/assets/97a3c7ca-49fd-4e96-9869-c0317210ccce" />

# Academic-Research-Survey-Data-Analysis-Tool

A complete Excel-based survey analytics project covering **1,002 students**, **15 teachers**, **5 teaching methods**, and the impact of **gamification** on learning outcomes. The workbook takes raw survey exports through cleaning, merging, KPI analysis, pivot tables, and an interactive dashboard — all built with native Excel formulas (no external scripts).

## Project Overview

The project analyzes a student survey to understand:
- How effective and engaging 5 teaching methods are (Lectures, Case Studies, Group Projects, Experiments, Online Tutorials)
- How gamification affects Engagement, Motivation, Participation, and Learning Outcomes
- How 15 teachers compare across experience, subject, and teaching method
- Which student segments (age, gender, performance tier) respond best to which methods

**Data flow:**

```
Raw Data → Cleaned Data → Master_Data (merged + enriched) → Analysis / KPI Analysis → Pivot Tables & Charts → Dashboard
```

## Workbook Structure

| Sheet | Description |
|---|---|
| **Home** | Navigation hub with links to every sheet |
| **Instructions** | Full documentation: data cleaning steps, formulas used, KPI & metric glossary |
| **Dashboard** | Interactive dashboard — KPI cards, charts, and slicers |
| **Overall Analysis** | Executive KPI summary and a live Top-5 Teacher ranking |
| **KPI Analysis** | Detailed KPI breakdown across Student, Teaching, Engagement, Gamification, and Teacher categories |
| **Master_Data** | The single source of truth — 1,002 students × 35 columns, merged and enriched |
| **Teacher_Master** | 15 teachers with method, subject, experience, and performance tier |
| **Summery Data** | Gamification vs. non-gamification summary across 4 metrics |
| **Raw Student Data** | Original survey export (contains whitespace issues) |
| **Clean Student Data** | Whitespace-trimmed version used downstream |
| **Raw Gamification Data** | Original gamification survey export |
| **Clean Gamification Data** | Whitespace-trimmed version used downstream |

### Pivot Tables & Charts

| Pivot Sheet | What It Shows |
|---|---|
| Gamification & Non-Gamification | Average impact % by metric |
| Teacher Count | Teacher count by experience band × teaching method |
| Teacher ID | Teacher count by teaching method |
| Gender Split | Student count by gender |
| Gamification Lift (Performance) | Average gamification lift by performance tier |
| Students per Teacher | Student demand pressure by teaching method |
| Effectiveness & Engagement | Effectiveness & engagement by age band × gender |
| Gamification Lift (Age) | Average gamification lift by age band |
| Gamification Lift (Gender) | Average gamification lift by gender |
| Effectiveness Distribution | Histogram of overall effectiveness ratings |
| Teaching Method (Effectiveness) | Best-method popularity by effectiveness |
| Teaching Method (Engagement) | Best-method popularity by engagement |
| Gamification Lift | Average gamification lift by performance tier |

Five slicers connect these pivots: **Gender**, **Age Band**, and **Performance Tier** (drive all Master_Data-based pivots/charts together); **Teaching Method** (drives Teacher_Master-based pivots); **Metric** (drives the Summary-based pivot). A slicer can only connect pivots built on the same source table.

## Data Cleaning

| Step | Action Taken | Why |
|---|---|---|
| Raw Student Data → Clean Student Data | Trimmed leading/trailing whitespace from Student Name (e.g. `"  Garima Sinha   "` → `"Garima Sinha"`) | Ensures names match exactly for lookups into Master_Data |
| Raw Gamification Data → Clean Gamification Data | Same whitespace trim applied to Student Name | Required so INDEX/MATCH lookups in Master_Data find every student |
| Clean Student Data | Added a "Round Overall Effectiveness" column | Rounded value used for cleaner display/reporting |

## Master_Data — Key Derived Columns

The Master_Data sheet (35 columns, 1,002 rows) merges the cleaned student and gamification data and adds the following calculated fields:

| Column | Formula Logic | Purpose |
|---|---|---|
| Age Band | `IF(Age<=22,"20-22",IF(Age<=25,"23-25",IF(Age<=28,"26-28","29-30")))` | Buckets student age into 4 bands for segmentation & slicers |
| Gamification % / Non-Gamification % (8 columns: Engagement, Motivation, Participation, Learning Outcomes) | `IFERROR(INDEX(...,MATCH(Name,...,0)),"Not Found")` | Pulls each student's gamification survey results by name (INDEX/MATCH used in place of XLOOKUP for cross-version Excel compatibility) |
| Engagement / Motivation / Participation / Learning Outcomes Lift (4 columns) | `Gamification % - Non-Gamification %` | Point-gain from gamification per student, per metric |
| Performance Tier | `IF(Overall Effectiveness<3.9,"Low",IF(<4.2,"Medium","High"))` | 3-tier segmentation used across pivots and as a slicer |
| Best Method (Effectiveness / Engagement) | `CHOOSE(MATCH(MAX(5 cols),...,0),"Lectures",...)` | The teaching method each student rated highest |
| Teachers Count - Best Method | `COUNTIF(Teacher_Master Teaching Method, Best Method)` | How many teachers teach that student's top method |
| Example Teacher - Best Method | `INDEX(Teacher_Master Name, MATCH(Best Method, Teaching Method,0))` | A sample teacher name for that method |

**Full column list:** Student ID, Student Name, Gender, Age, Age Band, effectiveness & engagement ratings per teaching method (10 columns), Overall Effectiveness, Round Overall Effectiveness, Overall Engagement, Gamification %/Non-Gamification % (8 columns), Lift (4 columns), Performance Tier, Best Method (Effectiveness), Best Method (Engagement), Teachers Count - Best Method, Example Teacher - Best Method.

## Teacher_Master — Key Derived Columns

| Column | Formula Logic | Purpose |
|---|---|---|
| Performance Tier | `IF(Experience<8,"Junior",IF(<12,"Mid","Senior"))` | Quick experience-based grouping |
| Experience Band | `IF(Experience<8,"Junior",IF(<=12,"Mid","Senior"))` | Used as the row/column field in Teacher pivot tables & slicers |

## Overall Analysis & KPI Analysis

- **KPI Analysis** aggregates five categories — Student, Teaching, Engagement, Gamification, and Teacher KPIs — using `AVERAGE`, `AVERAGEIFS`, `COUNTIF`, `MAX`, `MIN`, and `INDEX`/`MATCH` formulas referencing Master_Data and Teacher_Master.
- **Analysis** condenses these into an Executive KPI Summary: effectiveness/engagement overview, gamification impact & lift, a `CORREL()` check between effectiveness and engagement, the best teaching method, and a live Top-5 Teachers by Experience ranking (via `LARGE` + `INDEX`/`MATCH`).

## Tech Stack

- Microsoft Excel (formulas: `INDEX`/`MATCH`, `IFERROR`, `COUNTIF`/`COUNTIFS`, `AVERAGE`/`AVERAGEIFS`, `CHOOSE`, `CORREL`, `LARGE`)
- PivotTables & PivotCharts with connected slicers
- No external libraries or scripts — the entire analysis runs natively in Excel

## How to Use

1. Open the workbook and start on the **Home** sheet for navigation.
2. Read **Instructions** for the full data dictionary and formula reference.
3. Explore **Dashboard** for the interactive summary view, or drill into individual **Pivot Tables & Charts** sheets.
4. Use the slicers (Gender, Age Band, Performance Tier, Teaching Method, Metric) to filter views interactively.

## Repository Contents

```
├── student_survey_dataset.xlsx   # Full workbook (raw data, cleaning, analysis, dashboard)
└── README.md                     # This file
```

## License

Add a license of your choice (e.g. MIT) if you intend this to be reused by others.
