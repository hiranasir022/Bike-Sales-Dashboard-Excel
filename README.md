# 🚲 Bike Sales Dashboard (Excel)

An end-to-end Excel data analytics project that cleans raw customer data, builds PivotTables, and presents key business insights through an interactive dashboard — built to demonstrate practical Excel skills for data analysis roles (Data Analyst / Business Analyst / MIS).

---

## 📌 Project Overview

This project analyzes a **bike-buyers dataset (1,000 customer records)** to understand what factors — income, age, commute distance, gender, occupation — influence whether a customer purchases a bike. The workbook is structured across four sheets, moving from raw data → cleaned data → aggregated PivotTables → a final interactive dashboard.

---

## 🗂️ Workbook Structure

| Sheet | Purpose |
|---|---|
| **bike-buyers** | Raw, unprocessed source data (1,026 rows, with duplicates/inconsistencies) |
| **working sheet** | Cleaned dataset (1,000 unique records) with an added calculated column |
| **Pivot Table** | Three PivotTables summarizing income, commute distance, and age vs. purchase behavior |
| **Dashboard** | Final visual dashboard combining charts for reporting/presentation |

**Columns in the dataset:** ID, Marital Status, Gender, Income, Children, Education, Occupation, Home Owner, Cars, Commute Distance, Region, Age, Age Brackets, Purchased Bike.

---

## 🧹 Data Cleaning Process

1. Removed duplicate/invalid rows — raw data had 1,026 rows, reduced to **1,000 clean unique records**.
2. Standardized categorical shorthand into full readable labels for reporting clarity:
   - `M` → `Male`, `F` → `Female`
   - `M` (Marital Status) → `Married`, `S` → `Single`
   - `Yes` / `No` kept consistent for Home Owner & Purchased Bike fields
3. Added a new calculated column — **Age Brackets** — using a nested `IF` formula (explained below).

---

## 🧮 Formulas Used

### 1. Nested IF — Age Bracket Classification
Used to bucket each customer into an age group based on their `Age` column, so the data can be grouped/analyzed by generation in the PivotTables and charts.

```excel
=IF(L2>54,"old",IF(L2>=31,"Middle Age",IF(L2<31,"Adolescent","invalid")))
```

**How it works (step by step):**
| Condition | Result |
|---|---|
| `Age > 54` | **"old"** |
| `Age >= 31` (and ≤ 54) | **"Middle Age"** |
| `Age < 31` | **"Adolescent"** |
| Anything else (blank/error) | **"invalid"** (safety fallback) |

**To reuse it yourself:** copy the formula into the first data row of an Age Brackets column, replace `L2` with your Age cell reference, then drag-fill it down the entire column — Excel automatically adjusts the row reference for every row.

### 2. PivotTables (no formula typed — built via Excel's Insert → PivotTable)
Three PivotTables were created directly from the **working sheet** data:

| PivotTable | Rows | Columns | Values |
|---|---|---|---|
| Average Income by Gender | Gender | Purchased Bike | Average of Income |
| Purchases by Commute Distance | Commute Distance | Purchased Bike | Count of Purchased Bike |
| Purchases by Age Bracket | Age Brackets | Purchased Bike | Count of Purchased Bike |

**How to recreate:** select the cleaned data range → `Insert` → `PivotTable` → drag the relevant field into **Rows**, `Purchased Bike` into **Columns**, and the metric field (Income / Purchased Bike) into **Values**, then set the value field's summary type to **Average** or **Count** as needed.

---

## 📊 Dashboard

The **Dashboard** sheet pulls all three PivotTables together into a single visual report using native Excel charts:

- **Bar Chart** — *Average Income Per Purchase* (income comparison between buyers vs. non-buyers, split by gender)
- **Line Chart** — *Customer Commute* (purchase trend across commute distance ranges)
- **Line Chart** — *Customer Age Brackets* (purchase trend across Adolescent / Middle Age / Old groups)

All charts are linked directly to their PivotTables, so the dashboard updates automatically if the underlying data or PivotTable is refreshed (`Data` → `Refresh All`).

---

## 📈 Key Insights

- Out of 1,000 customers, **495 purchased a bike (49.5%)** and 531 did not.
- **Customers with a commute of 0–1 miles** are by far the most likely to buy a bike (200 buyers vs. 166 non-buyers) — short-distance commuters are the strongest target segment.
- **Middle-aged customers (31–54)** make up the largest share of both buyers and non-buyers, and have the highest bike purchase count overall.
- **Male customers** who purchased bikes have a slightly higher average income (~$60,124) than female buyers (~$55,774).

---

## 🛠️ Tools & Skills Used

- Microsoft Excel
- Data Cleaning (duplicate removal, text standardization)
- Logical Functions (`IF`, Nested IF)
- PivotTables & PivotCharts
- Dashboard Design & Data Visualization

---

## 📂 How to Use

1. Download/clone this repository.
2. Open `bike_sale_dashboard.xlsx` in Microsoft Excel.
3. Go to the **Dashboard** sheet to view the final report.
4. To explore the logic, check the **working sheet** (formulas) and **Pivot Table** sheet (aggregations).

---

## 👤 Author
**[Hira Nasir]**
 🔗 (https://www.linkedin.com/in/hira-nasir-448635a5/)

*If you found this project useful, feel free to ⭐ star the repo!*
