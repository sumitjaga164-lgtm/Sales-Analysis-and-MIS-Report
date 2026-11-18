# Automated Sales Analysis & MIS Dashboard (Excel)

## Project Overview
This project is a complete **Management Information System (MIS) Report** and **Sales Analysis Dashboard** built in Microsoft Excel. It demonstrates end-to-end data processing—from raw data cleaning and dynamic lookup modeling to automated KPI tracking and visual reporting.

---

## Key Features & Excel Functions Used

### 1. Data Cleaning & Standardization
* **Text Formatting:** Standardized inconsistent employee name casing using the `PROPER` function.
* **Dynamic Categorization:** Automatically classified transactions into `high value` and `regular` tiers using `IF` logic paired with **Conditional Formatting** for quick visual scanning.

### 2. Dynamic Lookup & Error Handling
* **Product Lookup (`VLOOKUP`):** Fetches product details dynamically based on a given `Order_ID`.
* **Two-Way Lookup (`INDEX + MATCH`):** Retrieves employee names accurately across rows and columns, wrapped with `IFERROR` to display `"INVALID ORDER ID"` instead of standard Excel error codes.

### 3. Automated KPI Summary Table
* **Core Metrics:** Calculated **Total Sales (`₹14,80,100`)**, **Average Sales (`₹98,673`)**, **Highest Sale (`₹2,80,500`)**, and **Lowest Sale (`₹22,500`)** using `SUM`, `AVERAGE`, `MAX`, and `MIN`.
* **Conditional Aggregations:**
  * **Total Orders Count:** `COUNTA`
  * **North Region Orders:** `COUNTIF`
  * **Sales Dept Total Revenue:** `SUMIF`
  * **Multi-Criteria Revenue (East Region + Laptop):** `SUMIFS` (`₹3,85,000`)

### 4. Visual Reporting (Pivot Table & Chart)
* Summarized product-wise net sales (`Keyboard`, `Laptop`, `Monitor`, `Mouse`) using a **Pivot Table**.
* Created a clean **Column Chart** (*Product-Wise Total Sales*) to highlight top-performing product categories.

---

## Files Included
* `Sales_Analysis_MIS_Report.xlsx` — Complete workbook containing raw data, KPI summary, and Pivot analysis.
* `screenshots/` — Preview images of the data sheet and visual chart.

---

## Tools & Skills
* **Microsoft Excel:** Data Cleaning, `VLOOKUP`, `INDEX-MATCH`, `SUMIFS`, `COUNTIF`, Conditional Formatting, Pivot Tables, Data Visualization.

