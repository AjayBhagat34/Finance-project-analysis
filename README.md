# 📊 Finance & Sales Analysis Dashboard – Excel Project

## 📌 Project Overview

This project is an **Excel-based Finance and Sales Analysis Dashboard** created to analyze sales transactions, identify sales trends, compare product performance, and generate meaningful business insights.

The project uses **Microsoft Excel**, including:

* Excel Tables
* Formulas
* PivotTables
* Data grouping and bucketing
* Round-off calculations
* GETPIVOTDATA
* Dashboard visualizations
* Product-wise and day-wise analysis

The objective of this project is to transform raw sales transaction data into structured analysis and an interactive dashboard that can support business decision-making.

---

## 📁 Workbook Structure

The Excel workbook contains four main worksheets:

### 1. `Finance`

This is the **main raw-data and calculation sheet**.

It contains **794 sales transactions** with the following fields:

| Column              | Description                               |
| ------------------- | ----------------------------------------- |
| Order ID            | Unique order identifier                   |
| Product ID          | Product identifier                        |
| Sale Date           | Date of the transaction                   |
| Amount in Sales     | Original sales amount                     |
| Discounted Value    | Discounted amount/value                   |
| Region              | Sales region                              |
| Workday             | Day of the week calculated from Sale Date |
| Roundoff            | Sales amount rounded to nearest ₹5        |
| Sales Amount Bucket | Sales grouped into predefined ranges      |
| Transaction         | Sequential transaction number             |

The data contains:

* **794 transactions**
* **6 products**
* **4 regions:** North, South, East and West
* Date range: **13 June 2022 – 6 September 2022**
* Total sales amount: **₹438,968**
* Average sales amount per transaction: **₹552.86**

---

### 2. `Finance Pivot`

This sheet contains the PivotTables and supporting calculations used for analysis.

The analysis includes:

#### 📅 Day-wise Sales Analysis

Sales are analyzed according to the day of the week.

Metrics include:

* Total sales by day
* Average sales by day

#### 🛍️ Product-wise Analysis

Product performance is analyzed using:

* Total sales
* Average sales

Products analyzed:

* PIZB0001
* PIZB0002
* PIZB0003
* PIZB0004
* PIZB0005
* PIZB0006

#### 💰 Sales Amount Buckets

Transactions are categorized into sales ranges such as:

* Below ₹300
* ₹300–₹500
* ₹500–₹700
* ₹700–₹900
* ₹900–₹1,100
* ₹1,100–₹1,300

This helps identify the distribution of transactions across different sales-value ranges.

#### 🔢 Product Count by Sales Bucket

The project also calculates the number of products/transactions falling within each sales bucket.

#### 🔄 Sales vs Round-off Comparison

The original sales amount is compared with the amount rounded to the nearest ₹5.

The project calculates the overall difference between:

**Actual Sales Amount vs Rounded Sales Amount**

---

### 3. `Finance Dashboard`

This worksheet contains the visual dashboard created from the underlying PivotTables.

The dashboard is designed to provide a quick overview of important sales metrics and trends.

It includes visual analysis based on:

* Day-wise sales
* Average sales
* Product-wise sales
* Sales amount buckets
* Actual vs rounded sales values

---

### 4. `Dashboard Questions`

This worksheet contains the analytical requirements/questions used to build the project.

Examples include:

* Find overall sales value for each day
* Find average sales value for each day
* Create round-off values for sales amounts
* Compare actual sales with rounded sales
* Calculate sales amount by different buckets
* Calculate number of sold products by sales buckets
* Compare each product's sales value with its average price

---

## 🧮 Excel Formulas & Techniques Used

### 1. Day of Week

The `TEXT` function is used to extract the day from the sale date.

```excel
=TEXT([@[Sale Date]],"DDDD")
```

Example:

`20-Jul-2022 → Wednesday`

---

### 2. Round Sales Amount

The `MROUND` function is used to round the sales amount to the nearest ₹5.

```excel
=MROUND([@[Amount in Sales]],5)
```

Example:

`₹197 → ₹195`

---

### 3. Sales Amount Bucketing

Sales transactions are grouped into predefined ranges to simplify analysis and visualization.

Example categories:

```text
Below 300
300-500
500-700
700-900
900-1100
1100-1300
```

---

### 4. GETPIVOTDATA

`GETPIVOTDATA` is used to retrieve specific values from PivotTables.

Example:

```excel
=GETPIVOTDATA("Amount in Sales",$B$16)
```

This allows dashboard metrics to remain connected to PivotTable calculations.

---

## 📊 Key Analysis

### Overall Sales

The dataset contains:

**Total Sales: ₹438,968**

**Average Sales per Transaction: ₹552.86**

### Regional Distribution

| Region | Transactions |
| ------ | -----------: |
| North  |          199 |
| South  |          199 |
| East   |          198 |
| West   |          198 |

The transaction distribution is almost evenly split across the four regions.

### Product Analysis

The project compares each product based on:

* Total sales generated
* Average transaction value

This helps identify differences in product-level sales performance.

### Day-wise Analysis

The project calculates both total and average sales for:

* Sunday
* Monday
* Tuesday
* Wednesday
* Thursday
* Friday
* Saturday

This provides insight into which days generate relatively higher or lower sales.

---

## 🎯 Project Objectives

The main objectives of this project are:

1. Analyze sales transaction data.
2. Calculate total and average sales.
3. Analyze sales performance by day.
4. Analyze product-wise sales performance.
5. Categorize transactions into sales-value buckets.
6. Compare actual sales with rounded sales values.
7. Use PivotTables for business analysis.
8. Build a Finance/Sales Dashboard.
9. Convert raw data into meaningful business insights.
10. Practice Excel data-analysis techniques used in real-world reporting.

---

## 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Tables
* PivotTables
* Pivot-based calculations
* Excel formulas
* Data analysis
* Dashboarding
* Data visualization

---

## 📈 Skills Demonstrated

This project demonstrates practical knowledge of:

* Data Cleaning & Preparation
* Excel Data Analysis
* PivotTables
* Dashboard Creation
* Business Reporting
* Sales Analysis
* Product Analysis
* Date Functions
* `TEXT()` function
* `MROUND()` function
* `GETPIVOTDATA()`
* Data Categorization
* KPI Analysis
* Data Visualization

---

## 💡 Business Insights

The analysis can help a business understand:

* Overall sales performance
* Average transaction value
* Sales trends across different days
* Product contribution to total sales
* Distribution of transactions by sales value
* Difference between actual and rounded sales values
* Regional transaction distribution

These insights can be used to identify sales patterns and support data-driven business decisions.

---

## 📂 Project Files

```text
Finance Excel Project/
│
├── Finance excel project.xlsx
└── README.md
```

---

## 🚀 How to Use

1. Download or clone this repository.
2. Open `Finance excel project.xlsx` in Microsoft Excel.
3. Go to the `Finance` sheet to view the transaction-level data.
4. Open `Finance Pivot` to explore the PivotTable analysis.
5. Open `Finance Dashboard` to view the dashboard.
6. Use the PivotTables and dashboard visuals to explore the sales analysis.

---

## 👨‍💻 Project Author

**Ajay Bhagat**

### Role

Data Analytics / Excel Project

### Focus Areas

`Excel` • `Data Analysis` • `PivotTables` • `Dashboarding` • `Business Intelligence`
