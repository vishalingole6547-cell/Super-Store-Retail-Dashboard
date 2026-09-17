# 📊 Excel Sales Dashboard | Power Query + Power Pivot + DAX

## 📌 Project Overview

This project is an interactive **Excel Sales Dashboard** developed to transform raw sales data into meaningful business insights.

The dashboard combines **Power Query** for data cleaning and transformation, **Power Pivot** for data modeling, **DAX** for advanced calculations, and **Pivot Charts, Slicers, and Filters** for interactive reporting.

The objective is to create a dynamic reporting solution that enables users to monitor sales performance, analyze trends, compare targets, and identify key business insights.

---

## 🎯 Project Objectives

* Clean and transform raw data using **Power Query**
* Build a structured relational **data model using Power Pivot**
* Create calculated KPIs using **DAX measures**
* Analyze sales performance across different dimensions
* Compare actual sales against targets
* Analyze monthly and regional sales trends
* Identify top-performing products and sales representatives
* Create an interactive and user-friendly dashboard
* Enable dynamic analysis using slicers and filters

---

## ❓ Business Questions

The dashboard is designed to answer the following business questions:

1. What is the total revenue generated?
2. What is the total quantity sold?
3. How does actual sales performance compare with the target?
4. What is the overall target achievement percentage?
5. How does revenue change month by month?
6. Which products generate the highest revenue?
7. Which regions contribute the most to sales?
8. Which sales representatives have the highest sales?
9. Which products have the highest quantity sold?
10. How does sales performance vary across different regions and months?
11. Which areas are performing below the sales target?
12. How do sales trends change when different filters are applied?

---

# 🔄 Project Workflow

```text
Raw Data
    ↓
Power Query
    ↓
Data Cleaning & Transformation
    ↓
Power Pivot
    ↓
Data Model & Relationships
    ↓
DAX Measures
    ↓
Pivot Tables
    ↓
Pivot Charts
    ↓
Slicers & Filters
    ↓
Interactive Dashboard
    ↓
Business Insights
```

### Workflow Steps

**1. Data Collection**

Raw sales data is imported into Excel.

**2. Data Transformation**

Power Query is used to clean and prepare the data.

Typical transformations include:

* Removing unnecessary columns
* Handling missing values
* Changing data types
* Removing duplicate records
* Standardizing text values
* Creating required calculated columns

**3. Data Modeling**

The transformed data is loaded into the **Power Pivot Data Model**.

Relationships are created between relevant tables to enable efficient analysis.

**4. DAX Calculations**

DAX measures are created to calculate important business KPIs.

**5. Visualization**

Pivot Tables and Pivot Charts are created from the data model.

**6. Dashboard Development**

Slicers and filters are connected to the dashboard to provide interactive analysis.

---

# 🗂️ Data Model

The dashboard uses **Power Pivot** to create a structured data model.

### Example Data Model

```text
                ┌─────────────────┐
                │   Date Table    │
                │─────────────────│
                │ Date            │
                │ Month           │
                │ Year            │
                │ Quarter         │
                └────────┬────────┘
                         │
                         │
                         ▼
┌──────────────┐    ┌─────────────────┐    ┌──────────────┐
│ Product      │    │   Sales Table   │    │ Salesperson  │
│──────────────│    │─────────────────│    │──────────────│
│ Product ID   │────│ Product ID      │    │ Rep ID       │
│ Product Name │    │ Date            │────│ Rep Name     │
│ Category     │    │ Region ID       │    │ Region       │
└──────────────┘    │ Rep ID          │    └──────────────┘
                    │ Quantity        │
                    │ Unit Price      │
                    │ Revenue        │
                    │ Target          │
                    └────────┬────────┘
                             │
                             ▼
                       ┌────────────┐
                       │  Region    │
                       │────────────│
                       │ Region ID  │
                       │ Region     │
                       └────────────┘
```

The model enables analysis across multiple dimensions such as:

* Date
* Product
* Region
* Salesperson
* Sales performance

---

# 🧮 DAX Model

DAX (**Data Analysis Expressions**) is used to create dynamic measures for the dashboard.

Measures respond automatically to slicers and filters, allowing users to perform dynamic analysis.

## 📐 DAX Measures

### Total Revenue

```DAX
Total Revenue =
SUM(Sales[Revenue])
```

### Total Quantity

```DAX
Total Quantity =
SUM(Sales[Quantity])
```

### Total Target

```DAX
Total Target =
SUM(Sales[Target])
```

### Achievement %

```DAX
Achievement % =
DIVIDE(
    [Total Revenue],
    [Total Target],
    0
)
```

### Average Revenue

```DAX
Average Revenue =
AVERAGE(Sales[Revenue])
```

### Number of Transactions

```DAX
Total Transactions =
COUNTROWS(Sales)
```

> **Note:** Measure and column names should be adjusted according to the actual names used in the workbook.

---

# 📊 Dashboard Features

The Excel dashboard includes the following features:

### 🔹 KPI Cards

The dashboard provides high-level KPIs such as:

* **Total Revenue**
* **Total Quantity**
* **Total Target**
* **Achievement %**
* **Total Transactions**

### 🔹 Sales Trend Analysis

A Pivot Chart displays sales/revenue trends across different months.

### 🔹 Product Analysis

The dashboard enables users to analyze:

* Revenue by product
* Quantity sold by product
* Product contribution to overall sales

### 🔹 Regional Analysis

Users can compare sales performance across different regions.

### 🔹 Salesperson Analysis

Sales performance can be analyzed by individual sales representatives.

### 🔹 Interactive Slicers

Slicers allow users to dynamically filter dashboard results by dimensions such as:

* Month
* Region
* Product
* Salesperson

### 🔹 Filters

Pivot Table and dashboard filters provide additional control over the analysis.

---

# 📸 Dashboard Screenshots

## Dashboard Overview

Add your dashboard screenshot here:

```markdown
![Sales Dashboard](screenshots/dashboard-overview.png)
```

## Sales Performance

```markdown
![Sales Performance](screenshots/sales-performance.png)
```

## Interactive Filters

```markdown
![Dashboard Filters](screenshots/dashboard-filters.png)
```

> Replace the image paths above with the actual screenshots uploaded to the repository.

---

# 🔍 Key Findings

The dashboard can be used to identify important sales patterns and performance indicators, including:

* Overall revenue and sales performance
* Monthly revenue trends
* Top-performing products
* Regional contribution to total revenue
* Salesperson performance
* Target achievement levels
* Products or regions requiring additional attention
* Changes in performance after applying filters

### Example Insight Format

```text
• The dashboard identifies the products contributing the highest revenue.
• Monthly analysis highlights periods of higher and lower sales activity.
• Regional analysis shows differences in sales contribution across regions.
• Target achievement measures help compare actual performance with planned targets.
• Salesperson analysis highlights variations in individual sales performance.
```

> Replace these example findings with the actual findings from your dataset.

---

# 🛠️ Tools & Technologies

| Tool                | Purpose                             |
| ------------------- | ----------------------------------- |
| **Microsoft Excel** | Dashboard development and reporting |
| **Power Query**     | Data cleaning and transformation    |
| **Power Pivot**     | Data modeling and relationships     |
| **DAX**             | KPI and business calculations       |
| **Pivot Tables**    | Data summarization                  |
| **Pivot Charts**    | Data visualization                  |
| **Slicers**         | Interactive filtering               |
| **Filters**         | Dynamic data analysis               |

---

# 📁 Repository Structure

```text
Excel-Sales-Dashboard/
│
├── 📊 Excel-Sales-Dashboard.xlsx
│
├── 📄 README.md
│
├── 📁 Dataset/
│   └── sales_data.csv
│
├── 📁 Screenshots/
│   ├── dashboard-overview.png
│   ├── sales-performance.png
│   └── dashboard-filters.png
│
└── 📁 Documentation/
    └── data-model.png
```

---

# 🚀 How to Use

1. Download or clone this repository.
2. Open `Excel-Sales-Dashboard.xlsx`.
3. Navigate to the **Dashboard** sheet.
4. Use the available **Slicers** and **Filters**.
5. Select different months, regions, products, or salespersons.
6. Review the updated KPIs and Pivot Charts.
7. Refresh the workbook when new data is added.

---

# 📚 Skills Demonstrated

This project demonstrates practical skills in:

* Excel Dashboard Development
* Data Cleaning
* Power Query
* Power Pivot
* Data Modeling
* DAX
* KPI Development
* Pivot Tables
* Pivot Charts
* Interactive Slicers
* Business Intelligence
* Data Visualization
* Business Analysis

---

# 🎓 Project Outcome

This project demonstrates how Excel can be used as a complete **Business Intelligence and Reporting solution**, starting from raw data and progressing through data transformation, data modeling, DAX calculations, visualization, and interactive dashboard development.

The final dashboard provides a centralized view of sales performance and enables users to explore the data through interactive filters and slicers.

---

## 👤 Author

**Vishal Ingole**

### 📌 Excel | Power Query | Power Pivot | DAX | Data Analytics
