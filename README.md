# PR. 2 Analyzer

## Project Overview

This Excel project analyzes customer sales data and demonstrates
practical Excel data-analysis, business-intelligence, and
data-visualization techniques.

The workbook uses Excel 2021 features including formulas, Conditional
Formatting, What-If Analysis, the Analysis ToolPak, regression,
descriptive statistics, customer analysis, Pivot Tables, charts, and
dashboard reporting.

## Workbook Structure

### 1. Project Instructions

Contains the project requirements and the list of analysis tasks to be
completed.

### 2. Dataset

Contains the main sales dataset with the following fields:

-   Customer_ID
-   Customer_Name
-   Region
-   Product_Category
-   Sales
-   Quantity
-   Discount
-   Order_Date
-   Profit

Additional analysis columns include:

-   Unique_customer
-   Total_sales per Customer
-   Rank per Customer
-   Discount analysis
-   Predicted profit

### 3. Descriptive-analysis

Contains descriptive statistics generated using the Excel Analysis
ToolPak.

The analysis includes measures such as:

-   Mean
-   Standard Error
-   Median
-   Mode
-   Standard Deviation
-   Sample Variance
-   Kurtosis
-   Skewness
-   Range
-   Minimum
-   Maximum
-   Sum
-   Count

### 4. Regression

Contains:

-   Linear Regression of Profit vs Sales
-   Regression statistics
-   ANOVA results
-   Regression coefficients
-   Monthly sales calculations
-   Monthly sales growth
-   High-value customer analysis

### 5. Dashboard

Provides a visual summary of the analysis using KPI cards and charts.

## Analysis Performed

### 1. Top 10 Customers

Conditional Formatting is used to identify high-value customers based on
total customer purchases.

Repeated customer names are handled by calculating total sales per
unique customer before ranking.

### 2. What-If Analysis

A What-If Data Table is used to examine how changing the discount
affects the calculated sales/profit result.

Different discount values can be tested without manually changing the
input for every scenario.

### 3. Linear Regression

A linear regression model is performed with:

-   Independent variable (X): Sales
-   Dependent variable (Y): Profit

The regression output includes R Square, Adjusted R Square, standard
error, ANOVA, coefficients, t-statistics, and p-values.

The model produced an R Square of approximately **0.9246**, meaning that
about **92.46% of the variation in Profit is explained by Sales** in
this linear model.

The estimated regression equation is approximately:

`Profit = 28.0683 + 0.8725 × Sales`

The Sales coefficient is statistically significant based on the very
small p-value reported by the regression output.

### 4. Descriptive Statistics

Key descriptive results include:

  Variable        Mean   Median   Standard Deviation
  ---------- --------- -------- --------------------
  Sales         976.09   976.72               299.27
  Quantity       10.00    10.00                 5.65
  Discount     0.09725     0.10              0.07233

The dataset contains **200 observations**.

### 5. Monthly Sales Growth

Monthly sales are calculated from the Order_Date and Sales columns using
date-based `SUMIFS()` formulas.

Monthly growth is calculated using:

`=(Current Month Sales - Previous Month Sales) / Previous Month Sales`

Conditional Formatting with up/down arrows is used to make positive and
negative growth easier to identify.

### 6. Timestamp

The workbook can use the `NOW()` function to create a dynamic timestamp
showing the current date and time.

Example:

`=NOW()`

### 7. High-Value Customers

Customer-level sales are aggregated using `SUMIF()`.

`INDEX()` and `MATCH()` are then used to retrieve customer names
associated with the highest sales values, while filtering can be used to
display the Top 5 high-value customers.

The highest-value customer in the dataset is:

**James Wilson --- 20,774.74**

Other high-value customers include:

1.  James Wilson --- 20,774.74
2.  Mary Davis --- 17,143.29
3.  David Brown --- 11,915.83
4.  Emma Davis --- 11,398.32
5.  Linda Wilson --- 9,970.20

### 8. Sales by Region and Product

Sales are summarized by region and product category to support business
comparisons and dashboard visualization.

Total sales by region:

  Region      Total Sales
  --------- -------------
  West          44,083.58
  Central       43,468.33
  South         40,248.79
  North         34,975.42
  East          32,441.64

Total sales by product category:

  Product Category     Total Sales
  ------------------ -------------
  Books                  47,002.46
  Clothing               41,927.59
  Electronics            41,044.00
  Furniture              36,612.89
  Office Supplies        28,630.82

### 9. Charts

The dashboard uses charts to communicate key performance indicators:

-   **Bar Chart:** compares sales across product categories.
-   **Line Chart:** shows monthly sales trends.
-   **Pie Chart:** shows the contribution of regions to total sales.

### 10. Dashboard

The Dashboard brings the main findings together in a single view.

Recommended KPI cards include:

-   Total Sales
-   Total Profit
-   Total Quantity
-   Average Sales

The dashboard also contains visualizations for:

-   Monthly sales trend
-   Sales by product category
-   Sales by region
-   Top high-value customers

## Key Business Insights

-   Total sales are approximately **195,217.76**.
-   Total profit is approximately **175,945.00** based on the workbook's
    current Profit values.
-   The dataset contains **200 transactions**.
-   **Books** is the highest-selling product category with approximately
    **47,002.46** in sales.
-   **West** is the highest-selling region with approximately
    **44,083.58** in sales.
-   **James Wilson** is the highest-value customer with approximately
    **20,774.74** in total sales.
-   Sales and Profit have a strong positive linear relationship, with an
    R Square of approximately **0.9246**.

## Excel Skills Demonstrated

This project demonstrates the following Excel skills:

-   `SUMIF()`
-   `SUMIFS()`
-   `LARGE()`
-   `INDEX()`
-   `MATCH()`
-   `RANK.EQ()`
-   `AVERAGE()`
-   `SUM()`
-   `NOW()`
-   `EDATE()`
-   Conditional Formatting
-   Icon Sets and arrows
-   What-If Analysis
-   Data Tables
-   Analysis ToolPak
-   Descriptive Statistics
-   Linear Regression
-   Pivot Tables
-   Bar Charts
-   Line Charts
-   Pie Charts
-   Dashboard design
-   Business insight and data storytelling

## Conclusion

The project transforms raw sales data into actionable business
information. It combines customer-level analysis, sales trends, regional
and product comparisons, statistical analysis, regression, and visual
reporting to provide a concise view of sales and profitability
performance.
