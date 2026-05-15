# Adventure Works Sales Dashboard — Power BI 

A quick business intelligence project built in Microsoft Power BI Desktop, analysing sales performance, customer behaviour, product returns, and regional trends for the fictional Adventure Works Cycles company - refreshed some skills, learned some new skills! Full write up below, [dashboard walkthrough here](https://www.youtube.com/watch?v=Yi4GuONy8R4): 

<img width="1356" height="771" alt="image" src="https://github.com/user-attachments/assets/f0ee9bfb-0fb1-4cbd-9a86-98379115e84c" />

> **Note:** This project was completed as part of the [Microsoft Power BI Desktop for Business Intelligence](https://www.udemy.com/course/microsoft-power-bi-up-running-with-power-bi-desktop/) course on Udemy (Maven Analytics). The dataset and brief were provided by the course; the data model, DAX logic, and dashboard design are my own work.

---

## Project Overview

| | |
|---|---|
| **Tool** | Microsoft Power BI Desktop |
| **Dataset** | Adventure Works Cycles (sales, products, customers, territories) |
| **Scope** | Sales performance, product analysis, customer segmentation, regional trends |
| **Key skills** | Data modelling, DAX, Power Query, dashboard design |

---

## Data Model

<img width="1443" height="668" alt="image" src="https://github.com/user-attachments/assets/5526b7f7-714a-4f95-afcd-688cec30b008" />

The report is built on a **star schema** — a central fact table surrounded by dimension tables connected via one-to-many relationships. This structure keeps the model clean, avoids data duplication, and ensures DAX filter context flows correctly from dimensions into the fact table.

### Tables

**Fact Tables**
- `Sales Data` — one row per order line, containing `OrderQuantity`, `OrderNumber`, `OrderDate`, `CustomerKey`, `ProductKey`, and `SalesTerritoryKey`
- `Returns Data` — one row per returned product, containing `ReturnQuantity`, `ReturnDate`, `ProductKey`, and `TerritoryKey`

**Dimension Tables**
- `Calendar Lookup` — a contiguous date table built in Power Query, with calculated columns for Day Name, Day of Week, Month Name, Month Name (DAX), Month Number, Month Short, Start of Month, Start of Quarter, Start of Week, and Start of Year
- `Customer Lookup` — customer demographics including `AnnualIncome`, `BirthYear`, `BirthDate`, `CustomerPriority`, `CustomerKey`, `DomainName`, and `EducationCategory`
- `Product Lookup` — product attributes including `ModelName`, `PricePoint`, `ProductColor`, `ProductCost`, `ProductKey`, `ProductName`, `ProductPrice`, `ProductSKU`, and `ProductStyle`; also contains a `Multiplication` calculated column used in revenue row-level calculations
- `Product Subcategories Lookup` — bridges products up to categories via `ProductSubcategoryKey` and `SubcategoryName`
- `Product Categories Lookup` — top-level grouping (Bikes, Accessories, Clothing) via `CategoryName` and `ProductCategoryKey`
- `Territory Lookup` — sales geography with `Continent`, `Country`, `Region`, and `SalesTerritoryKey`

### Relationships

All relationships are **one-to-many, single-direction** with the "one" side on the lookup/dimension table flowing down into the fact tables. Key relationships:

```
Calendar Lookup[Date]                               → Sales Data[OrderDate]
Calendar Lookup[Date]                               → Returns Data[ReturnDate]
Customer Lookup[CustomerKey]                        → Sales Data[CustomerKey]
Product Lookup[ProductKey]                          → Sales Data[ProductKey]
Product Lookup[ProductKey]                          → Returns Data[ProductKey]
Territory Lookup[SalesTerritoryKey]                 → Sales Data[SalesTerritoryKey]
Territory Lookup[SalesTerritoryKey]                 → Returns Data[TerritoryKey]
Product Subcategories Lookup[ProductSubcategoryKey] → Product Lookup[ProductSubcategoryKey]
Product Categories Lookup[ProductCategoryKey]       → Product Subcategories Lookup[ProductCategoryKey]
```

The snowflaked product hierarchy (Product Lookup → Product Subcategories Lookup → Product Categories Lookup) means filters applied at the category level flow through to the fact tables without needing to merge or flatten tables — keeping the model normalised and efficient.

---

## DAX Measures

All measures are stored in a dedicated **`_Measures`** table (not tied to any data table) to keep the fields pane organised and to signal clearly that these are calculated values, not raw columns.

### Core Revenue & Volume Measures

```dax
-- Total Revenue
Total Revenue = SUMX('Sales Data', 'Sales Data'[OrderQuantity] * RELATED('Product Lookup'[ProductPrice]))

-- Total Cost
Total Cost = SUMX('Sales Data', 'Sales Data'[OrderQuantity] * RELATED('Product Lookup'[ProductCost]))

-- Total Profit
Total Profit = [Total Revenue] - [Total Cost]

-- Profit Margin
Profit Margin = DIVIDE([Total Profit], [Total Revenue])

-- Total Orders
Total Orders = DISTINCTCOUNT('Sales Data'[OrderNumber])

-- Total Customers
Total Customers = DISTINCTCOUNT('Sales Data'[CustomerKey])

-- Revenue per Customer
Revenue per Customer = DIVIDE([Total Revenue], [Total Customers])
```

> `SUMX` is used rather than `SUM` because unit price and quantity live in separate tables — the iterator walks each row of `Sales Data`, fetches the related price from `Product Lookup` via `RELATED`, and multiplies before aggregating.

---

### Return Analysis

```dax
-- Total Returns
Total Returns = COUNT('Returns Data'[ReturnQuantity])

-- Return Rate
Return Rate = DIVIDE([Total Returns], [Total Orders], 0)

-- Total Returned Quantity
Total Returned Quantity = SUM('Returns Data'[ReturnQuantity])
```

---

### Time Intelligence

The Calendar table is the engine for all time intelligence. It must have no gaps in dates for these functions to work correctly.

```dax
-- Previous Month Revenue
Previous Month Revenue =
    CALCULATE(
        [Total Revenue],
        DATEADD('Calendar Lookup'[Date], -1, MONTH)
    )

-- Month-over-Month Revenue Change
MoM Revenue Change % =
    DIVIDE(
        [Total Revenue] - [Previous Month Revenue],
        [Previous Month Revenue]
    )

-- Revenue YTD
Revenue YTD =
    CALCULATE(
        [Total Revenue],
        DATESYTD('Calendar Lookup'[Date])
    )

-- Previous Year Revenue (for YoY comparison)
Previous Year Revenue =
    CALCULATE(
        [Total Revenue],
        DATEADD('Calendar Lookup'[Date], -1, YEAR)
    )
```

> All time intelligence functions rely on filter context from `Calendar Lookup`. Slicers and visual-level filters on date fields automatically shift the context window — `DATEADD` and `DATESYTD` then operate relative to whatever period is currently selected.

---

### Target & KPI Measures

```dax
-- Revenue Target (10% growth over previous month)
Revenue Target =
    [Previous Month Revenue] * 1.10

-- Target Gap
Target Gap = [Total Revenue] - [Revenue Target]

-- Orders Target
Orders Target = [Previous Month Orders] * 1.10
```

---

### Customer Segmentation

```dax
-- Top Customer by Revenue (used in card visual with TopN filter)
Top Customer Revenue =
    MAXX(
        TOPN(1, SUMMARIZE('Sales Data', 'Sales Data'[CustomerKey], "Rev", [Total Revenue]), [Rev], DESC),
        [Rev]
    )
```

---

## Power Query: Data Preparation

Key transformations applied before the data reached the model:

- **`Calendar Lookup` built in M** using `List.Dates` to generate a contiguous date sequence; extended with calculated columns for Day Name, Day of Week, Month Name, Month Name (DAX), Month Number, Month Short, Start of Month, Start of Quarter, Start of Week, and Start of Year — all computed in Power Query so the table can be marked as a Date Table in Power BI
- **`BirthYear` derived** in `Customer Lookup` from the `BirthDate` column, giving a clean integer field for age-bracket analysis without storing a volatile calculated age
- **`PricePoint` column added** to `Product Lookup` via a conditional column bucketing `ProductPrice` into price tiers — used to slice product performance without writing DAX
- **`ProductCost` and `ProductPrice` kept as separate columns** in `Product Lookup` rather than pre-computing margin in Power Query, so margin calculations can respond correctly to DAX filter context at report runtime
- **Column removal and data type enforcement** applied to all tables at load — particularly ensuring keys are integers and date fields are typed as Date (not DateTime), which is required for relationship matching and time intelligence to work correctly

---

## Key Business Insights

- **Bikes dominate revenue** but have a higher return rate than Accessories — margin analysis by subcategory surfaces this trade-off clearly
- **Revenue growth is seasonal**, with a consistent Q4 uplift driven by the North America territory
- **A small percentage of customers generate a disproportionate share of revenue** — visible through the revenue-per-customer distribution and the Top N customer analysis
- **Month-over-month KPIs** expose that while total orders trend upward, average order value fluctuates — suggesting promotional activity rather than organic growth in some periods

---

## Dashboard Pages

| Page | Focus |
|---|---|
| **Executive Summary** | High-level KPIs: revenue, orders, returns, profit — with MoM trend indicators |
| **Product Detail** | Drill-through page for individual product performance, return rate, and target tracking |
| **Customer Detail** | Customer count, revenue per customer, top customers, and demographic breakdown |
| **Map** | Regional sales distribution by territory |

> A full walkthrough of the dashboard design and the business story it tells is covered in the accompanying video.

---

## Files

```
├── AdventureWorks_Report.pbix    # Power BI report file
├── /data                         # Source CSV files
│   ├── AdventureWorks_Sales/     # Sales fact tables (by year)
│   ├── AdventureWorks_Returns.csv
│   ├── AdventureWorks_Customers.csv
│   ├── AdventureWorks_Products.csv
│   ├── AdventureWorks_Product_Subcategories.csv
│   ├── AdventureWorks_Product_Categories.csv
│   └── AdventureWorks_Territories.csv
└── /screenshots                  # Dashboard preview images
```

---

## What I (re-)learned

- Why **star schema beats flat tables** for DAX performance — filter context flows naturally from dimension to fact, and measures stay simple
- The difference between **calculated columns** (computed row-by-row at refresh, stored in the model) and **measures** (computed at query time in filter context) — and when to use each
- Why **`SUMX` is necessary** when a calculation requires data from multiple tables at the row level, and why `SUM` alone would fail or require a merged table
- How **`CALCULATE`** is the backbone of almost every non-trivial DAX measure — it modifies filter context before evaluating an expression
- How the **Calendar table** enables time intelligence — and why gaps in dates or missing Date Table marking break functions like `DATESYTD`
- The value of **organising measures into a dedicated table** early — it pays off immediately when the model grows past a handful of measures
