# Data-Analytics-Project

## Important Notes

- **Update the file path:** Replace the example Excel file path with the location of `Ecommerce.xlsx` on your computer.
- **Configure SQL Server:** Replace the placeholder server name with your own SQL Server instance name.
- **Create the database:** Ensure the destination database (`E-Commerce`) already exists before running the script.
- **Windows Authentication:** This script uses Windows Authentication. Make sure your Windows account has permission to access the SQL Server instance and destination database.
- **Install dependencies:** Install the required Python libraries and a compatible Microsoft SQL Server ODBC driver before execution.
- **Excel sheet names:** The worksheet names specified in the script must match the names in your Excel workbook.
- **Table replacement warning:** The script uses `if_exists="replace"`. If a destination table already exists, it will be dropped and recreated. Existing data and table-specific constraints or indexes may be lost.

## Security and Privacy

- The SQL Server instance name in this project is configuration information, not a password or secret credential.
- For this public repository, use a placeholder for your local computer name and avoid publishing unnecessary details about your development environment.
- Never commit database passwords, API keys, access tokens, or other sensitive credentials.
- Ensure SQL Server is configured securely and does not expose unnecessary remote access.
- Review all files before publishing them to a public repository.

## Disclaimer


# E-Commerce Data Import and SQL Server Data Quality Validation

## Project Overview

This project demonstrates how to import multiple sheets from an Excel workbook into Microsoft SQL Server using Python, pandas, SQLAlchemy, and pyodbc.

After importing the data, SQL queries are used to validate row counts, identify duplicate keys, check missing values, verify relationships between fact and dimension tables, and review SQL Server data types.

The project follows a practical data analytics workflow, from data ingestion to database validation and schema improvement recommendations.

## Project Objectives

- Import multiple Excel worksheets into SQL Server using Python.
- Exclude the Data Dictionary sheet from the database import.
- Understand how pandas creates SQL Server tables and infers data types.
- Validate table row counts and key uniqueness.
- Identify missing values and orphaned foreign keys.
- Review data types for financial, date, flag, and text columns.
- Recommend improvements for a production-oriented database schema.

## Technology Stack

- **Python** — Data ingestion
- **pandas** — Reading Excel worksheets
- **SQLAlchemy** — Database connectivity
- **pyodbc** — SQL Server connectivity
- **Microsoft SQL Server Express** — Database storage
- **SQL Server Management Studio (SSMS)** — SQL queries and validation

## Project Workflow

1. Read all worksheets from the Excel workbook.
2. Skip the Data Dictionary worksheet.
3. Import each remaining worksheet into SQL Server.
4. Validate the imported row counts.
5. Check duplicate keys and missing values.
6. Validate relationships between fact and dimension tables.
7. Inspect SQL Server data types.
8. Identify potential schema improvements.

---

## Part 1: Import Excel Sheets into SQL Server Using Python

### 1. Install Required Libraries

Run the following command in your terminal:

```bash
pip install pandas openpyxl sqlalchemy pyodbc
```

You also need a compatible Microsoft ODBC Driver for SQL Server installed on your machine.

### 2. Configure the Database Connection

The project uses SQL Server Express with Windows Authentication.

Update the Excel file path and SQL Server instance to match your environment.

### 3. Python Import Script

```python
import pandas as pd
from sqlalchemy import create_engine

# Excel file path
excel_file = r"D:\YT\SQL\Ecommerce.xlsx"

# SQL Server configuration
server = r"YOUR_SERVER_NAME\SQLEXPRESS"
database = "E-Commerce"

# Windows Authentication
connection_string = (
    f"mssql+pyodbc://{server}/{database}"
    "?driver=ODBC+Driver+17+for+SQL+Server"
    "&trusted_connection=yes"
)

engine = create_engine(connection_string)

try:
    # Read all Excel sheets
    all_sheets = pd.read_excel(
        excel_file,
        sheet_name=None
    )

    # Import each sheet except the Data Dictionary
    for sheet_name, df in all_sheets.items():

        if sheet_name == "Data Dictionary":
            continue

        table_name = sheet_name.strip().replace(" ", "_")

        df.to_sql(
            name=table_name,
            con=engine,
            schema="dbo",
            if_exists="replace",
            index=False
        )

        print(f"Imported: {sheet_name} -> dbo.{table_name}")

    print("All data tables imported successfully!")

finally:
    engine.dispose()
```

**Important:** Replace `YOUR_SERVER_NAME` with your actual SQL Server instance name. The script assumes the target database already exists.

### 4. Understanding the Import Parameters

| Parameter | Purpose |
|---|---|
| `sheet_name=None` | Reads all worksheets into a dictionary of DataFrames |
| `if sheet_name == "Data Dictionary"` | Skips the documentation worksheet |
| `strip()` | Removes leading and trailing whitespace from worksheet names |
| `replace(" ", "_")` | Replaces spaces in table names with underscores |
| `schema="dbo"` | Creates tables in the `dbo` schema |
| `if_exists="replace"` | Drops and recreates an existing target table |
| `index=False` | Prevents pandas from writing the DataFrame index as an extra column |
| `engine.dispose()` | Releases the SQLAlchemy engine's resources |

### Important Note About `if_exists="replace"`

The script uses `if_exists="replace"`. This is suitable for a simple demonstration or repeatable full reload when replacing the target tables is intentional.

However, replacing tables can remove existing data and database objects such as primary keys, indexes, and constraints. Pandas also infers SQL column types from the DataFrames.

For production pipelines, consider explicitly defining the SQL schema, loading into staging tables, validating the data, and using a controlled insert or merge strategy.

Using `if_exists="append"` instead can duplicate data if the same rows are loaded repeatedly without deduplication or other safeguards.

---

## Part 2: Validate the Imported Data

### 1. Check Table Row Counts

The following query checks the row counts of the three main tables.

```sql
SELECT
    'FactOrderLine' AS TableName,
    COUNT_BIG(*) AS TotalRows
FROM dbo.FactOrderLine

UNION ALL

SELECT
    'DimProduct',
    COUNT_BIG(*)
FROM dbo.DimProduct

UNION ALL

SELECT
    'DimCustomer',
    COUNT_BIG(*)
FROM dbo.DimCustomer;
```

#### Observed Results

| Table | Total Rows |
|---|---:|
| `FactOrderLine` | 9,109 |
| `DimProduct` | 60 |
| `DimCustomer` | 2,200 |

These counts should be compared with the corresponding Excel worksheets to confirm that the expected number of rows was imported.

### 2. Check Duplicate Keys

#### Fact Table: `LineID`

```sql
SELECT
    LineID,
    COUNT_BIG(*) AS DuplicateCount
FROM dbo.FactOrderLine
GROUP BY LineID
HAVING COUNT_BIG(*) > 1;
```

#### Product Dimension: `ProductKey`

```sql
SELECT
    ProductKey,
    COUNT_BIG(*) AS DuplicateCount
FROM dbo.DimProduct
GROUP BY ProductKey
HAVING COUNT_BIG(*) > 1;
```

#### Customer Dimension: `CustomerKey`

```sql
SELECT
    CustomerKey,
    COUNT_BIG(*) AS DuplicateCount
FROM dbo.DimCustomer
GROUP BY CustomerKey
HAVING COUNT_BIG(*) > 1;
```

**Observed result:** All three queries returned no rows.

This confirms that the tested key values are unique in their respective tables. It does not rule out duplicate records with different keys or establish that the keys are formally constrained by primary keys.

### 3. Check Missing Values in the Fact Table

```sql
SELECT
    SUM(CASE WHEN LineID IS NULL
        THEN 1 ELSE 0 END) AS MissingLineID,
    SUM(CASE WHEN OrderID IS NULL
        THEN 1 ELSE 0 END) AS MissingOrderID,
    SUM(CASE WHEN ProductKey IS NULL
        THEN 1 ELSE 0 END) AS MissingProductKey,
    SUM(CASE WHEN CustomerKey IS NULL
        THEN 1 ELSE 0 END) AS MissingCustomerKey
FROM dbo.FactOrderLine;
```

**Observed result:** All four missing-value counts were zero.

The required identifiers checked in the fact table contain no `NULL` values.

### 4. Check Orphaned Product Keys

An orphaned product key exists when a fact-table product key has no matching record in the product dimension.

```sql
SELECT
    COUNT_BIG(*) AS OrphanProductKeys
FROM dbo.FactOrderLine AS f
LEFT JOIN dbo.DimProduct AS p
    ON f.ProductKey = p.ProductKey
WHERE p.ProductKey IS NULL;
```

**Observed result:** `0`

### 5. Check Orphaned Customer Keys

```sql
SELECT
    COUNT_BIG(*) AS OrphanCustomerKeys
FROM dbo.FactOrderLine AS f
LEFT JOIN dbo.DimCustomer AS c
    ON f.CustomerKey = c.CustomerKey
WHERE c.CustomerKey IS NULL;
```

**Observed result:** `0`

Both checks indicate that every fact-table product and customer key has a matching dimension record.

---

## Part 3: Inspect SQL Server Data Types

### 1. Retrieve the Existing Schema

```sql
SELECT
    TABLE_NAME,
    COLUMN_NAME,
    DATA_TYPE,
    CHARACTER_MAXIMUM_LENGTH,
    NUMERIC_PRECISION,
    NUMERIC_SCALE,
    IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME IN (
      'FactOrderLine',
      'DimProduct',
      'DimCustomer'
  )
ORDER BY
    TABLE_NAME,
    ORDINAL_POSITION;
```

### Observations

The schema inspection revealed several areas worth reviewing:

- Many text columns were created as `VARCHAR(MAX)`.
- Financial amounts were stored as `FLOAT`.
- Date keys were stored as `BIGINT` or `FLOAT`.
- Binary flags were stored as `FLOAT` or `BIGINT`.
- Every inspected column allowed `NULL`.
- Primary keys and indexes needed separate inspection.

These results are consistent with tables created using automatically inferred data types. The exact suitability of each type depends on the source data and business requirements.

### 2. Check Primary Keys and Indexes

```sql
SELECT
    t.name AS TableName,
    i.name AS IndexName,
    i.is_primary_key AS IsPrimaryKey,
    i.type_desc AS IndexType
FROM sys.tables AS t
LEFT JOIN sys.indexes AS i
    ON t.object_id = i.object_id
   AND i.index_id > 0
WHERE t.schema_id = SCHEMA_ID('dbo')
  AND t.name IN (
      'FactOrderLine',
      'DimProduct',
      'DimCustomer'
  )
ORDER BY
    t.name,
    i.is_primary_key DESC,
    i.name;
```

This query helps determine whether primary-key indexes and other indexes exist on the imported tables.

---

## Part 4: Validate Flags and Date Keys

### 1. Validate `OnTimeFlag`

```sql
SELECT
    OnTimeFlag,
    COUNT_BIG(*) AS TotalRows
FROM dbo.FactOrderLine
GROUP BY OnTimeFlag
ORDER BY OnTimeFlag;
```

#### Observed Results

| Value | Rows |
|---|---:|
| `NULL` | 142 |
| `0` | 650 |
| `1` | 8,317 |
| **Total** | **9,109** |

The column contains only `0`, `1`, and `NULL`.

**Recommended type:** `BIT NULL`, provided the source definition confirms that it is a binary flag.

### 2. Validate `IsPromoted`

```sql
SELECT
    IsPromoted,
    COUNT_BIG(*) AS TotalRows
FROM dbo.FactOrderLine
GROUP BY IsPromoted
ORDER BY IsPromoted;
```

#### Observed Results

| Value | Rows |
|---|---:|
| `0` | 2,919 |
| `1` | 6,190 |
| **Total** | **9,109** |

The column contains only `0` and `1` in the current data.

**Recommended type:** `BIT NOT NULL`, if the source definition confirms that missing values are not allowed.

### 3. Validate Delivery and Return Date Keys

```sql
SELECT
    'DeliveryDateKey' AS ColumnName,
    COUNT_BIG(*) AS TotalRows,
    COUNT_BIG(DeliveryDateKey) AS NonNullRows,
    MIN(DeliveryDateKey) AS MinValue,
    MAX(DeliveryDateKey) AS MaxValue,
    SUM(
        CASE
            WHEN DeliveryDateKey IS NOT NULL
             AND DeliveryDateKey <> FLOOR(DeliveryDateKey)
            THEN 1 ELSE 0
        END
    ) AS FractionalRows
FROM dbo.FactOrderLine

UNION ALL

SELECT
    'ReturnDateKey',
    COUNT_BIG(*),
    COUNT_BIG(ReturnDateKey),
    MIN(ReturnDateKey),
    MAX(ReturnDateKey),
    SUM(
        CASE
            WHEN ReturnDateKey IS NOT NULL
             AND ReturnDateKey <> FLOOR(ReturnDateKey)
            THEN 1 ELSE 0
        END
    )
FROM dbo.FactOrderLine;
```

#### Observed Results

| Check | `DeliveryDateKey` | `ReturnDateKey` |
|---|---:|---:|
| Total rows | 9,109 | 9,109 |
| Non-null rows | 8,967 | 1,046 |
| NULL rows | 142 | 8,063 |
| Minimum | 20240104 | 20240113 |
| Maximum | 20260105 | 20260128 |
| Fractional rows | 0 | 0 |

The values resemble date keys in `YYYYMMDD` format.

A separate validation using `TRY_CONVERT(DATE, ..., 112)` returned zero invalid non-null dates for both columns.

**Recommended type:** `INT NULL` if these columns are intended to be integer date keys. Use `DATE` instead if the intended model stores actual calendar dates.

---

## Part 5: Review Financial Data Types

### 1. Inspect Financial Ranges

```sql
SELECT
    COUNT_BIG(*) AS TotalRows,
    MIN(NetSales) AS MinNetSales,
    MAX(NetSales) AS MaxNetSales,
    MIN(GrossSales) AS MinGrossSales,
    MAX(GrossSales) AS MaxGrossSales,
    MIN(ContributionMargin) AS MinContributionMargin,
    MAX(ContributionMargin) AS MaxContributionMargin,
    MIN(UnitCost) AS MinUnitCost,
    MAX(UnitCost) AS MaxUnitCost
FROM dbo.FactOrderLine;
```

#### Observed Results

| Column | Minimum | Maximum |
|---|---:|---:|
| `NetSales` | 0 | 446.43 |
| `GrossSales` | 0 | 485.25 |
| `ContributionMargin` | -270.64 | 168.3444 |
| `UnitCost` | 7.15 | 122.75 |

Negative contribution margins can be valid when costs exceed contribution revenue.

### 2. Check Values Beyond Two Decimal Places

```sql
SELECT
    SUM(CASE WHEN NetSales <> ROUND(NetSales, 2)
        THEN 1 ELSE 0 END) AS NetSalesMoreThan2Decimals,
    SUM(CASE WHEN GrossSales <> ROUND(GrossSales, 2)
        THEN 1 ELSE 0 END) AS GrossSalesMoreThan2Decimals,
    SUM(CASE WHEN ContributionMargin <> ROUND(ContributionMargin, 4)
        THEN 1 ELSE 0 END) AS MarginMoreThan4Decimals,
    SUM(CASE WHEN UnitCost <> ROUND(UnitCost, 4)
        THEN 1 ELSE 0 END) AS UnitCostMoreThan4Decimals
FROM dbo.FactOrderLine;
```

#### Observed Results

| Column | Count |
|---|---:|
| `NetSales` differing from 2-decimal rounding | 4,461 |
| `GrossSales` differing from 2-decimal rounding | 94 |
| `ContributionMargin` differing from 4-decimal rounding | 6,667 |
| `UnitCost` differing from 4-decimal rounding | 0 |

Additional checks found:

| Column | Rows differing from 2-decimal rounding |
|---|---:|
| `DiscountAmount` | 4,998 |
| `RefundAmount` | 669 |
| `ShippingRevenue` | 0 |
| `ProductCost` | 1,012 |
| `ContributionRevenue` | 4,616 |
| `ReturnLoss` | 990 |

A further check found that 91 `LostSalesValue` rows differed from four-decimal rounding. The other six inspected cost columns had zero such rows.

**Important:** These checks do not prove that all extra decimal places are meaningful. Because the columns currently use `FLOAT`, small floating-point representation differences can produce nonzero counts. Compare the source workbook and business definitions before selecting final decimal scales.

### 3. Recommended Financial Types

`DECIMAL` is generally preferable to `FLOAT` for financial calculations because it stores values using a defined decimal precision and scale.

Candidate types for further validation include:

| Column group | Candidate type |
|---|---|
| `NetSales`, `GrossSales`, `DiscountAmount`, `RefundAmount` | `DECIMAL(18,4)` |
| `ShippingRevenue`, `ContributionRevenue` | `DECIMAL(18,4)` |
| `ProductCost`, `UnitCost` | `DECIMAL(18,4)` |
| `ContributionMargin`, `ReturnLoss`, `LostSalesValue` | `DECIMAL(18,4)` |
| Other variable cost columns | `DECIMAL(18,4)` pending source validation |
| `AppliedDiscountPct`, `ReturnRecoveryPct` | `DECIMAL(7,4)` as a candidate |

These are provisional choices. The final types should preserve the source's intended precision, support expected future values, and avoid unintended rounding.

### 4. Validate Percentage Columns

```sql
SELECT
    'AppliedDiscountPct' AS ColumnName,
    MIN(AppliedDiscountPct) AS MinValue,
    MAX(AppliedDiscountPct) AS MaxValue,
    SUM(
        CASE
            WHEN AppliedDiscountPct <> ROUND(AppliedDiscountPct, 4)
            THEN 1 ELSE 0
        END
    ) AS RowsBeyond4Decimals
FROM dbo.FactOrderLine

UNION ALL

SELECT
    'ReturnRecoveryPct',
    MIN(ReturnRecoveryPct),
    MAX(ReturnRecoveryPct),
    SUM(
        CASE
            WHEN ReturnRecoveryPct <> ROUND(ReturnRecoveryPct, 4)
            THEN 1 ELSE 0
        END
    )
FROM dbo.FactOrderLine;
```

#### Observed Results

| Column | Minimum | Maximum | Rows beyond 4 decimals |
|---|---:|---:|---:|
| `AppliedDiscountPct` | 0 | 0.35 | 0 |
| `ReturnRecoveryPct` | 0 | 1 | 0 |

The values appear to be stored as fractions rather than percentage points. For example, `0.35` represents 35%.

`DECIMAL(7,4)` is a reasonable candidate, subject to validation against the source data.

---

## Part 6: Review Text Column Lengths

Many text columns were created as `VARCHAR(MAX)`. The following checks measured the maximum lengths currently stored.

### 1. Fact Table

| Column | Maximum observed length | Suggested type |
|---|---:|---|
| `LineID` | 8 | `VARCHAR(20)` |
| `OrderID` | 7 | `VARCHAR(20)` |
| `ProductKey` | 4 | `VARCHAR(20)` |
| `CustomerKey` | 6 | `VARCHAR(20)` |
| `GeographyKey` | 4 | `VARCHAR(20)` |
| `PromotionKey` | 4 | `VARCHAR(20)` |
| `FulfillmentKey` | 3 | `VARCHAR(20)` |
| `SalesChannelKey` | 4 | `VARCHAR(20)` |
| `ReturnReasonKey` | 3 | `VARCHAR(20)` |
| `OrderStatus` | 20 | `VARCHAR(50)` |
| `StockStatus` | 23 | `VARCHAR(50)` |
| `Currency` | 3 | `CHAR(3)` if the values are consistently three-character currency codes |

### 2. Product Dimension

| Column | Maximum observed length | Suggested type |
|---|---:|---|
| `ProductName` | 27 | `VARCHAR(100)` |
| `Brand` | 13 | `VARCHAR(50)` |
| `Category` | 17 | `VARCHAR(50)` |
| `Department` | 20 | `VARCHAR(50)` |
| `Subcategory` | 12 | `VARCHAR(50)` |
| `BrandTier` | 7 | `VARCHAR(20)` |
| `WeightClass` | 6 | `VARCHAR(20)` |
| `Seasonality` | 12 | `VARCHAR(30)` |

### 3. Customer Dimension

| Column | Maximum observed length | Suggested type |
|---|---:|---|
| `CustomerLabel` | 14 | `VARCHAR(50)` |
| `CustomerSegment` | 11 | `VARCHAR(30)` |
| `LoyaltyTier` | 8 | `VARCHAR(20)` |
| `AcquisitionChannel` | 14 | `VARCHAR(30)` |
| `CohortMonth` | 7 | `VARCHAR(20)` |
| `CohortQuarter` | 7 | `VARCHAR(20)` |
| `PreferredCategory` | 17 | `VARCHAR(50)` |
| `AgeBand` | 5 | `VARCHAR(20)` |

These are suggested capacities based on current data. Choose `NVARCHAR` instead where Unicode support is required, and verify that future source values will fit the selected lengths.

---

## Part 7: Production-Oriented Improvements

The investigation identified several improvements to consider before using the database in a production environment.

### 1. Define the Schema Explicitly

Create tables with intentional data types, primary keys, foreign keys, nullability rules, and indexes rather than relying entirely on pandas inference.

### 2. Protect Existing Tables

Avoid dropping and recreating production tables on every import. Consider staging tables and controlled loading or merge operations.

### 3. Prevent Duplicate Loads

If using `append`, implement an appropriate deduplication or upsert strategy. A primary key or unique constraint can help enforce the intended key rules.

### 4. Use Suitable Financial Types

Choose `DECIMAL(p,s)` based on the source data and calculation requirements. Verify that conversions do not unintentionally round values.

### 5. Validate Data Before Loading

Check source row counts, nulls, duplicate keys, date validity, value ranges, and referential integrity before accepting a load.

### 6. Improve Loading Performance

Test `fast_executemany=True` with the SQLAlchemy pyodbc engine and a suitable batch size. Performance depends on the driver, data, and insertion strategy.

### 7. Add Error Handling and Logging

Record successful imports, failed sheets, inserted row counts, and validation results. Do not report a complete load if one or more worksheets fail.

---

## Final Project Status

| Area | Status |
|---|---|
| Excel-to-SQL Server import | Completed |
| Row-count checks | Completed |
| Duplicate-key checks | Completed for the tested keys |
| Missing-value checks | Completed for selected fact-table keys |
| Product/customer relationship checks | Completed; zero orphan rows |
| Flag and date-key validation | Completed |
| Financial type review | Initial review completed; source precision still needs confirmation |
| Text-length review | Completed for the inspected columns |
| Production schema redesign | Recommended; not yet applied |

**Important:** The investigation and validation queries above were run against the existing database. The proposed data-type changes have not been applied as part of this work. The database should be backed up, the source schema reviewed, and the changes tested before implementation.

## Conclusion

This project demonstrates that data ingestion is only the first step in a reliable analytics workflow. After loading data into SQL Server, it is important to validate row counts, key uniqueness, relationships, nullability, date representations, financial precision, and column definitions.

The result is a stronger foundation for SQL analysis and subsequent Power BI reporting.


This project is intended for educational purposes and demonstrates how to import selected Excel worksheets into Microsoft SQL Server using Python, Pandas, SQLAlchemy, and PyODBC.

Always review the code and verify your database configuration before running it against an existing database. Use a test database when experimenting with table replacement or other potentially destructive operations.
