# DecodeLabs Data Analytics Track | Project 1: Data Cleaning & Integrity Pipeline

Name:Sandhia Kumari  
Program:DecodeLabs Industrial Training Program (Batch 2026)  
Project:Data Cleaning, Preparation, and Verification  

---

## What This Project Does

Before doing any analysis or building dashboards, data needs to be cleaned and validated. This project takes the raw e-commerce orders dataset (`Dataset_for_Data_Analytics.xlsx`) and cleans it using Python (`pandas` and `openpyxl`). 

Instead of deleting rows or making assumptions, I audited the file cell-by-cell, fixed formatting and precision issues, imputed missing data logically, and built a live Excel verification sheet to prove the dataset is 100% clean.

---

## Dataset Overview

* **Input File:** `Dataset_for_Data_Analytics.xlsx`
* **Size:** 1,200 rows × 14 columns
* **Date Range:** January 1, 2023 to June 30, 2025
* **Columns:** `OrderID`, `Date`, `CustomerID`, `Product`, `Quantity`, `UnitPrice`, `ShippingAddress`, `PaymentMethod`, `OrderStatus`, `TrackingNumber`, `ItemsInCart`, `CouponCode`, `ReferralSource`, `TotalPrice`

---

## Initial Findings & Data Issues

When checking the raw dataset, the core structure was mostly solid (no duplicate orders, no broken dates, and no trailing whitespace). However, I found two main issues:

1. **Missing Coupon Codes:** 309 rows (25.75% of the file) had blank `CouponCode` cells. This was the only column with missing values.
2. **Floating-Point & Number Formatting Issues:** 
   * 29 `TotalPrice` cells had floating-point artifacts (e.g. `769.3799999999999`).
   * 36 price entries (9 in `UnitPrice`, 27 in `TotalPrice`) were stored as plain integers instead of 2-decimal numbers.

> **Note on Row Retention:** I did not delete a single row. All 1,200 original records were preserved.

---

## Change Log

| ID | Task | What I Did | Status |
| :--- | :--- | :--- | :--- |
| **CR001** | Missing Values | Replaced 309 blank `CouponCode` entries with `"No Coupon"`. | Resolved |
| **CR002** | Floating-Point Fix | Rounded 29 `TotalPrice` values to 2 decimal places. | Resolved |
| **CR003** | Currency Format | Standardized all `UnitPrice` and `TotalPrice` entries to 2 decimal places. | Resolved |
| **CR004** | Date Check | Verified all dates are valid ISO 8601 (`YYYY-MM-DD`) with no time stamps. | Verified |
| **CR005** | Duplicate Check | Checked `OrderID`, `TrackingNumber`, and full rows for duplicates. Found 0. | Verified |
| **CR006** | Text Clean | Checked for stray whitespace and Proper Case formatting on text fields. | Verified |

*(Resolved = issue found and fixed. Verified = checked, no issue found.)*

---

## Why "No Coupon" Instead of Mode or Deletion?

When handling the 309 missing `CouponCode` values, I considered standard statistical imputation methods but decided against them:

* **Why not Mode?** Filling blank cells with the most common code (`FREESHIP`) would invent 309 fake promotional discounts that customers never actually used.
* **Why not Mean/Median?** `CouponCode` is text/categorical, so mathematical averages don't make sense.
* **Why not delete the rows?** Dropping 309 rows would delete over a quarter of the dataset (25.75%) for no good reason.
* **Why "No Coupon"?** A blank cell simply means the customer bought something without applying a discount code. Marking it `"No Coupon"` keeps the data honest and lets us compare coupon users vs. non-coupon users later.

---

## Validated Outliers & Repeat Customers

During the audit, two things looked unusual at first glance but turned out to be completely valid after verification:

* **11 Repeat Customers:** 11 `CustomerID`s appear twice. In every case, the `OrderID`, `TrackingNumber`, date, quantity, and unit prices were different—meaning these are genuine repeat orders, not duplicate entries.
* **8 Large Orders:** 8 order totals ranged between $3,334.00 and $3,456.40. Checking these manually confirmed that each order had `Quantity = 5` with unit prices around $667–$691, perfectly matching `TotalPrice = Quantity × UnitPrice`.

---

## Verification Gate (Live Excel Formulas)

DecodeLabs requires proving **zero duplicate IDs** and **zero improperly formatted dates** before moving to Project 2.

I added a `Verification` tab inside `Cleaned_Dataset.xlsx` that uses dynamic Excel formulas against the cleaned data:

| Test Name | Formula Used | Result |
| :--- | :--- | :--- |
| **Duplicate OrderIDs** | `=SUMPRODUCT(--(COUNTIF(OrderID, OrderID) > 1))` | **0 (PASS)** |
| **Duplicate Tracking Numbers** | `=SUMPRODUCT(--(COUNTIF(TrackingNumber, TrackingNumber) > 1))` | **0 (PASS)** |
| **Non-Date Values** | `=SUMPRODUCT(--NOT(ISNUMBER(Date)))` | **0 (PASS)** |
| **Time Component Check** | `=SUMPRODUCT(--(Date <> INT(Date)))` | **0 (PASS)** |
| **Out-of-Range Dates** | Dates outside 2023-01-01 to 2025-06-30 | **0 (PASS)** |
| **Blank Cells** | `=COUNTBLANK(Cleaned_Data_Range)` | **0 (PASS)** |
| **Price Arithmetic** | `=SUMPRODUCT(--(ABS(Quantity * UnitPrice - TotalPrice) > 0.005))` | **0 (PASS)** |
| **Stray Whitespace** | `=SUMPRODUCT(--(LEN(cell) <> LEN(TRIM(cell))))` | **0 (PASS)** |
| **Non-Positive Values** | Quantity or UnitPrice ≤ 0 | **0 (PASS)** |

**Overall Result:** **PASS**.  
I also negative-tested the formulas by intentionally pasting a duplicate ID and an unformatted date; the verification sheet immediately flagged them as `FAIL`, confirming the checks work as expected.

---

## Repository Files

```text
.
├── Dataset_for_Data_Analytics.xlsx   # Original raw dataset
├── Cleaned_Dataset.xlsx              # Cleaned data, Verification sheet, and Imputed_Rows tab
├── Change_Log_Project1.pdf           # Detailed change log report
├── clean_dataset.py                  # Python cleaning script
└── README.md                         # Project documentation
