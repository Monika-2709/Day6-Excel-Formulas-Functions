# 📊 Day 6 – Excel Formulas & Functions Fundamentals

## 📌 Project Overview

This project is part of my Data Analytics learning journey.

The objective of this task was to practice commonly used Microsoft Excel formulas and functions on a transactional sales dataset.

The project focuses on lookup functions, logical functions, conditional aggregation, counting, text manipulation, named ranges, and edge-case testing.

---

## 🎯 Objectives

- Practice commonly used Excel formulas for Data Analysis
- Understand lookup functions such as VLOOKUP and XLOOKUP
- Apply logical conditions using IF
- Perform conditional calculations using SUMIF and SUMIFS
- Count records using COUNTIF and COUNTIFS
- Practice text manipulation functions
- Understand INDEX and MATCH
- Use named ranges for better formula readability
- Test formulas against edge cases

---

## 🛠️ Tools Used

- Microsoft Excel
- Excel Tables
- Excel Formulas & Functions

---

## 📂 Dataset

The project uses a transactional sales dataset containing fields such as:

- Order ID
- Order Date
- Ship Date
- Customer Name
- Segment
- Category
- Region
- City
- Product Name
- Sales
- Quantity
- Discount
- Profit

The same dataset was continued from the previous learning task.

---

## 🔎 Functions Practiced

### Lookup Functions

- XLOOKUP
- VLOOKUP
- INDEX
- MATCH

### Logical Functions

- IF
- Nested IF

### Conditional Aggregation

- SUMIF
- SUMIFS

### Counting Functions

- COUNTIF
- COUNTIFS
- COUNTA

### Text Functions

- LEFT
- RIGHT
- LEN
- UPPER
- LOWER
- PROPER
- TEXTJOIN

### Other Concepts

- Named Ranges
- Excel Tables
- Edge Case Testing
- Input Validation

---

## 📊 Key Formula Examples

### XLOOKUP

```excel
=XLOOKUP(B5,SalesData[Order_ID],SalesData[Customer_Name],"Not Found")
