# 📊 Microsoft Excel Practice Labs

## 📌 Project Overview

This repository contains a structured collection of **Microsoft Excel practical labs** designed to build and demonstrate fundamental spreadsheet and data-management skills.

The exercises progress from **basic Excel operations and cell references** to more practical data-handling techniques such as **Paste Special, Excel Tables, sorting and filtering, and formula-based Data Validation**.

The labs emphasize hands-on practice with Excel's built-in features and formulas to improve **data organization, accuracy, consistency, and efficiency**. They also demonstrate how Excel can be used to apply rules to data, create controlled inputs, and prevent common data-entry errors such as duplicate values.

This repository serves as a practical record of my learning and provides examples of applying Excel concepts to structured datasets.

---

## 📁 Repository Contents

| File           | Description                                                             |
| -------------- | ----------------------------------------------------------------------- |
| `Lab 1.1.xlsx` | Excel fundamentals, formatting, formulas, and cell references           |
| `Lab 1.2.xlsx` | Data manipulation, Excel Tables, sorting/filtering, and Data Validation |

---

## 🧪 Lab 1.1 — Excel Fundamentals

The first lab focuses on developing a strong foundation in Microsoft Excel and understanding how formulas behave when applied across different cells.

### Topics Covered

* Creating and managing worksheets
* Renaming and organizing sheets
* Cell and data formatting
* Relative cell references
* Absolute cell references
* Mixed cell references
* Basic Excel formulas
* `SUM()` and `PRODUCT()` functions
* Creating structured tables
* Applying formulas across multiple cells
* Using absolute references for calculations

### Cell Reference Examples

**Relative Reference**

```excel
=B6*C6
```

The reference changes when the formula is copied to another cell.

**Absolute Reference**

```excel
=B6*$C$2
```

The referenced cell remains fixed when the formula is copied.

**Mixed Reference**

```excel
=B6*C$2
```

The row is fixed while the column can change.

---

## 🧪 Lab 1.2 — Data Management & Validation

The second lab focuses on practical data-management features available in Microsoft Excel.

### 📋 Paste Special

The lab includes practice with:

* Paste Values
* Paste Formulas
* Paste Formatting
* Transpose
* Mathematical operations using Paste Special

### 📊 Excel Tables

* Converting data ranges into Tables
* Applying Table Styles
* Sorting and filtering table data
* Using Total Row
* Creating calculated columns

### 🔍 Sorting & Filtering

* Ascending and descending sorting
* Multi-column sorting
* AutoFilter
* Custom filtering
* Filtering records based on conditions

### ✅ Data Validation

* Creating drop-down lists
* Restricting numerical input
* Creating custom validation rules
* Formula-based Data Validation
* Custom error messages
* Validating unique values
* Preventing duplicate entries

---

## 🧮 Key Excel Formulas

### Sum

```excel
=SUM(B6:C6)
```

Adds values from the specified cells.

### Product

```excel
=PRODUCT(B6,C6)
```

Multiplies the specified values.

### Range Validation

```excel
=AND(A2>=1,A2<=100)
```

Allows only values between 1 and 100.

### Unique Value Validation

```excel
=COUNTIF($A$2:A2,A2)=1
```

Ensures that a value is entered only once within the specified range.

---

## 🎯 Learning Objectives

Through these practical labs, I developed hands-on experience with:

* **Excel worksheet management**
* **Data formatting and organization**
* **Relative, absolute, and mixed references**
* **Excel formulas and functions**
* **Data manipulation**
* **Excel Tables**
* **Sorting and filtering**
* **Data Validation**
* **Drop-down lists**
* **Custom validation formulas**
* **Duplicate prevention**
* **Data accuracy and consistency**

---

## 🛠️ Tools & Technologies

| Tool / Feature          | Usage                                       |
| ----------------------- | ------------------------------------------- |
| **Microsoft Excel**     | Spreadsheet development and data management |
| **Excel Formulas**      | Calculations and logical validation         |
| **Data Validation**     | Controlling and validating data entry       |
| **Excel Tables**        | Structured data organization                |
| **Sorting & Filtering** | Data exploration and management             |
| **Paste Special**       | Controlled data transformation              |

---

## 📚 Skills Demonstrated

This project demonstrates practical knowledge of:

**Spreadsheet Fundamentals → Data Management → Formula Usage → Data Validation → Data Quality**

These skills provide a foundation for further work in **data analysis, reporting, dashboards, and business intelligence**.

---

## 🚀 Future Enhancements

Future Excel practice may include:

* Advanced Excel formulas
* Conditional Formatting
* Lookup functions such as `VLOOKUP`, `XLOOKUP`, and `INDEX-MATCH`
* Pivot Tables
* Pivot Charts
* Interactive dashboards
* Data cleaning
* Advanced filtering
* Excel-based data analysis

---

## 📂 Repository Structure

```text
Excel-Practice/
│
├── README.md
├── Lab 1.1.xlsx
└── Lab 1.2.xlsx
```

---

## ⭐ Purpose

This repository is part of my ongoing learning journey with **Microsoft Excel**, documenting practical exercises and demonstrating my ability to work with spreadsheets, formulas, structured data, and data-validation techniques.
