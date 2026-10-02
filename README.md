# EXCEL-PROJECT1
# 📊 Excel Formula & Data Analysis Project

> **From raw data to meaningful insights — one formula at a time.** ✨

This project is a practical Excel workbook created to demonstrate the use of **formulas, lookup functions, conditional logic, text manipulation, date analysis, and dynamic data filtering**.

The workbook contains three main datasets — **Student Performance, Product Sales, and Employee Information** — with formulas applied to analyze and transform the data efficiently.

---

## 🎯 Project Objective

The main objective of this project is to understand how Excel formulas can be used to:

* Organize and analyze structured data
* Classify student performance
* Search and retrieve information dynamically
* Calculate salaries and sales values
* Analyze employee joining dates
* Perform conditional calculations
* Extract specific records automatically
* Work with text, dates, and dynamic references

---

## 📁 Workbook Structure

### 🎓 1. Student Sheet

The **Student** sheet focuses on marks, grades, results, and performance analysis.

**Data included:**

* Student ID
* Student Name
* Marks
* Grade
* Result
* Performance
* Eligibility

### Formulas Used

**Nested IF – Grade Classification**

```excel
=IF(C4>=90,"A+",IF(C4>=80,"A",IF(C4>=70,"B",IF(C4>=60,"C","F"))))
```

**IF – Pass/Fail**

```excel
=IF(C4>=40,"PASS","FAIL")
```

**IF + OR – Performance**

```excel
=IF(OR(C4>=90,D4="A+"),"EXCELLENT","NORMAL")
```

**IF + AND – Eligibility**

```excel
=IF(AND(C4>=60,C4>=40),"ELIGIBLE","NOT ELIGIBLE")
```

**COUNTIFS – Students scoring above 60**

```excel
=COUNTIFS(C4:C13,">60")
```

**FILTER – Students scoring above 80**

```excel
=FILTER(A4:D13,C4:C13>80,"No students found")
```

---

### 🛍️ 2. Product Sheet

The **Product** sheet demonstrates product sales calculations and lookup operations.

**Data included:**

* Product ID
* Product Name
* Category
* Price
* Quantity
* Amount
* Total

### Formulas Used

**Amount Calculation**

```excel
=D2*E2
```

**VLOOKUP – Product Price**

```excel
=VLOOKUP(A13,A2:D8,4,FALSE)
```

This searches for a Product ID and returns its corresponding price.

---

### 👩‍💼 3. Employee Sheet

The **Employee** sheet demonstrates employee salary calculations, joining-date analysis, lookup functions, and text manipulation.

**Data included:**

* Employee ID
* Employee Name
* Department
* Salary
* Joining Date
* Year
* Month
* Day
* Years of Service
* Tax Rate
* Tax

### Date & Time Functions

**Year**

```excel
=YEAR(E4)
```

**Month**

```excel
=MONTH(E4)
```

**Day**

```excel
=DAY(E4)
```

**Years of Service**

```excel
=DATEDIF(E4,TODAY(),"Y")
```

---

## 🔎 Lookup Functions

### INDEX + MATCH

```excel
=INDEX($C$4:$C$10,MATCH(A13,$A$4:$A$10,0))
```

Used to dynamically find an employee's department.

### XMATCH

```excel
=XMATCH(D13,A4:A10,0)
```

Used to find the position of an employee within the employee list.

### XLOOKUP

XLOOKUP can be used to dynamically retrieve employee information such as salary:

```excel
=XLOOKUP(G2,A4:A10,D4:D10,"Not Found")
```

---

## 📝 Text Functions

The Employee sheet also demonstrates different text manipulation techniques.

| Function   | Purpose                            |
| ---------- | ---------------------------------- |
| `UPPER()`  | Converts text to uppercase         |
| `LOWER()`  | Converts text to lowercase         |
| `PROPER()` | Converts text to proper case       |
| `LEFT()`   | Extracts characters from the left  |
| `RIGHT()`  | Extracts characters from the right |
| `LEN()`    | Counts characters                  |

Example:

```excel
=UPPER(A17)
```

```excel
=LOWER(A17)
```

```excel
=PROPER(A17)
```

```excel
=LEFT(A17,3)
```

```excel
=RIGHT(A17,3)
```

```excel
=LEN(A17)
```

---

## 💰 Salary Analysis

### SUMIFS – Total IT Salary

```excel
=SUMIFS(D4:D10,C4:C10,"IT")
```

### AVERAGEIFS – Average IT Salary

```excel
=AVERAGEIFS(D4:D10,C4:C10,"IT")
```

### SUMIFS – Total HR Salary

```excel
=SUMIFS(D4:D10,C4:C10,"HR")
```

These functions allow salary information to be analyzed according to department.

---

## 🔄 Dynamic Reference Functions

### OFFSET

```excel
=OFFSET(B2,2,0)
```

OFFSET creates a dynamic reference by moving from a selected starting cell.

### INDIRECT

INDIRECT can convert text representing a cell reference into an actual reference.

Example:

```excel
=INDIRECT("B2")
```

---

## 🧠 Excel Skills Demonstrated

This project covers:

* ✅ Relative References
* ✅ Absolute References
* ✅ Cell Formatting
* ✅ Data Input
* ✅ IF Functions
* ✅ Nested IF
* ✅ IF with AND
* ✅ IF with OR
* ✅ COUNTIFS
* ✅ SUMIFS
* ✅ AVERAGEIFS
* ✅ VLOOKUP
* ✅ XLOOKUP
* ✅ XMATCH
* ✅ INDEX + MATCH
* ✅ TEXT Functions
* ✅ Date & Time Functions
* ✅ YEAR / MONTH / DAY
* ✅ DATEDIF
* ✅ FILTER
* ✅ INDIRECT
* ✅ OFFSET
* ✅ Basic Mathematical Calculations

---

## 📌 Main Practical Tasks

| Task                                | Excel Technique           |
| ----------------------------------- | ------------------------- |
| 🎓 Classify student grades          | IF / Nested IF            |
| 📈 Identify student performance     | IF + AND/OR               |
| 🔢 Count marks above 60             | COUNTIFS                  |
| 🛍️ Find product price              | VLOOKUP                   |
| 👩‍💼 Retrieve employee information | XLOOKUP                   |
| 📅 Analyze joining dates            | YEAR, MONTH, DAY, DATEDIF |
| 🔍 Search dynamically               | INDEX + MATCH / XMATCH    |
| ✨ Manipulate names                  | TEXT Functions            |
| 💰 Analyze departmental salaries    | SUMIFS / AVERAGEIFS       |
| 📋 Extract high-scoring students    | FILTER                    |
| 🔄 Create dynamic references        | INDIRECT / OFFSET         |

---

## 🌟 Key Learning Outcome

Through this project, I learned how Excel can transform simple tables into an **interactive data-analysis system** using formulas and functions.

Instead of manually calculating every value, formulas make the workbook **faster, more accurate, reusable, and dynamic**.

> **Input → Formula → Analysis → Insight 📊**

---

## 🛠️ Tools Used

* **Microsoft Excel**
* Excel Formulas & Functions
* Conditional Logic
* Lookup & Reference Functions
* Data Analysis
* Date & Time Functions
* Text Manipulation

---

## 📂 Project File

**Excel Workbook:** `prj1exl.xlsx`

### 👩‍🎓 Project Type

**Excel Practical / Formula & Data Analysis Project**

**author
NENSI RAY
---

### ✨ Conclusion

This project demonstrates the practical application of Excel functions for **student analysis, product lookup, employee management, salary analysis, and dynamic data extraction**.

It provides hands-on experience with both basic and advanced Excel formulas and shows how formulas can be combined to solve real-world data problems efficiently.
