# 📊 Excel Data Analysis Project
### 🎓 Student Performance • 💰 Sales Analysis • 👥 Employee Analytics

> **Created by:** Vishwas Solanki  
> **Guidance by:** Girish Sir  
> **Project Type:** Microsoft Excel | Formula-Based Data Analysis  
> **Workbook:** `vishu.xlsx`

---

## 🎥 Project Presentation Video

🔗 **Watch the project explanation:** [▶️ Add Your Video https://drive.google.com/file/d/1FDh9kUKgiizYIWo3zBCcjHc-6mNZG7Vb/view?usp=drivesdk

> Replace `PASTE_YOUR_VIDEO_LINK_HERE` with your actual YouTube, Google Drive, or other shareable video URL.

---

## 🌟 About the Project

This Excel project demonstrates how spreadsheet formulas can turn raw data into useful information. The workbook contains **three practical sheets**—Student Data, Sales Data, and Employees Data. Each sheet focuses on a different real-world use case and applies formulas for calculations, logical checks, lookups, text manipulation, dates, and dynamic ranges.

### 🎯 Project Objectives
- 📌 Apply Excel formulas to practical datasets.
- 🧮 Automate calculations and reduce manual work.
- 🔍 Retrieve specific records and match data efficiently.
- 📈 Summarize information using criteria-based calculations.
- 🧹 Manipulate text and work with dates.
- ⚡ Explore dynamic and modern lookup functions.

---

## 🗂️ Workbook Overview

| Sheet | Focus | What it demonstrates |
|---|---|---|
| 🎓 `student_data` | Student performance | Total, average, grade, rank, age, conditions, counts, filtering and text functions |
| 💰 `sales_data` | Sales operations | Discount calculation, product lookup, availability, conditional totals, month-wise sales and dynamic ranges |
| 👥 `employees_data` | Employee analytics | Current date, years of service, salary lookup and dynamic cell references |

---

## 🎓 Sheet 1 — Student Data

**Purpose:** Analyze student marks and make student information easier to evaluate.

### ✨ Key work completed
- 🧮 Calculate each student's total marks and average.
- 🏅 Assign a grade based on the average.
- 📊 Calculate rank from total marks.
- 🔎 Display students whose average is above 70.
- 🎂 Calculate age from date of birth.
- ✅ Check whether both Maths and Science marks are above 80.
- 🔢 Count students meeting marks-based conditions.
- ✍️ Extract first names and convert names to uppercase or lowercase.
- 🔍 Use a lookup formula to retrieve a name based on a selected value.

### 🧠 Formula Reference — Student Data

| Function | Why it is used |
|---|---|
| `SUM` | Adds Maths, Science and English marks to calculate total marks. |
| `AVERAGE` | Calculates the average of the three subjects. |
| `IF` | Applies grade rules and returns a result based on a condition. |
| Nested `IF` | Assigns grades across multiple average-score ranges. |
| `COUNTIF` | Counts values that meet one condition; it is also used to calculate rank from total marks. |
| `FILTER` | Returns student names whose average is above 70. |
| `TODAY` | Supplies the current date for age calculation. |
| `DATEDIF` | Calculates age in completed years from date of birth. |
| `AND` | Checks that both Maths and Science marks are greater than 80. |
| `COUNTIFS` | Counts students whose average is above 60. |
| `LEFT` | Extracts characters from the beginning of a name. |
| `FIND` | Locates the space between first name and surname for extraction. |
| `UPPER` | Converts a name to uppercase. |
| `LOWER` | Converts a name to lowercase. |
| `VLOOKUP` | Looks up a value in a table and returns a related value. |

---

## 💰 Sheet 2 — Sales Data

**Purpose:** Explore sales records, calculate discounts, check product availability and summarize sales by conditions.

### ✨ Key work completed
- 💸 Calculate discounts using sales-amount thresholds.
- 🔎 Retrieve product-related values using product names or product codes.
- 📦 Check availability for selected products.
- 🌍 Calculate sales for a particular region and product.
- 📅 Calculate sales for a selected salesperson and month.
- 🎯 Find a product's position in a list.
- ⚡ Create a dynamic sales-amount range that adjusts to the number of filled records.

### 🧠 Formula Reference — Sales Data

| Function | Why it is used |
|---|---|
| `IF` | Applies discount rules based on sales amount. |
| Nested `IF` | Selects the appropriate discount percentage from several thresholds. |
| `OR` | Marks a product as available if it matches either of the specified product names. |
| `VLOOKUP` | Searches for a product or product code and returns a related value from the selected table range. |
| `SUMIFS` | Adds sales amounts that meet multiple criteria, such as region and product. |
| `SUMPRODUCT` | Calculates sales for a selected salesperson and month by combining conditions and amounts. |
| `MONTH` | Extracts the month number from a sales date. |
| `XMATCH` | Returns the position of a product in a list. |
| `OFFSET` | Creates a reference to a range that can change dynamically. |
| `COUNTA` | Counts non-empty cells to determine the size of the dynamic sales range. |

---

## 👥 Sheet 3 — Employees Data

**Purpose:** Review employee records, calculate service duration and retrieve salaries using lookup techniques.

### ✨ Key work completed
- 📅 Display the current date automatically.
- 🗓️ Calculate completed years of service from the joining date.
- 💵 Retrieve employee salaries using employee IDs.
- 🔍 Look up a salary using `INDEX` and `MATCH`.
- ⚡ Use `INDIRECT` to refer to a range based on a text reference.

### 🧠 Formula Reference — Employees Data

| Function | Why it is used |
|---|---|
| `TODAY` | Returns the current date and updates with the date on which Excel recalculates. |
| `DATEDIF` | Calculates completed years between joining date and current date. |
| `XLOOKUP` | Searches for an employee ID and returns the matching salary, with an optional not-found result. |
| `INDEX` | Returns a value from a specified range at a particular position. |
| `MATCH` | Finds the position of an employee ID within a list. |
| `INDIRECT` | Converts a text reference such as `"D2:D21"` into a usable cell/range reference. |

---

## 📚 Complete Function Summary

The table below brings together the main Excel functions used across the three sheets.

| Function | Main use | Sheet |
|---|---|---|
| `SUM` | Add subject marks | Student Data |
| `AVERAGE` | Calculate average marks | Student Data |
| `IF` / Nested `IF` | Apply conditions and grade/discount rules | Student Data, Sales Data |
| `COUNTIF` | Count matches or help calculate rank | Student Data |
| `FILTER` | Return names that meet a condition | Student Data |
| `TODAY` | Get the current date | Student Data, Employees Data |
| `DATEDIF` | Calculate age or completed service years | Student Data, Employees Data |
| `AND` / `OR` | Test multiple logical conditions | Student Data, Sales Data |
| `COUNTIFS` | Count records meeting criteria | Student Data |
| `LEFT` / `FIND` | Extract first name from full name | Student Data |
| `UPPER` / `LOWER` | Change text case | Student Data |
| `VLOOKUP` | Retrieve related values from a lookup table | Student Data, Sales Data |
| `SUMIFS` | Add values matching multiple criteria | Sales Data |
| `SUMPRODUCT` | Calculate conditional, month-wise sales | Sales Data |
| `MONTH` | Get the month number from a date | Sales Data |
| `XMATCH` | Find a matching item's position | Sales Data |
| `OFFSET` / `COUNTA` | Build a dynamic range | Sales Data |
| `XLOOKUP` | Retrieve employee salaries by ID | Employees Data |
| `INDEX` / `MATCH` | Find a value using a matched position | Employees Data |
| `INDIRECT` | Convert text into a cell/range reference | Employees Data |

---

## 🛠️ Tools & Skills Demonstrated

- 🟩 Microsoft Excel
- 🧮 Formula-based calculations
- 🧠 Logical functions and conditional analysis
- 🔎 Lookup and reference functions
- 📆 Date and time functions
- 🔤 Text functions
- 📊 Criteria-based summaries
- ⚡ Dynamic array and dynamic-range concepts

> **Compatibility note:** Some functions—especially `FILTER`, `XLOOKUP`, and `XMATCH`—require a supported modern version of Excel. Dynamic-array formulas may spill into nearby cells, so keep the output area clear.

---

## 🚀 Learning Outcomes

By completing this project, I practiced how to:
- Organize and analyze data across multiple worksheets.
- Choose an appropriate formula for a given task.
- Use conditions to classify and summarize records.
- Search for related information without manually scanning every row.
- Work with dates, text, and dynamic references.
- Build a foundation for more advanced Excel reporting and data analysis.

---

## 🙏 Acknowledgement

A sincere thank you to **Girish Sir** for the valuable guidance, support, and encouragement throughout this project. This project helped me strengthen my practical understanding of Microsoft Excel and its real-world applications.

## 💙 Thank You!

**Thank you for taking the time to explore my Excel Data Analysis Project!**  
✨ *Every formula is a small step toward smarter analysis and better decisions.* 📊🚀

---

<div align="center">

### 👨‍💻 Created with dedication by **Vishwas Solanki**
**Guidance by Girish Sir** 🎓

⭐ *Learn • Practice • Analyze • Improve* ⭐

</div>
