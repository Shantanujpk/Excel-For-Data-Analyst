📘 Chapter 2: Basic Reporting and Data Entry Operations
🎯 Goal:
Learn foundational Excel functions for data entry, reporting, and formula management.
🧠 Topics Covered
1. Automatic Insertion of Numbers
Use fill handle (drag corner of a cell).
Use =ROW() or =SEQUENCE() for dynamic sequences.
Shortcut:
Ctrl + D – Fill Down
Ctrl + R – Fill Right

2. Date Formats
Format dates using built-in styles: short date, long date, or custom (dd-mmm-yyyy).
Shortcut:
Ctrl + ; – Insert current date
Ctrl + Shift + ; – Insert current time
Ctrl + 1 – Format Cells

3. Currency Change
Change number format to show currency (e.g., $, €, ₹).
Shortcut:
Ctrl + Shift + $ – Apply Currency Format
Ctrl + 1 – Open Format Options

4. Basic Math Formulas
=SUM(range) – Add values
=MIN(range) – Find minimum
=MAX(range) – Find maximum
=AVERAGE(range) – Calculate average
=COUNT(range) – Count numeric entries
=RANDBETWEEN(x, y) – Generate random number
Shortcut:
= – Start a formula
Ctrl + Shift + Enter – Array formula
F2 – Edit formula

5. Conditional Formatting
Apply rules to change cell color based on values (e.g., greater than, duplicate).
Shortcut:
Alt + H + L – Open Conditional Formatting Menu
6. Top Performers
Highlight top 10 items or top % using conditional formatting preset.

7. Advanced Conditional Formatting
Use custom formulas like =A1>AVERAGE($A$1:$A$10) to apply logic-driven formatting.

8. Share Data Without Sharing Formula
Use Paste Values to only share the result, not the formula.
Shortcut:
Ctrl + C – Copy
Ctrl + Alt + V → V – Paste Values
Ctrl + Z – Undo
Ctrl + Y – Redo

9. Relative and Absolute References
🔸 Relative Reference
Formula adjusts when copied.
Example: =A1 + B1 → Copied down becomes =A2 + B2.

🔹 Absolute Reference (F4)
Does not change when copied.
Use $ to lock:
$A$1 → Lock column and row
A$1 → Lock row only
$A1 → Lock column only
A1 → Relative (default)
Shortcut:
F4 – Toggle between reference types

✅ Bonus Navigation Tips
Tab – Move right
Shift + Tab – Move left
Ctrl + Arrow Key – Jump to edge of data
`Ctrl + `` – Show/Hide formulas
Ctrl + B – Bold
Ctrl + Shift + % – Apply Percentage Format


📘 Chapter 3: Data Cleaning and Preparation
Before performing any kind of meaningful data analysis, it is essential to clean and prepare the data. Dirty or inconsistent data can lead to inaccurate results and poor insights. This chapter outlines the key steps involved in preparing a dataset using Excel functions and formulas.

✅ Overview of Steps
The core actions involved in data cleaning and preparation are:
Remove Duplicates
Remove Blank Rows
Remove Blank Spaces
Remove Unbreakable (Non-breaking) Spaces
Fix Text Case
Fix Negative Stock Values
Split Data
Data Validation

Each step is described in detail below:
1. Remove Duplicates
Purpose: Eliminate duplicate rows to avoid double-counting or data skewing.
Excel Steps:
Go to the Data tab → Click on Remove Duplicates
Select the columns to check for duplication.

3. Remove Blank Rows
Purpose: Prevent gaps in your data table that can break formulas and analyses.
Excel Steps:
Use filters to find and delete empty rows, or apply sorting to bring blanks together and delete manually.

3. Remove Blank Spaces
Purpose: Clean up unwanted leading, trailing, or extra spaces that interfere with sorting, filtering, or lookups.
Function Used:
=TRIM(A1)
Use in a new column to clean the text, then paste as values.

4. Remove Unbreakable (Non-breaking) Spaces
Purpose: Some copy-pasted text contains non-breaking spaces (ASCII 160), which TRIM() won't remove.
Function Used:
=SUBSTITUTE(A1, CHAR(160), "")
This replaces non-breaking spaces with normal ones or removes them.

5. Fix Text Case (Uppercase/Lowercase/Proper Case)
Purpose: Standardize text for consistency (e.g., names or categories).

Function Used:
=PROPER(A1)
You can also use UPPER() or LOWER() as needed.

6. Fix Negative Stock Values
Purpose: Clean invalid or illogical data entries (like negative stock quantities).

Logic Used:
Use a conditional formula like:
=IF(A1 < 0, 0, A1)
This replaces negative values with 0 or a custom default.

7. Split Data
Purpose: Separate concatenated data (e.g., FirstName LastName or City, State).

Tools:
Flash Fill: Select the first cell → Press Ctrl + E
Formulas:
=LEFT(), =RIGHT(), =MID()
Or =TEXTSPLIT() in newer Excel versions
8. Data Validation
Purpose: Ensure that only valid data types or entries are allowed (e.g., dropdown lists, number limits).
Steps:
Select range → Go to Data tab → Click Data Validation

Set rules such as whole numbers only, date ranges, or custom formulas.

📌 Summary
Task	Tool/Formula
Remove Duplicates	Data → Remove Duplicates
Remove Blank Rows	Manual filter/sort/delete
Remove Blank Spaces	=TRIM()
Remove Non-breaking Spaces	=SUBSTITUTE() with CHAR(160)
Change Case	=PROPER(), =UPPER(), =LOWER()
Fix Negative Values	=IF(value < 0, fix, value)
Split Data	Ctrl + E, LEFT, RIGHT, TEXTSPLIT
Data Validation	Data → Data Validation


📘 Chapter 4: Excel Logic Functions - IF, IFS, AND, OR (Beginner-Friendly Guide)
✨ Purpose of This Sheet
This Excel sheet helps beginners understand how to use IF, IFS, AND, and OR functions in Excel to calculate incentives and logical decisions based on data.
🔹 Part 1: Incentive Calculation Using IF and IFS
📊 Data Overview:
We have a list of employees with:
Emp ID
Name
Department
Sales amount
Target status (Complete or Incomplete)

🎯 Goal:
Calculate Incentive based on whether the target is Complete or Incomplete.
✅ Logic Used:
1. IF Formula
=IF([Target Status]="Complete", [Sales]*15%, [Sales]*10%)
This means:
If the target is "Complete", give 15% of sales as incentive.
If not, give 10% of sales.
2. IFS Formula
=IFS([Target Status]="Complete", [Sales]*15%, [Target Status]="Incomplete", [Sales]*10%)
This does the same as above but uses IFS instead of nested IF
✅ Both formulas give the same output, just written differently.

🔹 Part 2: Using AND and OR for Logical Conditions
📊 Data Overview:
We have:
Department
Working years
Salary
Promotion Status (Yes/No)
Separate columns showing use of AND and OR

🎯 Goal:
Understand how AND and OR can be used in Excel to check multiple conditions.
✅ AND Function:
Returns TRUE only if ALL conditions are true
Example:
=IF(AND(Department="Finance", WorkingYears>=4), "Eligible", "Not Eligible")
✔️ This will return "Eligible" only if:
Department is Finance
AND Working Years is 4 or more

✅ OR Function:
Returns TRUE if ANY one condition is true

Example:
=IF(OR(Department="Finance", WorkingYears>5), "Eligible", "Not Eligible")
✔️ This will return "Eligible" if:
Department is Finance OR
Working Years is more than 5

🔄 Combined AND and OR:
You can also use both together:
=IF(AND(Department="Finance", OR(Promotion="Yes", WorkingYears>5)), "Yes", "No")
This checks:
Is department Finance
AND either Promotion is Yes or Working Years > 5

🧠 Summary Table
Function	Meaning	Example	Result
IF	One condition	=IF(A1>50, "Pass", "Fail")	Checks one condition
IFS	Multiple conditions	=IFS(A1>90, "A", A1>80, "B")	Multiple checks
AND	All must be true	=AND(A1>50, B1="Yes")	True only if both
OR	Any one is true	=OR(A1>50, B1="Yes")	True if any one is true



