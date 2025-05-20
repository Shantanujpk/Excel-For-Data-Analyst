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
