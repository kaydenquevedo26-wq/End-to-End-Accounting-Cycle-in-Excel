# End-to-End-Accounting-Cycle-in-Excel
**End-to-End Accounting Cycle in Excel**

A full accounting cycle for a simulated company, built entirely in Excel from raw, messy transaction data through a complete set of linked financial statements. Built to demonstrate hands-on Excel and accounting skills: Power Query, VBA macros, SUMIFS, INDEX/MATCH, and double-entry accounting. All data used in this project is simulated for demonstration purposes.

**What's in the workbook:**

The_Final_Project.xlsm walks through the full accounting cycle for a simulated company:

**Messy/raw data:** the original, inconsistent transaction data before any cleanup
Power Query cleaning: raw data is capitalized, trimmed, split, merged, and de-duplicated into one consistent format

**Transaction Log:** 115+ transactions in a clean, structured format, formatted automatically with a VBA macro

**T-Accounts:** beginning balances, March activity (via SUMIFS), and ending balances for every account

**Trial Balance:** pulls every account balance in automatically using absolute references and INDEX/MATCH, with adjusting entries applied

**Financial Statements:** a linked Income Statement, Balance Sheet, and Cash Flow Statement, all built with absolute references so nothing has to be re-entered by hand. Assets = Liabilities + Equity, and ending cash on the Cash Flow Statement matches the Balance Sheet.

The workbook has its own ReadMe tab with more in-depth, step-by-step explanations of how Power Query was used, how the functions work, and how the macro was built.
