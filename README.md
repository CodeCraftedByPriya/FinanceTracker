# OVERVIEW
- This VBA macro provides a Financial Tracker tool for recording and summarizing financial transactions. It manages both expenses and earnings in an Excel workbook, organizes them by categories, and generates summaries to aid financial management.

# FEATURES
- Data Entry: Allows users to input financial transactions (type, category, amount, and date).
- Dynamic Category Validation: Ensures that categories are valid before entry.
- Automated Summary: Generates reports for:
        - Category-wise Expenses and Earnings
        - Monthly Net Summary
- Color-coded Entries: Highlights earnings in green for easy distinction.
- Dynamic Table Formatting: Organizes data in structured, auto-formatted tables for clarity.

# HOW IT WORKS
### 1. Data Sheet Creation
- Ensures the Data worksheet exists, creating it if necessary.
- Sets up headers (Date, Type, Category, Amount) in the Data sheet.
### 2. User Input
- Prompts the user for:
        - Transaction Type: 'Expense' or 'Earnings.'
        - Category: Validates against a predefined list.
        - Amount: Ensures a positive value.
        - Date: Defaults to the current date if none is provided or validates user input for valid dates.
### 3. Entry Recording
- Appends validated inputs to the Data worksheet.
- Earnings amounts are highlighted in green.
### 4. Summary Generation
- Summarizes data in the Summary worksheet:
        - Category Totals: Lists total expenses and earnings by category.
        - Monthly Summary: Tracks net totals (earnings minus expenses) for each month.
### 5. Table Formatting
- Organizes data and summaries into Excel tables for better readability.
