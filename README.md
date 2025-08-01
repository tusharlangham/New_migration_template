# Task


# New_migration_template
Updating UK transaction capture template to LUX TCC


## Need to create a new migration template
from the earlier template the FX rates and the account code sheet is not pulled correctly, in order to rectify that we are doing certain upgradation in the TCC file.


# Process Know-how
## Understanding the process 
How UK template is generated through VBA

# Checking for communication 
  Any changes made in the addin file will give us an idea about the changes made in the TCC file which are generated.
Once we have this, we will know that how the changes are done,
Going Tab by tab, Making a list of requirements 

# Requirements Gathering

List of requirement - Requirement Gathering as per tab 
Get review from all of the requirements which we have gathered.

# Difference between a UK template, and a LUX template.
  where changes are needed.

# EXplaination of each modules


# VBAFunctions TC MigrateV2
Extracting data from sun journals to TCC migration template

Here’s a clear **README** for the VBA module you’ve written, structured in **3 concise steps** explaining its purpose, functionality, and usage timing:

---

# 📘 VBA Module: Hyperlink Generator & Sheet Cleaner

### ✅ **Why This Is Needed**

When working with long lists of file paths in Excel, you often need:

* Clickable hyperlinks to files or folders
* Clean worksheets free of unnecessary rows and columns
  Doing this manually is time-consuming and error-prone — this macro automates the process.

---

### ⚙️ **What It Does**

This module contains **3 macros**:

1. **`CreateHyperlink()`**

   * Converts the content of each selected cell (up to 2000 rows) into a hyperlink pointing to the same path.

2. **`CreateFolderHyperlink()`**

   * Converts file paths into hyperlinks pointing to their **parent folders** using a helper function `GetFolderpath()` (you need to define it).
   * Also replaces the cell content with the folder path.

3. **`DeleteRows()`**

   * Deletes all rows after row 1000 and all columns after column "Z" to reduce file size and improve performance.

---

### 📌 **When to Use It**

Use this module when:

* You receive a list of file/folder paths and want to make them clickable.
* You need links to **folder locations**, not individual files.
* Your Excel file is bloated due to extra unused rows and columns.
* You want to streamline your workflow and improve Excel performance.






