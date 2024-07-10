Excel Fuzzy Matcher with Learning Capability
Automated tool for fuzzy matching with learning elements

Description
Excel Fuzzy Matcher with Learning Capability is a tool for automatic data matching using fuzzy search and learning capabilities. The tool is designed to simplify the process of matching search queries with existing database records and can learn based on user corrections.

Key Features
Automatic Data Analysis: Excel automatically reads the most recently modified document in the specified folders.
Fuzzy Matching: The tool uses fuzzy search algorithms to match strings in the database.
Learnable Matching Table: The internal Power Query tool uses a "From-To" matching table to improve search accuracy. Manual search is simplified with a separate VBA form.
Learning from Corrections: Using VBA, the matching table is automatically updated based on user reviews and corrections, enhancing search quality for future queries.

Benefits
Simplifying Routine Tasks: Reduces manual data matching efforts.
Increased Accuracy: Fuzzy search algorithms improve result accuracy.
Automatic Learning: The tool remembers previous corrections and gets smarter with each use.

Use Cases
Matching search queries in price lists.
Processing specific product names for discount calculations.
Facilitating work with large volumes of data with various name variations.

Requirements
Microsoft Excel 2019 or later.
Activated Power Query feature.
Enabled macros for running VBA scripts.


How to Use
Ensure Tool is in Separate Directory:
Before using the tool, it is recommended to ensure that the tool is placed in a separate directory. It is important that the required folders are present according to the described structure.

Prepare Excel File in search Folder:
Place your Excel file in the search folder where each value to be searched is listed in the first column across any number of sheets.

Update Data in Excel File:
After placing the Excel file, update the data in the file(right-click on searching table -> refresh). The tool will automatically match against the most recently modified file in the search folder. For rows where matches are not found, double-click on the cell [code] to bring up the search form to find the required position. For the searching_eg version, all previously configured calculations will also be performed automatically.

Review and Adjust Matches in searching Table:
In the searching table, review the tool's results and ensure that pairs are correctly matched. If necessary, manually adjust corrections. Then, press the UpdateFromTo button or close the file. By default, the updateFromToOnClose setting on the "settings" sheet is set to 1, which triggers the table update function every time the file is closed.

Enjoy Simplified Matching:
Enjoy the benefit of no longer needing to manually match pairs that have been entered at least once.