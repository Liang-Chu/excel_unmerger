# Excel Unmerger

## Overview
The Excel Unmerger script processes Excel files to unmerge cells while preserving the original cell styles and values. It can ignore a specified number of rows and columns during the unmerging process.

## Features
- Unmerges cells in Excel files (.xlsx, .xls, .xlsm, .xlsb)
- Preserves cell styles and values
- Allows ignoring specified rows and columns

## Usage
1. Place all the excel (xlsx,xls,xlsm,xlsb) files needed to be unmerged in the same directory as the script.
2. Run the script.
3. Follow the prompts to:
   - Confirm the files to be processed.
   - Specify the number of rows and columns to ignore (optional).
4. The script will process each file and save the unmerged version with a prefix `unmerged_`.

## Example
```plaintext
Excel files found:
1. [UNLOADING_PLAN.xlsx](http://_vscodecontentref_/1)
Press Enter to confirm the files to be processed...       
Enter the number of rows to ignore (leave empty to process entire file):
Enter the number of columns to ignore (leave empty to process entire file):
Processing file: [UNLOADING_PLAN.xlsx](http://_vscodecontentref_/2)
Saved unmerged file as [unmerged_UNLOADING_PLAN.xlsx](http://_vscodecontentref_/3)
Press Enter to process the next file or close the program to exit...
```