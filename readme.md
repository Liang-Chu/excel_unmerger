# Excel Unmerger

## Overview
The Excel Unmerger script processes Excel files to unmerge cells while preserving the original cell styles and values. It can ignore a specified number of rows and columns during the unmerging process.
## Features
- Unmerges cells in Excel files (.xlsx, .xls, .xlsm, .xlsb)
- Preserves cell styles and values
- Allows ignoring specified rows and columns

## Usage
1. Place all the Excel files (xlsx, xls, xlsm, xlsb) that need to be unmerged in the same directory as the script. It is recommended to create a separate directory that only contains the script, and move the Excel files into this folder when you need to process them.
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
Before:
![image](https://github.com/user-attachments/assets/6da1817c-25d6-4647-b671-38d1e41c4209)
After:
![image](https://github.com/user-attachments/assets/7943d0c6-af7c-4f5c-868e-9e6e7b4add4b)
