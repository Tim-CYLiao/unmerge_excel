# unmerge_excel
Contain 2 files:
unmerge_excel_to_dataframe.py:

Function:
    unmerge_excel_to_df_dict(path): For unmerge excel into dictionary of dataframe.
    unmerge_ws_to_df(book,sheetnames): Unmerge specific worksheet of openpyxl's workbook object into a dataframe
Unmerged cells are filled by value of original merged cells.



&

umerge_excel.py :
Function:
    unmerge_excel(path): For unmerge excel and  save it.
    now some of merged cells couldn't be unmerged,need to fix it
Unmerged cells are filled by value of original merged cells.



