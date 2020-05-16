# -*- coding: utf-8 -*-
"""
Created on Fri May 15 14:26:28 2020

@author: Bharath
"""

import xlrd

workbook = xlrd.open_workbook(r'C:\Users\Bharath\Desktop\ExcelReadData.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
num_rows = worksheet.nrows - 1
curr_row = -1
# =============================================================================
print('Reading content row by row')
while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    print (row)
# =============================================================================
    
# =============================================================================
print('\n\nAccessing particular cell')
print(worksheet.cell(0, 0))
print(worksheet.cell(0, 1))
print(worksheet.cell(1, 0))
print(worksheet.cell(1, 1))
print(worksheet.cell(2, 0))
print(worksheet.cell(2, 1))

# =============================================================================
