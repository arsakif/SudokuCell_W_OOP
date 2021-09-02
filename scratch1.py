import random  # To Create random solution

import numpy
import xlwings as xw  # To read/write puzzles
import pandas as pd
from pandas import DataFrame as df
import time

start = time.time()

wb = xw.Book('sudoku.xlsx')  # Open the Excel Book
ws = wb.sheets('sudoku')  # Open the 'Sudoku Excel Sheet'
# Read sudoku puzzle in to a dataFrame
sdk_df = df(data=ws.range('A1:I9').value, columns=[0, 1, 2, 3, 4, 5, 6, 7, 8])

a = [[0, 0, 0]]

i = 0
a[i][0] = sdk_df
print(a[i][0])

aa = df(a[i][0])

print(aa)
