# RANK-3 SUDOKU SOLVER EXCERCISE
# Akif Arslan -- 08/22/2021 NY USA

""" This sudoku solver excersize is created to understand Classes in Python.
    The sudoku puzzle given to the solver from an Excel file. Solver write back the
    solution to the Excel File.
    If the puzzle is not solvable it return an error message that says it's not a solvable puzzle.
    The solver dosen't check the uniqueness of the solution, it will return one of the solutions
    randomly. Even if a completely empty sudoku grid is given, the solver will return a random
    solution. This feature may be utilized to create new sudoku puzzles."""

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
# if there is "NA" cells convert them to "0"
for i in range(0, 9):
    sdk_df.loc[:, i] = sdk_df.loc[:, i].apply(lambda x: int(0) if pd.isna(x) else int(x))


# Create a class to to represent a sudoku cell in a sudoku grid
class SudokuCell:
    def __init__(self, sdku_df, x_pos=0, y_pos=0):
        self.x_pos = x_pos  # x position in the grid
        self.y_pos = y_pos  # y position in the grid
        self.sdku_df = sdku_df  # sudoku puzzle

        self.rw = list(self.sdku_df.loc[self.x_pos, :])  # row values where the cell in
        self.clm = list(self.sdku_df.loc[:, self.y_pos])  # column values where the cell in
        self.bx = self.bx_vls()  # box values where the cell in
        self.bx_number = 0

        self.base_list = [1, 2, 3, 4, 5, 6, 7, 8, 9]  # base values of a rank-3 sudoku
        random.shuffle(self.base_list)  # shuffle values to create random solutions

        self.cndts = self.fnd_cndts()  # candidate values for the cell
        self.lncdts = len(self.cndts)  # number of the candidate values

    def bx_vls(self):  # Find the box values where the cell in
        if self.x_pos < 3 and self.y_pos < 3:
            self.bx_number = 1
            return [self.sdku_df.loc[i, j] for i in range(0, 3) for j in range(0, 3)]
        elif (2 < self.x_pos < 6) and self.y_pos < 3:
            self.bx_number = 2
            return [self.sdku_df.loc[i, j] for i in range(3, 6) for j in range(0, 3)]
        elif self.x_pos > 5 and self.y_pos < 3:
            self.bx_number = 3
            return [self.sdku_df.loc[i, j] for i in range(6, 9) for j in range(0, 3)]

        elif self.x_pos < 3 and (2 < self.y_pos < 6):
            self.bx_number = 4
            return [self.sdku_df.loc[i, j] for i in range(0, 3) for j in range(3, 6)]
        elif (2 < self.x_pos < 6) and (2 < self.y_pos < 6):
            self.bx_number = 5
            return [self.sdku_df.loc[i, j] for i in range(3, 6) for j in range(3, 6)]
        elif self.x_pos > 5 and (2 < self.y_pos < 6):
            self.bx_number = 6
            return [self.sdku_df.loc[i, j] for i in range(6, 9) for j in range(3, 6)]

        elif self.x_pos < 3 and self.y_pos > 5:
            self.bx_number = 7
            return [self.sdku_df.loc[i, j] for i in range(0, 3) for j in range(6, 9)]
        elif (2 < self.x_pos < 6) and self.y_pos > 5:
            self.bx_number = 8
            return [self.sdku_df.loc[i, j] for i in range(3, 6) for j in range(6, 9)]
        elif self.x_pos > 5 and self.y_pos > 5:
            self.bx_number = 9
            return [self.sdku_df.loc[i, j] for i in range(6, 9) for j in range(6, 9)]

    def fnd_cndts(self):  # Find the candidate values for the cell
        return [x for x in self.base_list if x not in set(self.rw + self.clm + self.bx)]

    # Override a comparison method for the SudokuCell class for finding the cell which has least candidates.
    def __gt__(self, other):
        return self.lncdts > other

    def __ge__(self, other):
        return self.lncdts >= other

    def __lt__(self, other):
        return self.lncdts < other

    def __le__(self, other):
        return self.lncdts <= other

    def __eq__(self, other):
        return self.lncdts == other


# =====================================================================================================================
class CheckLegit(SudokuCell):
    def __init__(self, sdku_df):
        super(CheckLegit, self).__init__(sdku_df)
        self.sdku_df = sdku_df
        self.legitSdk = True

    def row_check(self):
        for i in range(0, 9):
            unique_row = [self.sdku_df.loc[i, j] for j in range(0, 9) if self.sdku_df.loc[i, j] != 0]
            if len(set(unique_row)) < len(unique_row):
                print(f'Row {i + 1} has repetitive cells. Not a Legit Sudoku!')
                return 1

    def clm_check(self):
        for j in range(0, 9):
            unique_clm = [self.sdku_df.loc[i, j] for i in range(0, 9) if self.sdku_df.loc[i, j] != 0]
            if len(set(unique_clm)) < len(unique_clm):
                print(f'Column {j + 1} has repetitive cells. Not a Legit Sudoku!')
                return 1

    def box_check(self):
        for self.x_pos in range(0, 9, 3):
            for self.y_pos in range(0, 9, 3):
                unique_box = self.bx_vls()
                unique_box = [x for x in unique_box if x != 0]
                if len(set(unique_box)) < len(unique_box):
                    print(f'Box {self.bx_number} has repetitive cells. Not a Legit Sudoku!')
                    return 1

    def num_check(self):
        for i in range(0, 9):
            for j in range(0, 9):
                if self.sdku_df.loc[i, j] < 0 or self.sdku_df.loc[i, j] > 9:
                    print(f' {self.sdku_df.loc[i, j]} is not a legit number at ({i},{j})')
                    return 1

    def check_legit(self):
        if self.row_check() or self.clm_check() or self.box_check() or self.num_check():
            exit()


ss = sdk_df.copy()
CheckLegit(ss).check_legit()


# =====================================================================================================================
class SudokuSolve(SudokuCell):
    final_asgmnt_ls = []
    solved_sdk = sdk_df.copy()

    def __init__(self, sdku_df, assgm=None):
        super(SudokuSolve, self).__init__(sdku_df)
        if assgm is None:
            self.assgmt_hstry_ls = [[10, 10, []]]  # List to keep assignemt history for backtracking
        else:
            self.assgmt_hstry_ls = assgm
        self.no_solution_flag = False  # Flag to stop iteration in case not solvable puzzle
        self.cls_to_slvd = []

    # ========== ITERATION FOR SOLUTION ============
    def iterate_cells(self):
        for self.max_iteration in range(0, 500000):  # Start iteration
            if self.no_solution_flag:  # If no solution stop iteration
                break

            go_back_flag = False  # Flag to indicate start backtracking

            # Find the cells need to be solved
            self.cls_to_slvd = [SudokuCell(self.sdku_df, x_pos=i, y_pos=j) for i in range(0, 9) for j in range(0, 9) if
                                self.sdku_df.loc[i, j] == 0]

            if not self.cls_to_slvd:  # if there is no cell remained to solve stop iteration.
                print('SOLVED!!!')
                break

            self.cls_to_slvd.sort()  # Sort SudokuCells to start with the least candidate for less computation time.

            # Loop thourgh cells to check if there is any empty cell with zero candidate to start backtracking

            if not self.cls_to_slvd[0].cndts and self.sdku_df.loc[self.cls_to_slvd[0].x_pos, self.cls_to_slvd[0].y_pos] \
                    == 0:
                go_back_flag = True  # Backtracking flag is True if there is cell with zero candidate.

            if go_back_flag:  # Backtracking
                x_to_delete = self.assgmt_hstry_ls[-1][0]
                y_to_delete = self.assgmt_hstry_ls[-1][1]
                self.sdku_df.loc[x_to_delete, y_to_delete] = 0  # Delete the last assignment
                # If there is no candidate left at the current cell go back furter to prevous assignment
                while not self.assgmt_hstry_ls[-1][2]:
                    if 10 in self.assgmt_hstry_ls.pop():  # When reached to the beginning of the assignmet list stop
                        # iteration
                        print('Not a Solveable Puzzle')
                        no_solution_flag = True
                        return 0
                        # break
                    x_to_delete = self.assgmt_hstry_ls[-1][0]
                    y_to_delete = self.assgmt_hstry_ls[-1][1]
                    self.sdku_df.loc[x_to_delete, y_to_delete] = 0
                if self.no_solution_flag:
                    break
                self.sdku_df.loc[x_to_delete, y_to_delete] = self.assgmt_hstry_ls[-1][2].pop()
                continue

            # Assign Candidates to the Sudoku Solution
            x_to_asgn = self.cls_to_slvd[0].x_pos
            y_to_asgn = self.cls_to_slvd[0].y_pos
            cndt_ls_save = self.cls_to_slvd[0].cndts
            self.assgmt_hstry_ls.append(
                [x_to_asgn, y_to_asgn, cndt_ls_save])  # Save assignmet to assignmet history for backtracing
            self.sdku_df.loc[x_to_asgn, y_to_asgn] = self.assgmt_hstry_ls[-1][2].pop()

        if not self.no_solution_flag:  # If reached a solution write solution back to the Excel File
            ws.range('A12').options(index=False, header=False).value = self.sdku_df
            print(time.time() - start)
            SudokuSolve.final_asgmnt_ls = self.assgmt_hstry_ls
            SudokuSolve.solved_sdk = self.sdku_df
            return 1


SudokuSolve(sdk_df.copy()).iterate_cells()


# =====================================================================================================================

class CheckUniqueSolution(SudokuSolve):
    number_of_solutions = 1
    sdk_lst = []

    def __init__(self, sdku_df):
        super(CheckUniqueSolution, self).__init__(sdku_df)
        CheckUniqueSolution.sdk_lst.append(sdku_df)
        self.sdk = CheckUniqueSolution.sdk_lst[CheckUniqueSolution.number_of_solutions - 1]
        self.assgmt_hist = [SudokuSolve.final_asgmnt_ls[i] for i in range(0, len(SudokuSolve.final_asgmnt_ls))
                            if SudokuSolve.final_asgmnt_ls[i][2]]

    def check_another_solution(self):
        if self.assgmt_hist:
            candidates = self.assgmt_hist[0][2]
            x_pos = self.assgmt_hist[0][0]
            y_pos = self.assgmt_hist[0][1]

            while candidates:
                self.sdk.loc[x_pos, y_pos] = candidates.pop()
                sdk_temp = self.sdk.copy()
                if SudokuSolve(sdk_temp).iterate_cells():
                    CheckUniqueSolution.number_of_solutions += 1
                    CheckUniqueSolution.sdk_lst.append(self.sdk)
                    print(CheckUniqueSolution.number_of_solutions)
                    a = CheckUniqueSolution(self.sdk)
                    a.check_another_solution()

                else:
                    self.sdk.loc[x_pos, y_pos] = SudokuSolve.solved_sdk.loc[x_pos, y_pos]
                    del self.assgmt_hist[0]

                if (not candidates) and self.assgmt_hist:
                    candidates = self.assgmt_hist[0][2]




sss = sdk_df.copy()
CheckUniqueSolution(ss).check_another_solution()
