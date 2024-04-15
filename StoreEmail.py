#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Apr 15 18:07:25 2024

@author: George Audi
"""

"""this is a program to save emails to a txt file so that the emails are saved
and stops repeats from printing in the excel sheet"""

#use openpyxl to read the emails for repeats
import openpyxl

def EmailStorage(dic):
    sheet = openpyxl.load_workbook("List of Emails to be used.xlsx")
    sheetData = sheet.active
    maxRow = sheetData.max_row
    bodyList = []
    for i in range(1, maxRow):
        bodyCells = sheetData.cell(row = i, column = 4)
        bodyList.append(bodyCells.value)
    if dic["body"] in bodyList:
        return True
    return False
    
