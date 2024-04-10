#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''
Author: George Z Audi
Date: April 9th 2024'''

''' Code segment for pasting the email information 
into an excel sheet'''



import openpyxl
from openpyxl import Workbook, load_workbook

#creates excel and pasts emails in an orginized way



def Paster(data):
    wb = load_workbook('List of Emails to be used.xlsx')
    ws = wb.active
    ws.append(data)
    wb.save('List of Emails to be used.xlsx')
