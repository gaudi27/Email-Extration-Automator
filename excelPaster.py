#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''
Author: George Z Audi
Date: April 9th 2024'''

''' Code segment for pasting the email information 
into an excel sheet'''


import xlsxwriter


#creates excel and pasts emails in an orginized way

def infoPaster(data):
    
    workbook = xlsxwriter.Workbook("GmailsInfo.x1sx")
    worksheet = workbook.add_worksheet("firstSheet")
    
    worksheet.write(0, 0, "#") 
    worksheet.write(0, 1, "Sender") 
    worksheet.write(0, 3, "Subject") 
    worksheet.write(0, 4, "Body") 
    
    
    for index, entry in enumerate(data):
        worksheet.write(index+1, 0, str(index))
        worksheet.write(index+1, 1, entry["Sender"])
        worksheet.write(index+1, 3, entry["Subject"]) 
        worksheet.write(index+1, 4, entry["Body"]) 

    workbook.close()
