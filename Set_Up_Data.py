#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Feb 18 18:36:05 2020

@author: jalenwang
"""
import xlrd
import pandas as pd
import numpy as np 


stockname = []
allstockmonthlychanges = []

Mkt_RF = []
SMB = []
HML = []
RMW = []
CMA = []



def open_stock_excel():
    ExcelFileName= 'GoodData2.xlsx'
    workbook = xlrd.open_workbook(ExcelFileName)
    worksheet = workbook.sheet_by_name("Sheet1") # We need to read the data 
    
    #create stockname list
    for i in range (1,worksheet.nrows):
        stockname.append(worksheet.cell_value(i,0))

    #creast allstockmonthlychanges

    start_column = int(input("Start Column (3): ")) ##set start month
    end_column = int(input("End Columns (240): ")) ##set end month 
    
    for i in range (1,worksheet.nrows):
        monthlychanges = []
        for j in range (start_column, end_column-1):

            change = (worksheet.cell_value(i,j+1))/(worksheet.cell_value(i,j))-1
            monthlychanges.append(change)
        allstockmonthlychanges.append(monthlychanges)
    
    
    print (allstockmonthlychanges)

def open_factor_excel():
    ExcelFileName= 'factor.xlsx'
    workbook = xlrd.open_workbook(ExcelFileName)
    worksheet = workbook.sheet_by_name("Sheet1") # We need to read the data         
    
    start_row = int(input("Start Row (442): ")) ##set start month
    end_row = int(input("End Row (682: ")) ##set end month 
    
    
  
    for i in range (start_row, end_row):
        Mkt_RF.append(worksheet.cell_value(i,1))
        SMB.append(worksheet.cell_value(i,2))
        HML.append(worksheet.cell_value(i,3))
        RMW.append(worksheet.cell_value(i,4))
        CMA.append(worksheet.cell_value(i,5))
    
    print (Mkt_RF)
    print (SMB)
    print (HML)
    print (RMW)
    print (CMA)
    

def main ():
    open_stock_excel()
    open_factor_excel()
    
    
main()