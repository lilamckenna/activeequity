#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Mar  2 16:34:02 2020

@author: jalenwang
"""

import networkx as nx
import matplotlib.pyplot as plt
import xlrd
import pandas as pd
import numpy as np 
from sklearn import datasets
import xlsxwriter 
import statsmodels.api as sm



SMB = []
HML = []
RMW = []
CMA = []
Mkt_RF = []

AllStocksFactors = []
StockName = []


AllColumns = []

def set_up_stocks():
    ExcelFileName= 'cs89_results_export.xlsx'
    workbook = xlrd.open_workbook(ExcelFileName)
    worksheet = workbook.sheet_by_name("Sheet1") # We need to read the data  
    
    for i in range (1, worksheet.nrows): #I made this # 6 to do some testing 
        StockName.append(worksheet.cell_value(i,0))
        StockFactor = []
        if worksheet.cell_value(i,1) <= .05:
            StockFactor.append(worksheet.cell_value(i,1))
        else:
            StockFactor.append(0)
        if worksheet.cell_value(i,2) <= .05:
            StockFactor.append(worksheet.cell_value(i,2))
        else:
            StockFactor.append(0)  
        if worksheet.cell_value(i,3) <= .05:
            StockFactor.append(worksheet.cell_value(i,3))
        else:
            StockFactor.append(0)
        if worksheet.cell_value(i,4) <= .05:
            StockFactor.append(worksheet.cell_value(i,4))
        else:
            StockFactor.append(0)
        if worksheet.cell_value(i,5) <= .05:
            StockFactor.append(worksheet.cell_value(i,5))
        else:
            StockFactor.append(0)
        
        AllStocksFactors.append(StockFactor)

def data_for_excel():
    
    for i in range (0, len(AllStocksFactors)):
        for j in range (i+1, len(AllStocksFactors)):           
            if AllStocksFactors[i][0] != 0 and AllStocksFactors[j][0] != 0: 
                AllColumns.append([i,j+1])
                
    for i in range (0, len(AllStocksFactors)):
        for j in range (i+1, len(AllStocksFactors)):           
            if AllStocksFactors[i][1] != 0 and AllStocksFactors[j][1] != 0: 
                AllColumns.append([i,j+1])
                
    for i in range (0, len(AllStocksFactors)):
        for j in range (i+1, len(AllStocksFactors)):           
            if AllStocksFactors[i][2] != 0 and AllStocksFactors[j][2] != 0: 
                AllColumns.append([i,j+1])
                
    for i in range (0, len(AllStocksFactors)):
        for j in range (i+1, len(AllStocksFactors)):           
            if AllStocksFactors[i][3] != 0 and AllStocksFactors[j][3] != 0: 
                AllColumns.append([i,j+1])
                
    for i in range (0, len(AllStocksFactors)):
        for j in range (i+1, len(AllStocksFactors)):           
            if AllStocksFactors[i][4] != 0 and AllStocksFactors[j][4] != 0: 
                AllColumns.append([i,j+1])

  
def export_to_excel():
    with xlsxwriter.Workbook('test.xlsx') as workbook:
        worksheet = workbook.add_worksheet()

        for row_num, data in enumerate(AllColumns):
            worksheet.write_row(row_num, 0, data)
    
    
def main():
    set_up_stocks()
    data_for_excel()
    export_to_excel()
    

main()