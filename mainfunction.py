from sklearn import datasets
import numpy as np
import xlsxwriter
import statsmodels.api as sm
import xlrd
import pandas as pd
import numpy as np 

"""
@author: jalenwang and lilamckenna
"""

stockname = []
security_monthly_changes = []
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

    #create security_monthly_changes
    start_column = 3
    end_column = 243

    ##start_column = int(input("Start Column (3): ")) ##set start month
    ##end_column = int(input("End Columns (240): ")) ##set end month 
    
    for i in range (1,worksheet.nrows):
        monthlychanges = []
        for j in range (start_column, end_column):
            change = (worksheet.cell_value(i,j+1))/(worksheet.cell_value(i,j))-1
            monthlychanges.append(change)
        security_monthly_changes.append(monthlychanges)
    
    print (security_monthly_changes)
    print (len(security_monthly_changes))
    
def open_factor_excel():
    ExcelFileName= 'factor.xlsx'
    workbook = xlrd.open_workbook(ExcelFileName)
    worksheet = workbook.sheet_by_name("Sheet1") # We need to read the data         
    
    start_row = 442
    ##start_row = int(input("Start Row (442): ")) ##set start month
    end_row = 682
    ## end_row = int(input("End Row (682: ")) ##set end month 
    

    for i in range (start_row, end_row):
        Mkt_RF.append(worksheet.cell_value(i,1))
        SMB.append(worksheet.cell_value(i,2))
        HML.append(worksheet.cell_value(i,3))
        RMW.append(worksheet.cell_value(i,4))
        CMA.append(worksheet.cell_value(i,5))
    
def regressions():
    # print (len(factors))
    # print (len(security_monthly_changes))
    # print ("that was the monthly changes")
    # print (len(stockname))
    # print (len(SMB))
    # print (len((RMW))
              
    #setting up results Excel 
    workbook = xlsxwriter.Workbook('cs89_results_export.xlsx')
    worksheet = workbook.add_worksheet()
    #security_monthly_changes=list(security_monthly_changes)

    #This function regresses the time series data for 1 stock on 
    # time series data for 5 factors, and outputs their p values to an excel sheet. 
    for stock in security_monthly_changes:
    
        #X variables 
        factors = [ SMB, HML, RMW, CMA, Mkt_RF ]
        factors = sm.add_constant(factors)

        #Regression
            # model = sm.OLS(security_monthly_changes,factors)

        model2= sm.OLS( endog=security_monthly_changes, exog=factors)
        #formula='security_monthly_changes ~ SMB + HML + RMW + CMA + Mkt_RF',
        results = model2.fit()
        p_values = results.summary2().tables[1]['P>|t|']

        #Stock name in first columm
        worksheet.write(stockname.index(stock), 0, stock)
        
        #P values for regression of stock on each factor in next 5 columns
        for p in p_values:
            #parameters are row, column, item
            worksheet.write(security_monthly_changes.index(stock),p_values.index(p)+1, p)
            
    workbook.close()
            
            
def main ():
    open_stock_excel()
    open_factor_excel()
    regressions()
    
main() 
     
