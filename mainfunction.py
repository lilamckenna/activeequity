from sklearn import datasets
import numpy as np
import xlsxwriter 
import statsmodels.api as sm
import xlrd as xlrd
import pandas as pd
import numpy as np 

"""
@authors: jalenwang and lilamckenna
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
              
    #setting up results Excel 
    workbook = xlsxwriter.Workbook('cs89_results_export.xlsx')
    worksheet = workbook.add_worksheet()

    #This function regresses the time series data for 1 stock on 
    # time series data for 5 factors, and outputs their p values to an excel sheet. 
    for stock in security_monthly_changes:
       
        #X variables 
        factors = [ SMB, HML, RMW, CMA, Mkt_RF ]
        
        #Regression for each factor Y= Stock X= Factor
        model0= sm.OLS(stock, factors[0])
        model1= sm.OLS(stock, factors[1])
        model2= sm.OLS(stock, factors[2])
        model3= sm.OLS(stock, factors[3])
        model4= sm.OLS(stock, factors[4])
        
        #Fit regression for each factor
        results0 = model0.fit()
        results1 = model1.fit()
        results2 = model2.fit()
        results3 = model3.fit()
        results4 = model4.fit()
        
        #Get p-values for each factor, make into a list of all 
        p_values0 = results0.pvalues
        p_values1 = results1.pvalues
        p_values2 = results2.pvalues
        p_values3 = results3.pvalues
        p_values4 = results4.pvalues
        p_values_total= [p_values0[0],p_values1[0], p_values2[0], p_values3[0], p_values4[0]]
        
        #write name of stock in excel 
        worksheet.write(security_monthly_changes.index(stock), 0, stockname[security_monthly_changes.index(stock)])

        #write p values for stock in excel 
        for p in p_values_total:
            #parameters are row, column, item
            worksheet.write(security_monthly_changes.index(stock),p_values_total.index(p)+1, p)      
            
    workbook.close()
     
def main ():
    open_stock_excel()
    open_factor_excel()
    regressions()
    
main() 
     
