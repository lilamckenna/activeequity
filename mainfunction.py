from sklearn import datasets
import numpy as np
import xlsxwriter
import statsmodels.api as sm


#setting up results Excel 
workbook = xlsxwriter.Workbook('cs89_results_export.xlsx')
worksheet = workbook.add_worksheet()

#Import data into arrays from data excel sheet
stockname=[]
security_monthly_changes= [] 
SMB = [] 
HML = [] 
RMW = [] 
CMA = [] 
RF = [] 

#This function regresses the time series data for 1 stock on 
# time series data for 5 factors, and outputs their p values to an excel sheet. 
for stock in security_monthly_changes:
  
    #X variables 
    factors = np.array([SMB, HML, RMW, CMA, RF])
    
    #Regression
    model = sm.OLS(security_monthly_changes,factors).fit()
    results = model.fit()
    p_values = results.summary2().tables[1]['P>|t|']

    #Stock name in first columm
    worksheet.write(stockname.index(stock), 0, stock)
    
    #P values for regression of stock on each factor in next 5 columns
    for p in p_values:
        #parameters are row, column, item
        worksheet.write(security_monthly_changes.index(stock),p_values.index(p)+1, p)
            
       
workbook.close()