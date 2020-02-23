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

for stock in security_monthly_changes:
    
    factors = np.array([SMB, HML, RMW, CMA, RF])
    model = sm.OLS(security_monthly_changes,factors)
    results = model.fit()
    p_values = results.summary2().tables[1]['P>|t|']

    #Stock name in first columm
    worksheet.write(stockname.index(stock), 0, stock)
    
    #P values for regression of stock on each factor in next 5 columns
    for p in p_values:
        worksheet.write(security_monthly_changes.index(stock),p_values.index(p)+1, p)
            
       
workbook.close()