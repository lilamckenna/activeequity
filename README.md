# activeequity
**Active Edge Equity Investing. The next big thing.**

*Authors: Lila McKenna and Jalen Wang*

This project creates a network of the s&p 500 stocks, mapping relationships between stocks via a set of market tilts / factors.


Note: A lot of work for cleaning the data and pulling was run on Excel and Bloomberg terminal. 




**Files in this repo:**


mainfunction.py: Reads the bloomberg historical stock data & fama french data and runs the regression. It outputs the cs89_results_export file. 

cs89_reslts_export.xlsx:
This displays the statistical significance (p-values) of each of the five regressions for every stock in the S&p 500. 

FinalProject3: 
This code reads in "cs89_results_export", and displays the output as a graph using networkX.
We decided to use gephi for the presentation, but networkX displays the same networks, just not as easy to visualize. 

GoodData2: 
This is time-series stock data for the S&P 500. 

factor.xlsx:
This is time-series stock data for the fama french 5 factors. 

F-F_Research_Data_5_Factors_2x3.CSV:
This is also time-series stock data for fama french 5 factors, just stored in a csv. 

Walmart Financial Model.xlsx:
This is the financial model used to show the story of Walmart being a value stock (Quality tilt)

SetUpGrephi.py:
This sets up the edges and nodes csv for using Grephi 



Thanks for reading! 
-Lila and Jalen 
