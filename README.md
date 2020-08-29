# Automated_NAV_Calculation 

# Table of Content
### [Introduction](#introduction-1)
### [Assumptions](#assumptions-1)
### [Automated_Fund_NAV_Calculation_processflow diagram](#automated_fund_nav_calculation_processflow-diagram-1)
### [Steps to run the macro](#steps-to-run-the-macro-1)
### [Forward](#forward-1)

# Introduction
The purpose of [NAV Calculation.xlsm](https://github.com/Vanipreet/Automated_NAV_Calculation/blob/master/NAV%20Calculation.xlsm) is to create an automated process for the execution of Daily NAV calculation of a fund. There are multiple assumptions and pre-requisites that has been kept in place for the smooth execution of this macro.

# Assumptions
This macro is built with few assumptions in mind as listed below

1. For the macro to run, we need two external data inputs, Daily Pricing Sheet and Daily Liability Sheet
2. The macro is prepared for the purpose of daily NAV calculation and hence uses daily input files
3. Macro and the data inputs are available under "H:\NAV Cal" directory, however as per the drive parition, this can be appropriately mapped
4. Macro expects static input data. However with appropriate mapping, database can be also accurately utilized
5. For a security addition to the fund, security has to be manually added under [NAV Calculation.xlsm](https://github.com/Vanipreet/Automated_NAV_Calculation/blob/master/NAV%20Cal/NAV%20Calculation.xlsm) under "NAV Calculation" worksheet at the to the bottom of the currently available stocks list.
6. To create an automated stock list update, a third source of input will be required, which will pick the updated stock list and post under "NAV Calculation" worksheet.
Please note: Timing of the macro will have to updated for the above so that first macro picks up new stock list and then calculates NAV

# Automated_Fund_NAV_Calculation_processflow diagram
![alt text](https://github.com/Vanipreet/Automated_NAV_Calculation/blob/master/Automated%20NAV%20Calculation%20process.png)

# Steps to run the macro

1. Download the zip file [Automated_NAV_Calculation](https://github.com/Vanipreet/Automated_NAV_Calculation/archive/master.zip) onto your preferred location and unzip the file there
2. Open "NAV Calculation.xlsm" workbook
3. Right click on macro button "Portfolio Position"
4. From the drop down options click on "Assign Macro"
5. Select "StockList" and click on "Edit" button
6. Under module 3, map the Pricing and Liabilities file directory as per requirement
7. Repeat the step 6 for Module 1, Module 2, Module 4
8. For test run, change the file name "Liabilities_28082020.xlsx" to current Date, Month and Year in DDMMYYYY format. Save the file and close
9. Repeat the step 8 for "Pricing_28082020.xlsx" and "Transactions_28082020.xlsx"
10. Click on "Portfolio Position", "Fund Pricing" and "Calculate NAV" macro button under Control sheet of "NAV Calculation.xlsm" in this specific order
11. When the processing is complete on the third macro, message box will pop up advising the processing is complete. Click on "OK"
12. Calculated NAV file is available under "NAVs" folder with today's date

# Forward

This macro is a starting step for NAV calculation, utilizing some tweaks this can accurately fulfil NAV calculation requirement.

However the bigger picture is that instead of making this macro to work on a single fund, it can be accurately "Looped" for multiple funds.

This macro expects manually updating the stock list however utilizing a third input source can be provided, which will pick the updated stock list and post under "NAV Calculation" worksheet.

If this macro is to run on windows system, "Task Scheduler" can be utilized to automatically update the input files at a precribed time from the source drive and also to automatically run this macro so NAV can be calculated without any human intervention. (Apple systems might also have something similar to task scheduler that can be utilized)
