# Automated_NAV_Calculation 

# Table of Content
### Introduction
### Assumptions
### Automated_Fund_NAV_Calculation_processflow diagram
### Steps to run the macro
### [Foreword](#Foreword)

# Introduction
The purpose of [NAV Calculation.xlsm](https://github.com/Vanipreet/Automated_NAV_Calculation/blob/master/NAV%20Cal/NAV%20Calculation.xlsm) is to create an automated process for the execution of Daily NAV calculation of a fund. There are multiple assumptions and pre-requisites that has been kept in place for the smooth execution of this macro.

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
![alt text](https://github.com/Vanipreet/Automated_NAV_Calculation/blob/master/NAV%20Calculation%20Processflow.png)

# Steps to run the macro

1. Download the zip file [NAV Cal](https://github.com/Vanipreet/Automated_NAV_Calculation/tree/master/NAV%20Cal) onto your preferred location and unzip the file there
2. Open "NAV Calculation.xlsm" workbook
3. Right click on macro button "Calculate NAV"
4. From the drop down options click on "Assign Macro"
5. Select "Monthly_Recon" and click on "Edit" button
6. Under module 1, map the Pricing and Liabilities file directory as per requirement
Please Note: All the mapping is required only on the Module 1
7. For test run, change the file name "Liabilities_26072020.xlsx" to current Date, Month and Year in DDMMYYYY format. Save the file and close
8. Redo the Step 7 for "Pricing_26072020.xlsx" as well. 
9. Click on "Calculate NAV" macro button under Control sheet of "NAV Calculation.xlsm"

# Foreword

This macro is a starting step for NAV calculation, utilizing some tweaks this can accurately fulfil NAV calculation requirement.

However the bigger picture is that instead of making this macro to work on a single fund, it can be accurately "Looped" for multiple funds.

This macro expects manually updating the stock list however utilizing a third input source can be provided, which will pick the updated stock list and post under "NAV Calculation" worksheet.

If this macro is to run on windows system, "Task Scheduler" can be utilized to automatically update the input files at a precribed time from the source drive and also to automatically run this macro so NAV can be calculated without any human intervention. (Apple systems might also have something similar to task scheduler that can be utilized)
