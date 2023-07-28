# VBA-challenge
Module 2 Challenge - UWA/edX Data Analytics Bootcamp

Github repository at: https://github.com/alyssahondrade/VBA-challenge.git

## Table of Contents
1. [Introduction](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#introduction)
    1. [Goal]()
    2. [Source Code](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#source-code)
    3. [Technology](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#technology)
    4. [Dataset](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#dataset)
2. [Approach](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#approach)
3. [Results](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#results)
4. [References](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#references)

### Introduction
#### Goal
The goal of the project is to create a script that summarises a list of stock data to acquire the following:
- Yearly change, from the opening price at the start of the year, to the closing price at the end of the year.
- Percent change (as above with yearly change)
- Total stock volume

As well as the calculated values of the summarised results:
- Greatest percentage increase
- Greatest percentage decrease
- Greatest total stock volume

#### Source code
The source code for this project is **vba-challenge.bas**.

#### Technology
VBA code was written using **Microsoft Excel for Mac** (Version 16.75.2).

#### Dataset
Dataset was created by Trilogy Education Services (2U Inc. brand)

Initial testing for the code was conducted on **alphabetical_testing.xlsx** (available in the repository), with the final testing conducted on **multiple_year_stock_data.xlsx** (not provided due to file size).

### Approach
> NOTE: Summarise Module 2 Challenge notes
1. Dataset already provided, needed to understand the data prior to conducting any data wrangling. The following observations were made: 
- Multiple spreadsheets with identically structured data, meaning the code would need to be looped for all spreadsheets
- Each spreadsheet had a different number of rows, meaning a function is required to either
  - Loop through until the first "empty" row is found
  - Find the last row for each sheet
2. Prior to writing the VBA script, results were manually calculated in the initial test file to compare against script output.
- Confirmed the number of unique tickers using **Remove Duplicates** function
- Used **SUMIF** to get the total volume for each ticker
- Manually calculated samples of yearly change by referencing cells
3. Pseudocode produced to identify strategy:
- What variables are required? What data type for each?
- What outputs are required?
- What loop method is appropriate to acquire summarised results?
- What needs to be incremented through?
- What relevant equations are required (i.e. yearly change and percent change)?
4. Staged Process
  4.1 Find the last row of the spreadsheet.
    4.1.1 Initial method: while loop and **Not IsEmpty()**
    4.1.2 Final method: for loop and comparing an increment ahead
  4.2

### Results
Screenshots of the results, using the final test data, are given below:
(image_1: Analysis Results (2018))
(image_2: Calculated Values (2018))
(image_3: Calculated Values (2019))
(image_4: Calculated Values (2020))
> NOTE: Upload screenshots to repository and use relative links to get on README

### References
The **for loop** concept incrementing one counter ahead, and the equation to find the last row was derived from "Credit Charges" activity in Week 2, Day 3 of the bootcamp. The original solution (as per vba-challenge.bas history) utilised a while loop.

> NOTE: Gather all links accessed
