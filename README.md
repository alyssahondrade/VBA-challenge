# VBA-challenge
Module 2 Challenge - UWA/edX Data Analytics Bootcamp

Github repository at: https://github.com/alyssahondrade/VBA-challenge.git

## Table of Contents
1. [Introduction](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#introduction)
    1. [Goal](https://github.com/alyssahondrade/VBA-challenge/tree/main#goal)
    2. [Source Code](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#source-code)
    3. [Technology](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#technology)
    4. [Dataset](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#dataset)
2. [Approach](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#approach)
3. [Results](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#results)
4. [References](https://github.com/alyssahondrade/VBA-challenge/blob/main/README.md#references)

### Introduction
#### Goal
The goal of the project is to create a script that summarises a list of stock data to acquire the following:
1. Yearly change, from the opening price at the start of the year, to the closing price at the end of the year.
2. Percent change (as above with yearly change)
3. Total stock volume

As well as the calculated values of the summarised results:
1. Greatest percentage increase
2. Greatest percentage decrease
3. Greatest total stock volume

#### Source code
The source code for this project is [`vba-challenge.bas`](https://github.com/alyssahondrade/VBA-challenge/blob/main/vba-challenge.bas).

#### Technology
VBA code was written using **Microsoft Excel for Mac** (Version 16.75.2).

#### Dataset
Dataset was created by Trilogy Education Services (2U Inc. brand)

Initial testing for the code was conducted on [`alphabetical_testing.xlsx`](https://github.com/alyssahondrade/VBA-challenge/blob/main/alphabetical_testing.xlsx) (available in the repository), with the final testing conducted on `multiple_year_stock_data.xlsx` (not provided due to file size).

### Approach
1. Understand the provided dataset prior to conducting any data wrangling. The following observations were made: 
    - Multiple spreadsheets with identically structured data, meaning the code would need to be looped for all spreadsheets.
    - Each spreadsheet had a different number of rows, meaning a function is required to either.
          - Loop through until the first "empty" row is found.
          - Find the last row for each sheet.
2. Prior to writing the VBA script, results were manually calculated in the initial test file to compare against script output.
    - Confirmed the number of unique tickers using `Data > Remove Duplicates` function.
    - Used `SUMIF()` to get the total volume for each ticker.
    - Manually calculated samples of yearly change by referencing cells.
3. Pseudocode produced to identify strategy:
    - What variables are required? What data type for each?
    - What outputs are required?
    - What loop method is appropriate to acquire summarised results?
    - What needs to be incremented through?
    - What relevant equations are required (i.e. yearly change and percent change)?
4. Staged Process
    1. Find the last row of the spreadsheet.
        1. Initial method: while-loop and `Not IsEmpty()` as the condition.
        2. Final method: for-loop and comparing an increment ahead.
    2. Loop through each row and get the unique ticker name.
    3. Get the correct value for `open_price`.
    4. Get the correct value for `close_price`.
    5. Calculate and set `yearly_change` and `percent_change`.
    6. Write code for calculated values.
       1. If-block for `greatest_increase` and `greatest_decrease`, due to mutual exclusivity.
       2. If-block for `total_greatest_volume`.
    7. Conditional formatting on `yearly_change` and `percent_change` columns.
    8. Add formatting and headings.
    9. Alter the code to run for multiple spreadsheets, updating required references (i.e. `fy.Range()`).
    10. Run the code on the final test data.
       1. Update data type for counters to `Long` or `LongLong` as required.
       2. Confirm correct results and formatting in each spreadsheet.

### Results
Screenshots of the results, using the final test data, are given below:
(image_1: Analysis Results (2018))
(image_2: Calculated Values (2018))
(image_3: Calculated Values (2019))
(image_4: Calculated Values (2020))
> NOTE: Upload screenshots to repository and use relative links to get on README

### References
The for-loop concept incrementing one counter ahead, and the equation to find the last row was derived from "Credit Charges" activity in Week 2, Day 3 of the bootcamp. The original solution utilised a while-loop.

> NOTE: Gather all links accessed
