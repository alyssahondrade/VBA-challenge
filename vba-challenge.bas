Attribute VB_Name = "Module1"
Sub AnnualSummary():
    ' Declare variables
    ' 1) ticker_name to pull from spreadsheet
    ' 2) hold variable to track one ticker
    ' 3) open_price variable
    ' 4) close_price variable
    ' 5) year_change variable
    ' 6) percent_change variable
    ' 7) total_vol variable

    Dim ticker_name As String
    Dim hold_ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim total_vol As Long

    ' Use for while loop, condition: 'ticker_name' not empty, OR for/for each loop - just need last cell row
    ' Get 'ticker_name' from each row
    ' For the first, 'ticker_name' = 'hold'. For the rest, compare 'ticker_name' to hold'
    ' If TRUE
    ' 1) increment 'total_vol'
    ' 2) keep going until 'ticker_name' != 'hold'
    ' If FALSE, i.e. new ticker
    ' 1) calculate 'yearly_change' = 'close_price' - 'open_price'
    ' 2) calculate 'percent_change' = (('open_price' - 'close_price') / open_price) * 100
    
    ' BONUS
    ' Use Application.Worksheet.Max() ?
    
End Sub
