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
    Dim open_price As Double
    Dim close_price As Double
    Dim total_vol As Long

    Dim ticker As Range
    For Each ticker In Range("A1:A23000").Cells
        If IsEmpty(ticker) Then
            MsgBox ("This row is empty" & ticker.Address)
            Exit For
        End If
    Next ticker

    Dim counter As Integer ' counter to get all rows
    counter = 2 ' since start at A2

    Dim ticker_count As Integer ' counts unique ticker_names
    ticker_count = 2

    While (Not IsEmpty(Cells(counter, 1).Value))
        ticker_name = Cells(counter, 1).Value ' assign ticker_name
        Cells(ticker_count, 9).Value = ticker_name ' set I-column cells as unique tickers

        Dim within_ticker As Integer
        within_ticker = 0
        total_vol = 0
        While (ticker_name = Cells(counter, 1).Value)
            If (within_ticker = 0) Then
                open_price = Cells(counter, 3).Value ' because this is the first row of new ticker
            End If
            within_ticker = within_ticker + 1 ' loop through rest, otherwise open_price will keep updating
            close_price = Cells(counter, 6).Value ' always the last price before exiting while loop
            counter = counter + 1
        Wend
        Cells(ticker_count, 10).Value = open_price ' set just before ticker_count update
        Cells(ticker_count, 11).Value = close_price
        Cells(ticker_count, 12).Value = close_price - open_price
        Cells(ticker_count, 13).Value = FormatPercent((open_price - close_price) / open_price)
        ticker_count = ticker_count + 1 ' increment unique ticker, new since exited inner while loop
    Wend

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
