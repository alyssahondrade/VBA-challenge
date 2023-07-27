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
    Dim total_vol As LongLong

    Dim ticker As Range
    For Each ticker In Range("A1:A23000").Cells
        If IsEmpty(ticker) Then
            MsgBox ("This row is empty" & ticker.Address)
            Exit For
        End If
    Next ticker

    Dim counter As Long ' counter to get all rows
    counter = 2 ' since start at A2

    Dim ticker_count As Long ' counts unique ticker_names
    ticker_count = 2

    While (Not IsEmpty(Cells(counter, 1).Value))
        ticker_name = Cells(counter, 1).Value ' assign ticker_name
        Cells(ticker_count, 9).Value = ticker_name ' set I-column cells as unique tickers

        Dim within_ticker As Integer
        within_ticker = 0
        total_vol = 0
        While (ticker_name = Cells(counter, 1).Value)
            total_vol = total_vol + Cells(counter, 7).Value
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
        Cells(ticker_count, 13).Value = FormatPercent((close_price - open_price) / open_price)
        Cells(ticker_count, 14).Value = total_vol
        ticker_count = ticker_count + 1 ' increment unique ticker, new since exited inner while loop
    Wend

    ' BONUS
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As LongLong
    
    Dim max_count As Integer
    max_count = 2
    
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
    While (Not IsEmpty(Cells(max_count, 9).Value))
        If Cells(max_count, 13).Value > greatest_increase Then
            greatest_increase = Cells(max_count, 13).Value
            Cells(2, 22).Value = FormatPercent(greatest_increase, 2)
            Cells(2, 21).Value = Cells(max_count, 9).Value
        ElseIf Cells(max_count, 13).Value < greatest_decrease Then
            greatest_decrease = Cells(max_count, 13).Value
            Cells(3, 22).Value = FormatPercent(greatest_decrease, 2)
            Cells(3, 21).Value = Cells(max_count, 9).Value
        ElseIf Cells(max_count, 14).Value > greatest_volume Then
            greatest_volume = Cells(max_count, 14).Value
            Cells(4, 22).Value = greatest_volume
            Cells(4, 21).Value = Cells(max_count, 9).Value
        End If
        max_count = max_count + 1
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
