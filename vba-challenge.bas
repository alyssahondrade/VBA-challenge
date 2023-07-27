Attribute VB_Name = "Module1"
Sub AnnualSummary():
    Dim ticker_name As String
    Dim open_price As Double
    Dim close_price As Double
    Dim total_vol As LongLong
    
    ' Set headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Dim counter As Long ' counter to get all rows
    counter = 2
    
    Dim ticker_count As Long ' counts unique ticker_names
    ticker_count = 2
    
    While (Not IsEmpty(Cells(counter, 1).Value))
        ticker_name = Cells(counter, 1).Value
        Cells(ticker_count, 9).Value = ticker_name
        
        Dim within_ticker As Integer ' counts rows for a given ticker
        within_ticker = 0
        
        total_vol = 0
        While (ticker_name = Cells(counter, 1).Value)
            total_vol = total_vol + Cells(counter, 7).Value ' calculate total volume
            
            If (within_ticker = 0) Then
                open_price = Cells(counter, 3).Value ' get open_price, first value of new ticker
            End If
            within_ticker = within_ticker + 1 ' loop through rest, otherwise open_price will keep updating
            
            close_price = Cells(counter, 6).Value ' always the last price before exiting the while loop
            counter = counter + 1
        Wend
        
        ' Set calculated values before ticker_count update
        Cells(ticker_count, 10).Value = close_price - open_price
        Cells(ticker_count, 11).Value = FormatPercent((close_price - open_price) / open_price)
        Cells(ticker_count, 12).Value = total_vol
        
        ' Conditional formatting
        Dim yearly_change As Double
        yearly_change = Cells(ticker_count, 10).Value
        If (yearly_change > 0) Then
            Cells(ticker_count, 10).Interior.ColorIndex = 4
        Else
            Cells(ticker_count, 10).Interior.ColorIndex = 3
        ticker_count = ticker_count + 1 ' increment unique ticker
        End If

    Wend
    
    ' Bonus
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As LongLong
    Dim get_ticker As String
    
    Dim max_count As Integer
    max_count = 2
    
    ' initialise variables for comparison in while loop
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
    While (Not IsEmpty(Cells(max_count, 9).Value))
        get_ticker = Cells(max_count, 9).Value ' get current ticker
        If Cells(max_count, 11).Value > greatest_increase Then
            greatest_increase = Cells(max_count, 11).Value ' set new greatest_increase value
            Cells(2, 17).Value = FormatPercent(greatest_increase, 2)
            Cells(2, 16).Value = get_ticker
        ElseIf Cells(max_count, 11).Value < greatest_decrease Then
            greatest_decrease = Cells(max_count, 11).Value ' set new greatest_decrease value
            Cells(3, 17).Value = FormatPercent(greatest_decrease, 2)
            Cells(3, 16).Value = get_ticker
        ElseIf Cells(max_count, 12).Value > greatest_volume Then
            greatest_volume = Cells(max_count, 12).Value ' set new greatest_volume value
            Cells(4, 17).Value = greatest_volume
            Cells(4, 16).Value = get_ticker
        End If
        max_count = max_count + 1
    Wend
    
End Sub

