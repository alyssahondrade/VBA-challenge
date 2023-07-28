Sub AnnualSummary()
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
    
    ' Initialise counter to get all rows
    Dim counter As Long
    counter = 2
    
    ' Initialise counter to get unique tickers
    Dim ticker_count As Long
    ticker_count = 2
    
    ' Get the last row of the sheet
    Dim last_row As Long
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through all the rows
    For counter = 2 To last_row
        ticker_name = Cells(counter, 1).Value
                
        If (ticker_name <> Cells(counter + 1, 1).Value) Then
            ' Set unique ticker name
            Cells(ticker_count, 9).Value = ticker_name
            
            ' Increment the total volume for the last row
            total_vol = total_vol + Cells(counter, 7).Value
            Cells(ticker_count, 12).Value = total_vol
            
            ' Get the close price at unique ticker's last row
            close_price = Cells(counter, 6).Value
            
            ' Reset total volume
            total_vol = 0
            
            ' Calculate and set values
            Cells(ticker_count, 10).Value = close_price - open_price
            Cells(ticker_count, 11).Value = FormatPercent((close_price - open_price) / 100)
            ticker_count = ticker_count + 1
            
        Else
            ' Get the open price at unique ticker's first row
            If (total_vol = 0) Then
                open_price = Cells(counter, 3).Value
            End If
            
            ' Increment the total volume
            total_vol = total_vol + Cells(counter, 7).Value
            
        End If
        
    Next counter
    
    Range("A:Q").Columns.AutoFit
            
End Sub
