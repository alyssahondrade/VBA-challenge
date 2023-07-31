Sub AnnualSummary()
    Dim fy As Worksheet
    
    ' Loop through all the sheets
    For Each fy In Worksheets
        Dim ticker_name As String
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_vol As LongLong
        
        ' Set headings
        fy.Range("I1").Value = "Ticker"
        fy.Range("J1").Value = "Yearly Change"
        fy.Range("K1").Value = "Percent Change"
        fy.Range("L1").Value = "Total Stock Volume"
        fy.Range("O2").Value = "Greatest % Increase"
        fy.Range("O3").Value = "Greatest % Decrease"
        fy.Range("O4").Value = "Greatest Total Volume"
        fy.Range("P1").Value = "Ticker"
        fy.Range("Q1").Value = "Value"
        
        ' Initialise counter to get all rows
        Dim counter As Long
        counter = 2
        
        ' Initialise counter to get unique tickers
        Dim ticker_count As Long
        ticker_count = 2
        
        ' Get the last row of the sheet
        Dim last_row As Long
        last_row = fy.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all the rows
        For counter = 2 To last_row
            ticker_name = fy.Cells(counter, 1).Value
                    
            If (ticker_name <> fy.Cells(counter + 1, 1).Value) Then
                ' Set unique ticker name
                fy.Cells(ticker_count, 9).Value = ticker_name
                
                ' Increment the total volume for the last row
                total_vol = total_vol + fy.Cells(counter, 7).Value
                fy.Cells(ticker_count, 12).Value = total_vol
                
                ' Get the close price at unique ticker's last row
                close_price = fy.Cells(counter, 6).Value
                
                ' Reset total volume
                total_vol = 0
                
                ' Calculate and set values
                yearly_change = close_price - open_price
                percent_change = (close_price - open_price) / open_price

                ' Set and format yearly_change and percent_change
                fy.Cells(ticker_count, 10).Value = yearly_change
                fy.Cells(ticker_count, 10).NumberFormat = "0.00"
                fy.Cells(ticker_count, 11).Value = percent_change
                fy.Cells(ticker_count, 11).NumberFormat = "0.00%"
                
                ' Conditional formatting - yearly change
                If (yearly_change > 0) Then
                    fy.Cells(ticker_count, 10).Interior.ColorIndex = 4
                Else
                    fy.Cells(ticker_count, 10).Interior.ColorIndex = 3
                End If
                
                ' Conditional formatting - percent change
                If (percent_change > 0) Then
                    fy.Cells(ticker_count, 11).Interior.ColorIndex = 4
                Else
                    fy.Cells(ticker_count, 11).Interior.ColorIndex = 3
                End If
                
                ticker_count = ticker_count + 1
                
            Else
                ' Get the open price at unique ticker's first row
                If (total_vol = 0) Then
                    open_price = fy.Cells(counter, 3).Value
                End If
                
                ' Increment the total volume
                total_vol = total_vol + fy.Cells(counter, 7).Value
                
            End If
                
        Next counter
        
        ' Calculated value variables
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_volume As LongLong
        Dim get_ticker As String
        
        ' Initialise counter to get all rows
        Dim greatest_counter As Long
        greatest_counter = 2
        
        ' Get last row of unique tickers
        Dim last_unique As Long
        last_unique = fy.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Initialise variables for comparison
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0
        
        For greatest_counter = 2 To last_unique
            get_ticker = fy.Cells(greatest_counter, 9).Value
    
            ' Check percent_change column, single if-statement since mutually exclusive
            If (fy.Cells(greatest_counter, 11).Value > greatest_increase) Then
                ' Update the value
                greatest_increase = fy.Cells(greatest_counter, 11).Value
                
                ' Set new greatest_increase and get ticker
                fy.Range("Q2").Value = FormatPercent(greatest_increase)
                fy.Range("P2").Value = get_ticker
                
            ElseIf (fy.Cells(greatest_counter, 11).Value < greatest_decrease) Then
                ' Update the value
                greatest_decrease = fy.Cells(greatest_counter, 11).Value
                
                ' Set new greatest_decrease and get ticker
                fy.Range("Q3").Value = FormatPercent(greatest_decrease)
                fy.Range("P3").Value = get_ticker
                
            End If
            
            ' Find the greatest total volume
            If (fy.Cells(greatest_counter, 12).Value > greatest_volume) Then
                ' Update the value
                greatest_volume = fy.Cells(greatest_counter, 12).Value
                
                ' Set new greatest_volume and get ticker
                fy.Range("Q4").Value = greatest_volume
                fy.Range("P4").Value = get_ticker
            End If
            
        Next greatest_counter
        
        ' Autofit column formatting
        fy.Range("A:Q").Columns.AutoFit
        
    Next fy
    
End Sub
