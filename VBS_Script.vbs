Attribute VB_Name = "Module1"
Sub MultipleYearStockAnalysis()

    'Define list of Variables that will be used -Enitre variable list ketp here-
    Dim unique_ticker_count As Integer
    Dim ticker_name As String
    Dim previous_ticker_name As String
    Dim row_number As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim stock_volume As Double
    Dim max_increase(1 To 2) As Variant
    Dim max_decrease(1 To 2) As Variant
    Dim max_volume(1 To 2) As Variant
    Dim stock_volume_array(9999, 3) As Variant
    
    For Each ws In Worksheets
        
        'Sort list of Ticker symbols in Alphabetical order
        With ws.Sort
             .SortFields.Add Key:=Range("A1"), Order:=xlAscending
             .SortFields.Add Key:=Range("B1"), Order:=xlAscending
             .SetRange Columns("A:G")
             .Apply
        End With
    
        'Label new Summary table headers
        ws.Range("I1") = "Ticker"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("O1") = "Ticker"
        ws.Range("P1") = "Value"
        ws.Range("N2") = "Greatest % Increase"
        ws.Range("N3") = "Greatest % Decrease"
        ws.Range("N4") = "Greatest Total Volume"
        
       'Store "Greatest Value Summary Table"
        unique_ticker_count = 0
        row_number = 2
        ticker_name = ws.Range("A" & row_number)
        stock_volume_array(0, 0) = ticker_name

        
        ' Formatting sequence of Ticker list
        previous_ticker_name = ticker_name
        
        ' Storing maximum change and volume in index
        max_increase(1) = ""
        max_decrease(1) = ""
        max_volume(1) = ""
        max_increase(2) = 0
        max_decrease(2) = 0
        max_volume(2) = 0
        
        ' Looping through entire list of Tickers
        Do While (ticker_name <> "")
            
            If ticker_name = previous_ticker_name Then
                
                ' If duplicate ticker symbol is found, add to the total volume for that ticker symbol
                stock_volume_array(unique_ticker_count, 1) = stock_volume_array(unique_ticker_count, 1) + stock_volume
            
            Else
            
                ' Calculating yearly change and percent change
                ' Find Close price first
                close_price = ws.Range("F" & row_number - 1)
                
                ' Find difference of Close price and open price to calculate yearly change
                yearly_change = close_price - open_price
                
                ' Store yearly change in array
                stock_volume_array(unique_ticker_count, 2) = yearly_change
                
                ' After finding yearly change, calculate percentage change
                ' If change does not equal 0
                If open_price <> 0 Then
                    percent_change = (close_price - open_price) / open_price
                Else
                    percent_change = "n/a"
                End If
                
                ' Store percent change in array
                 stock_volume_array(unique_ticker_count, 3) = percent_change
                 
                ' Check to see is percent_change is the largest positive change in list , update array if it is
                If percent_change > max_increase(2) And percent_change <> "n/a" Then
                    ' Update the current max_incr
                    max_increase(1) = ticker_name
                    max_increase(2) = percent_change
                End If
                
                ' ' Check to see is percent_change is the largest negative change in list , update array if it is
                If percent_change < max_decrease(2) Then
                     ' Update the current max_decr
                    max_decrease(1) = ticker_name
                    max_decrease(2) = percent_change
                End If
                
                ' '--------------------------
                ' Prepare variables for new ticker
                ' '--------------------------
                ' Increment unique_ticker_count
                unique_ticker_count = unique_ticker_count + 1
                
                ' Store ticker_name and stock_volume
                stock_volume_array(unique_ticker_count, 0) = ticker_name
                stock_volume_array(unique_ticker_count, 1) = stock_volume
                
                ' Store new open price
                open_price = ws.Range("C" & row_number)
                
            End If
            
            ' Find the ticker that has the greatest  total volume traded in list
            ' Update ticker name if greatest total volume is found in row "A"
            previous_ticker_name = ticker_name
            
            ' Update stock volume Value for Greatest volume figure
            row_number = row_number + 1
            ticker_name = ws.Range("A" & row_number)
            stock_volume = CDbl(ws.Range("G" & row_number).Value)

        Loop
        
        
        ' Loop through list and print results
        For i = 0 To unique_ticker_count - 1
            
            ' Print ticker name
            ticker_name = stock_volume_array(i, 0)
            ws.Range("I" & i + 2) = ticker_name
            
            ' Print total volume of stock
            stock_volume = stock_volume_array(i, 1)
            ws.Range("L" & i + 2) = stock_volume
            ' Compare stock volume with max volume
            If stock_volume > max_volume(2) Then
                max_volume(1) = ticker_name
                max_volume(2) = stock_volume
            End If
            
            ' Print yearly change
            yearly_change = stock_volume_array(i, 2)
            ws.Range("J" & i + 2) = yearly_change
            
            If yearly_change > 0 Then
                ws.Range("J" & i + 2).Interior.ColorIndex = 4
            Else
                ws.Range("J" & i + 2).Interior.ColorIndex = 3
            End If
            
            ws.Range("K" & i + 2) = stock_volume_array(i, 3)
        Next i
        
        ws.Range("O2") = max_increase(1)
        ws.Range("P2") = max_increase(2)
        ws.Range("P2").NumberFormat = "0%"
        ' Greatest % decrease
        ws.Range("O3") = max_decrease(1)
        ws.Range("P3") = max_decrease(2)
        ws.Range("P3").NumberFormat = "0%"
        ' Greatest total volume
        ws.Range("O4") = max_volume(1)
        ws.Range("P4") = max_volume(2)
        
        ws.Range("K2:K" & (2 + unique_ticker_count)).NumberFormat = "0%"
        ' Format columns
        Cells.EntireColumn.AutoFit
        
        max_increase(2) = 0
        max_decrease(2) = 0
        max_volume(2) = 0
                
CountinueLoop:
    Next ws
       
    
End Sub
