Sub alphabtical_testing():

'Variables
Dim ticker As String 'ticker symbol
Dim ticker_counter As Integer 'ticker_counter
Dim final_row As Long 'final row in sheet
Dim o_price As Double 'opening price
Dim c_price As Double 'closing price
Dim yearly_change As Double 'yearly change in price
Dim per_change As Double 'percentage change (yearly change / opening price)
Dim total_sv As Double 'total stock volume
Dim greatest_per_change As Double 'greatest percentage change
Dim ticker_w_gc As String 'ticker with greatest percentage change
Dim greatest_per_decrease As Double 'greatest decrease percentage change
Dim ticker_w_gdc As String 'ticker with greatest decrese percentage change
Dim greatest_sv As Double 'greatest total stock volume
Dim ticker_w_gsv As String 'ticker with greatest total stock volume

'For each loop to loop through each worksheet
For Each ws In Worksheets

    'Activate current worksheet
    ws.Activate

    'Set final row to index of final row on current sheet
    final_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'Add the colums headers in each worksheet
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    ws.Range("O2:O4").Value = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    ws.Range("P1:Q1").Value = Array("Ticker", "Value")
    
    'Initialize variables being used
    ticker_counter = 0
    ticker = ""
    yearly_change = 0
    o_price = 0
    per_change = 0
    total_sv = 0
    
    'For loop to traverse through each row in worksheet
    For i = 2 To final_row
        ticker = Cells(i, 1).Value 'set ticker value to current cell
        
        'If statement to check for beggining of year
        If o_price = 0 Then
            o_price = Cells(i, 3).Value
        End If
        
        'Running total of total stock volume
        total_sv = total_sv + Cells(i, 7).Value
        
        'If statement runs until ticker value changes
        If Cells(i + 1, 1).Value <> ticker Then
            
            ticker_counter = ticker_counter + 1 'increment ticker counter
            Cells(ticker_counter + 1, 9) = ticker
    
            c_price = Cells(i, 6) 'set closing price
            yearly_change = c_price - o_price 'set yearly change
            
            Cells(ticker_counter + 1, 10).Value = yearly_change 'set desired cell with yearly change value
            
            'If statement checks if yearly change value is positive, sets cell green
            If yearly_change >= 0 Then
                Cells(ticker_counter + 1, 10).Interior.ColorIndex = 4
            'ElseIf statements checks if  yearly change value is megative, set cell red
            ElseIf yearly_change < 0 Then
                Cells(ticker_counter + 1, 10).Interior.ColorIndex = 3
            End If
            
            'If statement checks if opening price is 0
            If o_price = 0 Then
                per_change = 0
            'Else statement sets percentage change if opening price <> 0
            Else
                per_change = (yearly_change / o_price)
            End If
            
            Cells(ticker_counter + 1, 11).Value = Format(per_change, "Percent") 'Formats percentage change as 0.00
           
            
            o_price = 0 'Resets opening price when ticker changes
            

            Cells(ticker_counter + 1, 12).Value = total_sv 'Set desired cell with total stock volume value
            
            total_sv = 0 'Resets total stock volume when ticker changes
        End If
        
    Next i
    'Add headers for this section to each worksheet
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    final_row = ws.Cells(Rows.Count, "I").End(xlUp).Row 'sets final row to the last row in index of Column I
    
    'Initialize varibles for this section
    greatest_per_change = Cells(2, 11).Value
    ticker_w_gc = Cells(2, 9).Value
    greatest_per_decrease = Cells(2, 11).Value
    ticker_w_gdc = Cells(2, 9).Value
    greatest_sv = Cells(2, 12).Value
    ticker_w_gsv = Cells(2, 9).Value
    
    
    'For loop to traverse through each row in worksheet for Column I
    For i = 2 To final_row
    
        'If statement to compare values, greatest value is set to greatest percentage change
        If Cells(i, 11).Value > greatest_per_change Then
            greatest_per_change = Cells(i, 11).Value
            ticker_w_gc = Cells(i, 9).Value
        End If
        
        'If statement to compare values, greatest value is set to greatest percentage decrease change
        If Cells(i, 11).Value < greatest_per_decrease Then
            greatest_per_decrease = Cells(i, 11).Value
            ticker_w_gdc = Cells(i, 9).Value
        End If
        
        'If statement to compare values, greatest value is set to greatest stock volume
        If Cells(i, 12).Value > greatest_sv Then
            greatest_sv = Cells(i, 12).Value
            ticker_w_gsv = Cells(i, 9).Value
        End If
        
    Next i
    
    'Sets desired cells with desired variable values
    Range("P2").Value = Format(ticker_w_gc, "Percent")
    Range("Q2").Value = Format(greatest_per_change, "Percent")
    Range("P3").Value = Format(ticker_w_gdc, "Percent")
    Range("Q3").Value = Format(greatest_per_decrease, "Percent")
    Range("P4").Value = ticker_w_gsv
    Range("Q4").Value = greatest_sv
    
    'Note cells being filled in sheet P for unknown reason, this clears that unwanted contents
    ws.Columns(18).ClearContents
    ws.Columns(19).ClearContents
    
Next ws


End Sub
