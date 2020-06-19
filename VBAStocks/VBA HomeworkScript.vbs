Attribute VB_Name = "Module1"
Sub VBAHomeworkScript()

 Dim ws As Variant

'Loop through all of the stocks
For Each ws In Worksheets

    Dim i As Long
    Dim Ticker As String
    Dim Summary_Table_Row As Long
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim Ticker_volume As LongLong
    Dim lastrow As Long
    Dim o_price As Double
    Dim c_price As Double
    
    o_price = ws.Cells(2, 3).Value
    c_price = 0
    Ticker_volume = 0
    
    
    'Place the headers of our data collection table
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Year Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "TotalVolume"
    
    Summary_Table_Row = 2
    
    'Identify the worksheet name and the last row of the worksheet

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     'TEST-MsgBox ("Sheet " + wsname + " has" + Str(lastrow) + " rows")
     
     'Iterate through each row in the first column
    For i = 2 To lastrow
    
        'If the next cell in the 1st column is not equal to the current cell then
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
        
        'Assign the current value as the Ticker
        Ticker = ws.Cells(i, 1).Value
        
        'Push the Ticker Symbol into the summary table under "Ticker"
        ws.Range("J" & Summary_Table_Row) = Ticker
        
           
        'Add the volume amount of the current ticker into the total ticker value for that summary value
                Ticker_volume = Ticker_volume + Cells(i, 7).Value
        
            'Add if statement to get rid of the 0 after the last row
            If Ticker_volume <> 0 Then
            
                'Push the volume of the ticker into the column of the ticker volume
                ws.Range("M" & Summary_Table_Row) = Ticker_volume
                
                Else
                
                'Make the value in column M blank
                ws.Range("M" & Summary_Table_Row) = ""
                
            End If
        
        'Set the closing price of the current ticker
        c_price = ws.Cells(i, 6).Value
        
        'Put the difference of the opening price and the closing price in the "year change" cell
        ws.Range("K" & Summary_Table_Row) = c_price - o_price
        
        'Conditional Formatting year change cells
        If ws.Range("K" & Summary_Table_Row) > 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            
            Else
            
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            
        End If
        
        
        'Adding code to substite 0's for equation that breaks code with no opening or closing prices
            If o_price = 0 And c_price = 0 Then
                
                'Add 0 if the opening and closing price is 0 and cannot calculate a 0 % change
                
                ws.Range("L" & Summary_Table_Row) = 0
                
                Else
        
                    'If the opening price is ever 0, we need to substitute 1 as the denominator calculate % growth
                    If o_price = 0 Then
                    
                        ws.Range("L" & Summary_Table_Row) = (c_price - o_price) / 1
                        
                        Else
                        
                        ws.Range("L" & Summary_Table_Row) = (c_price - o_price) / o_price
                        
                    End If
             'Place the percent change from the opening price at the beginning of the year to the closing price of that year
             ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
        
            End If
        
        'Reset the Ticker volume to 0
        Ticker_volume = 0
        
        'Add the new ticker opening prise as the opening price variable
        o_price = ws.Cells(i + 1, 3).Value
        
        'Move to the next row in the summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Add the new ticker value as the new running total for the volume
        Ticker_volume = ws.Cells(i + 1, 7).Value
        
        'Push the volume of the new ticker into the column of the ticker volume
        ws.Range("M" & Summary_Table_Row) = Ticker_volume
        
        
        ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        
        'Add to the running total of the ticker volume
        Ticker_volume = Ticker_volume + ws.Cells(i, 7)
    
        
        End If
        
    Next i
    
    'Now let's return the stock with the greatest & Increase, Greatest % Decrease, & Greatest Total Volume
    'Declare our variables
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As LongLong
    
    'Create the table where we want to store our variables
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
 
    'Create variables to store values
   GreatestIncrease = 0
   GreatestDecrease = 0
    GreatestVolume = 0
    
    'Begin loop through summary table
    
    For i = 2 To lastrow
    
        'if the % change is greater than our stored "Greatest Increase" variable then
        If ws.Cells(i, 12).Value > GreatestIncrease Then
        
        'set the new value of the "Greater than" variable and the value of the ticker to the current value of the respective ticker
        Ticker = ws.Cells(i, 10).Value
        GreatestIncrease = ws.Cells(i, 12).Value
        
        'Place the ticker and its % change in the first row of the analysis table
        ws.Cells(2, 17).Value = Ticker
        ws.Cells(2, 18).Value = GreatestIncrease
       
        'if it's not greater than the greatest increase, let's check to see if it's lower than the greatest decrease
        ElseIf ws.Cells(i, 12).Value < GreatestDecrease Then
        
        'set the new value of the "Greatest Decrease" variable and the value of the ticker to the current value of the respective ticker
        Ticker = ws.Cells(i, 10).Value
        GreatestDecrease = ws.Cells(i, 12).Value
        
        'Place the ticker and its % change in the second row of the analysis table
        ws.Cells(3, 17).Value = Ticker
        ws.Cells(3, 18).Value = GreatestDecrease

        End If
        
        'Now since we're checking a different condition that can also apply to the same ticker, let's create another "if" statement to check the value and repeat the steps
        If ws.Cells(i, 13).Value > GreatestVolume Then
        
        'set the new value of the "Greatest Volume" variable and the value of the ticker to the current value of the respective ticker
        Ticker = ws.Cells(i, 10).Value
        GreatestVolume = ws.Cells(i, 13).Value
        
        'Place the ticker and its % change in the third (volume) row of the analysis table
        ws.Cells(4, 17).Value = Ticker
        ws.Cells(4, 18).Value = GreatestVolume
        
        End If
        
    Next i
    
     ws.Cells(2, 18).NumberFormat = "0.00%"
     ws.Cells(3, 18).NumberFormat = "0.00%"
     ws.Columns.AutoFit

     
'End the loop
    
Next ws

End Sub

