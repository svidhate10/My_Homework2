Sub my_procedure()
 
    
    Dim outputRow As Long
    Dim inputRow As Long
    Dim currentTicker As String
    Dim currentTickerVolume As Double
    Dim openvalue As Double
    Dim closevalue As Double
    Dim percentchange As Double
    
    Dim yearlyValuechange As Double
    Dim increase As Double
    Dim decrease As Double
    Dim maxtotalvolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
    
    'loop through each worksheet
    For Each ws In Worksheets
    
        'set column titles
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "YearlyChange"
        ws.Cells(1, 12).Value = "PercentChange"
        ws.Cells(1, 13).Value = "TotalStockvolume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'set counter for output rows
        outputRow = 2
        
         
        
        'set counter to read first input rows
        inputRow = 2
        
        'find out how many input rows to be read
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'initialize variables
        increase = 0
        decrease = 0
        maxtotalvolume = 0
        
        
        'Easy
        'loop through all input rows
        Do While inputRow < LastRow
        
            currentTicker = ws.Cells(inputRow, 1).Value
            currentTickerVolume = 0
            openvalue = ws.Cells(inputRow, 3).Value
            closevalue = 0
                      
            'read all input data rows for current ticker; read it until the ticker changes...
            For i = inputRow To LastRow
                currentTickerVolume = currentTickerVolume + ws.Cells(i, 7).Value
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    closevalue = ws.Cells(i, 6).Value
                    'ticker is changed and this indicates we have read all rows for current ticker... so exit the loop
                    Exit For
                End If
                
            Next i
            
            'Moderate
            yearlyValuechange = closevalue - openvalue
            
            If (openvalue = 0) Then
                percentchange = 0
            Else
                percentchange = (yearlyValuechange / openvalue) * 100
            End If
             
             ' print values in spreadsheet
            ws.Cells(outputRow, 10).Value = currentTicker
            ws.Cells(outputRow, 11).Value = yearlyValuechange
            ws.Cells(outputRow, 12).Value = percentchange & "%"
            ws.Cells(outputRow, 13).Value = currentTickerVolume
            'ws.Cells(outputRow, 20) = percentchange 'for test only
            
            If (ws.Cells(outputRow, 11).Value < 0) Then
              ws.Cells(outputRow, 11).Interior.ColorIndex = 3
        
            Else
            If (ws.Cells(outputRow, 11).Value > 0) Then
                ws.Cells(outputRow, 11).Interior.ColorIndex = 4
            End If
            End If
             
             'Hard
             'find Greatest Percent Increase
            If (percentchange > increase) Then
                increase = percentchange
                increaseTicker = ws.Cells(inputRow, 1).Value
            End If
            
                         
            'find greatest percentage decrease
            If (percentchange < decrease) Then
                decrease = percentchange
                decreaseTicker = ws.Cells(inputRow, 1).Value
            End If
            
             'find greatest volume change
            If (currentTickerVolume > maxtotalvolume) Then
                maxtotalvolume = currentTickerVolume
                maxTotalVolumeTicker = ws.Cells(inputRow, 1).Value
            End If
            
            'increment output row counter by 1
            outputRow = outputRow + 1
            
            'increment input row counter to next ticker row
            inputRow = i + 1
        Loop
        
        'Finding Greatest % Increase and Greatest % Decrease
        ws.Cells(2, 16).Value = increaseTicker
        ws.Cells(2, 17).Value = increase
        
        ws.Cells(3, 16).Value = decreaseTicker
        ws.Cells(3, 17).Value = decrease
        
        ws.Cells(4, 16).Value = maxTotalVolumeTicker
        ws.Cells(4, 17).Value = maxtotalvolume
        
        
    Next ws
    
End Sub




