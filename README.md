'code for VBA-Challenge

Sub MultipleQuarterlyStockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'Row
        Dim i As Long
        'Start ticker block row
        Dim j As Long
        Dim TickCount As Long
        'Last row column A
        Dim LastRowA As Long
        'Last row column I
        Dim LastRowI As Long
        'varaibles for percent changes
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        
        'Get the WorksheetName
        WorksheetName = ws.Name
        
        'Create Column Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker Counter
        TickCount = 2
        
        'Start row to 2
        j = 2
        
        'Find the last row with data in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (Last row in colum A is _ & LastRowA)
        
            'Loop through all rows
            For i = 2 To LastRowA
            
                'If ticker changes then print results
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'print the results
                    If ws.Cells(TickCount, 10).Value < 0 Then
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                    Else
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                    End If
                    
                    'Print the results
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    Else
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    End If
                    
                'calculate and write total volume in column
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                TickCount = TickCount + 1
                
                'New start row of ticker block
                j = i + 1
                
                End If
            
            Next i
            'Find last row of data in column i
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Setting up a section for summary of data
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'to find total volume
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                'to find greatest increase
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                'to find greatest decrease
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
            
            'Summary results
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
        'Adjust column with-in automatically
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
