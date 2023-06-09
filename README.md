# Analyzing Stock Data with VBA
Sub Stock_Data()
    For Each ws In Worksheets
        Dim WorksheetName As String
        Dim Ticker As String
        Dim Total As Double
        Dim i As Long
        Dim j As Long
        Dim Ticker_Counter As Long
        Dim LastRow As Long
        Dim percent_change As Double
        
        'Set title row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Get worksheet name
        WorksheetName = ws.Name
        MsgBox (WorksheetName)
        'Set variables
        Ticker_Counter = 2
        j = 0
        Total = 0
        Change = 0
        start_price = 0
        end_price = 0
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Loop through rows to make calculations
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Total = Total + ws.Cells(i, 7).Value
                end_price = ws.Cells(i, 6).Value
                Change = end_price - start_price
                ws.Cells(Ticker_Counter, 10).Value = Change
                ws.Cells(Ticker_Counter, 11).Value = Change / start_price
                'Put ticker in new table
                ws.Range("I" & Ticker_Counter).Value = Ticker
                'Put volume in new table
                ws.Range("L" & Ticker_Counter).Value = Total
                If ws.Cells(Ticker_Counter, 10).Value < 0 Then
                    ws.Cells(Ticker_Counter, 10).Interior.ColorIndex = 3
                    Else
                    ws.Cells(Ticker_Counter, 10).Interior.ColorIndex = 4
                End If
                Ticker_Counter = Ticker_Counter + 1
                Total = 0
                Change = 0
                
                ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    start_price = ws.Cells(i, 3).Value
                    Ticker = ws.Cells(i, 1).Value
                    Total = Total + ws.Cells(i, 7).Value
            'If cell immediately following a row is the same ticker
            Else
            Total = Total + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        'For new table
        'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        Greatest_Vol = ws.Cells(2, 12).Value
        Greatest_Incr = ws.Cells(2, 11).Value
        Greatest_Decr = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'If another value in the volume column is bigger, populate the new greatest volume
                If ws.Cells(i, 12).Value > Greatest_Vol Then
                Greatest_Vol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                Else
                Greatest_Vol = Greatest_Vol
                End If
                
                'If another value is bigger, populate new greatest increase
                If ws.Cells(i, 11).Value > Greatest_Incr Then
                Greatest_Incr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                Else
                Greatest_Incr = Greatest_Incr
                End If
                
                'If another value is bigger in percent_change, populate new greatest decrease
                If ws.Cells(i, 11).Value < Greatest_Decr Then
                Greatest_Decr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                Else
                Greatest_Decr = Greatest_Decr
                End If
            ws.Cells(2, 17).Value = Format(Greatest_Incr, "Percent")
            ws.Cells(3, 17).Value = Format(Greatest_Decr, "Percent")
            ws.Cells(4, 17).Value = Format(Greatest_Vol, "Scientific")
            
            Next i
        
    Next ws
End Sub
