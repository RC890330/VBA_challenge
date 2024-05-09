Sub multiple_year_stock_data():

    For Each ws In Worksheets
        
        'Define variables'
        Dim Ticker As String
        Dim Row As Integer
        Dim QuaChange As Double
        Dim PerChange As Double
        Dim TotalV As LongLong
        Dim ClosingPrice As Double
        Dim OpenPrice As Double
        
        'Setting for variables
        TotalV = 0
        TotalRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        Row = 2
        
        'Heading for the result columns'  
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        
        OpenPrice = Cells(2, 3).Value
        
        'Listing out all the variable present in each ws'
        For i = 2 To TotalRow
        
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                
                'This Part is for Ticker & Total Volume'
                Ticker = ws.Cells(i, 1).Value
                
                TotalV = TotalV + Cells(i, 7).Value
            
                ws.Range("I" & Row).Value = Ticker
                
                ws.Range("L" & Row).Value = TotalV
                
                'This part is for QuarterlyChange'
                ClosingPrice = Cells(i, 6).Value
                
                QuaChange = ClosingPrice - OpenPrice
                
                ws.Range("J" & Row).Value = QuaChange
                
                'Conditional formatting highlight negative value w/ red and positive w/ green'
                If (QuaChange < 0) Then
                    
                    ws.Range("J" & Row).Interior.ColorIndex = 3
                
                Else
                    ws.Range("J" & Row).Interior.ColorIndex = 4
                
                End If
                
                'Code for PercentChange'
                PerChange = (ClosingPrice / OpenPrice) - 1
                
                ws.Range("K" & Row).Value = PerChange
                
                OpenPrice = ws.Cells(i + 1, 3).Value
            
                Row = Row + 1
                
                TotalV = 0
            
            Else
                
                TotalV = TotalV + Cells(i, 7).Value
            
            End If
            
        Next i
        
        'Script for "Greatest % increase", "Greatest % decrease", and "Greatest total volume"'
        'Heading for the table'
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volumn"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Total row for percent change (TrPc) & Total Stock Volume (Tsv) & ticker'
        TrPc = ws.Cells(Rows.Count, "K").End(xlUp).Row
        Tsv = ws.Cells(Rows.Count, "L").End(xlUp).Row
        TickRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        'Find Greatest % increase'
        ws.Range("Q2").Value = WorksheetFunction.max(ws.Range("K2:K" & TrPc))
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
        
        'Greatest % decrease'
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & TrPc))
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Greatest total volume'
        ws.Range("Q4").Value = WorksheetFunction.max(ws.Range("L2:L" & Tsv))
        
        'Lookup ticker for Greatest % increase'
        ws.Range("P2").Value = WorksheetFunction.Lookup(ws.Range("Q2").Value, ws.Range("K2:K" & TrPc), ws.Range("I2:I" & TrPc))
        'ws.Range("P2").Value = WorksheetFunction.Lookup(ws.Range("Q2").Value, ws.Range("K2:K" & TrPc), ws.Range("I2:I" & TickRow))
    
        'Lookup ticker for Greatest % decrease'
        'ws.Range("P3").Value = WorksheetFunction.Lookup(ws.Range("Q3").Value, ws.Range("K2:K" & TrPc), ws.Range("I2:I" & TickRow))
        
        'Lookup ticker for Greatest total stock volume'
        'ws.Range("P4").Value = WorksheetFunction.Lookup(ws.Range("Q4").Value, ws.Range("K2:K" & TrPc), ws.Range("I2:I" & TickRow))
    
        'FORMAT'
        'AutoFit all columns from I to Q'
        ws.Columns("I:Q").AutoFit
        
        'Format the percent change column to %'
        ws.Range("K2:K" & TrPc).Style = "Percent"
        ws.Range("K2:K" & TrPc).NumberFormat = "0.00%"
        
    Next ws

End Sub


