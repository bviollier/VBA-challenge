Sub Stock()

For Each ws In Worksheets

    Dim ticker As String
    Dim first_open_price As Double
    Dim last_open_price As Double
    Dim yearly_change As Double
    Dim percentage_change As Double
    

    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("N1").Value = "Initial Price"
    ws.Range("O1").Value = "Final Price"
    ws.Range("Q1").Value = "Greatest % increase"
    ws.Range("Q2").Value = "Greatest % decrease"
    ws.Range("Q3").Value = "Greatest total volume"

    first_open_price = 0
    last_open_price = 0
    percentage_change = 0
    total_volume = 0

'For Ticker, Yearly Change, Percent Change (and Initial and Final Price)
    For i = 2 To 797711
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            last_open_price = ws.Cells(i, 3).Value
            
            If ws.Cells(i, 3).Value = 0 Then
                last_open_price = 1
            End If
            
            ticker = ws.Cells(i, 1).Value
            
            yearly_change = last_open_price - first_open_price
            percentage_change = yearly_change / first_open_price
            
            ws.Range("I" & Summary_Table_Row).Value = ticker
            ws.Range("O" & Summary_Table_Row).Value = last_open_price
            ws.Range("N" & Summary_Table_Row).Value = first_open_price
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
            ws.Range("K" & Summary_Table_Row).Value = percentage_change
            
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf ws.Range("J" & Summary_Table_Row).Value > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            first_open_price = ws.Cells(i, 3).Value
             If ws.Cells(i, 3).Value = 0 Then
                first_open_price = 1
            End If
            
        

        
        End If
    Next i
    
'For Total Stock Volume
    Summary_Table_Row = 2
    
    For i = 2 To 797711
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            ws.Range("I" & Summary_Table_Row).Value = ticker
            ws.Range("L" & Summary_Table_Row).Value = total_volume
            
            Summary_Table_Row = Summary_Table_Row + 1
            total_volume = 0
        Else
            total_volume = total_volume + ws.Cells(i, 7).Value
        End If
        
    Next i
    
'For the Table
        result1 = WorksheetFunction.Max(ws.Range("K2:K3200"))
        result2 = WorksheetFunction.Min(ws.Range("K2:K3200"))
        result3 = WorksheetFunction.Max(ws.Range("L2:L3200"))
        
        ws.Range("R1").Value = result1
        ws.Range("R2").Value = result2
        ws.Range("R3").Value = result3
        
        ws.Range("R1").NumberFormat = "0.00%"
        ws.Range("R2").NumberFormat = "0.00%"
    
Next ws
    
End Sub