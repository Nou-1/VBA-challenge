Attribute VB_Name = "Module1"
Sub stock_market()
    
    For Each ws In Worksheets
    
        Dim Ticker As String
        
        Dim Summary_Row As Integer
        Summary_Row = 2
        
        Dim Yearly_Change As Double
        Yearly_Change = 0
        
        Dim Total_volume As Double
        Total_volume = 0
        
        Dim j As Long
        j = 2
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ws.Range("i" & Summary_Row).Value = Ticker
                
                Yearly_Change = ws.Cells(i, 6) - ws.Cells(j, 3).Value
                ws.Range("J" & Summary_Row).Value = Yearly_Change
                
                If ws.Cells(j, 3).Value <> 0 Then
                    Percent_Change = Yearly_Change / ws.Cells(j, 3).Value
                Else: Percent_Change = 0
                End If
                
                ws.Range("K" & Summary_Row).Value = Format(Percent_Change, "Percent")
                
                Total_volume = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                ws.Range("L" & Summary_Row) = Total_volume
                
                    If Yearly_Change > 0 Then
                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                    ElseIf Yearly_Change < 0 Then
                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                    End If
                
                Summary_Row = Summary_Row + 1
                Yearly_Change = 0
                Percent_Change = 0
                Total_volume = 0
                j = i + 1
            End If
            
        Next i
    Next ws
        



End Sub
