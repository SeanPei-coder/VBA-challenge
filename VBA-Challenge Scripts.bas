Attribute VB_Name = "Module1"

Sub yearly_change_and_percentchange()

Dim i As Long, j As Integer, open_value As Double, close_value As Double

For Each ws In Worksheets

i = 2
j = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If ws.Cells(i, 1).Value = ws.Cells(j, 9).Value And ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
        open_value = ws.Cells(i, 3).Value
        On Error Resume Next
        ws.Cells(j, "K").Value = ws.Cells(j, "J").Value / ws.Cells(i, 3).Value
        ws.Cells(j, "K").Style = "Percent"
        ws.Cells(j, "K").NumberFormat = "0.00%"
        
    ElseIf ws.Cells(i, 1).Value = ws.Cells(j, 9).Value And ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
        close_value = ws.Cells(i, 6).Value
        ws.Cells(j, "J").Value = close_value - open_value
            If ws.Cells(j, "J").Value > 0 Then
                ws.Cells(j, "J").Interior.Color = vbGreen
            Else
                ws.Cells(j, "J").Interior.Color = vbRed
            End If
        j = j + 1
    End If
    
Next i

Next ws


End Sub
Sub Ticker()

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim tickers As String
Dim j As Long

j = 2

For i = 2 To lastrow
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        tickers = ws.Cells(i, 1).Value
        ws.Cells(j, "I").Value = tickers
        j = j + 1
        
    End If
Next i
Next ws


End Sub

Sub vol_add()

For Each ws In Worksheets

Dim vol As Double, j As Long
j = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        vol = vol + ws.Cells(i, 7).Value
    Else
        vol = vol + ws.Cells(i, 7).Value
        ws.Cells(j, "L").Value = vol
        vol = 0
        j = j + 1
    End If
Next i
Next ws

End Sub

Sub header()

For Each ws In Worksheets
    ws.Cells(1, "I").Value = " Ticker"
    ws.Cells(1, "J").Value = " Yearly Change"
    ws.Cells(1, "K").Value = " Percnet Change"
    ws.Cells(1, "L").Value = " Total Stock Volume"
    
    
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
Next ws
End Sub

Sub challengs()

For Each ws In Worksheets

Dim max As Double, min As Double, max_vol As Double

max = ws.Application.WorksheetFunction.max(ws.Range("K:K"))

min = ws.Application.WorksheetFunction.min(ws.Range("K:K"))

max_vol = ws.Application.WorksheetFunction.max(ws.Range("L:L"))


lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    For i = 2 To lastrow
        If ws.Cells(i, "K").Value = max Then
            ws.Range("P2").Value = ws.Cells(i, "K").Offset(, -2).Value
            ws.Range("Q2").Value = ws.Cells(i, "K").Value
            ws.Range("Q2").Style = "Percent"
            ws.Range("Q2").NumberFormat = "0.00%"
            
        ElseIf ws.Cells(i, "K").Value = min Then
            ws.Range("P3").Value = ws.Cells(i, "K").Offset(, -2).Value
            ws.Range("Q3").Value = ws.Cells(i, "K").Value
            ws.Range("Q3").Style = "Percent"
            ws.Range("Q3").NumberFormat = "0.00%"
            
        ElseIf ws.Cells(i, "L").Value = max_vol Then
            ws.Range("P4").Value = ws.Cells(i, "L").Offset(, -3).Value
            ws.Range("Q4").Value = ws.Cells(i, "L").Value
            
        End If
    Next i
Next ws

           
End Sub


