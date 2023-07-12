Sub Module2()
    Dim totVol As Double
    Dim curSymbol As String
    Dim openYear As Double
    Dim closeYear As Double
    Dim row As Long
    Dim inputRow As Long
    Dim diff As Double
    Dim percent As Double
    Dim ws As Worksheet

    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"

        totVol = 0
        curSymbol = ""
        openYear = -1
        closeYear = -1
        row = 2
        inputRow = 2
        While Not IsEmpty(ws.Cells(row, 1))
            If Right(ws.Cells(row, 2), 4) = "0102" Then
                curSymbol = ws.Cells(row, 1).Value
                openYear = ws.Cells(row, 3).Value
                totVol = totVol + ws.Cells(row, 7).Value
            ElseIf Right(ws.Cells(row, 2), 4) = "1231" Then
                closeYear = ws.Cells(row, 6).Value
                totVol = totVol + ws.Cells(row, 7).Value
                diff = closeYear - openYear
                ws.Cells(inputRow, 9).Value = curSymbol
                ws.Cells(inputRow, 10).Value = diff
                If diff < 0 Then
                    ws.Cells(inputRow, 10).Interior.Color = RGB(255, 0, 0)
                Else
                    ws.Cells(inputRow, 10).Interior.Color = RGB(0, 255, 0)
                End If
                percent = diff / openYear
                ws.Cells(inputRow, 11).Value = percent
                ws.Cells(inputRow, 11) = Format(ws.Cells(inputRow, 11), "0.00%")
                ws.Cells(inputRow, 12).Value = totVol
                totVol = 0
                inputRow = inputRow + 1
            Else
                totVol = totVol + ws.Cells(row, 7).Value
            End If
            row = row + 1
        Wend

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        row = 2
        Dim tickers(2) As String
        Dim values(2) As Double
        
        values(0) = -1
        values(1) = 1
        values(2) = -1
        
        While Not IsEmpty(ws.Cells(row, 9))
            If ws.Cells(row, 11) > values(0) Then
                tickers(0) = ws.Cells(row, 9).Value
                values(0) = ws.Cells(row, 11).Value
            End If
            If ws.Cells(row, 11) < values(1) Then
                tickers(1) = ws.Cells(row, 9).Value
                values(1) = ws.Cells(row, 11).Value
            End If
            If ws.Cells(row, 12) > values(2) Then
                tickers(2) = ws.Cells(row, 9).Value
                values(2) = ws.Cells(row, 12).Value
            End If
            row = row + 1
        Wend
        
        For i = 2 to 4
            ws.Cells(i, 16).Value = tickers(i-2)
            ws.Cells(i, 17).Value = values(i-2)
            If i <> 4 Then
                ws.Cells(i, 17) = Format(ws.Cells(i, 17), "0.00%")
            End If
        Next i
        
    Next
End Sub

