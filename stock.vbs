Sub stockanalysis()
'declaring variables
Dim ticker As String
Dim openprice, closeprice, total As Double
Dim tablerow As Integer


For Each ws In Worksheets
'defining variables
total = 0
openprice = ws.Cells(2, "C").Value
tablerow = 2
'defining variables
ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    total = total + ws.Cells(i, "G")
    If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
        ws.Cells(tablerow, "L") = total
        ws.Cells(tablerow, "I") = ws.Cells(i, "A")
        closeprice = ws.Cells(i, "F")
        ws.Cells(tablerow, "J") = closeprice - openprice
        If ws.Cells(tablerow, "J") > 0 Then
            ws.Cells(tablerow, "J").Interior.Color = RGB(0, 255, 0)
        Else
            ws.Cells(tablerow, "J").Interior.Color = RGB(255, 0, 0)
        End If



        If openprice <> 0 Then
            ws.Cells(tablerow, "K") = FormatPercent((closeprice - openprice) / openprice, 2)
        Else
            ws.Cells(tablerow, "K") = 0
        End If

        tablerow = tablerow + 1
        openprice = ws.Cells(i + 1, "C")
        total = 0
    End If
Next i


Next ws

End Sub
