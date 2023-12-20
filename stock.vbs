Sub stockanalysis()
Dim ticker As String
Dim openprice, closeprice, total As Double
Dim tablerow As Integer


For Each ws In Worksheets

total = 0
openprice = ws.Cells(2, "C").Value
tablerow = 2

'create headers for stock table
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
'can end the first part here with next ws and end sub

'create summary table
ws.Cells(1, "P") = "Ticker"
ws.Cells(1, "Q") = "Value"
ws.Cells(2, "O") = "Greatest % Increase"
ws.Cells(3, "O") = "Greatest % Decrease"
ws.Cells(4, "O") = "Greatest Total Volume"

'declare summary table variables

Dim maxvalue As Double
Dim minvalue As Double
Dim summarytotal As Double
Dim maxticker As String
Dim minticker As String
Dim totalticker As String
Dim summarytablerow As Integer


'define initial values of summary table variables

maxvalue = ws.Cells(2, "K")
minvalue = ws.Cells(2, "K")
summarytotal = ws.Cells(2, "L")
maxticker = ws.Cells(2, "I")
minticker = ws.Cells(2, "I")
totalticker = ws.Cells(2, "I")
summarytablerow = 2

For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
'check if maxvalue has changed, if so redefine

If ws.Cells(i, "K") > maxvalue Then
maxvalue = ws.Cells(summarytablerow, "K")
maxticker = ws.Cells(summarytablerow, "I")

End If
'check if minvalue has changed, if so redefine
If ws.Cells(i, "K") < minvalue Then
minvalue = ws.Cells(summarytablerow, "K")
minticker = ws.Cells(summarytablerow, "I")

End If

'check if minvalue has changed, if so redefine
If ws.Cells(i, "L") > summarytotal Then
summarytotal = ws.Cells(summarytablerow, "L")
totalticker = ws.Cells(summarytablerow, "I")

End If
summarytablerow = summarytablerow + 1


Next i
ws.Cells(2, "P") = maxticker
ws.Cells(2, "Q") = maxvalue
ws.Cells(3, "P") = minticker
ws.Cells(3, "Q") = minvalue
ws.Cells(4, "P") = totalticker
ws.Cells(4, "Q") = summarytotal


Next ws

End Sub
