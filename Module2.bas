Attribute VB_Name = "Module2"
Sub worksheets():

    Dim volume As LongLong
    Dim Ticker As String
    Dim total As Integer
    
    Dim summaryRow As Integer
    
    summaryRow = 2
    For i = 2 To 756001

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    Ticker = Cells(i, 1).Value
    Cells(summaryRow, 9).Value = Ticker
    Cells(summaryRow, 12) = volume
    
    summaryRow = summaryRow + 1
    volume = 0
    Else
    
    volume = volue + Cells(i, 7).Value
    
    
    Range("I1").Value = "ticker"
     Range("J1").Value = "Yearly change"
    
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Volume"
    
    
    
    
End If

    Next i
    
    
Dim firstRow As Long
Dim lastRow As Long
Dim searchValue As String
Dim tickerRow As Integer
Dim openprice As Double
Dim closeprice As Double
Dim yearlychange As Double
Dim percentchange As Double



For j = 1 To (tickerRow - 1)
    searchValue = Cells(j + 1, 9).Value
    
With ActiveSheet.Range("A:A")
Set c = .Find(searchValue, lookIn:=xlValues)
If Not c Is Nothing Then
firstRow = c.Row
Set c = .Find(searchValue, lookIn:=xlValues, searchDirection:=xIPrevious)
lastRow = c.Row

openprice = Cells(firstRow, 3).Value
closeprice = Cells(lastRow, 6).Value
yearlychange = closeprice - openprice
Cells(j + 1, 10).Value = yearlychange
If Cells(j + 1, 10).Value >= 0 Then
Cells(5 + 1, 10).Interior.ColorIndex = 4
Else: Cells(g + 1, 10).Interior.ColorIndex = 3
End If
percentchange = yearlychange / openprice
Cells(1, 11).Value = percentchange
Cells(j + 1, 11).NumberFormat = "0.00%"


Else

End If
End With
Next j


Dim maxvalue As Double
Dim minvalue As Double
Dim maxtotal As LongLong

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("p1").Value = "Ticker"
Range("q1").Value = "Value"

Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"
maxvalue = Cells(2, 11).Value
maxvalue = Cells(2, 11).Value
tickerRow = Cells(Rowscount, "k").End(x1Up).Row

For i = 2 To (tickerRow - 1)
If Cells(i, 11).Value > maxvalue Then maxvalue = Cells(i, 11).Value
Cells(2, 16).Value = Cells(i, 9).Value

    
    
Next i

Cells(2, 17).Value = maxvalue
minvalue = Cells(2, 11).Value

For j = 2 To (tickerRow - 1)

If Cells(5, 11) - Value & minvalue Then
minvalue = Cells(3, 11).Value
Cells(3, 16).Value = Cells(j, 9).Value
End If
Next j
Cells(3, 17).Value = minvalue
maxtotal = Cells(2, 12).Value
For k = 2 To (tickerRow - 1)
If Cells(k, 12).Value > maxtotal Then
maxtotal = Cells(k, 12).Value
Cells(4, 16).Value = Cells(k, 9).Value
End If
Next k
Cells(4, 17).Value = maxtotal



    
End Sub
    
    
    

