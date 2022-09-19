Function NS_DOC(ar As Range, x As Double, n As Byte) 'bien n de noi suy trong truong hop co nhieu hon 3 cot
Dim i As Integer
If ar.Cells(2, 1) > ar.Cells(1, 1) Then
    i = 1
    Do While ar.Cells(i, 1) <= x And i < ar.Rows.Count
        i = i + 1
    Loop
Else
    i = 1
    Do While ar.Cells(i, 1) >= x And i < ar.Rows.Count
        i = i + 1
    Loop
End If
NS_DOC = ar.Cells(i - 1, n) + (ar.Cells(i, n) - ar.Cells(i - 1, n)) * (x - ar.Cells(i - 1, 1)) / (ar.Cells(i, 1) - ar.Cells(i - 1, 1))
End Function

Function NS2C(ar As Range, x As Double, y As Double)
Dim i, j As Integer
Dim a1, a2 As Double
'xac dinh chi so i voi cells(i-1,1)<x<cells(i,1)
If ar.Cells(3, 1) > ar.Cells(2, 1) Then
    i = 2
    Do While ar.Cells(i, 1) <= x And i < ar.Rows.Count
        i = i + 1
    Loop
Else
    i = 2
    Do While ar.Cells(i, 1) >= x And i < ar.Rows.Count
        i = i + 1
    Loop
End If
'xac dinh chi so j voi cells(1,j-1)<y<cells(1,j)
If ar.Cells(1, 3) > ar.Cells(1, 2) Then
    j = 2
    Do While ar.Cells(1, j) <= y And j < ar.Columns.Count
        j = j + 1
    Loop
Else
     j = 2
    Do While ar.Cells(1, j) >= y And j < ar.Columns.Count
        j = j + 1
    Loop
End If
'xac dinh 2 gia tri a1, a2 tu noi suy 1 chieu voi x truoc
a1 = ar.Cells(i - 1, j - 1) + (ar.Cells(i, j - 1) - ar.Cells(i - 1, j - 1)) * (x - ar.Cells(i - 1, 1)) / (ar.Cells(i, 1) - ar.Cells(i - 1, 1))
a2 = ar.Cells(i - 1, j) + (ar.Cells(i, j) - ar.Cells(i - 1, j)) * (x - ar.Cells(i - 1, 1)) / (ar.Cells(i, 1) - ar.Cells(i - 1, 1))
'noi suy 1 chieu theo cot tu 2 gia tri a1 va a2 o tren
NS2C = a1 + (a2 - a1) * (y - ar.Cells(1, j - 1)) / (ar.Cells(1, j) - ar.Cells(1, j - 1))
End Function

Function NS_NGANG(ar As Range, x As Double, n As Byte)
Dim j As Integer
If ar.Cells(1, 2) > ar.Cells(1, 1) Then
    j = 1
    Do While ar.Cells(1, j) <= x And j < ar.Columns.Count
        j = j + 1
    Loop
Else
     j = 1
    Do While ar.Cells(1, j) >= x And j < ar.Columns.Count
        j = j + 1
    Loop
End If
NS_NGANG = ar.Cells(n, j - 1) + (ar.Cells(n, j) - ar.Cells(n, j - 1)) * (x - ar.Cells(1, j - 1)) / (ar.Cells(1, j) - ar.Cells(1, j - 1))
End Function
' End Function
' End Function
