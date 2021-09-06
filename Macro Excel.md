- 👋 Hi, I’m @thephuoc
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...

<!---
thephuoc/thephuoc is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->

=XLOOKUP(C3,'TU DONG DAP TRAN LOGGER'!$C$3:$C$14514,'TU DONG DAP TRAN LOGGER'!$D$3:$D$14514,0,-1,-1)


=XLOOKUP(C3,'[Auto Source.xlsx]2015'!$C$2:$C$2507,'[Auto Source.xlsx]2015'!$D$2:$D$2507,,-1,1)


=TIME(HOUR(NOW())+12,MINUTE(NOW()),SECOND(NOW()))

=TIME(HOUR(NOW())+RANDBETWEEN(1,23),MINUTE(NOW())+RANDBETWEEN(1,59),SECOND(NOW()))

=DATE(2019,3,RANDBETWEEN(1,14))

=DATE(2020,2,RANDBETWEEN(1,14))

=DATE(2021,2,RANDBETWEEN(1,14))











Sub Valuu()
'
' Valuu Macro
'

'
    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-21
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    ActiveWindow.SmallScroll Down:=-39
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    ActiveWindow.SmallScroll Down:=-18
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-3
    Range("B3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
End Sub




