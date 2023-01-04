Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim UndoString As String, srce As Range
    On Error GoTo err_handler
    UndoString = Application.CommandBars("Standard").Controls("&Undo").List(1)
    If Left(UndoString, 5) <> "Paste" And UndoString <> "Auto Fill" Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Undo
    If UndoString = "Auto Fill" Then
        Set srce = Selection
        srce.Copy
        Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.SendKeys "{ESC}"
        Union(Target, srce).Select
    Else
        Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
err_handler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

                
                'edit 
                '
                
                '
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim UndoString As String, srce As Range
    On Error GoTo err_handler
    UndoString = Application.CommandBars("Standard").Controls("&Undo").List(1)
    If Left(UndoString, 5) <> "Paste" And UndoString <> "Auto Fill" Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Undo
    If UndoString = "Auto Fill" Then
        Set srce = Selection
        srce.Copy
        Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.SendKeys "{ESC}"
        Union(Target, srce).Select
    Else
        Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
err_handler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

                                End Sub
                            End Sub
                        End Sub
                    End Sub
