Sub VBA100_14()

    Dim ws As Worksheet

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationAutomatic '一度自動計算走らせてすべて計算させておく
        .Calculation = xlCalculationManual
    End With

    For Each ws In Worksheets
        If ws.Name Like "*社外秘*" Then
            Call deleteSheet(ws)
        Else
            Call pasteValues(ws)
        End If
    Next ws
    
    Worksheets(1).Activate

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic '自動計算に戻す
    End With

End Sub

Sub deleteSheet(ByVal ws As Worksheet)

    Application.DisplayAlerts = False

    With ws
        .Visible = xlSheetVisible 'VeryHidden対策
        .Delete
    End With

    Application.DisplayAlerts = True
    
End Sub

Sub pasteValues(ByVal ws As Worksheet)

    With ws
        .Visible = xlSheetVisible
        .Cells.Copy
        .Cells.PasteSpecial Paste:=xlPasteValues
    End With
    
    Application.CutCopyMode = False

    Call setCellA1(ws)

End Sub

Sub setCellA1(ByVal ws As Worksheet)

    ws.Activate

    Range("A1").Select
    With ActiveWindow
        .ScrollRow = 1
        .ScrollColumn = 1
        .Zoom = 100
    End With

End Sub
