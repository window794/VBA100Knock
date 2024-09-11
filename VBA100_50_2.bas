Sub VBA100_50_1()
On Error GoTo 終了
    
    Application.ScreenUpdating = False

    Dim n1 As LongLong: n1 = 0
    Dim n2 As LongLong: n2 = 0
    Dim n3 As LongLong: n3 = 1
    Dim nTmp As LongLong
    Dim r As Long: r = 4
    
    Cells(1, 1).Value = "'" & n1
    Cells(2, 1).Value = "'" & n2
    Cells(3, 1).Value = "'" & n3
    
    Do
        nTmp = n1 + n2 + n3
        Cells(r, 1).Value = "'" & nTmp
        r = r + 1
        
        n1 = n2
        n2 = n3
        n3 = nTmp
    Loop
    
終了:
    Application.ScreenUpdating = True
    
End Sub
