Sub VBA100_22()
'https://excel-ubara.com/vba100/VBA100_022.html

    Dim i As Long, ret As String
    Cells.Clear

    With Range("A1")
        .Value = 1
        .AutoFill Destination:=Range("A1:A100"), Type:=xlFillSeries
    End With

    For i = 1 To 100
        ret = Switch(i Mod 15 = 0, "FizzBuzz", i Mod 5 = 0 , "Buzz", i Mod 3 = 0 , "Fizz", _
                    i Mod 15 <> 0, i Mod 5 <> 0, i Mod 3 <> 0, "")
        If ret = "" Then GoTo Continue '該当なしの場合は列番号取得しない

        col = Switch(ret = "FizzBuzz", 4, ret = "Buzz", 3, ret = "Fizz", 2)
        Cells(i, col).Value = ret
Continue:
    Next i

End Sub

Sub VBA100_22_2()

    Dim i As Long
    Cells.Clear

    With Range("A1")
        .Value = 1
        .AutoFill Destination:= Range("A1:A100"), Type:=xlFillSeries
    End With

    For i = 1 To 100
        Select Case True
            Case i Mod 15 = 0
                Cells(i, 4).Value = "FizzBuzz"

            Case i Mod 5 = 0
                Cells(i, 3).Value = "Buzz"
            
            Case i Mod 3 = 0
                Cells(i, 2).Value = "Fizz"
        End Select
    Next i

End Sub
