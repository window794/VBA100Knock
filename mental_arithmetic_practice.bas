Option Explicit
Option Base 1

Sub VBA100_41()
Rem https://excel-ubara.com/vba100/VBA100_041.html

    Dim i As Long
    Dim score As Long
    Dim strFormula As String
    Dim seigo(1 To 10) As Variant
    Dim answer As String
    Dim startTime As Double, endTime As Double, middleTime As Double, processTime As Double
    Dim msg As String
    
    For i = 1 To 10
        strFormula = generateFormula '数式生成
        
        startTime = Timer '開始時間取得
        
        answer = InputBox(i & "/10問目" & vbCrLf & strFormula, "暗算練習") '回答入力
        
        endTime = Timer '終了時間取得
        middleTime = endTime - startTime
        processTime = processTime + middleTime
        
        If StrPtr(answer) = 0 Then '入力がキャンセルされたら
            seigo(i) = i & ": " & Evaluate(strFormula) & ": 未回答" '未回答
            GoTo Continue
        End If
        
        If IsNumeric(answer) Then
            If answer = Evaluate(strFormula) Then
                seigo(i) = i & ": " & Evaluate(strFormula) & ": 〇" '正解
                score = score + 1
            Else
                seigo(i) = i & ": " & Evaluate(strFormula) & ": ×" '不正解
            End If
        End If
Continue:
    Next i
    
    msg = generateMsg(seigo, score, processTime)
    MsgBox msg

End Sub

Function generateFormula() As String
    Dim i As Long
    Dim num1 As Long, num2 As Long
    Dim symbols As Variant: symbols = Array("+", "-", "*", "/")
    Dim strFormula(1 To 3) As Variant
    
    Randomize
    i = Int(4 * Rnd + 1)
    
    Select Case symbols(i)
        Case "+"
            Randomize
            num1 = Int(101 * Rnd)
            num2 = Int(101 * Rnd)
            
            strFormula(1) = num1
            strFormula(2) = "+"
            strFormula(3) = num2
        
        Case "-"
            Do
                Randomize
                num1 = Int(101 * Rnd)
                num2 = Int(101 * Rnd)
            Loop Until num1 > num2 'num1がnum2より大きくなるまで乱数生成
            
            strFormula(1) = num1
            strFormula(2) = "-"
            strFormula(3) = num2
        
        Case "*"
            Randomize
            '計算のしやすさを考慮して、100までの数と10までの数を生成
            num1 = Int(101 * Rnd)
            num2 = Int(11 * Rnd)
            
            strFormula(1) = num1
            strFormula(2) = "*"
            strFormula(3) = num2
        
        Case "/"
            '計算のしやすさを考慮して、割られる数を50まで、割る数を10までとし
            '計算結果に小数点を含まないもののみ（答えが割り切れて、かつ整数）式を生成する
            '0除算防止のため、1～の乱数を生成する
            Do
                Randomize
                num1 = Int(50 * Rnd + 1)
                num2 = Int(10 * Rnd + 1)
            Loop Until InStr(num1 / num2, ".") = 0
            
            strFormula(1) = num1
            strFormula(2) = "/"
            strFormula(3) = num2
            
    End Select
    
    generateFormula = Join(strFormula, " ")
            
End Function

Function generateMsg(ByRef seigo As Variant, score As Long, processTime As Double) As String
    Dim msg As String

    Select Case score
        Case 0: msg = "やる気あるんかワレ？"
        Case 1 To 3: msg = "もっとがんばりましょう"
        Case 4 To 5: msg = "もうすこしがんばりましょう"
        Case 6 To 7: msg = "がんばりました！"
        Case 8 To 9: msg = "よくできました！"
        Case 10: msg = "花丸満点！！"
    End Select
    
    generateMsg = Join(seigo, vbCrLf) & vbCrLf & vbCrLf & _
                            "時間: " & WorksheetFunction.Round(processTime, 2) & "秒" & vbCrLf & _
                            "得点: " & score & vbCrLf & _
                            "ひと言: " & msg
End Function
