Option Explicit

Sub TestFraction()
    Dim f1 As New Fraction
    Dim f2 As New Fraction
    Dim ans As New Fraction
    
    ' 2/6 で初期化（自動で 1/3 に約分される）
    f1.Init 2, 6
    
    ' 1/2 で初期化
    f2.Init 1, 2
    
    ' 足し算を実行 (1/3 + 1/2 = 5/6)
    Set ans = f1.Add(f2)
    
    ' 結果をイミディエイトウィンドウに出力
    Debug.Print f1.ToString() & " + " & f2.ToString() & " = " & ans.ToString()
End Sub
