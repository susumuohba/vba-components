Option Explicit

Private pNumerator As Long
Private pDenominator As Long

Public Property Get Numerator() As Long
    Numerator = pNumerator
End Property

Public Property Get Denominator() As Long
    Denominator = pDenominator
End Property

Public Sub Init(ByVal num As Long, ByVal den As Long)
    If den = 0 Then Err.Raise 11, , "分母に0は指定できません"
    
    If den < 0 Then
        num = -num
        den = -den
    End If
    
    Dim gcd As Long
    gcd = GetGCD(Abs(num), Abs(den))
    
    pNumerator = num / gcd
    pDenominator = den / gcd
End Sub

Public Function Add(ByVal target As Fraction) As Fraction
    Dim result As New Fraction
    Dim newNum As Long
    Dim newDen As Long
    
    newNum = (pNumerator * target.Denominator) + (target.Numerator * pDenominator)
    newDen = pDenominator * target.Denominator
    
    result.Init newNum, newDen
    Set Add = result
End Function

Public Function Multiply(ByVal target As Fraction) As Fraction
    Dim result As New Fraction
    result.Init pNumerator * target.Numerator, pDenominator * target.Denominator
    Set Multiply = result
End Function

Public Function ToString() As String
    If pDenominator = 1 Then
        ToString = CStr(pNumerator)
    Else
        ToString = CStr(pNumerator) & "/" & CStr(pDenominator)
    End If
End Function

Private Function GetGCD(ByVal a As Long, ByVal b As Long) As Long
    Do While b <> 0
        Dim temp As Long
        temp = b
        b = a Mod b
        a = temp
    Loop
    GetGCD = a
End Function
