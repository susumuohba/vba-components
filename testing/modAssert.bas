Attribute VB_Name = "modAssert"
Option Explicit

' ============================================================
' 自作の軽量単体テストフレームワーク
' アドイン不要・このモジュールをコピーするだけで動作する
' 実行結果はイミディエイトウィンドウ（Ctrl+G）に出力される
' ============================================================

Private passCount As Long
Private failCount As Long

' 概要: テストの実行結果カウンタをリセットする。各テストスイートの先頭で呼ぶ
Public Sub ResetCounters()
    passCount = 0
    failCount = 0
End Sub

' 概要: 実測値と期待値が等しいことを検証する
' 引数: testName - テスト名／expected - 期待値／actual - 実測値
Public Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If expected = actual Then
        passCount = passCount + 1
        Debug.Print "  [OK]   " & testName
    Else
        failCount = failCount + 1
        Debug.Print "  [NG]   " & testName & " (期待値: " & expected & " / 実測値: " & actual & ")"
    End If
End Sub

' 概要: 条件がTrueであることを検証する
' 引数: testName - テスト名／condition - 検証する条件式
Public Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    AssertEqual testName, True, condition
End Sub

' 概要: 条件がFalseであることを検証する
' 引数: testName - テスト名／condition - 検証する条件式
Public Sub AssertFalse(ByVal testName As String, ByVal condition As Boolean)
    AssertEqual testName, False, condition
End Sub

' 概要: これまでの実行結果サマリをイミディエイトウィンドウに出力する
Public Sub PrintSummary()
    Debug.Print "----------------------------------------"
    Debug.Print "合計: " & (passCount + failCount) & "件 / 成功: " & passCount & "件 / 失敗: " & failCount & "件"
    If failCount = 0 Then
        Debug.Print "=> ALL GREEN"
    Else
        Debug.Print "=> " & failCount & "件の失敗があります。上記[NG]行を確認してください。"
    End If
End Sub
