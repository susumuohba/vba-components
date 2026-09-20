Attribute VB_Name = "modAggregateTests"
Option Explicit

' ============================================================
' modAggregate.bas の単体テスト
'
' 実行方法:
'   1. このモジュールとmodAssert.bas、modAggregate.bas、modFileUtil.bas を
'      同一ブックにインポートする
'   2. カーソルを RunAllTests() 内に置き、F5キーで実行する
'   3. イミディエイトウィンドウ（表示 > イミディエイトウィンドウ、または Ctrl+G）
'      で結果を確認する
'
' 改修時は必ずこのテストを実行し、既存の挙動が壊れていないことを
' 確認してから納品・引き渡しを行うこと。
' テストケース一覧と実施結果は、5点セットのドキュメント側
' （4.6 テストケース一覧）にも記録すること。
' ============================================================

Public Sub RunAllTests()
    modAssert.ResetCounters
    Debug.Print "=== modAggregate 単体テスト開始 ==="

    Test_IsSelfFile_同一パスならTrue
    Test_IsSelfFile_大文字小文字を無視して比較する
    Test_IsSelfFile_異なるパスならFalse
    Test_GetLastDataRow_データが複数行ある場合
    Test_GetLastDataRow_ヘッダーのみの場合
    Test_GetLastDataRow_途中に空行があっても一番下まで取得できる
    Test_IsBlankRow_全列空欄ならTrue
    Test_IsBlankRow_一部でもデータがあればFalse

    modAssert.PrintSummary
End Sub

Private Sub Test_IsSelfFile_同一パスならTrue()
    Dim result As Boolean
    result = modAggregate.IsSelfFile("C:\work\集計.xlsm", "C:\work\集計.xlsm")
    modAssert.AssertTrue "IsSelfFile: 完全一致のパスはTrueになる", result
End Sub

Private Sub Test_IsSelfFile_大文字小文字を無視して比較する()
    Dim result As Boolean
    result = modAggregate.IsSelfFile("C:\WORK\集計.xlsm", "c:\work\集計.xlsm")
    modAssert.AssertTrue "IsSelfFile: 大文字小文字が異なってもTrueになる", result
End Sub

Private Sub Test_IsSelfFile_異なるパスならFalse()
    Dim result As Boolean
    result = modAggregate.IsSelfFile("C:\work\4月実績_東京.xlsx", "C:\work\集計.xlsm")
    modAssert.AssertFalse "IsSelfFile: 異なるパスはFalseになる", result
End Sub

' GetLastDataRow はワークシートを引数に取るため、テスト実行時に
' 一時シートを作成してテストデータを流し込み、後片付けまで行う
Private Sub Test_GetLastDataRow_データが複数行ある場合()
    Dim sheet As Worksheet
    Set sheet = CreateTempSheet()

    sheet.Range("A1").Value = "日付"
    sheet.Range("A2").Value = "2026/09/01"
    sheet.Range("A3").Value = "2026/09/02"
    sheet.Range("A4").Value = "2026/09/03"

    modAssert.AssertEqual "GetLastDataRow: 3行のデータがあれば4を返す", 4, modAggregate.GetLastDataRow(sheet, 1)

    DeleteTempSheet sheet
End Sub

Private Sub Test_GetLastDataRow_ヘッダーのみの場合()
    Dim sheet As Worksheet
    Set sheet = CreateTempSheet()

    sheet.Range("A1").Value = "日付"

    modAssert.AssertEqual "GetLastDataRow: ヘッダーのみなら1を返す", 1, modAggregate.GetLastDataRow(sheet, 1)

    DeleteTempSheet sheet
End Sub

' 1つのシートに、空行を挟んで2つの表がある状態を再現し、
' 一番下の表の末尾まで正しく最終行を検出できることを確認する
Private Sub Test_GetLastDataRow_途中に空行があっても一番下まで取得できる()
    Dim sheet As Worksheet
    Set sheet = CreateTempSheet()

    sheet.Range("A1").Value = "日付" ' 1つ目の表のヘッダー
    sheet.Range("A2").Value = "2026/09/01"
    sheet.Range("A3").Value = "2026/09/02"
    ' A4は空行（表と表の区切り）
    sheet.Range("A5").Value = "日付" ' 2つ目の表のヘッダー
    sheet.Range("A6").Value = "2026/09/10"
    sheet.Range("A7").Value = "2026/09/11"

    modAssert.AssertEqual "GetLastDataRow: 途中に空行があっても最下段の表の末尾（7行目）を返す", 7, modAggregate.GetLastDataRow(sheet, 1)

    DeleteTempSheet sheet
End Sub

Private Sub Test_IsBlankRow_全列空欄ならTrue()
    Dim sheet As Worksheet
    Set sheet = CreateTempSheet()

    ' A4〜D4はすべて空欄のまま

    modAssert.AssertTrue "IsBlankRow: A〜D列が全て空欄ならTrue", modAggregate.IsBlankRow(sheet, 4)

    DeleteTempSheet sheet
End Sub

Private Sub Test_IsBlankRow_一部でもデータがあればFalse()
    Dim sheet As Worksheet
    Set sheet = CreateTempSheet()

    sheet.Range("B4").Value = "東京" ' 拠点名だけ入っている状態

    modAssert.AssertFalse "IsBlankRow: 1列でもデータがあればFalse", modAggregate.IsBlankRow(sheet, 4)

    DeleteTempSheet sheet
End Sub

Private Function CreateTempSheet() As Worksheet
    Dim sheet As Worksheet
    Application.DisplayAlerts = False
    Set sheet = ThisWorkbook.Worksheets.Add
    sheet.Name = "__TestTemp_" & Format(Now, "hhmmss")
    Application.DisplayAlerts = True
    Set CreateTempSheet = sheet
End Function

Private Sub DeleteTempSheet(ByVal sheet As Worksheet)
    Application.DisplayAlerts = False
    sheet.Delete
    Application.DisplayAlerts = True
End Sub
