Attribute VB_Name = "modAggregate"
Option Explicit

Private Const DATA_SHEET_NAME As String = "Sheet1"
Private Const RESULT_SHEET_NAME As String = "集計結果"
Private Const ERROR_SHEET_NAME As String = "エラーログ"
Private Const HEADER_ROW As Long = 1

' 概要: 対象フォルダ内の全ブックを集計結果シートへ書き出す（メイン処理）
' 引数: folderPath - 対象フォルダの絶対パス
' 戻り値: 正常に集計できたファイル件数
Public Function AggregateWorkbooks(ByVal folderPath As String) As Long

    Dim aggSheet As Worksheet
    Dim errSheet As Worksheet
    Dim files As Collection
    Dim file As Variant
    Dim successCount As Long
    Dim outRow As Long

    Set aggSheet = PrepareResultSheet()
    Set errSheet = PrepareErrorSheet()

    Set files = modFileUtil.GetTargetFiles(folderPath)
    outRow = HEADER_ROW + 1
    successCount = 0

    For Each file In files

        If Not IsSelfFile(CStr(file), ThisWorkbook.FullName) Then

            On Error GoTo CopyError

            outRow = CopyOneWorkbook(CStr(file), aggSheet, outRow)
            successCount = successCount + 1

            On Error GoTo 0
        End If

NextFile:
    Next file

    AggregateWorkbooks = successCount
    Exit Function

CopyError:
    errSheet.Cells(errSheet.Cells(errSheet.Rows.Count, 1).End(xlUp).Row + 1, 1).Value = file
    errSheet.Cells(errSheet.Cells(errSheet.Rows.Count, 2).End(xlUp).Row + 1, 2).Value = Err.Description
    Err.Clear
    Resume NextFile

End Function

' 概要: 1ファイル分のデータシートを読み取り、集計シートへ追記する
' 引数: filePath - 読み込むブックの絶対パス／aggSheet - 書き込み先シート／startRow - 書き込み開始行
' 戻り値: 書き込み後の次回開始行
Private Function CopyOneWorkbook(ByVal filePath As String, ByVal aggSheet As Worksheet, ByVal startRow As Long) As Long

    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim outRow As Long

    Set sourceBook = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceSheet = sourceBook.Worksheets(DATA_SHEET_NAME)

    lastRow = GetLastDataRow(sourceSheet, 1)
    outRow = startRow

    For row = HEADER_ROW + 1 To lastRow
        If Not IsBlankRow(sourceSheet, row) Then
            aggSheet.Cells(outRow, 1).Value = sourceSheet.Cells(row, 1).Value ' 日付
            aggSheet.Cells(outRow, 2).Value = sourceSheet.Cells(row, 2).Value ' 拠点名
            aggSheet.Cells(outRow, 3).Value = sourceSheet.Cells(row, 3).Value ' 担当者名
            aggSheet.Cells(outRow, 4).Value = sourceSheet.Cells(row, 4).Value ' 金額
            aggSheet.Cells(outRow, 5).Value = sourceBook.Name                ' 取得元ファイル名
            outRow = outRow + 1
        End If
    Next row

    sourceBook.Close SaveChanges:=False

    CopyOneWorkbook = outRow

End Function

' 概要: A〜D列（日付・拠点名・担当者名・金額）が すべて空欄の行かどうかを判定する（純粋関数・テスト容易）
' 引数: sheet - 対象シート／row - 判定対象の行番号
' 戻り値: A〜D列がすべて空欄であればTrue
Public Function IsBlankRow(ByVal sheet As Worksheet, ByVal row As Long) As Boolean
    IsBlankRow = (sheet.Cells(row, 1).Value = "") And _
                 (sheet.Cells(row, 2).Value = "") And _
                 (sheet.Cells(row, 3).Value = "") And _
                 (sheet.Cells(row, 4).Value = "")
End Function

' 概要: 指定列を基準に、シートの最終データ行を求める（純粋関数・テスト容易）
' 引数: sheet - 対象シート／checkColumn - 基準列番号
' 戻り値: 最終データ行の行番号。データがヘッダーのみの場合はHEADER_ROWを返す
Public Function GetLastDataRow(ByVal sheet As Worksheet, Optional ByVal checkColumn As Long = 1) As Long
    Dim row As Long
    row = sheet.Cells(sheet.Rows.Count, checkColumn).End(xlUp).Row
    If row < HEADER_ROW Then row = HEADER_ROW
    GetLastDataRow = row
End Function

' 概要: 対象ファイルが集計ブック自身（このマクロを含むブック）かどうかを判定する（純粋関数・テスト容易）
' 引数: filePath - 判定対象ファイルの絶対パス／thisWorkbookFullName - 集計ブック自身のフルパス
' 戻り値: 自分自身であればTrue
Public Function IsSelfFile(ByVal filePath As String, ByVal thisWorkbookFullName As String) As Boolean
    IsSelfFile = (LCase$(filePath) = LCase$(thisWorkbookFullName))
End Function

Private Function PrepareResultSheet() As Worksheet
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets(RESULT_SHEET_NAME)
    On Error GoTo 0

    If sheet Is Nothing Then
        Set sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sheet.Name = RESULT_SHEET_NAME
    Else
        sheet.Cells.Clear
    End If

    sheet.Range("A1:E1").Value = Array("日付", "拠点名", "担当者名", "金額", "取得元ファイル名")
    Set PrepareResultSheet = sheet
End Function

Private Function PrepareErrorSheet() As Worksheet
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets(ERROR_SHEET_NAME)
    On Error GoTo 0

    If sheet Is Nothing Then
        Set sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sheet.Name = ERROR_SHEET_NAME
    Else
        sheet.Cells.Clear
    End If

    sheet.Range("A1:B1").Value = Array("ファイルパス", "エラー内容")
    Set PrepareErrorSheet = sheet
End Function
