Attribute VB_Name = "modLoggerTests"
Option Explicit

' ============================================================
' modLogger.bas の単体テスト
'
' 実行方法:
'   1. testing/modAssert.bas と、このモジュール、modLogger.bas を
'      同一ブックにインポートする
'   2. カーソルを RunAllTests() 内に置き、F5キーで実行する
'   3. イミディエイトウィンドウ（Ctrl+G）で結果を確認する
' ============================================================

Public Sub RunAllTests()
    modAssert.ResetCounters
    Debug.Print "=== modLogger 単体テスト開始 ==="

    Test_LevelToText_各レベルを文字列に変換できる
    Test_BuildLogLine_フォーマットが正しい
    Test_Write_しきい値未満のログは出力されない
    Test_Write_しきい値以上のログは出力される
    Test_Write_追記モードで複数回書き込める

    modAssert.PrintSummary
End Sub

Private Sub Test_LevelToText_各レベルを文字列に変換できる()
    modAssert.AssertEqual "LevelToText: LogDebug", "DEBUG", modLogger.LevelToText(LogDebug)
    modAssert.AssertEqual "LevelToText: LogInfo", "INFO", modLogger.LevelToText(LogInfo)
    modAssert.AssertEqual "LevelToText: LogWarn", "WARN", modLogger.LevelToText(LogWarn)
    modAssert.AssertEqual "LevelToText: LogError", "ERROR", modLogger.LevelToText(LogError)
End Sub

Private Sub Test_BuildLogLine_フォーマットが正しい()
    Dim actual As String
    actual = modLogger.BuildLogLine(DateSerial(2026, 9, 20) + TimeSerial(13, 5, 9), "modAggregate", LogWarn, "対象ファイルが見つかりません")

    modAssert.AssertEqual "BuildLogLine: 日時+レベル+モジュール名+メッセージの形式になる", _
        "2026/09/20 13:05:09 [WARN] modAggregate: 対象ファイルが見つかりません", actual
End Sub

Private Sub Test_Write_しきい値未満のログは出力されない()
    Dim tempPath As String
    tempPath = GetTempLogPath()
    DeleteFileIfExists tempPath

    modLogger.Threshold = LogWarn
    modLogger.Write "modTest", LogInfo, "これはINFOなので出力されないはず", tempPath

    modAssert.AssertFalse "Write: しきい値未満（INFO<WARN）のログはファイルに出力されない", FileExistsWithContent(tempPath)

    DeleteFileIfExists tempPath
End Sub

Private Sub Test_Write_しきい値以上のログは出力される()
    Dim tempPath As String
    tempPath = GetTempLogPath()
    DeleteFileIfExists tempPath

    modLogger.Threshold = LogWarn
    modLogger.Write "modTest", LogError, "これはERRORなので出力されるはず", tempPath

    Dim content As String
    content = ReadFileAllText(tempPath)

    modAssert.AssertTrue "Write: しきい値以上（ERROR>=WARN）のログはファイルに出力される", _
        (InStr(content, "これはERRORなので出力されるはず") > 0)

    DeleteFileIfExists tempPath
End Sub

Private Sub Test_Write_追記モードで複数回書き込める()
    Dim tempPath As String
    tempPath = GetTempLogPath()
    DeleteFileIfExists tempPath

    modLogger.Threshold = LogDebug
    modLogger.Write "modTest", LogInfo, "1回目", tempPath
    modLogger.Write "modTest", LogInfo, "2回目", tempPath

    Dim content As String
    content = ReadFileAllText(tempPath)

    modAssert.AssertTrue "Write: 1回目の内容が残っている（上書きされていない）", (InStr(content, "1回目") > 0)
    modAssert.AssertTrue "Write: 2回目の内容も追記されている", (InStr(content, "2回目") > 0)

    DeleteFileIfExists tempPath
End Sub

' ---- テスト用ヘルパー ----

Private Function GetTempLogPath() As String
    GetTempLogPath = Environ$("TEMP") & "\modLoggerTests_temp.log"
End Function

Private Function FileExistsWithContent(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExistsWithContent = (Len(Dir(filePath)) > 0)
    On Error GoTo 0
End Function

Private Function ReadFileAllText(ByVal filePath As String) As String
    Dim fileNum As Integer
    Dim line As String
    Dim result As String

    If Len(Dir(filePath)) = 0 Then
        ReadFileAllText = ""
        Exit Function
    End If

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        result = result & line & vbCrLf
    Loop
    Close #fileNum

    ReadFileAllText = result
End Function

Private Sub DeleteFileIfExists(ByVal filePath As String)
    On Error Resume Next
    If Len(Dir(filePath)) > 0 Then Kill filePath
    On Error GoTo 0
End Sub
