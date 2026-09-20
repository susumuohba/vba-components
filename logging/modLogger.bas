Attribute VB_Name = "modLogger"
Option Explicit

' ============================================================
' 軽量ロガー
' ・レベル（DEBUG/INFO/WARN/ERROR）としきい値による出力制御
' ・出力先はファイル固定・追記モード
' ・フォーマット：日時＋レベル＋呼び出し元モジュール名＋メッセージ
'
' 【重要な制約】
' VBAには呼び出し元を自動的に辿る仕組みがないため、
' 「呼び出し元モジュール名」は呼び出し側が引数として明示的に渡す。
' ============================================================

Public Enum LogLevel
    LogDebug = 0
    LogInfo = 1
    LogWarn = 2
    LogError = 3
End Enum

' しきい値。既定値は LogDebug（0）＝すべて出力。
' 呼び出し側で「Logger.Threshold = LogLevel.LogWarn」のように上書きして使う。
Public Threshold As LogLevel

Private Const LOG_FILE_NAME As String = "log.txt"

' 概要: ログを1行出力する（しきい値未満のレベルは出力しない）
' 引数: moduleName - 呼び出し元のモジュール名（呼び出し側が明示的に渡す）
'       level      - ログレベル
'       message    - 出力するメッセージ
'       filePath   - 省略時はブックと同じフォルダの log.txt。テスト等で差し替える場合のみ指定
' 戻り値: なし
Public Sub Write(ByVal moduleName As String, ByVal level As LogLevel, ByVal message As String, Optional ByVal filePath As String = "")

    If level < Threshold Then Exit Sub

    If filePath = "" Then filePath = GetLogFilePath()

    Dim line As String
    line = BuildLogLine(Now, moduleName, level, message)

    WriteLineToFile filePath, line

End Sub

' 概要: ログ1行分の文字列を組み立てる（純粋関数・テスト容易）
' 引数: timestamp - 出力日時／moduleName - モジュール名／level - レベル／message - メッセージ
' 戻り値: "yyyy/mm/dd hh:nn:ss [LEVEL] moduleName: message" 形式の文字列
Public Function BuildLogLine(ByVal timestamp As Date, ByVal moduleName As String, ByVal level As LogLevel, ByVal message As String) As String
    BuildLogLine = Format(timestamp, "yyyy/mm/dd hh:nn:ss") & " [" & LevelToText(level) & "] " & moduleName & ": " & message
End Function

' 概要: LogLevelを文字列表記に変換する（純粋関数・テスト容易）
' 引数: level - ログレベル
' 戻り値: "DEBUG" / "INFO" / "WARN" / "ERROR"
Public Function LevelToText(ByVal level As LogLevel) As String
    Select Case level
        Case LogDebug: LevelToText = "DEBUG"
        Case LogInfo: LevelToText = "INFO"
        Case LogWarn: LevelToText = "WARN"
        Case LogError: LevelToText = "ERROR"
        Case Else: LevelToText = "UNKNOWN"
    End Select
End Function

' 概要: 既定のログファイルパスを返す（ブックと同じフォルダ、固定ファイル名）
' 引数: なし
' 戻り値: ログファイルの絶対パス
Public Function GetLogFilePath() As String
    GetLogFilePath = ThisWorkbook.Path & "\" & LOG_FILE_NAME
End Function

' 概要: 指定ファイルへ1行追記する（副作用のある処理。ロジックはBuildLogLine側でテストする）
' 引数: filePath - 出力先ファイルパス／line - 出力する1行
' 戻り値: なし
Private Sub WriteLineToFile(ByVal filePath As String, ByVal line As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Append As #fileNum
    Print #fileNum, line
    Close #fileNum
End Sub
