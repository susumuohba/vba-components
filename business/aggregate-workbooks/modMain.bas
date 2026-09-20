Attribute VB_Name = "modMain"
Option Explicit

' 概要: 「集計実行」ボタンから呼ばれるエントリーポイント
' 引数: なし
' 戻り値: なし
Public Sub RunAggregate()

    Dim folderPath As String
    Dim count As Long

    folderPath = modFileUtil.PickFolder()
    If folderPath = "" Then
        Exit Sub ' キャンセル時は何もしない
    End If

    count = modAggregate.AggregateWorkbooks(folderPath)

    MsgBox count & " 件のファイルを集計しました。" & vbCrLf & _
           "エラーが発生したファイルがある場合は「エラーログ」シートを確認してください。", _
           vbInformation, "集計完了"

End Sub
