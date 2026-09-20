Attribute VB_Name = "modFileUtil"
Option Explicit

' 概要: フォルダ選択ダイアログを表示し、選択されたフォルダの絶対パスを返す
' 引数: なし
' 戻り値: 選択されたフォルダパス（未選択時は空文字）
Public Function PickFolder() As String

    Dim folderDialog As FileDialog
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)

    folderDialog.Title = "集計対象フォルダを選択してください"

    If folderDialog.Show = -1 Then
        PickFolder = folderDialog.SelectedItems(1)
    Else
        PickFolder = ""
    End If

End Function

' 概要: 指定フォルダ直下（サブフォルダは含まない）の .xlsx / .xlsm ファイル一覧を取得する
' 引数: folderPath - 検索対象フォルダの絶対パス
' 戻り値: ファイルの絶対パスを格納したコレクション。該当なしの場合は空のコレクション
Public Function GetTargetFiles(ByVal folderPath As String) As Collection

    Dim files As New Collection
    Dim fileName As String
    Dim ext As String

    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        ext = LCase$(Right$(fileName, 5))
        If ext Like "*.xlsx" Or ext Like "*.xlsm" Then
            files.Add folderPath & fileName
        End If
        fileName = Dir
    Loop

    Set GetTargetFiles = files

End Function
