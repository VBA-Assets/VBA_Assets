Option Explicit

'================================================================================
' ファイル入出力 (UTF-8 BOMなし / ADODB.Stream)
'================================================================================

Public Sub WriteTextFileUTF8(ByVal filePath As String, ByVal content As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2    ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText content

    ' BOM（先頭3バイト）を除去してバイナリ保存
    Dim stmBin As Object
    Set stmBin = CreateObject("ADODB.Stream")
    stmBin.Type = 1 ' adTypeBinary
    stmBin.Open
    stm.Position = 3
    stm.CopyTo stmBin
    stmBin.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    stmBin.Close
    stm.Close
End Sub

Public Function ReadTextFileUTF8(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2    ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileUTF8 = stm.ReadText
    stm.Close
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir(filePath)) > 0)
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    FolderExists = (Len(Dir(folderPath, vbDirectory)) > 0)
End Function

'================================================================================
' フォルダ選択ダイアログ
'================================================================================

Public Function SelectFolder(ByVal title As String) As String
    Dim dialog As Object
    Set dialog = Application.FileDialog(4) ' msoFileDialogFolderPicker
    dialog.Title = title
    dialog.AllowMultiSelect = False
    If dialog.Show = -1 Then
        SelectFolder = dialog.SelectedItems(1)
    Else
        SelectFolder = ""
    End If
End Function

'================================================================================
' プロジェクトごとのフォルダパス設定
'================================================================================

Public Function GetProjectFolderPath(ByVal projFileName As String) As String
    GetProjectFolderPath = GetAppSetting("FolderPath_" & projFileName)
End Function

Public Sub SaveProjectFolderPath(ByVal projFileName As String, ByVal folderPath As String)
    SaveAppSetting "FolderPath_" & projFileName, folderPath
End Sub

'================================================================================
' 設定保存・読み込み (ThisWorkbook.CustomDocumentProperties)
' アドインはユーザーが手動保存しないため、設定後に明示的に保存する
'================================================================================

Public Sub SaveAppSetting(ByVal key As String, ByVal Value As String)
    Dim props As Object
    Set props = ThisWorkbook.CustomDocumentProperties

    On Error Resume Next
    props.item(key).Value = Value
    If Err.Number <> 0 Then
        Err.Clear
        props.Add Name:=key, LinkToContent:=False, Type:=4, Value:=Value
    End If
    On Error GoTo 0

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
End Sub

Public Function GetAppSetting(ByVal key As String) As String
    Dim props As Object
    Set props = ThisWorkbook.CustomDocumentProperties
    On Error Resume Next
    GetAppSetting = props.item(key).Value
    On Error GoTo 0
End Function

'================================================================================
' ヘルパー関数
'================================================================================

Public Function GetProjectFileName(vbProj As Object) As String
    ' vbProj.fileName はOneDriveのAutoSave時に一時ファイルパスを返すため
    ' Application.Workbooks から対応するブックを特定して名前を取得する
    Dim wb As Workbook
    On Error Resume Next
    For Each wb In Application.Workbooks
        If wb.VBProject Is vbProj Then
            GetProjectFileName = wb.Name
            Exit Function
        End If
    Next wb
    On Error GoTo 0
End Function

Public Function GetFileNameFromPath(ByVal path As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    Dim p As Long
    p = InStrRev(path, sep)
    If p > 0 Then
        GetFileNameFromPath = Mid(path, p + 1)
    Else
        GetFileNameFromPath = path
    End If
End Function
