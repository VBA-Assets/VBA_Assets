Option Explicit

'================================================================================
' クリップボード操作
'================================================================================

'--------------------------------------------------------------------------------
' 概要 : クリップボードにテキストを設定します
'--------------------------------------------------------------------------------
Public Sub SetClipboardText(ByVal Text As String)
    Dim DataObj As Object
    ' MSForms.DataObject (参照設定がなくても動くようにLate Binding)
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    On Error Resume Next
    DataObj.SetText Text
    DataObj.PutInClipboard
    
    If Err.Number <> 0 Then
        MsgBox "クリップボードへのコピーに失敗しました。" & vbCrLf & _
               "他のアプリがクリップボードを使用中の可能性があります。", vbExclamation
    End If
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' 概要 : クリップボードからテキストを取得します
'--------------------------------------------------------------------------------
Public Function GetClipboardText() As String
    Dim DataObj As Object
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    On Error Resume Next
    DataObj.GetFromClipboard
    GetClipboardText = DataObj.GetText
    
    If Err.Number <> 0 Then
        GetClipboardText = "" ' 取得失敗時は空文字
    End If
    On Error GoTo 0
End Function

'================================================================================
' 設定保存・読み込み (ThisDocument.CustomDocumentProperties)
'================================================================================

'--------------------------------------------------------------------------------
' 概要 : アドイン自身のユーザープロパティに設定を保存します
'--------------------------------------------------------------------------------
Public Sub SaveAppSetting(ByVal key As String, ByVal Value As String)
    Dim props As Object
    
    ' このコードが記述されている文書（アドイン自身）のプロパティを操作
    ' Wordの場合は ThisDocument
    Set props = ThisDocument.CustomDocumentProperties

    On Error Resume Next
    ' 既存プロパティがあれば値を更新
    props.item(key).Value = Value
    
    If Err.Number <> 0 Then
        ' 存在しなければ新規作成 (Type 4 = msoPropertyTypeString)
        Err.Clear
        props.Add Name:=key, LinkToContent:=False, Type:=4, Value:=Value
    End If
    On Error GoTo 0

    ' 設定を永続化するためにアドイン自身を上書き保存
    ' ※アドインが読み取り専用の場所に配置されている場合はエラーになるため無視する
    On Error Resume Next
    ThisDocument.Save
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' 概要 : アドイン自身のユーザープロパティから設定を読み込みます
'--------------------------------------------------------------------------------
Public Function GetAppSetting(ByVal key As String) As String
    Dim props As Object
    ' Wordの場合は ThisDocument
    Set props = ThisDocument.CustomDocumentProperties
    
    On Error Resume Next
    GetAppSetting = props.item(key).Value
    On Error GoTo 0
    ' エラー時（プロパティ未定義時）は空文字が返る
End Function

'================================================================================
' ヘルパー関数
'================================================================================

'--------------------------------------------------------------------------------
' 概要 : プロジェクトのファイル名を取得します（UI表示・設定保存用）
'--------------------------------------------------------------------------------
Public Function GetProjectFileName(vbProj As Object) As String
    Dim fName As String
    On Error Resume Next
    fName = vbProj.fileName
    On Error GoTo 0
    
    ' パスからファイル名を抽出
    If fName <> "" Then
        GetProjectFileName = GetFileNameFromPath(fName)
    Else
        GetProjectFileName = ""
    End If
End Function

'--------------------------------------------------------------------------------
' 概要 : パスからファイル名を取得 (文字列操作のみ)
'--------------------------------------------------------------------------------
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
