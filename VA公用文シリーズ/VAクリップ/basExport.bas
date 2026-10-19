Option Explicit

'--------------------------------------------------------------------------------
' 概要 : 全モジュールをJSON形式に変換してクリップボードへ送ります
' 引数 : TargetProj - エクスポート対象のVBProjectオブジェクト
'--------------------------------------------------------------------------------
Public Sub ExportModulesToClipboardJSON(ByVal TargetProj As Object)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim jsonParts As New Collection
    Dim jsonStr As String
    Dim codeContent As String
    Dim compTypeStr As String
    Dim jsonObj As String
    Dim exportCount As Long
    
    On Error GoTo ErrorHandler
    
    ' 1. プロジェクト確定
    If TargetProj Is Nothing Then
        MsgBox "対象プロジェクトが指定されていません。", vbExclamation
        Exit Sub
    End If
    Set vbProj = TargetProj
    
    exportCount = 0
    
    ' 2. 各コンポーネントをループ
    For Each vbComp In vbProj.VBComponents
        ' コードモジュールの内容取得
        If vbComp.CodeModule.CountOfLines > 0 Then
            codeContent = vbComp.CodeModule.lines(1, vbComp.CodeModule.CountOfLines)
        Else
            codeContent = ""
        End If
        
        ' タイプの文字列表現
        Select Case vbComp.Type
            Case 1: compTypeStr = "Standard"
            Case 2: compTypeStr = "Class"
            Case 3: compTypeStr = "Form"
            Case 100: compTypeStr = "Document" ' WordではThisDocument
            Case Else: compTypeStr = "Unknown"
        End Select
        
        ' JSONオブジェクト文字列の構築
        jsonObj = "{" & vbCrLf & _
                  "  ""name"": """ & vbComp.Name & """," & vbCrLf & _
                  "  ""type"": """ & compTypeStr & """," & vbCrLf & _
                  "  ""code"": """ & basJson.EscapeString(codeContent) & """" & vbCrLf & _
                  "}"
        
        jsonParts.Add jsonObj
        exportCount = exportCount + 1
    Next vbComp
    
    ' 3. JSON配列として結合
    jsonStr = "[" & vbCrLf
    Dim i As Long
    For i = 1 To jsonParts.count
        jsonStr = jsonStr & jsonParts(i)
        If i < jsonParts.count Then
            jsonStr = jsonStr & "," & vbCrLf
        Else
            jsonStr = jsonStr & vbCrLf
        End If
    Next i
    jsonStr = jsonStr & "]"
    
    ' 4. クリップボードへコピー
    basUtility.SetClipboardText jsonStr
    
    MsgBox exportCount & " 個のモジュールをクリップボードにコピーしました。" & vbCrLf & _
           "生成AIに貼り付けて使用してください。", _
           vbInformation, "エクスポート完了"
    Exit Sub

ErrorHandler:
    MsgBox "エクスポート中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub
