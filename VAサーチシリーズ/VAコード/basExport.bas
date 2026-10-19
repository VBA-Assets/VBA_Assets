Option Explicit

'--------------------------------------------------------------------------------
' 概要 : 全モジュールをファイルにエクスポートします
' 引数 : TargetProj - エクスポート対象のVBProjectオブジェクト
'         FolderPath - 出力先フォルダパス
'--------------------------------------------------------------------------------
Public Sub ExportModulesToFiles(ByVal TargetProj As Object, ByVal FolderPath As String)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeContent As String
    Dim compTypeStr As String
    Dim fileExt As String
    Dim filePath As String
    Dim exportCount As Long
    Dim manifestModules As String
    Dim firstModule As Boolean

    On Error GoTo ErrorHandler

    If TargetProj Is Nothing Then
        MsgBox "対象プロジェクトが指定されていません。", vbExclamation
        Exit Sub
    End If
    Set vbProj = TargetProj

    ' フォルダ末尾にセパレータを付加
    If Right(FolderPath, 1) <> Application.PathSeparator Then
        FolderPath = FolderPath & Application.PathSeparator
    End If

    ' フォルダが存在しない場合は作成
    If Not basUtility.FolderExists(FolderPath) Then
        MkDir FolderPath
    End If

    exportCount = 0
    firstModule = True
    manifestModules = ""

    For Each vbComp In vbProj.VBComponents
        ' コードの取得
        If vbComp.CodeModule.CountOfLines > 0 Then
            codeContent = vbComp.CodeModule.lines(1, vbComp.CodeModule.CountOfLines)
        Else
            codeContent = ""
        End If

        ' タイプと拡張子の決定
        Select Case vbComp.Type
            Case 1:   compTypeStr = "Standard":  fileExt = ".bas"
            Case 2:   compTypeStr = "Class":     fileExt = ".cls"
            Case 3:   compTypeStr = "Form":      fileExt = ".frm"
            Case 100: compTypeStr = "Document":  fileExt = ".bas"
            Case Else: compTypeStr = "Unknown":  fileExt = ".bas"
        End Select

        ' ファイルに書き込み
        filePath = FolderPath & vbComp.Name & fileExt
        basUtility.WriteTextFileUTF8 filePath, codeContent

        ' manifest エントリ追加
        If Not firstModule Then
            manifestModules = manifestModules & "," & vbCrLf
        End If
        manifestModules = manifestModules & _
            "    {""name"": """ & vbComp.Name & """, " & _
            """type"": """ & compTypeStr & """, " & _
            """file"": """ & vbComp.Name & fileExt & """}"
        firstModule = False
        exportCount = exportCount + 1
    Next vbComp

    ' manifest.json の書き込み
    Dim manifestStr As String
    manifestStr = "{" & vbCrLf & _
                  "  ""project"": """ & vbProj.Name & """," & vbCrLf & _
                  "  ""modules"": [" & vbCrLf & _
                  manifestModules & vbCrLf & _
                  "  ]" & vbCrLf & _
                  "}"
    basUtility.WriteTextFileUTF8 FolderPath & "manifest.json", manifestStr

    MsgBox exportCount & " 個のモジュールをエクスポートしました。" & vbCrLf & _
           "フォルダ: " & FolderPath, vbInformation, "エクスポート完了"
    Exit Sub

ErrorHandler:
    MsgBox "エクスポート中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub
