Option Explicit

'--------------------------------------------------------------------------------
' 概要 : フォルダ内のファイルからモジュールをインポートします
' 引数 : TargetProj - インポート対象のVBProjectオブジェクト
'         FolderPath - 入力元フォルダパス（manifest.json が存在すること）
'--------------------------------------------------------------------------------
Public Sub ImportModulesFromFiles(ByVal TargetProj As Object, ByVal FolderPath As String)
    Dim vbProj As Object
    Dim manifestPath As String
    Dim manifestStr As String
    Dim modules As Collection
    Dim i As Long
    Dim successCount As Long

    On Error GoTo ErrorHandler

    If TargetProj Is Nothing Then
        MsgBox "対象プロジェクトが指定されていません。", vbExclamation
        Exit Sub
    End If

    ' 自プロジェクトへのインポートは禁止（実行中のコードが破壊される）
    If TargetProj Is ThisWorkbook.VBProject Then
        MsgBox "VAコード自身へのインポートはできません。" & vbCrLf & _
               "別のプロジェクトを選択してください。", vbExclamation
        Exit Sub
    End If

    Set vbProj = TargetProj

    ' フォルダ末尾にセパレータを付加
    If Right(FolderPath, 1) <> Application.PathSeparator Then
        FolderPath = FolderPath & Application.PathSeparator
    End If

    ' manifest.json の読み込み
    manifestPath = FolderPath & "manifest.json"
    If Not basUtility.FileExists(manifestPath) Then
        MsgBox "manifest.json が見つかりません。" & vbCrLf & _
               "フォルダ: " & FolderPath, vbExclamation
        Exit Sub
    End If

    manifestStr = basUtility.ReadTextFileUTF8(manifestPath)
    Set modules = basJson.ParseManifestModules(manifestStr)

    If modules.Count = 0 Then
        MsgBox "manifest.json にモジュール定義が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' インポート処理
    successCount = 0
    For i = 1 To modules.Count
        Dim modInfo As Object
        Set modInfo = modules(i)

        Dim modName As String, modType As String, modFile As String
        On Error Resume Next
        modName = modInfo("name")
        modType = modInfo("type")
        modFile = modInfo("file")
        On Error GoTo 0

        Dim codeFilePath As String
        codeFilePath = FolderPath & modFile

        If Not basUtility.FileExists(codeFilePath) Then
            Debug.Print "ファイルが見つかりません: " & codeFilePath
        Else
            Dim codeContent As String
            codeContent = basUtility.ReadTextFileUTF8(codeFilePath)
            If ProcessSingleModule(vbProj, modName, modType, codeContent) Then
                successCount = successCount + 1
            End If
        End If
    Next i

    MsgBox "処理完了" & vbCrLf & _
           "インポート成功: " & successCount & " / " & modules.Count & vbCrLf & vbCrLf & _
           "Ctrl+Sを押して、ファイルを保存してください。", vbInformation, "処理完了"
    Exit Sub

ErrorHandler:
    MsgBox "インポート中にエラーが発生しました。" & vbCrLf & _
           "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

'--------------------------------------------------------------------------------
' 内部関数: 単一モジュールの処理
'--------------------------------------------------------------------------------
Private Function ProcessSingleModule(vbProj As Object, modName As String, _
                                     modType As String, modCode As String) As Boolean
    Dim targetComp As Object

    If modName = "" Then Exit Function

    Set targetComp = Nothing
    On Error Resume Next
    Set targetComp = vbProj.VBComponents(modName)
    On Error GoTo 0

    ' Document型（シート / ThisWorkbook）
    If IsDocumentModule(targetComp, modType) Then
        If targetComp Is Nothing Then
            Debug.Print "Skip: Document " & modName & " not found."
            ProcessSingleModule = False
            Exit Function
        End If
        With targetComp.CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromString modCode
        End With
        ProcessSingleModule = True
        Exit Function
    End If

    ' 標準 / クラス / フォーム
    If Not targetComp Is Nothing Then
        If targetComp.Type = 3 Then
            ' フォームはコードのみ更新（レイアウト維持）
            With targetComp.CodeModule
                If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                .AddFromString modCode
            End With
        Else
            ' 標準・クラスは削除して再作成
            On Error Resume Next
            vbProj.VBComponents.Remove targetComp
            On Error GoTo 0
            Set targetComp = vbProj.VBComponents.Add(GetVBComponentType(modType))
            targetComp.Name = modName
            With targetComp.CodeModule
                If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                .AddFromString modCode
            End With
        End If
    Else
        ' 新規作成
        Set targetComp = vbProj.VBComponents.Add(GetVBComponentType(modType))
        targetComp.Name = modName
        With targetComp.CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromString modCode
        End With
    End If

    ProcessSingleModule = True
End Function

Private Function IsDocumentModule(comp As Object, typeStr As String) As Boolean
    If Not comp Is Nothing Then
        If comp.Type = 100 Then
            IsDocumentModule = True
            Exit Function
        End If
    End If
    Select Case LCase(typeStr)
        Case "document", "sheet", "workbook"
            IsDocumentModule = True
        Case Else
            IsDocumentModule = False
    End Select
End Function

Private Function GetVBComponentType(typeStr As String) As Integer
    Select Case LCase(typeStr)
        Case "standard":  GetVBComponentType = 1
        Case "class":     GetVBComponentType = 2
        Case "form":      GetVBComponentType = 3
        Case Else:        GetVBComponentType = 1
    End Select
End Function
