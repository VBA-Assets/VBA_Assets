Option Explicit

'--------------------------------------------------------------------------------
' 概要 : クリップボードのJSONデータからモジュールをインポートします
' 引数 : TargetProj - インポート対象のVBProjectオブジェクト
'--------------------------------------------------------------------------------
Public Sub ImportModulesFromClipboardJSON(ByVal TargetProj As Object)
    Dim jsonStr As String
    Dim vbProj As Object
    Dim modules As Collection
    Dim modInfo As Variant ' Collection or Dictionary
    Dim i As Long
    Dim successCount As Long
    
    On Error GoTo ErrorHandler
    
    ' 1. プロジェクト確定
    If TargetProj Is Nothing Then
        MsgBox "対象プロジェクトが指定されていません。", vbExclamation
        Exit Sub
    End If
    Set vbProj = TargetProj
    
    ' 2. クリップボード取得
    jsonStr = basUtility.GetClipboardText()
    If Len(jsonStr) = 0 Then
        MsgBox "クリップボードが空か、テキストデータの取得に失敗しました。", vbExclamation
        Exit Sub
    End If
    
    ' 3. JSONパース
    Set modules = basJson.ParseJSON(jsonStr)
    
    If modules.count = 0 Then
        MsgBox "有効なJSONデータが見つかりませんでした。" & vbCrLf & _
               "クリップボードの内容が正しい形式（生成AIからの出力）か確認してください。", vbExclamation
        Exit Sub
    End If
    
    ' 4. インポート処理
    successCount = 0
    For i = 1 To modules.count
        Set modInfo = modules(i)
        
        If ProcessSingleModule(vbProj, modInfo) Then
            successCount = successCount + 1
        End If
    Next i
    
    MsgBox "処理完了" & vbCrLf & _
           "インポート成功: " & successCount & " / " & modules.count, vbInformation
           
    Exit Sub

ErrorHandler:
    MsgBox "インポート中にエラーが発生しました。" & vbCrLf & _
           "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

'--------------------------------------------------------------------------------
' 内部関数: 単一モジュールの処理分岐
'--------------------------------------------------------------------------------
Private Function ProcessSingleModule(vbProj As Object, modInfo As Variant) As Boolean
    Dim modName As String
    Dim modType As String
    Dim modCode As String
    Dim targetComp As Object
    
    On Error Resume Next
    modName = modInfo("name")
    modType = modInfo("type")
    modCode = modInfo("code")
    On Error GoTo 0
    
    If modName = "" Then Exit Function
    
    ' 既存コンポーネントの検索
    Set targetComp = Nothing
    On Error Resume Next
    Set targetComp = vbProj.VBComponents(modName)
    On Error GoTo 0
    
    ' --- 分岐ロジック ---
    
    ' Case A: ThisDocument (Document型)
    If IsDocumentModule(targetComp, modType) Then
        If targetComp Is Nothing Then
            ' 存在しないDocumentへのインポートはスキップ
            Debug.Print "Skip: Document " & modName & " not found."
            ProcessSingleModule = False
            Exit Function
        End If
        
        ' コードの全置換
        With targetComp.CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromString modCode
        End With
        ProcessSingleModule = True
        Exit Function
    End If
    
    ' Case B: 標準モジュール / クラスモジュール / フォーム
    If Not targetComp Is Nothing Then
        ' 既存がある場合
        If targetComp.Type = 3 Then ' vbext_ct_MSForm
            ' フォームの場合はコードのみ更新（レイアウト維持）
            With targetComp.CodeModule
                If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                .AddFromString modCode
            End With
        Else
            ' 標準・クラスは削除して再作成（クリーンインストール）
            On Error Resume Next
            vbProj.VBComponents.Remove targetComp
            On Error GoTo 0
            
            Set targetComp = vbProj.VBComponents.Add(GetVBComponentType(modType))
            targetComp.Name = modName
            
            ' VBEが自動挿入したOption Explicit等を削除してから書き込む
            With targetComp.CodeModule
                If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                .AddFromString modCode
            End With
        End If
    Else
        ' 新規作成
        Set targetComp = vbProj.VBComponents.Add(GetVBComponentType(modType))
        targetComp.Name = modName
        
        ' VBEが自動挿入したOption Explicit等を削除してから書き込む
        With targetComp.CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromString modCode
        End With
    End If
    
    ProcessSingleModule = True
End Function

'--------------------------------------------------------------------------------
' ヘルパー: コンポーネントタイプの判定
'--------------------------------------------------------------------------------
Private Function IsDocumentModule(comp As Object, typeStr As String) As Boolean
    ' コンポーネント実体がある場合はTypeプロパティ優先 (100 = Document)
    If Not comp Is Nothing Then
        If comp.Type = 100 Then
            IsDocumentModule = True
            Exit Function
        End If
    End If
    
    ' JSON文字列判定 (Word用に調整)
    Select Case LCase(typeStr)
        ' Excel互換のためにsheet等も残していますが、Wordでは document / thisdocument が主になります
        Case "document", "thisdocument", "sheet", "workbook", "class interface"
            IsDocumentModule = True
        Case Else
            IsDocumentModule = False
    End Select
End Function

Private Function GetVBComponentType(typeStr As String) As Integer
    Select Case LCase(typeStr)
        Case "standard", "standardmodule", "module"
            GetVBComponentType = 1 ' vbext_ct_StdModule
        Case "class", "classmodule"
            GetVBComponentType = 2 ' vbext_ct_ClassModule
        Case "form", "userform", "msform"
            GetVBComponentType = 3 ' vbext_ct_MSForm
        Case Else
            GetVBComponentType = 1 ' Default to Standard
    End Select
End Function
