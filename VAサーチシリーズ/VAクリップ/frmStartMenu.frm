Option Explicit

' プロジェクトオブジェクトを内部で管理するためのコレクション
Private m_Projects As Collection

' フォーム初期化
Private Sub UserForm_Initialize()
    ' コレクション初期化
    Set m_Projects = New Collection
    
    ' コンボボックスの設定 (※フォーム上に cboProjectList を配置してください)
    With cboProjectList
        .Style = fmStyleDropDownList
    End With
    
    ' プロジェクト一覧読み込み
    LoadProjects
End Sub

' プロジェクト一覧読み込みロジック
Private Sub LoadProjects()
    Dim vbProjs As Object
    Dim proj As Object
    Dim i As Long
    Dim defaultFileName As String
    Dim fName As String
    Dim dispName As String
    Dim validCount As Long
    
    Set vbProjs = Application.VBE.VBProjects
    
    ' 1. 前回保存されたプロジェクト名（ファイル名）を取得
    defaultFileName = basUtility.GetAppSetting("LastSelectedProject")
    
    ' 2. 保存値がない場合は、アクティブなブック名をデフォルトにする
    If defaultFileName = "" Then
        On Error Resume Next
        If Not ActiveWorkbook Is Nothing Then
            defaultFileName = ActiveWorkbook.Name
        End If
        On Error GoTo 0
    End If
    
    validCount = 0
    cboProjectList.Clear
    Set m_Projects = New Collection
    
    For i = 1 To vbProjs.Count
        Set proj = vbProjs(i)
        ' プロジェクトが保護されていないか確認 (0 = vbext_pp_none)
        If proj.Protection = 0 Then
            validCount = validCount + 1
            fName = basUtility.GetProjectFileName(proj)
            If fName = "" Then fName = "(未保存のブック)"
            
            ' 表示名: プロジェクト名 (ファイル名)
            dispName = proj.Name & "  (" & fName & ")"
            
            m_Projects.Add proj
            cboProjectList.AddItem dispName
        End If
    Next i
    
    If validCount = 0 Then
        cboProjectList.AddItem "操作可能なプロジェクトなし"
        cboProjectList.Enabled = False
    Else
        ' デフォルト選択を実行
        SetDefaultSelection defaultFileName
    End If
End Sub

' デフォルト選択の設定（ファイル名で部分一致検索）
Private Sub SetDefaultSelection(fileName As String)
    Dim i As Long
    Dim itemText As String
    
    If fileName = "" Then Exit Sub
    
    ' リストを走査してファイル名を含む項目を探す
    For i = 0 To cboProjectList.ListCount - 1
        itemText = cboProjectList.List(i)
        ' 表示形式が "ProjName (FileName)" なので、FileNameが含まれるかチェック
        If InStr(1, itemText, fileName, vbTextCompare) > 0 Then
            cboProjectList.ListIndex = i
            Exit Sub
        End If
    Next i
    
    ' 見つからなければ先頭を選択
    If cboProjectList.ListCount > 0 Then cboProjectList.ListIndex = 0
End Sub

' 選択されたプロジェクトを取得
Private Function GetSelectedProject() As Object
    If cboProjectList.ListIndex = -1 Then
        MsgBox "プロジェクトを選択してください。", vbExclamation
        Set GetSelectedProject = Nothing
        Exit Function
    End If
    
    ' コレクションは1ベース、ListIndexは0ベース
    Set GetSelectedProject = m_Projects(cboProjectList.ListIndex + 1)
End Function

'================================================================================
' エクスポート実行 (基本案・確認メッセージ)
'================================================================================
Private Sub CommandButton1_Click()
    Dim proj As Object
    Set proj = GetSelectedProject()
    
    If Not proj Is Nothing Then
        ' --- 確認メッセージ ---
        Dim msg As String
        msg = "プロジェクト名: " & proj.Name & vbCrLf & vbCrLf & _
              "このプロジェクトから全モジュールをクリップボードにエクスポートします。" & vbCrLf & _
              "よろしいですか？"
              
        If MsgBox(msg, vbOKCancel + vbQuestion, "エクスポート確認") <> vbOK Then
            Exit Sub
        End If
        ' --------------------
        
        ' ★設定保存: 選択したプロジェクトのファイル名を保存
        Dim fName As String
        fName = basUtility.GetProjectFileName(proj)
        If fName <> "" And fName <> "(未保存のブック)" Then
            basUtility.SaveAppSetting "LastSelectedProject", fName
        End If
        
        Me.Hide
        ' 選択されたプロジェクトを渡して実行
        Call basExport.ExportModulesToClipboardJSON(proj)
        Unload Me
    End If
End Sub

'================================================================================
' インポート実行 (案A・内容プレビュー確認)
'================================================================================
Private Sub CommandButton2_Click()
    Dim proj As Object
    Set proj = GetSelectedProject()
    
    If proj Is Nothing Then Exit Sub

    ' 1. クリップボードの中身を簡易チェック
    Dim clipText As String
    clipText = basUtility.GetClipboardText()
    
    ' 空チェック
    If Len(clipText) = 0 Then
        MsgBox "クリップボードが空です。", vbExclamation
        Exit Sub
    End If

    ' JSONっぽくない([ で始まっていない)場合は警告して中止
    If Left(Trim(clipText), 1) <> "[" Then
        MsgBox "クリップボードの内容が有効なJSON形式ではないようです。" & vbCrLf & _
               "処理を中止します。", vbExclamation, "形式エラー"
        Exit Sub
    End If

    ' 2. JSONをパースしてモジュール名を取得・プレビュー
    Dim jsonColl As Collection
    On Error Resume Next
    Set jsonColl = basJson.ParseJSON(clipText)
    On Error GoTo 0

    If jsonColl Is Nothing Then
        MsgBox "JSONの解析に失敗しました。", vbCritical
        Exit Sub
    End If
    
    If jsonColl.Count = 0 Then
        MsgBox "JSON内にモジュール定義が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 3. インポート対象のモジュール名を列挙
    Dim modList As String
    Dim i As Long
    Dim item As Object ' basJsonのパース結果はCollectionの中にCollectionが入る構造
    Dim modName As String
    Dim displayCount As Long
    
    displayCount = 0
    For i = 1 To jsonColl.Count
        Set item = jsonColl(i)
        modName = ""
        On Error Resume Next
        modName = item("name")
        On Error GoTo 0
        
        If modName <> "" Then
            If displayCount < 10 Then ' 表示は10個までにする
                modList = modList & "・" & modName & vbCrLf
            End If
            displayCount = displayCount + 1
        End If
    Next i
    
    If displayCount > 10 Then
        modList = modList & "他 " & (displayCount - 10) & " 個..." & vbCrLf
    End If

    ' 4. 詳細な確認メッセージを表示
    Dim msg As String
    msg = "以下の " & displayCount & " 個のモジュールを" & vbCrLf & _
          "プロジェクト [" & proj.Name & "] にインポートします。" & vbCrLf & vbCrLf & _
          modList & vbCrLf & _
          "※警告: 既存の同名モジュールは上書きされます。" & vbCrLf & _
          "実行してよろしいですか？"

    If MsgBox(msg, vbOKCancel + vbExclamation + vbDefaultButton2, "インポート実行の最終確認") <> vbOK Then
        Exit Sub
    End If

    ' ★設定保存: 選択したプロジェクトのファイル名を保存
    Dim fName As String
    fName = basUtility.GetProjectFileName(proj)
    If fName <> "" And fName <> "(未保存のブック)" Then
        basUtility.SaveAppSetting "LastSelectedProject", fName
    End If

    ' 5. 実行
    Me.Hide
    Call basImport.ImportModulesFromClipboardJSON(proj)
    Unload Me
End Sub
