Option Explicit

'================================================================================
' フォームに必要なコントロール（VBEフォームデザイナーで設定してください）
'
' cboProjectList   ComboBox   プロジェクト選択
' lblFolderPath    Label      フォルダパス表示（BorderStyle=1, BackColor=薄いグレー）
' CommandButton3   Button     フォルダ変更（Caption: "変更"）
' lblVBAProject    Label      Caption: "VBAプロジェクト"（小さいフォント・グレー）
' CommandButton1   Button     Caption: "エクスポート →"
' CommandButton2   Button     Caption: "← インポート"
' lblFolderFiles   Label      Caption: "フォルダ（ファイル）"（小さいフォント・グレー）
'================================================================================

Private m_Projects As Collection

'--------------------------------------------------------------------------------
' フォーム初期化
'--------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Set m_Projects = New Collection

    With cboProjectList
        .Style = fmStyleDropDownList
    End With

    LoadProjects
End Sub

'--------------------------------------------------------------------------------
' プロジェクト一覧の読み込み
'--------------------------------------------------------------------------------
Private Sub LoadProjects()
    Dim vbProjs As Object
    Dim proj As Object
    Dim i As Long
    Dim defaultFileName As String
    Dim fName As String
    Dim validCount As Long

    Set vbProjs = Application.VBE.VBProjects
    defaultFileName = basUtility.GetAppSetting("LastSelectedProject")

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
        If proj.Protection = 0 Then
            validCount = validCount + 1
            fName = basUtility.GetProjectFileName(proj)
            If fName = "" Then fName = "(未保存のブック)"
            m_Projects.Add proj
            cboProjectList.AddItem proj.Name & "  (" & fName & ")"
        End If
    Next i

    If validCount = 0 Then
        cboProjectList.AddItem "操作可能なプロジェクトなし"
        cboProjectList.Enabled = False
    Else
        SetDefaultSelection defaultFileName
        If cboProjectList.ListIndex >= 0 Then
            UpdateFolderDisplay m_Projects(cboProjectList.ListIndex + 1)
        End If
    End If
End Sub

'--------------------------------------------------------------------------------
' デフォルト選択
'--------------------------------------------------------------------------------
Private Sub SetDefaultSelection(fileName As String)
    Dim i As Long
    If fileName = "" Then Exit Sub

    For i = 0 To cboProjectList.ListCount - 1
        If InStr(1, cboProjectList.List(i), fileName, vbTextCompare) > 0 Then
            cboProjectList.ListIndex = i
            Exit Sub
        End If
    Next i

    If cboProjectList.ListCount > 0 Then cboProjectList.ListIndex = 0
End Sub

'--------------------------------------------------------------------------------
' プロジェクト変更時にフォルダ表示を更新
'--------------------------------------------------------------------------------
Private Sub cboProjectList_Change()
    Dim proj As Object
    Set proj = GetSelectedProject()
    If Not proj Is Nothing Then UpdateFolderDisplay proj
End Sub

'--------------------------------------------------------------------------------
' フォルダパスの表示を更新（末尾のフォルダ名が見えるよう短縮表示）
'--------------------------------------------------------------------------------
Private Sub UpdateFolderDisplay(proj As Object)
    Dim fName As String
    Dim folderPath As String
    fName = basUtility.GetProjectFileName(proj)
    folderPath = basUtility.GetProjectFolderPath(fName)

    If folderPath = "" Then
        lblFolderPath.Caption = "（未設定）"
        lblFolderPath.ForeColor = RGB(180, 180, 180)
    Else
        lblFolderPath.Caption = ShortenPath(folderPath)
        lblFolderPath.ForeColor = RGB(80, 80, 80)
    End If
End Sub

'--------------------------------------------------------------------------------
' パスを短縮表示（親フォルダ名\フォルダ名 の形式）
'--------------------------------------------------------------------------------
Private Function ShortenPath(ByVal folderPath As String) As String
    Dim sep As String
    sep = Application.PathSeparator

    ' 末尾セパレータを除去
    If Right(folderPath, 1) = sep Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If

    Dim p1 As Long, p2 As Long
    p1 = InStrRev(folderPath, sep)
    If p1 = 0 Then
        ShortenPath = folderPath
        Exit Function
    End If

    Dim lastName As String
    lastName = Mid(folderPath, p1 + 1)

    Dim parent As String
    parent = Left(folderPath, p1 - 1)
    p2 = InStrRev(parent, sep)

    If p2 > 0 Then
        ShortenPath = "...\" & Mid(parent, p2 + 1) & sep & lastName
    Else
        ShortenPath = parent & sep & lastName
    End If
End Function

'================================================================================
' フォルダ変更 (CommandButton3)
'================================================================================
Private Sub CommandButton3_Click()
    Dim proj As Object
    Set proj = GetSelectedProject()
    If proj Is Nothing Then Exit Sub

    Dim fName As String
    fName = basUtility.GetProjectFileName(proj)

    Dim folderPath As String
    folderPath = basUtility.SelectFolder("フォルダを選択してください")
    If folderPath = "" Then Exit Sub

    basUtility.SaveProjectFolderPath fName, folderPath
    UpdateFolderDisplay proj
End Sub

'================================================================================
' エクスポート実行 (CommandButton1)
'================================================================================
Private Sub CommandButton1_Click()
    Dim proj As Object
    Set proj = GetSelectedProject()
    If proj Is Nothing Then Exit Sub

    Dim fName As String
    fName = basUtility.GetProjectFileName(proj)

    Dim folderPath As String
    folderPath = basUtility.GetProjectFolderPath(fName)

    If folderPath = "" Then
        folderPath = basUtility.SelectFolder("エクスポート先フォルダを選択してください")
        If folderPath = "" Then Exit Sub
        basUtility.SaveProjectFolderPath fName, folderPath
        UpdateFolderDisplay proj
    End If

    Dim msg As String
    msg = "プロジェクト: " & proj.Name & vbCrLf & _
          "フォルダ: " & folderPath & vbCrLf & vbCrLf & _
          "全モジュールをファイルにエクスポートします。" & vbCrLf & _
          "よろしいですか？"
    If MsgBox(msg, vbOKCancel + vbQuestion, "エクスポート確認") <> vbOK Then Exit Sub

    If fName <> "" And fName <> "(未保存のブック)" Then
        basUtility.SaveAppSetting "LastSelectedProject", fName
    End If

    Me.Hide
    Call basExport.ExportModulesToFiles(proj, folderPath)
    Unload Me
End Sub

'================================================================================
' インポート実行 (CommandButton2)
'================================================================================
Private Sub CommandButton2_Click()
    Dim proj As Object
    Set proj = GetSelectedProject()
    If proj Is Nothing Then Exit Sub

    Dim fName As String
    fName = basUtility.GetProjectFileName(proj)

    Dim folderPath As String
    folderPath = basUtility.GetProjectFolderPath(fName)

    If folderPath = "" Then
        folderPath = basUtility.SelectFolder("インポート元フォルダを選択してください")
        If folderPath = "" Then Exit Sub
        basUtility.SaveProjectFolderPath fName, folderPath
        UpdateFolderDisplay proj
    End If

    Dim sep As String
    sep = Application.PathSeparator
    If Right(folderPath, 1) <> sep Then folderPath = folderPath & sep

    If Not basUtility.FileExists(folderPath & "manifest.json") Then
        MsgBox "manifest.json が見つかりません。" & vbCrLf & _
               "フォルダ: " & folderPath, vbExclamation
        Exit Sub
    End If

    Dim manifestStr As String
    manifestStr = basUtility.ReadTextFileUTF8(folderPath & "manifest.json")

    Dim modules As Collection
    Set modules = basJson.ParseManifestModules(manifestStr)

    If modules.Count = 0 Then
        MsgBox "インポート可能なモジュールが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim modList As String
    Dim displayCount As Long
    Dim i As Long
    displayCount = 0
    For i = 1 To modules.Count
        Dim item As Object
        Set item = modules(i)
        Dim modName As String
        modName = ""
        On Error Resume Next
        modName = item("name")
        On Error GoTo 0
        If modName <> "" Then
            displayCount = displayCount + 1
            If displayCount <= 10 Then modList = modList & "・" & modName & vbCrLf
        End If
    Next i
    If displayCount > 10 Then
        modList = modList & "他 " & (displayCount - 10) & " 個..." & vbCrLf
    End If

    Dim msg As String
    msg = "以下の " & displayCount & " 個のモジュールを" & vbCrLf & _
          "プロジェクト [" & proj.Name & "] にインポートします。" & vbCrLf & vbCrLf & _
          modList & vbCrLf & _
          "フォルダ: " & folderPath & vbCrLf & vbCrLf & _
          "※警告: 既存の同名モジュールは上書きされます。" & vbCrLf & _
          "実行してよろしいですか？"

    If MsgBox(msg, vbOKCancel + vbExclamation + vbDefaultButton2, "インポート実行の最終確認") <> vbOK Then
        Exit Sub
    End If

    If fName <> "" And fName <> "(未保存のブック)" Then
        basUtility.SaveAppSetting "LastSelectedProject", fName
    End If

    Me.Hide
    Call basImport.ImportModulesFromFiles(proj, folderPath)
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' 選択中のプロジェクトを返す
'--------------------------------------------------------------------------------
Private Function GetSelectedProject() As Object
    If cboProjectList.ListIndex = -1 Then
        MsgBox "プロジェクトを選択してください。", vbExclamation
        Set GetSelectedProject = Nothing
        Exit Function
    End If
    Set GetSelectedProject = m_Projects(cboProjectList.ListIndex + 1)
End Function
