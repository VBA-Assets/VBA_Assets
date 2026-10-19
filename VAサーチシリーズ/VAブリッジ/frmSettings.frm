Option Explicit

' =============================================================================
' フォーム概要:
' マスタファイルのパス設定および、ローカルシートとマスタシートの紐付けを管理する。
' Last Updated: 2026/02/11 10:00 (Ver415: Added GetLocalPath for initialization)
' =============================================================================

' -----------------------------------------------------------------------------
' フォーム初期化
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    
    ' ローカルブックの全シート名をコンボボックスに追加
    cmbSheets.Clear
    For Each sh In ActiveWorkbook.Worksheets
        cmbSheets.AddItem sh.Name
    Next sh
    
    ' ブックに保存されているマスタパスを取得して表示
    Dim currentPath As String
    ' ★修正: 保存されているパスを現在のOS環境に合わせて解決(変換)してから表示する
    currentPath = basUtility.GetLocalPath(basConfig.GetBookSetting(cstSettings.PROP_MASTER_PATH))
    txtMasterPath.text = currentPath
    
    ' アクティブシートをデフォルトで選択状態にする
    cmbSheets.value = ActiveSheet.Name
    
    ' マスタパスが設定済みでファイルが存在する場合、マスタ側のシート一覧を取得
    If currentPath <> "" And basUtility.FileExists(currentPath) Then
        RefreshMasterSheetList currentPath
    End If
End Sub

' -----------------------------------------------------------------------------
' 「参照」ボタンクリック時
' ファイル選択ダイアログを表示してマスタファイルを選択させる
' -----------------------------------------------------------------------------
Private Sub btnBrowse_Click()
    Dim selectedFile As Variant
    
    ' OSに応じてファイル選択ダイアログを表示
    #If Mac Then
        selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm),*.xlsx;*.xlsm")
    #Else
        selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm),*.xlsx;*.xlsm", Title:="マスタファイルを選択してください")
    #End If
    
    ' ファイルが選択された場合
    If selectedFile <> False And selectedFile <> "False" Then
        txtMasterPath.text = selectedFile
        ' 選択されたファイルのシート一覧を更新
        RefreshMasterSheetList CStr(selectedFile)
    End If
End Sub

' -----------------------------------------------------------------------------
' シート選択コンボボックス変更時
' 選択されたシートの設定済みマスタシート名を表示する
' -----------------------------------------------------------------------------
Private Sub cmbSheets_Change()
    If cmbSheets.value = "" Then Exit Sub
    
    Dim sh As Worksheet
    Set sh = ActiveWorkbook.Worksheets(cmbSheets.value)
    
    ' 紐付いているマスタシート名を取得
    Dim masterShName As String
    masterShName = basConfig.GetSheetSetting(sh, cstSettings.PROP_SUFFIX_SHEET)
    
    ' マスタシート名コンボボックスに反映
    On Error Resume Next
    cmbMasterSheets.value = masterShName
    If Err.Number <> 0 Then cmbMasterSheets.value = ""
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' 「設定保存」ボタンクリック時
' 現在の設定内容をブックおよびシートのカスタムプロパティに保存する
' -----------------------------------------------------------------------------
Private Sub btnSave_Click()
    ' 入力チェック: マスタパス空欄チェック
    If txtMasterPath.text = "" Then
        MsgBox "マスタファイルのパスを指定してください。", vbExclamation, cstSettings.MSG_TITLE
        Exit Sub
    End If
    
    ' 【追加】入力チェック: マスタパス実在チェック
    ' 指定されたパスにファイルが確実に存在するか確認する
    ' ※テキストボックス内のパス(現在のOS形式)でチェックする
    If Not basUtility.FileExists(txtMasterPath.text) Then
        MsgBox "指定されたマスタファイルが見つかりません。" & vbCrLf & _
               "正しいパスを入力してください。", vbExclamation, cstSettings.MSG_TITLE
        Exit Sub
    End If
    
    ' 入力チェック: シート選択
    If cmbSheets.value = "" Or cmbMasterSheets.value = "" Then
        MsgBox "対象シートとマスタ側のシートを両方選択してください。", vbExclamation, cstSettings.MSG_TITLE
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = ActiveWorkbook.Worksheets(cmbSheets.value)
    
    ' マスタパス保存(ブック共有設定)
    ' ※テキストボックスの値をそのまま保存する
    ' (Windowsで保存すればWindowsパス、Macで保存すればMacパスとなるが、
    '  GetLocalPathが読み込み時に自動解決するため問題ない)
    basConfig.SaveBookSetting cstSettings.PROP_MASTER_PATH, txtMasterPath.text
    
    ' シート紐付け保存(シート個別設定)
    basConfig.SaveSheetSetting sh, cstSettings.PROP_SUFFIX_SHEET, cmbMasterSheets.value
    
    ' 設定変更時は強制的に再同期させるため、最終同期時刻を初期化(リセット)
    basConfig.SaveBookSetting cstSettings.PROP_LAST_SYNC, "1900/01/01", sh.Parent
    
    MsgBox "設定を保存しました。" & vbCrLf & "「データ読み込み」を実行して同期を開始してください。", vbInformation, cstSettings.MSG_TITLE
End Sub

' -----------------------------------------------------------------------------
' 「削除」ボタンクリック時
' 選択中のシートの設定を削除する
' -----------------------------------------------------------------------------
Private Sub btnDelete_Click()
    If cmbSheets.value = "" Then Exit Sub
    
    Dim sh As Worksheet
    Set sh = ActiveWorkbook.Worksheets(cmbSheets.value)
    
    If MsgBox(sh.Name & " の管理設定を削除しますか?", vbQuestion + vbYesNo, cstSettings.MSG_TITLE) = vbYes Then
        ' 設定削除処理
        basConfig.DeleteSheetSetting sh, cstSettings.PROP_SUFFIX_SHEET
        cmbMasterSheets.value = ""
        MsgBox "設定を削除しました。", vbInformation, cstSettings.MSG_TITLE
    End If
End Sub

' -----------------------------------------------------------------------------
' 「閉じる」ボタンクリック時
' -----------------------------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' ヘルパー: マスタファイルのシート一覧取得
' マスタファイルを読み取り専用で一時的に開き、シート名を取得する
' -----------------------------------------------------------------------------
Private Sub RefreshMasterSheetList(ByVal path As String)
    If Not basUtility.FileExists(path) Then Exit Sub
    
    Dim wbMaster As Workbook, sh As Worksheet
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' マスタを開く(リンク更新なし)
    Set wbMaster = Workbooks.Open(fileName:=path, ReadOnly:=True, UpdateLinks:=False)
    
    ' 【削除】ファイルオープン直後のDoEventsは削除(OnTimeによる遅延実行があるため不要)
    
    ' シート名をコンボボックスに追加
    cmbMasterSheets.Clear
    For Each sh In wbMaster.Worksheets
        cmbMasterSheets.AddItem sh.Name
    Next sh
    
    wbMaster.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "マスタファイルの読み込みに失敗しました。", vbCritical, cstSettings.MSG_TITLE
End Sub
