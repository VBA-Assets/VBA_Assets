Option Explicit

' フォーム概要: 動作モード選択ダイアログ
' 読み込み開始前に表示され、ユーザーの選択結果（編集/閲覧）を呼び出し元に返す
' Last Updated: 2026/02/06 05:55 (Ver336: Added Esc key handling)

' =============================================================================
' 公開プロパティ
' =============================================================================
' 編集モードが選択されたかどうか（True=編集, False=閲覧）
' ただし IsCancelled=True の場合はこの値は無視される
Public ResultEditMode As Boolean
' キャンセル（×ボタン/Esc）されたかどうか
Public IsCancelled As Boolean

' 対象ブック（表示用などに使用する場合）
Public TargetBook As Workbook

' -----------------------------------------------------------------------------
' 初期化
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ResultEditMode = False
    IsCancelled = True ' デフォルトはキャンセル扱い
End Sub

' -----------------------------------------------------------------------------
' 「閲覧モード」ボタンクリック時
' -----------------------------------------------------------------------------
Private Sub btnViewMode_Click()
    ' 結果を設定: 閲覧モード
    ResultEditMode = False
    IsCancelled = False
    ' フォームを閉じるのではなく隠す（呼び出し元に制御を戻すため）
    Me.Hide
End Sub

' -----------------------------------------------------------------------------
' 「編集モード」ボタンクリック時
' -----------------------------------------------------------------------------
Private Sub btnEditMode_Click()
    ' 結果を設定: 編集モード
    ResultEditMode = True
    IsCancelled = False
    ' フォームを閉じるのではなく隠す
    Me.Hide
End Sub

' -----------------------------------------------------------------------------
' フォームが閉じられる時の処理（×ボタン等）
' -----------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' ユーザーがウィンドウの×ボタンをクリックした場合
    If CloseMode = vbFormControlMenu Then
        ' フォームを破棄せずに隠すだけに留める
        Cancel = True
        IsCancelled = True
        Me.Hide
    End If
End Sub

' -----------------------------------------------------------------------------
' Escキー対応
' 各コントロールにフォーカスがある場合でもEscキーを拾ってキャンセル処理を行う
' -----------------------------------------------------------------------------
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then CancelProcessing
End Sub

Private Sub btnViewMode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then CancelProcessing
End Sub

Private Sub btnEditMode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then CancelProcessing
End Sub

' -----------------------------------------------------------------------------
' 共通キャンセル処理
' -----------------------------------------------------------------------------
Private Sub CancelProcessing()
    IsCancelled = True
    Me.Hide
End Sub