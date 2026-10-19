Option Explicit

'================================================================================
' 検索フォーム用コード
'================================================================================

' --- 検索実行ボタン (CommandButton1) ---
Private Sub CommandButton1_Click()
    Dim strMode As String
    
    ' 検索モードの判定
    If Me.OptionButton1.Value = True Then
        strMode = "AND"
    Else
        strMode = "OR"
    End If
    
    ' メインモジュールの検索処理を呼び出し
    ' ※TextBox1に入力された文字列を渡します
    Call SearchAddin_Main.ExecuteSearch(Me.TextBox1.Text, strMode)
    
    ' 検索実行後、フォームを閉じる場合は以下を有効化してください
    ' Unload Me
End Sub

' --- 解除ボタン (CommandButton2) ---
Private Sub CommandButton2_Click()
    ' メインモジュールの解除処理を呼び出し
    Call SearchAddin_Main.ClearSearch
    
    ' フォームを閉じる
    Unload Me
End Sub