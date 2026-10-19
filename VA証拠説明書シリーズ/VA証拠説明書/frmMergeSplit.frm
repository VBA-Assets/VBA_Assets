Option Explicit

' 結合ボタンが押されたときの処理
Private Sub CommandButton1_Click()
    Me.Hide
    Call ThisDocument.ExecuteTableFormat("Merge")
    Unload Me
End Sub

' 分割ボタンが押されたときの処理
Private Sub CommandButton2_Click()
    Me.Hide
    Call ThisDocument.ExecuteTableFormat("Split")
    Unload Me
End Sub

Private Sub Label1_Click()

End Sub