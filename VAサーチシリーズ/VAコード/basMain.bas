Option Explicit

'================================================================================
' ★エントリーポイント開発用★
' このプロシージャ内にカーソルを置き、[F5]キーを押して実行してください。
'================================================================================
Public Sub Main()
    Call ShowImportExportMenu
End Sub

'--------------------------------------------------------------------------------
' メインメニュー（ユーザーフォーム呼び出し）
'--------------------------------------------------------------------------------
Public Sub ShowImportExportMenu()
    On Error Resume Next
    frmStartMenu.Show
    If Err.Number <> 0 Then
        MsgBox "フォーム 'frmStartMenu' が見つかりません。", vbExclamation, strAppTitle
    End If
    On Error GoTo 0
End Sub

'================================================================================
' ★リボン用コールバック★
'================================================================================
' 他のアドイン（VAクリップ等）と重ならない固有の名前にしています。
Public Sub VACode_OnRibbonButton_Click(ByVal control As Object)
    Call ShowImportExportMenu
End Sub
