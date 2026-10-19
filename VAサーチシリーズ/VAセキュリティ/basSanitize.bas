Option Explicit

'================================================================================
' プロシージャ: SanitizeWorkbookCompletely
' 機能:       ドキュメントのプロパティなどを削除
'================================================================================
Sub SanitizeWorkbookCompletely()
    '実行の有無を選択する
    Dim result As VbMsgBoxResult
    Dim msgConfirm As String
    Dim msgComplete As String
    
    ' OSによる分岐
    #If Mac Then
        msgConfirm = "Mac版Excelでは標準の検査機能(RemoveDocumentInformation)がサポートされていません。" & vbCrLf & _
                     "代替として、以下の簡易削除処理を実行します。" & vbCrLf & _
                     "・全シートのコメント削除" & vbCrLf & _
                     "・作成者情報の削除" & vbCrLf & vbCrLf & _
                     "この操作は元に戻せません。継続しますか？"
        msgComplete = "コメントと作成者情報を削除しました。"
    #Else
        msgConfirm = "この操作により、ブックからすべてのコメント、ドキュメントのプロパティ、個人情報などが完全に削除されます。" & vbCrLf & _
                     "この操作は元に戻せません。継続しますか"
        msgComplete = "ブックからすべてのコメント、プロパティなどが削除されました。"
    #End If
    
    result = MsgBox(msgConfirm, vbYesNo + vbExclamation, "VAセキュリティ")
    
    If result = vbYes Then
        'エラー発生時も処理を継続する（保護されたシート等の対策）
        On Error Resume Next
        
        #If Mac Then
            ' --- Mac用 簡易サニタイズ処理 ---
            Dim ws As Object
            ' 全シートのコメントを削除
            For Each ws In ActiveWorkbook.Worksheets
                ws.Cells.ClearComments
            Next ws
            
            ' 主要なプロパティをクリア
            With ActiveWorkbook.BuiltinDocumentProperties
                .Item("Author").Value = ""
                .Item("Last Author").Value = ""
                .Item("Company").Value = ""
            End With
            
        #Else
            ' --- Windows用 完全サニタイズ処理 ---
            With ActiveWorkbook
                'すべての情報を削除する (Excelの定数 xlRDIAll = 99)
                .RemoveDocumentInformation 99
            End With
        #End If
        
        '終了メッセージを表示する
        MsgBox msgComplete, vbInformation
        
        'エラー処理を戻す
        On Error GoTo 0
    End If
End Sub


'================================================================================
' プロシージャ: SanitizeWorkbookFromRibbon
' 機能:       リボンから呼び出されるコールバック関数
'================================================================================
Sub SanitizeWorkbookFromRibbon(Optional control As IRibbonControl)
    ' 既存の実行プロシージャを呼び出し
    SanitizeWorkbookCompletely
End Sub
