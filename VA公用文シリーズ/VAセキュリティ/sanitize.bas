Option Explicit

' メイン処理：リボン依存に変更
' 引数を削除し、他の場所からも呼び出しやすくしました
Sub SanitizeDocumentCompletely()
    '実行の有無を選択する
    Dim result As VbMsgBoxResult
    Dim confirmMsg As String
    Dim completeMsg As String
    Dim i As Long
    
    ' OSによるメッセージと処理の分岐
    #If Mac Then
        confirmMsg = "この操作により、文書からすべての蛍光ペンと文書変数が削除されます。" & vbCrLf & _
                     "(文書変数は、アドインなどが文書に保存した、目に見えない情報です)" & vbCrLf & _
                     "(注: Mac版では個人情報・変更履歴の一括削除機能はサポートされていません)" & vbCrLf & _
                     "この操作は元に戻せません。継続しますか？"
        completeMsg = "文書からすべての蛍光ペンと文書変数が削除されました。"
    #Else
        confirmMsg = "この操作により、文書からすべての蛍光ペン、コメント、変更履歴、プロパティ、文書変数などが完全に削除されます。" & vbCrLf & _
                     "(文書変数は、アドインなどが文書に保存した、目に見えない情報です)" & vbCrLf & _
                     "この操作は元に戻せません。継続しますか？"
        completeMsg = "文書からすべての蛍光ペン、コメント、変更履歴、プロパティ、文書変数などが削除されました。"
    #End If

    result = MsgBox(confirmMsg, vbYesNo + vbExclamation, "VAセキュリティ")
    
    If result = vbYes Then
        With ActiveDocument
            ' すべての情報を削除する (Windowsのみ実行)
            ' RemoveDocumentInformationはMac版Wordでは使用不可のためスキップ
            #If Win64 Or Win32 Then
                On Error Resume Next '念のためエラー回避
                .RemoveDocumentInformation wdRDIAll
                On Error GoTo 0
            #End If
        
            ' すべての蛍光ペンを削除する (Mac/Win共通)
            .Content.HighlightColorIndex = wdNoHighlight

            ' すべての文書変数を削除する (Mac/Win共通)
            ' RemoveDocumentInformation wdRDIAll では文書変数が消えないため、個別に削除する
            ' (VA仮名化の対応表も消えるが、仮名化した状態で配布する場合もあるため、確認せずに削除する)
            For i = .Variables.Count To 1 Step -1
                .Variables(i).Delete
            Next i

            ' 終了メッセージを表示する
            MsgBox completeMsg, vbInformation
        End With
    End If
End Sub

' リボンからのコールバック用プロシージャ
Sub OnSanitize(control As IRibbonControl)
    ' メイン処理を呼び出す
    SanitizeDocumentCompletely
End Sub
