Option Explicit
'==================================================================================================
'   概要        : 文書から、個人情報（プロパティ）・コメント・変更履歴・蛍光ペン・文書変数などを削除します。
'
'   備考        : Mac版Wordでは、個人情報などの一括削除（RemoveDocumentInformation）が使えないため、
'                 蛍光ペンと文書変数だけを削除します。
'==================================================================================================

'--------------------------------------------------------------------------------------------------
'   概要        : 確認のメッセージを出してから、作業中の文書の情報を削除します（メイン処理）。
'                 リボンのほか、クイックアクセスツールバーなどからも呼べるよう、引数を持ちません。
'--------------------------------------------------------------------------------------------------
Sub SanitizeDocumentCompletely()
    Dim result As VbMsgBoxResult
    Dim confirmMsg As String
    Dim completeMsg As String
    Dim i As Long

    On Error GoTo HandleErr

    '--- 前処理 ---
    '開いている文書がなければ、案内して終わります
    If Documents.Count = 0 Then
        MsgBox "文書を開いてから、［情報削除］を押してください。", vbExclamation, strAppTitle
        Exit Sub
    End If

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

    '実行の有無を選択する
    result = MsgBox(confirmMsg, vbYesNo + vbQuestion, strAppTitle)
    If result <> vbYes Then Exit Sub

    '--- メイン処理 ---
    With ActiveDocument
        ' すべての情報を削除する (Windowsのみ実行)
        ' RemoveDocumentInformationはMac版Wordでは使用不可のためスキップ
        #If Win64 Or Win32 Then
            '文書の種類によっては削除できない情報があっても、残りの削除を続けるため、この1行だけエラーを無視する
            On Error Resume Next
            .RemoveDocumentInformation wdRDIAll
            On Error GoTo HandleErr
        #End If

        ' すべての蛍光ペンを削除する (Mac/Win共通)
        .Content.HighlightColorIndex = wdNoHighlight

        ' すべての文書変数を削除する (Mac/Win共通)
        ' RemoveDocumentInformation wdRDIAll では文書変数が消えないため、個別に削除する
        ' (VA仮名化の対応表も消えるが、仮名化した状態で配布する場合もあるため、確認せずに削除する)
        For i = .Variables.Count To 1 Step -1
            .Variables(i).Delete
        Next i
    End With

    ' 終了メッセージを表示する
    MsgBox completeMsg, vbInformation, strAppTitle
    Exit Sub

HandleErr:
    MsgBox "情報の削除の途中でエラーが起きました。" & vbCrLf & _
           "番号：" & Err.Number & vbCrLf & Err.Description, vbExclamation, strAppTitle
End Sub

'--------------------------------------------------------------------------------------------------
'   概要        : リボンの［情報削除］ボタンから呼ばれます。
'   引数        : control … 押されたボタン（参照設定に頼らないよう Object で受けます）
'--------------------------------------------------------------------------------------------------
Sub OnSanitize(ByVal control As Object)
    ' メイン処理を呼び出す
    SanitizeDocumentCompletely
End Sub
