Option Explicit
'==================================================================================================
'   概要        : 作業中のブックから、コメント・メモ・ドキュメントのプロパティ・個人情報などを削除します。
'
'   備考        : Mac版Excelでは、個人情報などの一括削除（RemoveDocumentInformation）が使えないため、
'                 全シートのコメントと、主要なドキュメントのプロパティ（作成者など）だけを削除します。
'==================================================================================================

'--------------------------------------------------------------------------------------------------
'   概要        : 確認のメッセージを出してから、作業中のブックの情報を削除します（メイン処理）。
'                 リボンのほか、クイックアクセスツールバーなどからも呼べるよう、引数を持ちません。
'--------------------------------------------------------------------------------------------------
Sub SanitizeWorkbookCompletely()
    Dim result As VbMsgBoxResult
    Dim msgConfirm As String
    Dim msgComplete As String

    On Error GoTo HandleErr

    '--- 前処理 ---
    '開いているブックがなければ、案内して終わります
    If ActiveWorkbook Is Nothing Then
        MsgBox "ブックを開いてから、［情報削除］を押してください。", vbExclamation, strAppTitle
        Exit Sub
    End If

    ' OSによるメッセージの分岐
    #If Mac Then
        msgConfirm = "Mac版Excelでは標準の検査機能(RemoveDocumentInformation)がサポートされていません。" & vbCrLf & _
                     "代替として、以下の簡易削除処理を実行します。" & vbCrLf & _
                     "・全シートのコメント削除" & vbCrLf & _
                     "・作成者情報の削除" & vbCrLf & vbCrLf & _
                     "この操作は元に戻せません。継続しますか？"
        msgComplete = "コメントと作成者情報を削除しました。"
    #Else
        msgConfirm = "この操作により、ブックからすべてのコメント、ドキュメントのプロパティ、個人情報などが完全に削除されます。" & vbCrLf & _
                     "この操作は元に戻せません。継続しますか？"
        msgComplete = "ブックからすべてのコメント、プロパティなどが削除されました。"
    #End If

    '実行の有無を選択する
    result = MsgBox(msgConfirm, vbYesNo + vbQuestion, strAppTitle)
    If result <> vbYes Then Exit Sub

    '--- メイン処理 ---
    #If Mac Then
        ' --- Mac用 簡易サニタイズ処理 ---
        Dim ws As Object
        ' 全シートのコメントを削除する
        ' （保護されたシートなど、削除できないシートがあっても残りのシートを続けるため、この1行だけエラーを無視する）
        For Each ws In ActiveWorkbook.Worksheets
            On Error Resume Next
            ws.Cells.ClearComments
            On Error GoTo HandleErr
        Next ws

        ' 主要なプロパティを空にする
        With ActiveWorkbook.BuiltinDocumentProperties
            .Item("Author").Value = ""
            .Item("Last Author").Value = ""
            .Item("Company").Value = ""
        End With
    #Else
        ' --- Windows用 完全サニタイズ処理 ---
        ' すべての情報を削除する (Excelの定数 xlRDIAll = 99)
        ' （ブックの種類によっては削除できない情報があっても、完了の知らせまで進めるため、この1行だけエラーを無視する）
        On Error Resume Next
        ActiveWorkbook.RemoveDocumentInformation 99
        On Error GoTo HandleErr
    #End If

    '終了メッセージを表示する
    MsgBox msgComplete, vbInformation, strAppTitle
    Exit Sub

HandleErr:
    MsgBox "情報の削除の途中でエラーが起きました。" & vbCrLf & _
           "番号：" & Err.Number & vbCrLf & Err.Description, vbExclamation, strAppTitle
End Sub

'--------------------------------------------------------------------------------------------------
'   概要        : リボンの［情報削除］ボタンから呼ばれます。
'   引数        : control … 押されたボタン（参照設定に頼らないよう Object で受けます）
'--------------------------------------------------------------------------------------------------
Sub SanitizeWorkbookFromRibbon(ByVal control As Object)
    ' メイン処理を呼び出す
    SanitizeWorkbookCompletely
End Sub
