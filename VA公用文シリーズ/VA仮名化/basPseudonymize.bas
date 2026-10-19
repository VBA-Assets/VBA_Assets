Option Explicit
'==================================================================================================
'   概要        : VA仮名化のリボンのコールバックと、主な処理（仮名化・実名に戻す・語の登録）を行います。
'                 処理の対象は、文書の本文（表を含む）です。脚注・ヘッダー・フッター・テキストボックスは
'                 対象外です。
'==================================================================================================

'共通の前処理（BeginProcess）で開始・退避し、共通の後処理（EndProcess）で終了・復元する状態
Private docProcess As Document ' 処理の対象の文書
Private blnUndoRecordStarted As Boolean ' 元に戻す（Undo）の記録を開始したかどうか
Private blnTrackRevisions As Boolean ' 退避した変更履歴の記録の設定値
Private blnTrackRevisionsSaved As Boolean ' 変更履歴の記録の設定値を退避できたかどうか

'語の登録で前回選んだ種別（次の登録で、種別を選ぶ画面に最初から選んでおく。Word を終了すると消える）
Private strLastType As String

'================================================================================
' リボンコールバック (Ribbon Callbacks)
'================================================================================

' 機能: リボンから「仮名」（文書全体を仮名化）を実行するコールバック
Sub rxPseudonymize(control As IRibbonControl)
    Call Pseudonymize
End Sub

' 機能: リボンから「実名」（文書全体を実名に戻す）を実行するコールバック
Sub rxRestore(control As IRibbonControl)
    Call Restore
End Sub

' 機能: リボンから「選んだ語を仮名化」を実行するコールバック
Sub rxRegisterWord(control As IRibbonControl)
    Call RegisterSelectedWord
End Sub


'================================================================================
' メイン処理プロシージャ (Main Procedures)
'================================================================================

'================================================================================
' 機能:         文書の本文の登録した語を、記号に置き換えます（仮名化）。
'               長い語から先に置き換えます（「山田太郎」を「山田」より先に置き換えるため）。
'               仮名化中の文書でもう一度実行した場合は、残っている語だけが置き換わります。
'================================================================================
Sub Pseudonymize()
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    Dim doc As Document ' 対象の文書
    Dim strTokens() As String ' 記号
    Dim strWords() As String ' 語
    Dim strVarNames() As String ' 文書変数の名前
    Dim lngCount As Long ' 項目の数
    Dim lngOrder() As Long ' 語の長い順の並び
    Dim blnDone As Boolean ' 処理を終えたかどうか
    Dim i As Long

    '文書が開かれていない場合は終了する
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument

    'ファイルから消して覚えておいた辞書があれば、書き戻す
    Call basDictionary.RestoreFromMemory(doc)

    '辞書を読み込む
    lngCount = basDictionary.LoadEntries(doc, strTokens, strWords, strVarNames)
    '登録した語がない場合は、登録の方法を知らせて終了する
    If lngCount = 0 Then
        MsgBox "この文書には、仮名化した語がありません。" & vbCrLf & _
               "伏せたい語を文書の中で選んで、［選んだ語を仮名化］を押してください。", vbInformation, strAppTitle
        Exit Sub
    End If

    '--- 前処理 ---
    Call BeginProcess(doc, "仮名化")

    '--- メイン処理 ---
    '語の長い順の並びを求める
    lngOrder = SortByLengthDesc(strWords, lngCount)
    '長い語から順に、語を記号に置き換える
    For i = 1 To lngCount
        Call ReplaceAllInBody(doc, strWords(lngOrder(i)), strTokens(lngOrder(i)))
    Next i
    blnDone = True

Cleanup:
    '--- 後処理 ---
    Call EndProcess
    '本文に記号がなければ、辞書をファイルから消して覚えておく
    Call StoreDictionary(doc)

    '処理を終えた場合は、伏せ漏れを確かめるよう知らせる
    If blnDone Then
        MsgBox "仮名化しました。" & vbCrLf & vbCrLf & _
               "AIに渡す前に、伏せたい語が残っていないか確かめてください。" & vbCrLf & _
               "残っていた場合は、その語を選んで［選んだ語を仮名化］を押してください。", _
               vbInformation, strAppTitle
    End If

    Set doc = Nothing
    Exit Sub

HandleErr:
    '--- エラー処理 ---
    MsgBox "仮名化の途中でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, strAppTitle
    blnDone = False
    Resume Cleanup
End Sub

'================================================================================
' 機能:         文書の本文の記号を、登録した語に戻します。
'               英字が半角になった記号（〔人名A〕）も戻します。
'               戻せなかった記号（辞書にない記号など）があれば、その件数を知らせます。
'               仮名化中でない文書でも実行できます（AIの出力を後から貼った場合のため）。
'================================================================================
Sub Restore()
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    Dim doc As Document ' 対象の文書
    Dim strTokens() As String ' 記号
    Dim strWords() As String ' 語
    Dim strVarNames() As String ' 文書変数の名前
    Dim lngCount As Long ' 項目の数
    Dim lngRemain As Long ' 戻せなかった記号の数
    Dim blnCleared As Boolean ' 辞書をファイルから消したかどうか
    Dim blnDone As Boolean ' 処理を終えたかどうか
    Dim i As Long

    '文書が開かれていない場合は終了する
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument

    'ファイルから消して覚えておいた辞書があれば、書き戻す（実名に戻した後に、AIの出力を貼った場合のため）
    Call basDictionary.RestoreFromMemory(doc)

    '辞書を読み込む
    lngCount = basDictionary.LoadEntries(doc, strTokens, strWords, strVarNames)
    '登録した語がない場合は終了する
    If lngCount = 0 Then
        MsgBox "この文書には、仮名化した語がありません。" & vbCrLf & _
               "AIの出力は、仮名化した文書（元の文書）に貼り付けてから、［実名］を押してください。", _
               vbInformation, strAppTitle
        Exit Sub
    End If

    '--- 前処理 ---
    Call BeginProcess(doc, "実名に戻す")

    '--- メイン処理 ---
    'すべての項目を繰り返す
    For i = 1 To lngCount
        Call RestoreToken(doc, strTokens(i), strWords(i))
    Next i
    '戻せなかった記号を数える
    lngRemain = basDictionary.CountTokens(doc.Content.Text)
    blnDone = True

Cleanup:
    '--- 後処理 ---
    Call EndProcess
    '本文に記号がなければ、辞書をファイルから消して覚えておく
    blnCleared = StoreDictionary(doc)

    '処理を終えた場合は、結果を知らせる
    If blnDone Then
        If lngRemain = 0 Then
            If blnCleared Then
                MsgBox "実名に戻しました。" & vbCrLf & vbCrLf & _
                       "仮名化した語の一覧は、この文書のファイルからは消しました（実名がファイルに隠れて残らないようにするため）。" & vbCrLf & _
                       "この文書を閉じるまでは、［仮名］を押すと同じ記号で仮名化できます。", vbInformation, strAppTitle
            Else
                MsgBox "実名に戻しました。", vbInformation, strAppTitle
            End If
        Else
            MsgBox "実名に戻しました。" & vbCrLf & vbCrLf & _
                   "ただし、実名に戻せなかった記号が " & lngRemain & " 件あります。" & vbCrLf & _
                   strTokenOpen & "　" & strTokenClose & "で囲まれた記号を探して、直してください。", _
                   vbExclamation, strAppTitle
        End If
    End If

    Set doc = Nothing
    Exit Sub

HandleErr:
    '--- エラー処理 ---
    MsgBox "実名に戻す途中でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, strAppTitle
    blnDone = False
    Resume Cleanup
End Sub

'================================================================================
' 機能:         選択範囲の語を、フォームで選んだ種別で辞書に登録し、文書のその語をすぐ記号に置き換えます。
'               （登録した語が記号に変わるので、どの語を登録したかが文書を見て分かる）
'               登録済みの語（または記号）を選んだ場合は、登録の解除を確認し、解除する語の記号を元の語に戻します。
'================================================================================
Sub RegisterSelectedWord()
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    Dim doc As Document ' 対象の文書
    Dim strText As String ' 選択範囲の文字列（前後の空白などを取り除いたもの）
    Dim strType As String ' 種別（例：人名）
    Dim strToken As String ' 記号
    Dim strWord As String ' 語
    Dim strVarName As String ' 文書変数の名前
    Dim strReason As String ' 登録できない理由
    Dim frm As frmTokenType ' 種別を選ぶフォーム
    Dim strStatus As String ' 処理を終えた後にステータスバーに表示する文字列

    '文書が開かれていない場合は終了する
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument

    '本文以外（脚注・ヘッダーなど）で選んだ場合は終了する
    If Selection.StoryType <> wdMainTextStory Then
        MsgBox "本文の中で、伏せたい語を選んでください。", vbInformation, strAppTitle
        GoTo Cleanup
    End If

    '選択範囲の文字列を取り出す
    If Selection.Type = wdSelectionIP Then
        strText = ""
    Else
        strText = basDictionary.TrimSelectionText(Selection.Text)
    End If
    '何も選んでいない場合は終了する
    If strText = "" Then
        MsgBox "伏せたい語を文書の中で選んでから、［選んだ語を仮名化］を押してください。", vbInformation, strAppTitle
        GoTo Cleanup
    End If

    'ファイルから消して覚えておいた辞書があれば、書き戻す（登録済みの語の判定と、記号の英字の続き番号に使う）
    Call basDictionary.RestoreFromMemory(doc)

    '--- 登録済みの語（または記号）の場合は、登録の解除 ---
    If basDictionary.FindEntry(doc, strText, strToken, strWord, strVarName) Then
        '解除するかどうかを確認する
        If MsgBox("「" & strWord & "」（" & strToken & "）は仮名化済みです。" & vbCrLf & _
                  "仮名化を解除しますか？（文書のこの記号は、元の語に戻ります）", vbYesNo + vbQuestion, strAppTitle) = vbNo Then GoTo Cleanup

        '前処理
        Call BeginProcess(doc, "仮名化の解除")
        '文書の記号を元の語に戻す（解除すると元に戻せなくなるため）
        Call RestoreToken(doc, strToken, strWord)
        '辞書から削除する
        Call basDictionary.DeleteEntry(doc, strVarName)
        strStatus = "「" & strWord & "」（" & strToken & "）の仮名化を解除しました。"
        GoTo Cleanup
    End If

    '--- 新しい語の登録 ---
    '登録できない語の場合は、理由を知らせて終了する
    If Not basDictionary.IsRegistrableWord(strText, strReason) Then
        MsgBox strReason, vbInformation, strAppTitle
        GoTo Cleanup
    End If

    '種別を選ぶフォームを表示する（前回選んだ種別を、最初から選んでおく）
    Set frm = New frmTokenType
    strType = frm.SelectType(strText, strLastType)
    Unload frm
    Set frm = Nothing
    'キャンセルした場合は終了する
    If strType = "" Then GoTo Cleanup
    '選んだ種別を、次の登録のために覚えておく
    strLastType = strType
    '種別が一覧にない場合は終了する（定数の誤りに備える）
    If Not basDictionary.IsValidType(strType) Then
        MsgBox "種別「" & strType & "」は使えません。", vbExclamation, strAppTitle
        GoTo Cleanup
    End If

    '前処理
    Call BeginProcess(doc, "選んだ語の仮名化")
    '辞書に追加する
    strToken = basDictionary.AddEntry(doc, strType, strText)
    '文書のその語を、すぐ記号に置き換える
    Call ReplaceAllInBody(doc, strText, strToken)
    strStatus = "「" & strText & "」を " & strToken & " に仮名化しました。"

Cleanup:
    '--- 後処理 ---
    Call EndProcess
    '本文に記号がなければ、辞書をファイルから消して覚えておく
    Call StoreDictionary(doc)

    'フォームが残っている場合（フォームの表示中にエラーが起きた場合）は破棄する
    If Not frm Is Nothing Then
        On Error Resume Next
        Unload frm
        Set frm = Nothing
        On Error GoTo 0
    End If

    '処理を終えた場合は、ステータスバーに結果を表示する
    '（登録は繰り返し行う操作のため、メッセージは出さない。Mac版で未対応の場合に備え、エラーは無視する）
    If strStatus <> "" Then
        On Error Resume Next
        Application.StatusBar = strStatus
        On Error GoTo 0
    End If

    Set doc = Nothing
    Exit Sub

HandleErr:
    '--- エラー処理 ---
    MsgBox "選んだ語の仮名化の途中でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, strAppTitle
    strStatus = ""
    Resume Cleanup
End Sub


'================================================================================
' 共通の処理 (Common Procedures)
'================================================================================

'================================================================================
' 機能:         辞書を Word の中に覚え、本文に記号がなければ辞書をファイル（文書変数）から消します。
'               各処理の後処理から呼びます。エラーが起きても、呼び出し元の後処理を止めないようにします。
' 引数:         doc - 対象の文書（Nothing の場合は何もしない）
' 戻り値:       辞書をファイルから消した場合は True
'================================================================================
Private Function StoreDictionary(ByVal doc As Document) As Boolean
    StoreDictionary = False
    If doc Is Nothing Then Exit Function
    On Error Resume Next
    StoreDictionary = basDictionary.StoreToMemory(doc)
    On Error GoTo 0
End Function
'================================================================================
' 機能:         記号を元の語に戻します。英字が半角になった記号（〔人名A〕）も戻します。
' 引数:         doc - 対象の文書
'               strToken - 記号
'               strWord - 元の語
'================================================================================
Private Sub RestoreToken(ByVal doc As Document, ByVal strToken As String, ByVal strWord As String)
    Dim strNarrow As String ' 英字を半角にした記号

    '記号を元の語に置き換える
    Call ReplaceAllInBody(doc, strToken, strWord)
    '英字を半角にした記号も、元の語に置き換える
    strNarrow = basDictionary.ToNarrowLetters(strToken)
    If strNarrow <> strToken Then Call ReplaceAllInBody(doc, strNarrow, strWord)
End Sub

'================================================================================
' 機能:         文書の本文（表を含む）の文字列を、すべて置き換えます。
'               書式は検索・置換の対象にせず、文字列だけを完全一致で置き換えます。
' 引数:         doc - 対象の文書
'               strFind - 検索する文字列
'               strReplace - 置き換える文字列
'================================================================================
Private Sub ReplaceAllInBody(ByVal doc As Document, ByVal strFind As String, ByVal strReplace As String)
    Dim rng As Range ' 検索の範囲

    '本文全体を範囲にする
    Set rng = doc.Content
    With rng.Find
        '前回の検索・置換の条件を消す
        .ClearFormatting
        .Replacement.ClearFormatting
        '検索・置換の文字列（「^」は特殊文字の記号のため「^^」にする）
        .Text = Replace(strFind, "^", "^^")
        .Replacement.Text = Replace(strReplace, "^", "^^")
        '文字列だけを完全一致で検索する
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        '日本語版Wordの検索オプション（あいまい検索をオフ、全角と半角を区別、空白・句読点を無視しない）
        '（Mac版などで未対応の場合に備え、エラーは無視する）
        On Error Resume Next
        .MatchFuzzy = False
        .MatchByte = True
        .IgnoreSpace = False
        .IgnorePunct = False
        'エラーのトラップを呼び出し元に戻す（置換のエラーは呼び出し元で処理する）
        On Error GoTo 0
        'すべて置き換える
        .Execute Replace:=wdReplaceAll
    End With

    Set rng = Nothing
End Sub

'================================================================================
' 機能:         語の長い順の並び（配列の添字）を返します。
' 引数:         strWords - 語の配列（1始まり）
'               lngCount - 語の数
' 戻り値:       添字の配列（1始まり）
'================================================================================
Private Function SortByLengthDesc(ByRef strWords() As String, ByVal lngCount As Long) As Long()
    Dim lngOrder() As Long ' 添字の並び
    Dim lngTemp As Long ' 入れ替え用
    Dim i As Long
    Dim j As Long

    '添字を順に並べる
    ReDim lngOrder(1 To lngCount)
    For i = 1 To lngCount
        lngOrder(i) = i
    Next i

    '挿入ソートで、語の長い順に並べ替える
    For i = 2 To lngCount
        lngTemp = lngOrder(i)
        j = i - 1
        Do While j >= 1
            If Len(strWords(lngOrder(j))) >= Len(strWords(lngTemp)) Then Exit Do
            lngOrder(j + 1) = lngOrder(j)
            j = j - 1
        Loop
        lngOrder(j + 1) = lngTemp
    Next i

    SortByLengthDesc = lngOrder
End Function

'================================================================================
' 機能:         仮名化・実名に戻す・語の登録に共通の前処理を行います。
'               1. 元に戻す（Undo）の記録の開始（処理全体を Ctrl+Z 1回で戻せるようにする）
'               2. 変更履歴の記録の設定の退避と無効化（置き換えが変更履歴に残らないようにする）
'               3. マウスポインタのウエイト表示と画面表示の停止
'               呼び出し元の Cleanup で、必ず EndProcess を呼んでください。
' 引数:         doc - 対象の文書
'               strProcessName - 元に戻す（Undo）の一覧に表示する処理名
'================================================================================
Private Sub BeginProcess(ByVal doc As Document, ByVal strProcessName As String)
    '対象の文書を退避する（EndProcess で変更履歴の記録の設定を戻すため）
    Set docProcess = doc

    '設定の取得・変更のエラーは無視して処理を続行する（Mac版で未対応の場合に備える）
    On Error Resume Next

    '--- 1. 元に戻す（Undo）の記録の開始 ---
    If blnUseUndoRecord Then
        Application.UndoRecord.StartCustomRecord "VA仮名化：" & strProcessName
        blnUndoRecordStarted = (Err.Number = 0)
        Err.Clear
    End If

    '--- 2. 変更履歴の記録の設定の退避と無効化 ---
    blnTrackRevisions = doc.TrackRevisions
    blnTrackRevisionsSaved = (Err.Number = 0)
    Err.Clear
    doc.TrackRevisions = False
    Err.Clear

    '--- 3. マウスポインタと画面表示 ---
    'マウスポインタをウエイト表示にする
    Application.System.Cursor = wdCursorWait
    '画面表示を停止する
    Application.ScreenUpdating = False
    Err.Clear
End Sub

'================================================================================
' 機能:         BeginProcess に対応する共通の後処理を行います。
'               正常終了・エラー終了の両方で、呼び出し元の Cleanup から必ず呼んでください。
'               （BeginProcess を呼ぶ前に Cleanup に進んだ場合も、問題なく終了します）
'               1. 変更履歴の記録の設定の復元
'               2. 元に戻す（Undo）の記録の終了
'               3. 画面表示とマウスポインタの復帰
'================================================================================
Private Sub EndProcess()
    '後処理の途中でエラーが起きても、残りの後処理を必ず行う
    On Error Resume Next

    '--- 1. 変更履歴の記録の設定の復元 ---
    '設定値を退避できた場合のみ元に戻す
    If blnTrackRevisionsSaved Then
        docProcess.TrackRevisions = blnTrackRevisions
        blnTrackRevisionsSaved = False
    End If

    '--- 2. 元に戻す（Undo）の記録の終了 ---
    '記録を開始した場合のみ終了する
    If blnUndoRecordStarted Then
        Application.UndoRecord.EndCustomRecord
        blnUndoRecordStarted = False
    End If

    '--- 3. 画面表示とマウスポインタ ---
    '画面表示を再開する
    Application.ScreenUpdating = True
    'マウスポインタを標準に戻す
    Application.System.Cursor = wdCursorNormal

    '対象の文書の参照を解放する
    Set docProcess = Nothing
End Sub
