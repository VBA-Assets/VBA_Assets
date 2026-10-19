Option Explicit

'外部ツール由来要素の一括整形を確認するメッセージ
Private Const strCleanupMessage As String = "外部ツール由来の要素（手動改行・リンク・WikiLink・フロントマター・水平線・強調記号）が検知されました。Wordに最適な形式へ一括整形しますか？（先頭のフロントマターと水平線の段落は削除され、強調記号は取り除かれます）"

'共通の前処理（BeginProcess）で開始・退避し、共通の後処理（EndProcess）で終了・復元する状態
Private blnUndoRecordStarted As Boolean ' 元に戻す（Undo）の記録を開始したかどうか
Private blnSmartCutPaste As Boolean ' 退避したスマート カット アンド ペーストの設定値
Private blnSmartCutPasteSaved As Boolean ' スマート カット アンド ペーストの設定値を退避できたかどうか

'================================================================================
' リボンコールバック (Ribbon Callbacks)
'================================================================================

' 機能: リボンから「ニュートラル」を実行するコールバック
Sub rxChangeToNeutral(control As IRibbonControl)
    Call ChangeToNeutral
End Sub

' 機能: リボンから「マークダウン」を実行するコールバック
Sub rxChangeToMarkdown(control As IRibbonControl)
    Call ChangeToMarkdown
End Sub

' 機能: リボンから「アウトライン」を実行するコールバック
Sub rxChangeToOutline(control As IRibbonControl)
    Call ChangeToOutline
End Sub


'================================================================================
' メイン処理プロシージャ (Main Procedures)
'================================================================================

'================================================================================
' 機能:         文書全体の書式をインデントのない「ニュートラル」な状態に
'               変換します。
'================================================================================
Sub ChangeToNeutral()
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    '選択範囲の「行間詰め」スタイルの適用を行うかどうかを判定する
    '（「行間詰め」スタイルの適用を行う設定で、範囲が選択されている場合）
    Dim blnBlockStyle As Boolean
    blnBlockStyle = (blnSelectionBlock = True And Selection.Type <> wdSelectionIP)

    '--- 前処理 ---
    '「行間詰め」スタイルの適用を行う場合は、表のスタイルと行頭の空白の処理を行わない
    Call BeginProcess("ニュートラル", Not blnBlockStyle)

    '--- 事前処理 ---
    '選択範囲の「行間詰め」スタイルの適用を行う場合
    If blnBlockStyle Then
        '選択範囲のスタイルを「行間詰め」に変更
        Call ApplyBlockStyle
        '処理をスキップする
        GoTo Cleanup
    End If

    '--- メイン処理 ---
    Dim pars As Paragraphs
    Set pars = ActiveDocument.Paragraphs

    Dim par As Paragraph
    'すべての段落を繰り返す
    'アウトラインまたはリストの段落が1つでもある場合はインデントの解除を行います。
    For Each par In pars
        '段落を対象
        With par
            'スタイル名を把握する
            Dim strStyle As String
            strStyle = .Style.NameLocal

            'スタイルが見出しまたはリスト段落の場合
            If Left(strStyle, 3) = "見出し" Or strStyle = "リスト段落" Then
                'インデントを解除する (全段落が対象)
                Call basIndent.ProcessIndent("Reset")

                '繰り返しを終了する
                Exit For
            End If
       End With
    Next

    'すべての段落を繰り返す
    'マークダウンの段落が1つでもある場合はアウトラインへの変換を行います。
    For Each par In pars
        '段落を対象
        With par
            '文字列を取得する
            Dim strText As String '段落内テキスト
            strText = .Range.Text

            '文字列が見出しまたはリストの場合
            If Left(strText, 1) = strHead Or Left(strText, 1) = strList Or basConvert.GetNumListLength(strText) > 0 Then
                'アウトラインへの変換を行う
                Call basConvert.MarkdownToOutline

                '繰り返しを終了する
                Exit For
            End If
        End With
    Next

Cleanup:
    '--- 後処理 ---
    Call EndProcess

    Set par = Nothing
    Set pars = Nothing
    Exit Sub

HandleErr:
    '--- エラー処理 ---
    Resume Cleanup
End Sub

'================================================================================
' 機能:         文書のアウトライン構造をプレーンテキストベースの
'               「マークダウン」形式に変換します。
'================================================================================
Sub ChangeToMarkdown()
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    '--- 前処理 ---
    Call BeginProcess("マークダウン", True)

    '--- メイン処理 ---
    Dim pars As Paragraphs
    Set pars = ActiveDocument.Paragraphs

    'すべての段落を繰り返す
    'アウトラインまたはリストが1つでも設定されている場合は、マークダウンへの変換を行います。
    Dim par As Paragraph
    For Each par In pars
        '段落を対象
        With par
            'スタイル名を把握する
            Dim strStyle As String
            strStyle = .Style.NameLocal

            'スタイルが見出しまたはリスト段落の場合
            If Left(strStyle, 3) = "見出し" Or strStyle = "リスト段落" Then
                'マークダウンへの変換を行う
                Call basConvert.OutlineToMarkdown

                '繰り返しを終了する
                Exit For
            End If
        End With
    Next
Cleanup:
    '--- 後処理 ---
    Call EndProcess

    Set par = Nothing
    Set pars = Nothing
    Exit Sub

HandleErr:
    Resume Cleanup
End Sub

'================================================================================
' 機能:         文書をインデントが適切に設定された「アウトライン」状態に
'               変換または更新します。
'================================================================================
Sub ChangeToOutline()
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    '--- 前処理 ---
    Call BeginProcess("アウトライン", True)

    '--- メイン処理 ---
    Dim pars As Paragraphs
    Set pars = ActiveDocument.Paragraphs

    'すべての段落を繰り返す
    'マークダウンの段落が1つでもある場合はアウトラインへの変換を行います。
    Dim par As Paragraph
    For Each par In pars
        '段落を対象
        With par
            '文字列を取得する
            Dim strText As String '段落内テキスト
            strText = .Range.Text

            'マークダウンの見出しまたはリストの場合
            If Left(strText, 1) = strHead Or Left(strText, 1) = strList Or basConvert.GetNumListLength(strText) > 0 Then
                'アウトラインへの変換を行う
                Call basConvert.MarkdownToOutline

                '繰り返しを終了する
                Exit For
            End If
        End With
    Next

    'すべての段落を繰り返す
    'アウトラインまたはリストが1つでも設定されている場合は、インデントの設定を行います。
    For Each par In pars
        '段落を対象
        With par
            'スタイル名を把握する
            Dim strStyle As String
            strStyle = .Style.NameLocal
            'スタイルが見出しまたはリスト段落の場合
            If Left(strStyle, 3) = "見出し" Or strStyle = "リスト段落" Then
                'インデントを設定する
                Call basIndent.ProcessIndent("Aline")

                '繰り返しを終了する
                Exit For
            End If
        End With
    Next

Cleanup:
    '--- 後処理 ---
    Call EndProcess

    Set par = Nothing
    Set pars = Nothing
    Exit Sub

HandleErr:
    Resume Cleanup
End Sub

'================================================================================
' 機能:         3つの変換処理（ニュートラル、マークダウン、アウトライン）に共通の
'               前処理を行います。
'               1. 外部ツール由来要素の検知と、一括整形の確認
'               2. 元に戻す（Undo）の記録の開始（処理全体を Ctrl+Z 1回で戻せるようにする）
'               3. スマート カット アンド ペーストの設定の退避と無効化
'               4. マウスポインタのウエイト表示と画面表示の停止
'               5. 外部ツール由来要素の一括整形
'               6. 表内の段落への「行間詰め」スタイルの適用と、行頭の空白のインデント化
'               エラーが起きた場合は、呼び出し元へエラーを伝えます。
'               （On Error GoTo 0 で抜けると、呼び出し元のエラー処理に戻らず
'                 マクロが停止し、EndProcess が呼ばれないため）
'               呼び出し元の Cleanup で、必ず EndProcess を呼んでください。
' 引数:         strProcessName - 元に戻す（Undo）の一覧に表示する処理名
'               blnPreFormat - True の場合は、6. の処理を行う
'================================================================================
Private Sub BeginProcess(ByVal strProcessName As String, ByVal blnPreFormat As Boolean)
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    Dim formatter As clsPreFormatter ' 検知・整形を行うクラス
    Dim blnCleanup As Boolean ' 外部ツール由来要素の一括整形を行うかどうか
    Dim lngErrNumber As Long ' エラー番号（呼び出し元へ伝えるため退避する）
    Dim strErrSource As String ' エラーの発生元
    Dim strErrDescription As String ' エラーの内容

    'インスタンスを生成する
    Set formatter = New clsPreFormatter

    '--- 1. 外部ツール由来要素の検知と、一括整形の確認 ---
    blnCleanup = False
    '外部ツール由来要素の一括整形が有効な場合のみ処理を実行する
    If blnEnableExternalCleanup Then
        'アクティブドキュメント内に外部ツール由来の特殊要素が存在する場合
        If formatter.DetectExternalElements(ActiveDocument) Then
            '一括整形を実行するかどうかをユーザーに確認する
            blnCleanup = (MsgBox(strCleanupMessage, vbYesNo + vbQuestion, "VA公用文") = vbYes)
        End If
    End If

    '--- 2. 元に戻す（Undo）の記録の開始 ---
    '確認メッセージの後、整形の前に開始し、整形も含めて Ctrl+Z 1回で戻せるようにする
    '（Mac版で未対応の場合に備え、エラーは無視して処理を続行する）
    If blnUseUndoRecord Then
        On Error Resume Next
        Application.UndoRecord.StartCustomRecord "VA公用文：" & strProcessName
        blnUndoRecordStarted = (Err.Number = 0)
        Err.Clear
        '自前のエラー処理に戻す（On Error GoTo 0 にすると、呼び出し元へ戻らずに停止する）
        On Error GoTo HandleErr
    End If

    '--- 3. スマート カット アンド ペーストの設定の退避と無効化 ---
    '有効なままだと、Range.Delete で行頭の半角空白が削除されないことがあるため
    '（Mac版で未対応の場合に備え、取得・設定のエラーは無視する）
    On Error Resume Next
    blnSmartCutPaste = Options.SmartCutPaste
    blnSmartCutPasteSaved = (Err.Number = 0)
    Options.SmartCutPaste = False
    Err.Clear
    '自前のエラー処理に戻す
    On Error GoTo HandleErr

    '--- 4. マウスポインタと画面表示 ---
    'マウスポインタをウエイト表示にする
    Application.System.Cursor = wdCursorWait
    '画面表示を停止する
    Application.ScreenUpdating = False

    '--- 5. 外部ツール由来要素の一括整形 ---
    If blnCleanup Then
        '一括整形の処理を呼び出し、不要な要素を除去・変換する
        formatter.ExecuteSimpleCleanup ActiveDocument
    End If

    '--- 6. 表のスタイルと行頭の空白 ---
    If blnPreFormat Then
        '表内の段落に「行間詰め」スタイルを適用する
        If blnApplyTableStyle = True Then
            formatter.ApplyTableStyle "行間詰め"
        End If

        '行頭の空白をインデントに置き換える
        If blnReplaceSpaceWithIndent = True Then
            '変数オブジェクトを生成する（アドバンス幅を取得するため）
            Dim Variables As clsVariablesManager
            Set Variables = New clsVariablesManager
            formatter.ReplaceSpaceWithIndent Variables.sngAdvWidth, "行間詰め"
            Set Variables = Nothing
        End If
    End If

    '使用したオブジェクト変数を解放する
    Set formatter = Nothing
    Exit Sub

HandleErr:
    '--- エラー処理 ---
    'エラーの内容を退避する（オブジェクトの解放で Err が消えないようにするため）
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    '使用したオブジェクト変数を解放する
    Set formatter = Nothing
    '呼び出し元へエラーを伝える
    '（呼び出し元の HandleErr → Cleanup → EndProcess で、退避した状態を元に戻す）
    Err.Raise lngErrNumber, strErrSource, strErrDescription
End Sub

'================================================================================
' 機能:         BeginProcess に対応する共通の後処理を行います。
'               正常終了・エラー終了の両方で、呼び出し元の Cleanup から必ず呼んでください。
'               1. 元に戻す（Undo）の記録の終了
'               2. スマート カット アンド ペーストの設定の復元
'               3. 画面表示とマウスポインタの復帰
'================================================================================
Private Sub EndProcess()
    '後処理の途中でエラーが起きても、残りの後処理を必ず行う
    On Error Resume Next

    '--- 1. 元に戻す（Undo）の記録の終了 ---
    '記録を開始した場合のみ終了する
    If blnUndoRecordStarted Then
        Application.UndoRecord.EndCustomRecord
        blnUndoRecordStarted = False
    End If

    '--- 2. スマート カット アンド ペーストの設定の復元 ---
    '設定値を退避できた場合のみ元に戻す
    If blnSmartCutPasteSaved Then
        Options.SmartCutPaste = blnSmartCutPaste
        blnSmartCutPasteSaved = False
    End If

    '--- 3. 画面表示とマウスポインタ ---
    '画面表示を再開する
    Application.ScreenUpdating = True
    'マウスポインタを標準に戻す
    Application.System.Cursor = wdCursorNormal
End Sub

'================================================================================
' 機能:         選択範囲の段落スタイルを「標準」と「行間詰め」の間で切り替えます
'================================================================================
Private Sub ApplyBlockStyle()
    Dim pars As Paragraphs
    Set pars = Selection.Paragraphs
    
    '選択範囲がない場合は終了
    If pars.count = 0 Then Exit Sub
    
    '変更後のスタイル名を格納する変数
    Dim strNewStyle As String
    
    '--- 判定処理（繰り返し処理の前に一度だけ実行） ---
    '選択範囲の最初の段落のスタイルを確認して、動作を決定する
    If pars(1).Style.NameLocal = "行間詰め" Then
        'メッセージを表示する
        If MsgBox("選択範囲の段落スタイルを「標準」に変更します。", vbOKCancel, "VA公用文") = vbCancel Then Exit Sub
        '変更スタイルを「標準」に設定する
        strNewStyle = "標準"
    Else
        'メッセージを表示する
        If MsgBox("選択範囲の段落スタイルを「行間詰め」に変更します。", vbOKCancel, "VA公用文") = vbCancel Then Exit Sub
        '変更スタイルを「行間詰め」に設定する
        strNewStyle = "行間詰め"
    End If
    
    '--- 繰り返し処理 ---
    Dim par As Paragraph
    Dim fntBefore As Font
    Dim wdpAlignBefore As WdParagraphAlignment
    Dim sngLeftIndentBefore As Single
    Dim sngFirstLineIndentBefore As Single
    
    '範囲内のすべての段落を繰り返す
    For Each par In pars
        With par
            '元のフォント、配置およびインデントを保存する
            Set fntBefore = .Range.Font.Duplicate
            wdpAlignBefore = .alignment
            sngLeftIndentBefore = .LeftIndent
            sngFirstLineIndentBefore = .FirstLineIndent
            
            '段落を変更する
            .Style = strNewStyle
    
            '元のフォント、配置およびインデントを復元する
            .Range.Font = fntBefore
            .alignment = wdpAlignBefore
            .LeftIndent = sngLeftIndentBefore
            .FirstLineIndent = sngFirstLineIndentBefore
        End With
    Next par

Cleanup:
    'オブジェクト変数を解放
    Set fntBefore = Nothing
    Set par = Nothing
    Set pars = Nothing
End Sub
