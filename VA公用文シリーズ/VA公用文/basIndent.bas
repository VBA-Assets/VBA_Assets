Option Explicit

'==================================================================================================
' 概要:         インデントの設定または解除を行う共通処理プロシージャ。
' 引数:         processMode (String) - 処理モードを指定 ("Aline" または "Reset")。
'==================================================================================================
'インデント処理プロシージャを定義する
Public Sub ProcessIndent(ByVal processMode As String)
    '処理時間を計測する
    'Dim dblStart As Double
    'dblStart = Timer

    'エラーのトラップを有効にする
    On Error GoTo HandleErr
    
    '変数オブジェクトを宣言する
    Dim Variables As clsVariablesManager
    '変数オブジェクトを生成する
    Set Variables = New clsVariablesManager
    
    'リストフォーマッターを宣言する
    Dim ListFormatter As clsListFormatter
    'アウトラインを再設定するか判定する
    If blnReconfigOutline Then
        'リストフォーマッターの準備をする
        Set ListFormatter = New clsListFormatter
    '判定を終了する
    End If
    
    '--- 初期設定 ---
    'ステータスバーに表示する文言の変数を宣言する
    Dim strStatusMessage As String
    '処理モードがAlineの場合
    If processMode = "Aline" Then
        '文言を設定する
        strStatusMessage = "設定"
    'それ以外の場合
    Else
        '文言を解除に設定する
        strStatusMessage = "解除"
    '条件判定を終了する
    End If
    '進捗状況を表示する
    Application.StatusBar = "インデントの" & strStatusMessage & "を準備しています。"

    '配列の型を変換する
    'インデントの値は計算に用いるため、数値に変換します。
    'Publicでは宣言できません。
    'カウンタ変数を宣言する
    Dim k As Long
    '見出しの左インデント用の配列を宣言する
    Dim dblHeadingLeftIndents(10) As Double
    '見出しの最初の行のインデント用の配列を宣言する
    Dim dblHeadingFirstLineIndents(10) As Double
    '本文の左インデント用の配列を宣言する
    Dim dblTextLeftIndents(10) As Double
    '本文の最初の行のインデント用の配列を宣言する
    Dim dblTextFirstLineIndents(10) As Double
    '配列の要素数だけ繰り返す
    For k = 0 To 9
        '見出しの左インデントの値を数値に変換する
        dblHeadingLeftIndents(k) = Val(Variables.varHeadingLeftIndents(k))
        '見出しの最初の行のインデントの値を数値に変換する
        dblHeadingFirstLineIndents(k) = Val(Variables.varHeadingFirstLineIndents(k))
        '本文の左インデントの値を数値に変換する
        dblTextLeftIndents(k) = Val(Variables.varTextLeftIndents(k))
        '本文の最初の行のインデントの値を数値に変換する
        dblTextFirstLineIndents(k) = Val(Variables.varTextFirstLineIndents(k))
    '次の要素へ進む
    Next

    '処理用の変数を定義する
    'セクション内の最初の見出しレベルを格納する変数を宣言する
    Dim xFirst As Long
    '見出し通過フラグを宣言する
    Dim blnAfterHeading As Boolean
    '右揃え通過フラグを宣言する
    Dim blnAfterAlignRight As Boolean
    '見出しレベルを格納する変数を宣言する
    Dim x As Long
    '見出し判定フラグを宣言する
    Dim isHeading As Boolean

    'セクションオブジェクトを宣言する
    Dim sec As Section
    '段落コレクションを宣言する
    Dim pars As Paragraphs
    '段落オブジェクトを宣言する
    Dim par As Paragraph

    '総段落数を格納する変数を宣言する
    Dim lngTotalParagraphs As Long
    '処理済み段落数を格納する変数を宣言する
    Dim lngProcessedCount As Long

    'スタイル名を格納する変数を宣言する
    Dim strStyle As String
    '段落内テキストを格納する変数を宣言する
    Dim strText As String
    
    '段落数を把握する
    lngTotalParagraphs = ActiveDocument.Paragraphs.count
    '処理済み段落数を初期化する
    lngProcessedCount = 0
    
    'すべてのセクションを繰り返す
    For Each sec In ActiveDocument.Sections
        
        '最初の見出しレベルをリセットする
        xFirst = 0
        '見出し通過フラグを解除する
        blnAfterHeading = False
        '右揃え通過フラグを解除する
        blnAfterAlignRight = False
        '見出しレベルをリセットする
        x = 0
        
        'セクション内の段落コレクションを取得する
        Set pars = sec.Range.Paragraphs
        
        'すべての段落を繰り返す
        For Each par In pars
            '段落番号をカウントする
            lngProcessedCount = lngProcessedCount + 1
            
            '50段落ごとに処理
            'スキップや繰り返し処理以降でなければなりません。
            If lngProcessedCount Mod 50 = 0 Then
                '進捗率を算定する
                '進捗率を表示する
                Application.StatusBar = "インデントを" & strStatusMessage & "しています。　" & _
                                        Int(lngProcessedCount / lngTotalParagraphs * 100) & "%完了"
            '条件判定を終了する
            End If
            
            'スタイル名を把握する
            strStyle = par.Style.NameLocal
            
            '右揃えの場合はスキップする
            If blnAfterAlignRight Then GoTo Continue
            
            'スタイルが見出しであるかどうかを判定する
            isHeading = (Left(strStyle, 3) = "見出し")
            
            '最初の見出しまたはリスト以前の場合
            If blnAfterHeading = False Then
                'スタイルが「見出し」の場合
                If isHeading Then
                    ' アウトラインを設定する
                    If blnReconfigOutline Then
                         'アウトラインの設定を適用する
                         ListFormatter.ConfigOutline par
                    '条件判定を終了する
                    End If
                
                    '見出し通過フラグを設定する
                    blnAfterHeading = True
                'スタイルが「リスト段落」の場合
                ElseIf strStyle = "リスト段落" Then
                    '見出し通過フラグを設定する
                    blnAfterHeading = True
                'それ以外の場合
                Else
                    '文字列を取得する
                    strText = par.Range.Text
                    
                    'マークダウンの見出しまたはリストの場合
                    If Left(strText, 1) = strHead Or _
                       Left(strText, 1) = strList Or _
                       basConvert.GetNumListLength(strText) > 0 Then
                       
                       '見出し通過フラグを設定する
                       blnAfterHeading = True
                    '条件判定を終了する
                    End If
                '条件判定を終了する
                End If
            '条件判定を終了する
            End If
            
            ' 最初の見出しまたはリスト以前の場合
            If Not blnAfterHeading Then
                '処理をスキップする
                If blnSkipUntil Then GoTo Continue
            
            ' 見出し通過後、かつ「停止設定」が有効な場合
            ElseIf blnStopAfter Then
                ' スタイルが「行間詰め」以外かつ右揃えの場合
                If (strStyle <> "行間詰め") And (par.Format.alignment = wdAlignParagraphRight) Then
                    '右揃えフラグを設定する
                    blnAfterAlignRight = True
                    '処理をスキップする
                    GoTo Continue
                '条件判定を終了する
                End If
            '条件判定を終了する
            End If
            
            'スタイルが「見出し」の場合
            If isHeading Then
                '見出しレベルを把握する
                x = Val(Right$(strStyle, 1))
                
                '最初の見出しレベルを把握する
                '途中で見出しレベルが高くなった場合も再設定します。
                If xFirst = 0 Or xFirst > x Then
                    '最初の見出しレベルを保存する
                    xFirst = x
                '条件判定を終了する
                End If
            '条件判定を終了する
            End If
            
            '各段落を対象に処理
            With par
                '--- モードに応じたインデント処理の実行 ---
                ' インデント「設定」の場合
                If processMode = "Aline" Then
                    '== インデント設定処理 ==
                    'スタイルが「見出し」の場合
                    If isHeading Then
                        '左インデント（ポイント単位）を設定する
                        '配列から見出しレベルに応じた値を読みだし、最初の見出しレベルに応じた増加量を加えています。
                        '文字単位でインデントを設定するとアウトラインを「右揃え」にしないとインデントが揃わず、
                        'その場合「均等割り付け」時にインデントがずれるのでポイント単位で設定しています。
                        .LeftIndent = (dblHeadingLeftIndents(x) + Variables.lngLeftIndentsAdd * (x - xFirst)) * Variables.sngAdvWidth
                        '最初の行のインデント（ポイント単位）を設定する
                        .FirstLineIndent = dblHeadingFirstLineIndents(x) * Variables.sngAdvWidth
                        
                    'スタイルが「標準」の場合
                    ElseIf strStyle = "標準" Then
                        '左インデント（ポイント単位）を設定する
                        '配列から見出しレベルに応じた値を読みだし、最初の見出しレベルに応じた増加量を加えています。
                        .LeftIndent = (dblTextLeftIndents(x) + Variables.lngLeftIndentsAdd * (x - xFirst)) * Variables.sngAdvWidth
                        '最初の行のインデント（ポイント単位）を設定する
                        '配列から見出しレベルに応じた値を読みだしています。
                        .FirstLineIndent = dblTextFirstLineIndents(x) * Variables.sngAdvWidth
                        
                    'スタイルが「リスト段落」の場合
                    ElseIf strStyle = "リスト段落" Then
                        '左インデント（ポイント単位）を設定する
                        '配列から見出しレベルに応じた値を読みだし、最初の見出しレベルに応じた増加量を加えています。
                        '見出しとの関係は、オプションのリスト段落用のインデントで設定しています。
                        .LeftIndent = (dblHeadingLeftIndents(x) + Variables.lngLeftIndentsAdd * (x - xFirst) + Variables.lngListLeftIndent) * Variables.sngAdvWidth
                        '最初の行のインデント（ポイント単位）を設定する
                        'オプションのリスト段落用のインデントに設定しています。
                        .FirstLineIndent = Variables.lngListFirstLineIndent * Variables.sngAdvWidth
                    '条件判定を終了する
                    End If
                    
               'インデント「解除」の場合
                Else
                    '== インデント解除処理 ==
                    '「標準」の段落は、段落に直接設定された書式（行頭の空白から変換したインデントを含む）を
                    'すべて解除してスタイルの定義に戻すため、.Reset を使います。
                    '「見出し」と「リスト段落」の段落は、.Reset すると見出し番号や段落番号の設定が壊れるため、
                    'インデントだけを個別に0にします。
                    'スタイルが「見出し」の場合
                    If isHeading Then
                        'インデントが0になるように調節する
                        'リセットすると見出し番号の振り直しも解除されてしまいます。
                        .LeftIndent = 0
                        '最初の行のインデントを0にする
                        .FirstLineIndent = 0
                        '最初の行の文字単位のインデントを0にする
                        .CharacterUnitFirstLineIndent = 0
                        '文字単位の左インデントを0にする
                        .CharacterUnitLeftIndent = 0
                    'スタイルが「標準」の場合
                    ElseIf strStyle = "標準" Then
                        '段落の設定を解除する（直接設定された書式をすべて解除し、スタイルの定義に戻す）
                        .Reset
                    'スタイルが「リスト段落」の場合
                    ElseIf strStyle = "リスト段落" Then
                        'インデントを0に設定する
                        '.Resetするとリストの設定まで解除されてしまいます。
                        .LeftIndent = 0
                        '最初の行のインデントを0にする
                        .FirstLineIndent = 0
                        '文字単位の左インデントを0にする
                        .CharacterUnitLeftIndent = 0
                        '最初の行の文字単位のインデントを0にする
                        .CharacterUnitFirstLineIndent = 0
                    '条件判定を終了する
                    End If
                '条件判定を終了する
                End If
            'Withブロックを終了する
            End With
            
Continue:
        '次の段落へ進む
        Next par
    '次のセクションへ進む
    Next sec

Cleanup:
    '進捗状況を解除する
    Application.StatusBar = ""

    '処理時間を計測する
    'Dim dblEnd As Double
    'dblEnd = Timer
    'MsgBox dblEnd - dblStart

    '段落オブジェクト変数を解放
    Set par = Nothing
    '段落コレクションオブジェクト変数を解放
    Set pars = Nothing
    'セクションオブジェクト変数を解放
    Set sec = Nothing
    'リストフォーマッターオブジェクト変数を解放
    Set ListFormatter = Nothing

    '変数オブジェクトを解放する
    Set Variables = Nothing
    'プロシージャを正常終了する
    Exit Sub

'エラー発生時の処理を行う
HandleErr:
    'エラーメッセージを表示する
    MsgBox lngProcessedCount & "番目の段落付近でエラーが発生したため、処理を停止しました。", vbOKOnly + vbExclamation, "VA公用文"
    'CleanUpセクションにジャンプして後処理を実行
    Resume Cleanup
'インデント処理プロシージャを終了する
End Sub

