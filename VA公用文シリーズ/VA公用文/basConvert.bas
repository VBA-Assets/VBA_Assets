Option Explicit
'変数管理クラスを格納する変数を宣言する
Dim Variables As clsVariablesManager

'==================================================================================================
' 概要:          現在開いているWord文書のスタイル(見出し、リスト)を解析し、
'               Markdown形式のプレーンテキストに変換します。
' 引数:          blnDummy (Boolean, Optional):
'               Wordの「マクロ」ダイアログボックスにこのマクロが表示されるのを
'               防ぐためのダミー引数です。
'==================================================================================================
Sub OutlineToMarkdown(Optional blnDummy As Boolean)
    '処理時間を計測する
    'Dim dblStart As Double
    'dblStart = Timer

    'エラーのトラップを有効にする
    On Error GoTo HandleErr
    
    '変数オブジェクトを生成する
    Set Variables = New clsVariablesManager
    
    '進捗状況を表示する
    Application.StatusBar = "マークダウンへの変換を準備しています。"

    '処理用の変数を定義する
    Dim blnAfterHeading As Boolean ' 見出しフラグ
    Dim blnAfterAlignRight As Boolean ' 右揃えフラグ
    
    Dim sec As Section ' セクションオブジェクト
    Dim pars As Paragraphs ' 段落コレクション
    Dim par As Paragraph ' 段落オブジェクト
    
    Dim lngTotalParagraphs As Long ' 総段落数
    Dim lngProcessedCount As Long ' 処理済み段落数
    
    Dim strStyle As String ' スタイル名
    Dim strText As String ' 段落内テキスト
    
    Dim strHeads As String ' マークダウンの見出し用記号セット
    Dim x As Long ' 見出しレベル
    
    '段落数を把握する
    lngTotalParagraphs = ActiveDocument.Paragraphs.count
    lngProcessedCount = 0
    
    'すべてのセクションを繰り返す
    For Each sec In ActiveDocument.Sections
        
        '見出しフラグを解除する
        blnAfterHeading = False
        '右揃えフラグを解除する
        blnAfterAlignRight = False
        
        Set pars = sec.Range.Paragraphs
        
        'すべての段落を繰り返す
        For Each par In pars
            '段落番号をカウントする
            lngProcessedCount = lngProcessedCount + 1
            
            '50段落ごとに処理
            If lngProcessedCount Mod 50 = 0 Then
                '進捗率を表示する
                Application.StatusBar = "マークダウンを設定しています。　" & _
                    Int(lngProcessedCount / lngTotalParagraphs * 100) & "%完了"
            End If
            
            'スタイル名を把握する
            strStyle = par.Style.NameLocal
            '文字列を取得する
            strText = par.Range.Text
            
            '右揃えの場合はスキップする
            If blnAfterAlignRight Then GoTo Continue
            
            ' 最初の見出しまたはリスト済みでない場合のみ処理
            If Not blnAfterHeading Then
                
                ' スタイルが「見出し」または「リスト段落」の場合
                If (Left(strStyle, 3) = "見出し") Or (strStyle = "リスト段落") Then
                    blnAfterHeading = True
                    
                ' マークダウンが「見出し」、「箇条書き」または「段落番号」の場合
                ElseIf (Left(strText, 1) = strHead) Or _
                       (Left(strText, 1) = strList) Or _
                       (GetNumListLength(strText) > 0) Then
                    blnAfterHeading = True
                End If
            End If
            
            '最初の見出しまたはリスト以前の場合
            If blnAfterHeading = False Then
                '「最初の見出し、箇条書きまたは段落番号までの処理を停止する」の定数がTrueの場合はスキップする
                If blnSkipUntil = True Then GoTo Continue
            '最初の見出しまたはリスト以降の場合
            Else
                '「配置が右揃えで処理を停止する」の定数がtrueの場合
                If blnStopAfter = True Then
                    'スタイルが「行間詰め」以外で配置が右揃えの場合
                    If strStyle <> "行間詰め" And par.Format.alignment = wdAlignParagraphRight Then
                        '右揃えフラグを設定する
                        blnAfterAlignRight = True
                        GoTo Continue
                    End If
                End If
            End If
            
            '各段落を対象に処理
            With par
                'スタイルが「見出し」の場合
                If Left(strStyle, 3) = "見出し" Then
                    'マークダウンの記号セットをリセットする
                    strHeads = ""
                    '見出しレベルに応じてマークダウンを追加する
                    x = Val(Right$(strStyle, 1))
                    If x > 0 Then
                          strHeads = String(x, strHead)
                    End If
                    
                    '標準スタイルを適用する
                    '標準スタイルを適用することで、インデントも解除されます。
                    .Style = "標準"
                    
                    '記号後の空白が定数で指定されている場合は空白を追加する
                    If blnMarkdownSpace = True Then strHeads = strHeads + " "
                    
                    '記号セットを行頭に挿入する
                    .Range.InsertBefore strHeads
                    
                'スタイルが「リスト段落」の場合
                ElseIf strStyle = "リスト段落" Then
                    '番号書式に応じて記号を行頭に加える
                    Select Case .Range.ListFormat.ListType
                        Case wdListBullet '「・」の場合
                            '記号後に空白を入れる場合
                            If blnMarkdownSpace = True Then
                                .Range.InsertBefore strList + " "
                            '記号後に空白を入れない場合
                            Else
                                .Range.InsertBefore strList
                            End If
                            
                        Case wdListSimpleNumbering '「①」などの場合
                            '記号後に空白を入れる場合
                            If blnMarkdownSpace = True Then
                                .Range.InsertBefore strNumList + " "
                            '記号後に空白を入れない場合
                            Else
                                .Range.InsertBefore strNumList
                            End If
                    End Select
                    
                    '標準スタイルを適用する
                    '標準スタイルを適用することで、インデントも解除されます。
                    .Style = "標準"
                    
                'スタイルが「標準」の場合
                ElseIf strStyle = "標準" Then
                    '標準スタイルを適用する
                    'インデントを解除するため、改めて適用しなおしています。
                    .Style = "標準"
                End If
            End With
            
Continue:
        Next par
    Next sec

Cleanup:
    '進捗状況を解除する
    Application.StatusBar = ""
    
    'オブジェクト変数を解放する
    Set par = Nothing
    Set pars = Nothing
    Set sec = Nothing
    
    '変数オブジェクトを解放する
    Set Variables = Nothing
    Exit Sub

'エラー発生時の処理を行う
HandleErr:
    MsgBox lngProcessedCount & "番目の段落付でエラーが発生したため、処理を停止しました。", vbOKOnly + vbExclamation, "VA公用文"
    Resume Cleanup
End Sub

'==================================================================================================
' 概要:          現在開いているWord文書内のMarkdown形式のテキスト(見出し、リスト)を解析し、
'               対応するWordのスタイルや書式に変換します。
' 引数:          blnDummy (Boolean, Optional):
'               Wordの「マクロ」ダイアログボックスにこのマクロが表示されるのを
'               防ぐためのダミー引数です。
'==================================================================================================
Sub MarkdownToOutline(Optional blnDummy As Boolean)
    '処理時間を計測する
    'Dim dblStart As Double
    'dblStart = Timer
    
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    '変数オブジェクトを生成する
    Set Variables = New clsVariablesManager
    
    'リストフォーマッターの準備
    Dim ListFormatter As clsListFormatter
    Set ListFormatter = New clsListFormatter
    
    '進捗状況を表示する
    Application.StatusBar = "アウトラインの設定を準備しています。"

    '処理用の変数を定義する
    Dim blnAfterHeading As Boolean ' 見出しフラグ
    Dim blnAfterAlignRight As Boolean ' 右揃えフラグ
    Dim blnOutlineSet As Boolean ' アウトライン設定済みフラグ
    
    Dim sec As Section ' セクションオブジェクト
    Dim pars As Paragraphs ' 段落コレクション
    Dim par As Paragraph ' 段落オブジェクト
    
    Dim lngTotalParagraphs As Long ' 総段落数
    Dim lngProcessedCount As Long ' 処理済み段落数
    
    Dim strStyle As String ' スタイル名
    Dim strText As String ' 段落内テキスト
    Dim x As Long ' 見出しレベル
    Dim lngNumListLength As Long ' 段落番号の記号の文字数
    Dim rngDelete As Range ' 削除用範囲オブジェクト
    
    '段落数を把握する
    lngTotalParagraphs = ActiveDocument.Paragraphs.count
    lngProcessedCount = 0
    
    'すべてのセクションを繰り返す
    For Each sec In ActiveDocument.Sections
        
        '見出しフラグを解除する
        blnAfterHeading = False
        '右揃えフラグを解除する
        blnAfterAlignRight = False
        'アウトライン設定済みフラグを解除する
        blnOutlineSet = False
        
        Set pars = sec.Range.Paragraphs
        
        'すべての段落を繰り返す
        For Each par In pars
            '段落番号をカウントする
            lngProcessedCount = lngProcessedCount + 1
            
            '50段落ごとに処理
            If lngProcessedCount Mod 50 = 0 Then
                '進捗率を表示する
                Application.StatusBar = "アウトラインを設定しています。　" & _
                                            Int(lngProcessedCount / lngTotalParagraphs * 100) & "%完了"
            End If
            
            '文字列を取得する
            strText = par.Range.Text
            'スタイル名を把握する
            strStyle = par.Style.NameLocal
            
            '目次スタイルはスキップする
            If Left(strStyle, 2) = "目次" Then GoTo Continue
            
            '右揃えの場合はスキップする
            If blnAfterAlignRight Then GoTo Continue
            
            '最初の見出しまたはリスト以前の場合
            If blnAfterHeading = False Then
                '文字列が見出しまたはリストの場合
                If Left(strText, 1) = strHead Or _
                   Left(strText, 1) = strList Or _
                   GetNumListLength(strText) > 0 Then
                    '見出しフラグを設定する
                    blnAfterHeading = True
                '見出しまたはリストの場合
                ElseIf Left(strStyle, 3) = "見出し" Or strStyle = "リスト段落" Then
                    '見出しフラグを設定する
                    blnAfterHeading = True
                End If
            End If
            
            '最初の見出しまたはリスト以前の場合
            If blnAfterHeading = False Then
                '「最初の見出し、箇条書きまたは段落番号までの処理を停止する」の定数がTrueの場合はスキップする
                If blnSkipUntil = True Then GoTo Continue
            '最初の見出しまたはリスト以降の場合
            Else
                '「配置が右揃えで処理を停止する」の定数がtrueの場合
                If blnStopAfter = True Then
                    'スタイルが「行間詰め」以外で配置が右揃えの場合
                    If strStyle <> "行間詰め" And par.Format.alignment = wdAlignParagraphRight Then
                        '右揃えフラグを設定する
                        blnAfterAlignRight = True
                        GoTo Continue
                    End If
                End If
            End If
            
            '段落番号の記号の文字数を把握する（段落番号でない場合は0）
            lngNumListLength = GetNumListLength(strText)
            
            '--- 各段落を対象に処理 ---
            With par
                '文字列が見出しの場合
                If Left(strText, 1) = strHead Then
                    
                    'マークダウン記号が9個から1個までチェック
                    For x = 9 To 1 Step -1
                        '該当する段落の場合
                        If Left(strText, x) = String(x, strHead) Then
                            
                            'マークダウン記号を削除する
                            Set rngDelete = .Range
                            rngDelete.End = rngDelete.Start + x
                            rngDelete.Delete
                            
                            'スタイルを適用する
                            'アウトラインよりも先に適用しないと番号の振り直しが解除されます。
                            .Style = "見出し " & x
                            
                            '繰り返しを終了する
                            Exit For
                        End If
                    Next x
                    
                    'アウトラインを設定していない場合
                    If blnOutlineSet = False Then
                        'アウトラインを適用する
                        ListFormatter.ConfigOutline par
                        'アウトライン設定済みフラグを設定する
                        blnOutlineSet = True
                    End If
                    
                    '1文字目(マークダウン記号の直後)が空白の場合は削除する
                    If Left(.Range.Text, 1) = " " Then
                        Set rngDelete = .Range
                        rngDelete.End = rngDelete.Start + 1
                        rngDelete.Delete
                    End If
                    
                '文字列が箇条書きの場合
                ElseIf Left(strText, 1) = strList Then
                    
                    '最初の1文字を削除する
                    Set rngDelete = .Range
                    rngDelete.End = rngDelete.Start + 1
                    rngDelete.Delete
                    
                    '箇条書きを設定する
                    'スタイルが段落番号と共通なので直接リストを設定しています。
                    ListFormatter.ConfigList wdBulletGallery, par
                    
                    '1文字目(マークダウン記号の直後)が空白の場合は削除する
                    If Left(.Range.Text, 1) = " " Then
                        Set rngDelete = .Range
                        rngDelete.End = rngDelete.Start + 1
                        rngDelete.Delete
                    End If
                    
                '文字列が段落番号の場合
                ElseIf lngNumListLength > 0 Then
                    
                    '段落番号の記号（数字の桁数＋1文字）を削除する
                    Set rngDelete = .Range
                    rngDelete.End = rngDelete.Start + lngNumListLength
                    rngDelete.Delete
                    
                    '段落番号を設定する
                    'スタイルが箇条書きと共通なので直接リストを設定しています。
                    ListFormatter.ConfigList wdNumberGallery, par
                    
                    '1文字目(マークダウン記号の直後)が空白の場合は削除する
                    If Left(.Range.Text, 1) = " " Then
                        Set rngDelete = .Range
                        rngDelete.End = rngDelete.Start + 1
                        rngDelete.Delete
                    End If
                    
                End If
            End With
Continue:
        Next par
    Next sec

Cleanup:
    '進捗状況を解除する
    Application.StatusBar = ""
    
    'オブジェクト変数を解放する
    Set rngDelete = Nothing
    Set par = Nothing
    Set pars = Nothing
    Set sec = Nothing
    
    '変数オブジェクトを解放する
    Set Variables = Nothing
    Set ListFormatter = Nothing
    Exit Sub

HandleErr:
    MsgBox lngProcessedCount & "番目の段落付でエラーが発生したため、処理を停止しました。", vbOKOnly + vbExclamation, "VA公用文"
    Resume Cleanup
End Sub

'==================================================================================================
' 概要:          段落の先頭が段落番号のマークダウン（行頭から連続する半角数字1桁以上＋"."）
'               かどうかを判定し、記号部分の文字数を返します。
'               "10." 以降の番号も認識します。"1)"、"(1)"、"①"、"1、" は対象外です。
' 引数:          strText (String): 判定する段落のテキスト
' 戻り値:        Long: 段落番号の場合は「数字の桁数＋1」、段落番号でない場合は 0
'==================================================================================================
Public Function GetNumListLength(ByVal strText As String) As Long
    Dim lngPos As Long ' 文字位置
    Dim lngCode As Long ' 文字コード
    
    '戻り値を初期化する
    GetNumListLength = 0
    
    '行頭から連続する半角数字の位置を進める
    lngPos = 1
    Do While lngPos <= Len(strText)
        lngCode = AscW(Mid$(strText, lngPos, 1))
        '半角数字でない場合は終了する
        If lngCode < 48 Or lngCode > 57 Then Exit Do
        lngPos = lngPos + 1
    Loop
    
    '数字が1桁以上あり、直後が "." の場合は段落番号とする
    If lngPos > 1 And Mid$(strText, lngPos, 1) = "." Then
        GetNumListLength = lngPos
    End If
End Function
