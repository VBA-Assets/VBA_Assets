Option Explicit
'================================================================================
' クラスモジュール: clsPreFormatter
' 機能: 文書の書式整形および外部ツール由来要素の検知・整形を行う
'================================================================================


'================================================================================
' 機能:         文書内のすべての表の段落に指定したスタイルを適用します。
'================================================================================
Public Sub ApplyTableStyle(ByVal targetStyleName As String, Optional ByVal targetDoc As Document = Nothing)
    If targetDoc Is Nothing Then Set targetDoc = ActiveDocument
    On Error GoTo HandleErr
    
    If targetDoc.Tables.count = 0 Then Exit Sub
    
    Dim isScreenUpdating As Boolean
    isScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Dim tbl As Table
    Dim pars As Paragraphs
    Dim par As Paragraph
    Dim fntBefore As Font
    Dim parAlign As WdParagraphAlignment
    
    For Each tbl In targetDoc.Tables
        Set pars = tbl.Range.Paragraphs
        For Each par In pars
            If par.Style.NameLocal <> targetStyleName Then
                Set fntBefore = par.Range.Font.Duplicate
                parAlign = par.alignment
                
                On Error Resume Next
                par.Style = targetStyleName
                
                If Err.Number = 0 Then
                    par.Range.Font = fntBefore
                    par.alignment = parAlign
                End If
                On Error GoTo HandleErr
                Set fntBefore = Nothing
            End If
        Next par
    Next tbl

Cleanup:
    If isScreenUpdating Then Application.ScreenUpdating = True
    Set par = Nothing
    Set pars = Nothing
    Set tbl = Nothing
    Exit Sub

HandleErr:
    Resume Cleanup
End Sub

'================================================================================
' 機能:         文書内の全段落を対象に、行頭の不要なスペースをインデントに自動変換します。
'================================================================================
Public Sub ReplaceSpaceWithIndent(ByVal sngAdvWidth As Single, ByVal excludeStyleName As String, Optional ByVal targetDoc As Document = Nothing)
    If targetDoc Is Nothing Then Set targetDoc = ActiveDocument
    On Error GoTo HandleErr

    Dim isScreenUpdating As Boolean
    isScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    '行頭の半角空白を Range.Delete で削除するため、呼び出し元（basControl.BeginProcess）で
    'スマート カット アンド ペーストを無効にしておくこと（有効なままだと削除されないことがある）
    Dim pars As Paragraphs
    Dim par As Paragraph
    Dim sngSpaceCnt As Single
    Dim lngLenBefore As Long ' 空白を削除する前の段落の文字数
    Set pars = targetDoc.Paragraphs

    For Each par In pars
        If Len(par.Range.Text) > 1 Then
            If par.Style.NameLocal = excludeStyleName Then GoTo Continue

            With par
                If sngAdvWidth > 0 Then
                    sngSpaceCnt = .FirstLineIndent / sngAdvWidth
                Else
                    sngSpaceCnt = 0
                End If

                Do Until .Range.Characters(1) <> "　" And .Range.Characters(1) <> " "
                    lngLenBefore = Len(.Range.Text)
                    If .Range.Characters(1) = "　" Then
                        sngSpaceCnt = sngSpaceCnt + 1
                        .Range.Characters(1).Delete
                    ElseIf .Range.Characters(1) = " " Then
                        sngSpaceCnt = sngSpaceCnt + 0.5
                        .Range.Characters(1).Delete
                    End If
                    '空白を削除できなかった場合は、無限ループを避けるため繰り返しを終了する
                    If Len(.Range.Text) >= lngLenBefore Then Exit Do
                Loop
                .FirstLineIndent = sngSpaceCnt * sngAdvWidth
            End With
        End If
Continue:
    Next par

Cleanup:
    If isScreenUpdating Then Application.ScreenUpdating = True
    Set par = Nothing
    Set pars = Nothing
    Exit Sub

HandleErr:
    Resume Cleanup
End Sub

'================================================================================
' 機能:         外部ツール由来の要素（手動改行、リンク、WikiLink、フロントマター、水平線、強調記号）を検知する
' 戻り値:       Boolean - 検知された場合は True
'================================================================================
Public Function DetectExternalElements(Optional ByVal targetDoc As Document = Nothing) As Boolean
    '引数が省略された場合はアクティブなドキュメントを対象とする
    If targetDoc Is Nothing Then Set targetDoc = ActiveDocument
    '検知結果を保持する変数を宣言する
    Dim blnDetected As Boolean
    '初期値をFalseに設定する
    blnDetected = False
    
    '1. 手動改行 (^l) の検知
    '対象ドキュメントの検索オブジェクトを操作する
    With targetDoc.Range.Find
        '前回の検索フォーマット設定をクリアする
        .ClearFormatting
        '検索文字列として手動改行(^l)を指定する
        .Text = "^l"
        '文書の末尾に向かって検索を進める
        .Forward = True
        '見つからなければ検索を終了する
        .Wrap = wdFindStop
        '検索を実行し、一致する箇所があった場合
        If .Execute Then
            '検知フラグをTrueに設定する
            blnDetected = True
            '以降のチェックは不要なため、終了処理へジャンプする
            GoTo EndDetection
        End If
    End With
    
    '2. ハイパーリンクの検知
    'ドキュメント内に設定されているハイパーリンクの数を調べる
    If targetDoc.Hyperlinks.count > 0 Then
        '1つ以上存在する場合は検知フラグをTrueに設定する
        blnDetected = True
        '以降のチェックは不要なため、終了処理へジャンプする
        GoTo EndDetection
    End If
    
    '3. WikiLink ([[]) の検知
    '対象ドキュメントの検索オブジェクトを操作する
    With targetDoc.Range.Find
        '前回の検索フォーマット設定をクリアする
        .ClearFormatting
        '検索文字列としてWikiLinkの開始記号([[[)を指定する
        .Text = "[["
        'ワイルドカードを使用しない
        .MatchWildcards = False
        '文書の末尾に向かって検索を進める
        .Forward = True
        '見つからなければ検索を終了する
        .Wrap = wdFindStop
        '検索を実行し、一致する箇所があった場合
        If .Execute Then
            '検知フラグをTrueに設定する
            blnDetected = True
            '以降のチェックは不要なため、終了処理へジャンプする
            GoTo EndDetection
        End If
    End With
    
    '4. フロントマター・水平線 (---、***、___ の段落) の検知
    'フロントマターの開始・終了記号(---)も水平線と同じ形なので、まとめて検知する
    Dim par As Paragraph
    'すべての段落を繰り返す
    For Each par In targetDoc.Paragraphs
        '段落が水平線の形である場合
        If IsHorizontalRule(GetPlainText(par)) Then
            '表内の段落は対象外とする
            If Not par.Range.Information(wdWithInTable) Then
                '検知フラグをTrueに設定する
                blnDetected = True
                '以降のチェックは不要なため、繰り返しを終了する
                Exit For
            End If
        End If
    Next par
    '検知済みの場合は、以降のチェックは不要なため、終了処理へジャンプする
    If blnDetected Then GoTo EndDetection

    '5. 強調記号 (**、__、*、_ の対) の検知
    If ProcessEmphasis(targetDoc, False) Then blnDetected = True

EndDetection:
    '検知結果（TrueまたはFalse）を関数の戻り値として返す
    DetectExternalElements = blnDetected
End Function

'================================================================================
' 機能:         外部ツール由来の特殊要素を一括整形する
'================================================================================
Public Sub ExecuteSimpleCleanup(Optional ByVal targetDoc As Document = Nothing)
    '引数が省略された場合はアクティブなドキュメントを対象とする
    If targetDoc Is Nothing Then Set targetDoc = ActiveDocument
    
    '現在の画面更新状態を保存するための変数を宣言する
    Dim isScreenUpdating As Boolean
    '現在の画面更新状態を取得する
    isScreenUpdating = Application.ScreenUpdating
    '処理速度向上のため、画面更新を一時的に停止する
    Application.ScreenUpdating = False
    
    'エラー発生時に終了処理ブロックへジャンプするように設定する
    On Error GoTo CleanupError

    '1. 手動改行 (^l) を 段落記号 (^p) に置換
    '対象ドキュメント全体の検索・置換オブジェクトを操作する
    With targetDoc.Range.Find
        '検索文字列のフォーマットをクリアする
        .ClearFormatting
        '置換後のフォーマットをクリアする
        .Replacement.ClearFormatting
        '検索対象として手動改行を指定する
        .Text = "^l"
        '置換対象として段落記号を指定する
        .Replacement.Text = "^p"
        '文書の末尾に向かって検索を進める
        .Forward = True
        '文書の末尾に達しても全体をくまなく検索する
        .Wrap = wdFindContinue
        '書式に基づく検索は行わない
        .Format = False
        '大文字と小文字を区別しない
        .MatchCase = False
        '完全に一致する単語のみを検索しない
        .MatchWholeWord = False
        '半角と全角を区別しない
        .MatchByte = False
        '英単語の異なる活用形を区別しない
        .MatchAllWordForms = False
        '発音が似ている単語を検索しない
        .MatchSoundsLike = False
        'ワイルドカードを使用しない
        .MatchWildcards = False
        '条件に一致するすべてを一括で置換する
        .Execute Replace:=wdReplaceAll
    End With
    
    '2. ハイパーリンクを解除（表示テキストは保持）
    'ドキュメント内にハイパーリンクが存在する限りループを続ける
    Do While targetDoc.Hyperlinks.count > 0
        'コレクションの先頭にあるハイパーリンクを削除（リンク解除）する
        targetDoc.Hyperlinks(1).Delete
    Loop
    
    '3. WikiLink ([[テキスト]] -> テキスト, [[A|B]] -> B) の置換
    '対象ドキュメントの検索オブジェクトを操作する
    With targetDoc.Range.Find
        '検索文字列のフォーマットをクリアする
        .ClearFormatting
        '検索文字列としてWikiLinkの開始記号([[[)を指定する
        .Text = "[["
        'ワイルドカードを使用しない
        .MatchWildcards = False
        '文書の末尾に向かって検索を進める
        .Forward = True
        '見つからなければ検索を終了する
        .Wrap = wdFindStop
        '検索を実行し、一致箇所が見つかる間ループを繰り返す
        Do While .Execute
            '見つかった箇所（範囲）を操作するための変数を宣言する
            Dim rngMatch As Range
            '検索結果の範囲オブジェクトを複製して割り当てる
            Set rngMatch = .Parent.Duplicate
            '閉じ記号(])が見つかるまで範囲の末尾を広げる
            rngMatch.MoveEndUntil Cset:="]", count:=wdForward
            'さらに閉じ記号(]])を含めるため、範囲の末尾を2文字進める
            rngMatch.MoveEnd wdCharacter, 2
            
            '取得した範囲の末尾2文字が正しく「]]」であるか確認する
            If Right(rngMatch.Text, 2) = "]]" Then
                'リンク内部のテキストを保持する変数を宣言する
                Dim innerText As String
                '先頭の「[[」と末尾の「]]」を除いた文字列を抽出する
                innerText = Mid(rngMatch.Text, 3, Len(rngMatch.Text) - 4)
                
                'エイリアス区切りのパイプ(|)の位置を保持する変数を宣言する
                Dim pipePos As Long
                '内部テキストからパイプ(|)の位置を検索する
                pipePos = InStr(innerText, "|")
                'パイプ(|)が存在する場合（エイリアス指定がある場合）
                If pipePos > 0 Then
                    'パイプ以降のテキスト（表示名）で元のWikiLink全体を置換する
                    rngMatch.Text = Mid(innerText, pipePos + 1)
                'パイプ(|)が存在しない場合
                Else
                    '内部テキストそのもので元のWikiLink全体を置換する
                    rngMatch.Text = innerText
                End If
                '次の検索が正しく行われるよう、選択範囲を末尾に畳む
                .Parent.Collapse wdCollapseEnd
            '末尾が「]]」でない（書式が壊れている）場合
            Else
                '何も置換せず、選択範囲を末尾に畳んで次の検索へ進む
                .Parent.Collapse wdCollapseEnd
            End If
        Loop
    End With
    
    '4. フロントマター (先頭の --- で囲まれたブロック) の削除
    'フロントマターの終了記号(---)の段落番号を取得する（フロントマターでない場合は0）
    Dim lngEndPara As Long
    lngEndPara = GetFrontMatterEnd(targetDoc)
    'フロントマターである場合
    If lngEndPara > 0 Then
        '削除対象となる範囲を保持する変数を宣言する
        Dim rngDel As Range
        '1段落目の先頭から終了段落の末尾までの範囲を設定する
        Set rngDel = targetDoc.Range(targetDoc.Paragraphs(1).Range.Start, targetDoc.Paragraphs(lngEndPara).Range.End)
        '設定したフロントマター範囲全体を削除する
        rngDel.Delete
    End If

    '5. 水平線 (---、***、___ の段落) の削除
    'フロントマターと判定されなかった先頭の --- も、ここで水平線として削除する
    '削除対象の段落の範囲を格納するコレクションを用意する
    Dim colRules As Collection
    Set colRules = New Collection
    Dim par As Paragraph
    'すべての段落を繰り返す
    For Each par In targetDoc.Paragraphs
        '段落が水平線の形である場合
        If IsHorizontalRule(GetPlainText(par)) Then
            '表内の段落は対象外とし、それ以外の段落の範囲を削除対象に加える
            If Not par.Range.Information(wdWithInTable) Then colRules.Add par.Range
        End If
    Next par
    '段落番号がずれないよう、後ろから順に削除する
    Dim lngIndex As Long
    For lngIndex = colRules.count To 1 Step -1
        colRules(lngIndex).Delete
    Next lngIndex

    '6. 強調記号 (**、__、*、_ の対) の除去
    '記号だけを削除し、中の文字列は残す（太字などの書式には変換しない）
    Call ProcessEmphasis(targetDoc, True)

CleanupError:
    '処理開始前に画面更新が有効だった場合は、再度有効に戻す
    If isScreenUpdating Then Application.ScreenUpdating = True
End Sub

'================================================================================
' 機能:         文書の先頭がフロントマターかどうかを判定する
'               次の両方を満たす場合にフロントマターとみなす
'               ・1段落目が「---」
'               ・2段落目から数えて10段落以内に終了記号「---」があり、その間の段落が
'                 すべて「key: value」形式（半角コロンを含む）または空行
'               満たさない場合、先頭の「---」は水平線として扱う
' 戻り値:       Long - フロントマターの終了記号の段落番号（フロントマターでない場合は 0）
'================================================================================
Private Function GetFrontMatterEnd(ByVal targetDoc As Document) As Long
    Dim lngLast As Long ' 終了記号を探す最後の段落番号
    Dim i As Long ' 段落番号
    Dim strText As String ' 段落のテキスト

    '戻り値を初期化する
    GetFrontMatterEnd = 0

    '段落が2つ未満の場合、または1段落目が「---」でない場合は終了する
    If targetDoc.Paragraphs.count < 2 Then Exit Function
    If GetPlainText(targetDoc.Paragraphs(1)) <> "---" Then Exit Function

    '終了記号を探す範囲を、2段落目から数えて10段落以内に限定する
    lngLast = targetDoc.Paragraphs.count
    If lngLast > 11 Then lngLast = 11

    '2段落目から順に確認する
    For i = 2 To lngLast
        strText = GetPlainText(targetDoc.Paragraphs(i))
        '終了記号が見つかった場合は、その段落番号を返す
        If strText = "---" Then
            GetFrontMatterEnd = i
            Exit Function
        '空行でも「key: value」形式でもない段落がある場合は、フロントマターではない
        ElseIf strText <> "" And InStr(strText, ":") = 0 Then
            Exit Function
        End If
    Next i
End Function

'================================================================================
' 機能:         段落が水平線かどうかを判定する
'               「-」「*」「_」のいずれか1種類の記号だけが3文字以上続く場合に水平線とみなす
'               記号の間の空白（「* * *」など）は無視する
' 引数:         strText - 段落記号・手動改行・前後の空白を除いた段落のテキスト
'================================================================================
Private Function IsHorizontalRule(ByVal strText As String) As Boolean
    Dim strMark As String ' 先頭の記号

    '戻り値を初期化する
    IsHorizontalRule = False

    '半角空白・全角空白・タブを取り除く
    strText = Replace(Replace(Replace(strText, " ", ""), "　", ""), vbTab, "")

    '3文字未満の場合は水平線ではない
    If Len(strText) < 3 Then Exit Function

    '先頭の文字が水平線の記号でない場合は水平線ではない
    strMark = Left$(strText, 1)
    If strMark <> "-" And strMark <> "*" And strMark <> "_" Then Exit Function

    'すべての文字が先頭と同じ記号の場合に水平線とする
    IsHorizontalRule = (strText = String(Len(strText), strMark))
End Function

'================================================================================
' 機能:         対になった強調記号（**x**、__x__、*x*、_x_）を検知または除去する
'               長い記号から先に処理する（** を先に除去しないと、* の対として誤って処理されるため）
' 引数:         blnRemove - True の場合は除去する、False の場合は検知のみ行う
' 戻り値:       Boolean - 対象となる強調記号が見つかった場合は True
'================================================================================
Private Function ProcessEmphasis(ByVal targetDoc As Document, ByVal blnRemove As Boolean) As Boolean
    Dim blnFound As Boolean ' 強調記号が見つかったかどうか

    blnFound = False
    If ProcessEmphasisMark(targetDoc, "\*\*([!\*^13]@)\*\*", 2, False, blnRemove) Then blnFound = True
    If ProcessEmphasisMark(targetDoc, "__([!_^13]@)__", 2, True, blnRemove) Then blnFound = True
    If ProcessEmphasisMark(targetDoc, "\*([!\*^13]@)\*", 1, False, blnRemove) Then blnFound = True
    If ProcessEmphasisMark(targetDoc, "_([!_^13]@)_", 1, True, blnRemove) Then blnFound = True
    ProcessEmphasis = blnFound
End Function

'================================================================================
' 機能:         1種類の強調記号の対を検索し、検知または除去する
'               中の文字列に段落記号を含むものは対象外とする（検索文字列で除外）
'               次の場合も対象外とする
'               ・記号の直後が半角空白の場合（行頭の箇条書き記号「* 」など）
'               ・blnCheckAlnum が True で、記号の前後が英数字の場合（file_name など）
' 引数:         strFind - 記号の対と中の文字列に一致するワイルドカードの検索文字列
'               lngMarkLen - 片側の記号の文字数（** なら 2、* なら 1）
'               blnCheckAlnum - True の場合は、記号の前後が英数字かどうかを確認する（_ 系用）
'               blnRemove - True の場合は除去する、False の場合は検知のみ行う
' 戻り値:       Boolean - 対象となる強調記号が見つかった場合は True
'================================================================================
Private Function ProcessEmphasisMark(ByVal targetDoc As Document, ByVal strFind As String, _
        ByVal lngMarkLen As Long, ByVal blnCheckAlnum As Boolean, ByVal blnRemove As Boolean) As Boolean
    Dim rngFind As Range ' 検索範囲（検索後は一致した範囲）
    Dim strText As String ' 一致した文字列
    Dim lngStart As Long ' 一致した範囲の開始位置
    Dim blnSkip As Boolean ' 対象外とするかどうか

    '戻り値を初期化する
    ProcessEmphasisMark = False

    '文書全体を検索範囲とする
    Set rngFind = targetDoc.Content
    With rngFind.Find
        .ClearFormatting
        .Text = strFind
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        '一致する箇所が見つかる間、繰り返す
        Do While .Execute
            strText = rngFind.Text
            lngStart = rngFind.Start

            '記号の直後が半角空白の場合は対象外とする
            blnSkip = (Mid$(strText, lngMarkLen + 1, 1) = " ")
            '記号の前後が英数字の場合は対象外とする
            If Not blnSkip And blnCheckAlnum Then
                blnSkip = IsAlnumAt(targetDoc, lngStart - 1) Or IsAlnumAt(targetDoc, rngFind.End)
            End If

            '対象外の場合
            If blnSkip Then
                '1文字先から検索を続ける
                rngFind.SetRange lngStart + 1, targetDoc.Content.End
            '対象の場合
            Else
                ProcessEmphasisMark = True
                '検知のみの場合は終了する
                If Not blnRemove Then Exit Function
                '記号を除き、中の文字列だけにする
                rngFind.Text = Mid$(strText, lngMarkLen + 1, Len(strText) - lngMarkLen * 2)
                '置き換えた文字列の後から検索を続ける
                rngFind.SetRange rngFind.End, targetDoc.Content.End
            End If
        Loop
    End With
End Function

'================================================================================
' 機能:         文書の指定位置の文字が半角英数字かどうかを判定する
'               位置が文書の範囲外の場合は False を返す
'================================================================================
Private Function IsAlnumAt(ByVal targetDoc As Document, ByVal lngPos As Long) As Boolean
    Dim lngCode As Long ' 文字コード

    '戻り値を初期化する
    IsAlnumAt = False

    '位置が文書の範囲外の場合は終了する
    If lngPos < targetDoc.Content.Start Or lngPos >= targetDoc.Content.End Then Exit Function

    '半角数字、半角英大文字、半角英小文字のいずれかの場合に True とする
    lngCode = AscW(targetDoc.Range(lngPos, lngPos + 1).Text)
    IsAlnumAt = (lngCode >= 48 And lngCode <= 57) Or _
                (lngCode >= 65 And lngCode <= 90) Or _
                (lngCode >= 97 And lngCode <= 122)
End Function

'================================================================================
' 機能:         段落のテキストから、段落記号・手動改行・前後の半角空白を除いた文字列を返す
'================================================================================
Private Function GetPlainText(ByVal par As Paragraph) As String
    Dim strText As String ' 段落のテキスト

    '末尾の段落記号(vbCr)と手動改行(Chr(11))を削除する
    strText = Replace(par.Range.Text, vbCr, "")
    strText = Replace(strText, Chr(11), "")

    '前後の空白を取り除く
    GetPlainText = Trim(strText)
End Function
