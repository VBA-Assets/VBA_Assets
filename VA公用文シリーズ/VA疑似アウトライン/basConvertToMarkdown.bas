Option Explicit

'================================================================================
' モジュールレベル変数
'================================================================================
' 設定値を管理するクラス等のインスタンスを保持する変数
Private Variables As VariablesManager

'================================================================================
' ConvertToMarkdown（メインプロシージャ）
' 機能: 現在開いているWord文書の疑似アウトライン（インデントや特定の記号で
'       表現された階層構造）を解析し、マークダウン形式のテキストに変換します。
'       同時に、本文のフォントや段落書式を標準スタイルにリセットします。
' 引数: blnDammy (Boolean, Optional):
'       Wordの「マクロ」ダイアログボックスにこのマクロが表示されるのを
'       防ぐためのダミー引数です。
'================================================================================
Sub ConvertToMarkdown(Optional blnDammy As Boolean)
    ' 変数管理クラス（VariablesManager）の新しいインスタンスを生成する
    Set Variables = New VariablesManager

    ' 確認メッセージに表示する、元に戻せるかどうかの案内を設定する
    Dim strUndoMessage As String
    If blnUseUndoRecord Then
        ' 元に戻す（Undo）の記録を行う場合
        strUndoMessage = "Ctrl+Zで元に戻せます。"
    Else
        ' 元に戻す（Undo）の記録を行わない場合（Mac版の既定）
        strUndoMessage = "元に戻せない可能性があります。"
    End If

    ' ユーザーに処理の開始を確認する（キャンセルされた場合は処理を抜ける）
    If MsgBox("マークダウンへの変換およびフォント・段落書式のリセットを開始します。" & vbCr & _
        strUndoMessage & vbCr & vbCr & _
        "よろしいですか？", vbOKCancel, "マークダウンへの変換") = vbCancel Then Exit Sub

    ' ステータスバーに進捗状況を表示する（ユーザーへのフィードバック）
    Application.StatusBar = "マークダウンへの変換を準備しています。"

    ' 実行時エラーが発生した場合に、ラベル HandleErr へジャンプするよう設定
    On Error GoTo HandleErr

    ' 元に戻す（Undo）の記録を開始し、変換全体を Ctrl+Z 1回で戻せるようにする
    Dim blnUndoRecordStarted As Boolean
    If blnUseUndoRecord Then
        ' Mac版で未対応の場合に備え、エラーは無視して処理を続行する
        On Error Resume Next
        Application.UndoRecord.StartCustomRecord "VA疑似アウトライン：マークダウン"
        ' 記録を開始できたかどうかを記憶する（終了処理で記録を閉じるため）
        blnUndoRecordStarted = (Err.Number = 0)
        Err.Clear
        ' ラベル HandleErr へジャンプする設定に戻す
        On Error GoTo HandleErr
    End If

    ' 処理速度向上のため、画面の描画更新を停止する
    Application.ScreenUpdating = False
    
    ' マウスポインタを待機状態（砂時計/回転円）にする
    Application.System.Cursor = wdCursorWait
    
    ' 設定に応じて、表内のスタイルを変更する
    If blnApplyTableStyle = True Then
        Dim objPreFormatter As clsPreFormatter
        Set objPreFormatter = New clsPreFormatter
        objPreFormatter.ApplyTableStyle "行間詰め", ActiveDocument
        Set objPreFormatter = Nothing
    End If
    
    ' ページ設定に基づき、1文字あたりの幅（アドバンス幅）を計算して変数に格納する
    ' 計算式: (用紙幅 - 左右余白) / 1行あたりの文字数
    With ActiveDocument.PageSetup
        Dim sngAdvWidth As Single
        sngAdvWidth = (.PageWidth - .LeftMargin - .RightMargin) / .CharsLine
    End With
    
    ' 設定に応じて、行頭空白をインデントに置換する
    If blnReplaceSpaceWithIndent = True Then
        Set objPreFormatter = New clsPreFormatter
        ' "行間詰め" スタイルを除外して実行
        objPreFormatter.ReplaceSpaceWithIndent sngAdvWidth, "行間詰め", ActiveDocument
        Set objPreFormatter = Nothing
    End If
    
    '--- 処理1: 複合見出しの処理を実行 ---
    Call ProcessCompositeHeadings
    
    '--- 処理2: 構造を解析してマークダウン記号へ変換する処理を実行 ---
    Call ConvertToMarkdownSyntax
    
Cleanup:
    ' 終了処理ラベル
    ' 終了処理の途中でエラーが起きても、残りの終了処理を必ず行う
    On Error Resume Next
    ' 元に戻す（Undo）の記録を開始した場合のみ、記録を終了する
    If blnUndoRecordStarted Then Application.UndoRecord.EndCustomRecord
    ' 画面の描画更新を再開する
    Application.ScreenUpdating = True
    ' ステータスバーの表示をクリアする
    Application.StatusBar = ""
    ' マウスポインタを通常の状態に戻す
    Application.System.Cursor = wdCursorNormal
    ' クラスインスタンスを解放する
    Set Variables = Nothing
    ' プロシージャを終了する
    Exit Sub

' エラー発生時の処理ラベル
HandleErr:
    ' エラーメッセージを表示する
    MsgBox "エラーが発生したため、処理を停止しました。", vbOKOnly, "VA書式修正"
    ' Cleanupラベルへ移動して終了処理を行う（Resumeによりエラー状態をクリア）
    Resume Cleanup
End Sub

'================================================================================
' 複合見出しを処理
' 機能: 複合見出しの改行処理を行います。
'================================================================================
Private Sub ProcessCompositeHeadings()
    ' ステータスバーに進捗状況を表示する
    Application.StatusBar = "複合見出しを処理しています。"
       
    Dim i As Long       ' 処理中の段落番号（カウンタ）
    Dim totalPars As Long ' 文書全体の段落総数
    
    ' 全段落数を取得する
    totalPars = ActiveDocument.Paragraphs.count
    
    Dim sec As Section  ' セクション用オブジェクト変数
    Dim par As Paragraph ' 段落用オブジェクト変数
    
    ' 文書内のすべてのセクションに対してループ処理を行う
    For Each sec In ActiveDocument.Sections
        
        ' 最初のセクション以外で特定のフラグをリセットする場合の処理（現在はコメントアウトまたはプレースホルダ）
        If i > 0 Then
            ' 必要に応じて見出し通過フラグや右揃えフラグの解除ロジックをここに記述
        End If
        
        ' 現在のセクション内のすべての段落に対してループ処理を行う
        For Each par In sec.Range.Paragraphs
            ' 段落カウンタをインクリメント
            i = i + 1
            
            ' 現在の段落に対する処理
            With par
                ' フォントの倍率を100%にリセットする（長体・平体を解除）
                .Range.Font.Scaling = 100
                
                ' 設定で行間詰めを除外しない設定かつ、スタイルが「行間詰め」の場合は処理をスキップする
                If Variables.blnNoParagraphSpace = False And .Style = "行間詰め" Then GoTo ContinueLoop
                
Repeat:
                '--- 行頭の空白文字を処理するブロック ---
                
                ' 設定「行頭空白をインデントに置換しない」場合（＝空白を削除する処理）
                ' ※置換する場合は冒頭のクラス呼び出しで処理済みのため、ここでのElseブロックは不要
                If blnReplaceSpaceWithIndent = False Then
                    ' 段落内の文字列（Characters）を操作
                    With .Range.Characters
                        Dim j As Long       ' 文字列走査用カウンタ
                        ' 1文字目から全文字数までループ（※動的に削除するため注意が必要だが、ここでは先頭文字をチェック）
                        For j = 1 To .count
                            ' 1文字目が全角スペースの場合
                            If .First = "　" Then
                                ' 1文字目を削除する
                                .First.Delete
                            ' 1文字目が半角スペースの場合
                            ElseIf .First = " " Then
                            ' Wordの仕様対策: 「(」の前の半角スペースは .Delete メソッドで消えない場合があるため、
                            ' カーソル操作とBackSpaceキー入力で強制的に削除する
                            ' 参考: https://vba-assets.net/delete-half-width-space/
                                ' 1文字目を選択する
                                .First.Select
                                
                                ' カーソルを右（文字の後ろ）に移動する
                                Selection.MoveRight
                                
                                ' BackSpaceキー入力をシミュレートして削除する
                                Selection.TypeBackspace
                            Else
                                ' 空白以外が現れたらループを抜ける（行頭の空白処理終了）
                                Exit For
                            End If
                        Next
                    End With
                End If
                
                '--- 複合見出し（"第1章" 等）や括弧書きの直前に改行を入れるか判定する処理 ---
                Dim blnReturn As Boolean ' 改行挿入フラグ
                blnReturn = False
                
                With .Range
                    ' ケース1: 「第」で始まる場合（例：第1条）
                    If .Characters(1).Text = "第" Then
                        ' 4～6文字目をチェック（数字2桁＋スペースなどを考慮）
                        For j = 4 To 6
                            If j > .Characters.count Then Exit For
                            ' 指定位置に数字がある場合、見出しパターンとみなす
                            If .Characters(j).Text Like "#" Then
                                blnReturn = True
                                Exit For
                            End If
                        Next
                    
                    ' ケース2: 数字で始まる場合（例：1 (1) ）
                    ElseIf .Characters(1).Text Like "#" Then
                        ' 2～4文字目をチェック
                        For j = 2 To 4
                            If j > .Characters.count Then Exit For
                            ' 括弧書きが続く場合、複合パターンとみなす
                            If .Characters(j).Text = "(" Or .Characters(j).Text = "（" Then
                                blnReturn = True
                                Exit For
                            End If
                        Next
                    
                    ' ケース3: カタカナ/アルファベットで始まる場合
                    ElseIf .Characters(1).Text Like "[ア-ン]" _
                        Or .Characters(1).Text Like "[ｱ-ﾝ]" _
                        Or .Characters(1).Text Like "[ａ-ｚ]" _
                        Or .Characters(1).Text Like "[a-z]" Then
                        
                        ' 2～3文字目の括弧をチェック
                        For j = 2 To 3
                            If j > .Characters.count Then Exit For
                            If .Characters(j).Text = "(" Or .Characters(j).Text = "（" Then
                                blnReturn = True
                                Exit For
                            End If
                        Next
                    
                    ' ケース4: 開き丸括弧で始まる場合（例：(1) ア ）
                    ElseIf .Characters(1).Text = "（" Or .Characters(1).Text = "(" Then
                        ' 2文字目が数字の場合
                        If .Characters(2).Text Like "#" Then
                            ' 3～4文字目をチェック
                            For j = 3 To 4
                                If j > .Characters.count Then Exit For
                                ' 閉じ括弧を見つけた場合
                                If .Characters(j).Text = "）" Or .Characters(j).Text = ")" Then
                                    ' 閉じ括弧の後ろの文字を確認（カタカナ等）
                                    Dim k As Long
                                    For k = j + 1 To j + 2
                                        ' 段落の文字数を超えたらループを抜ける
                                        ' （「(1)」「(10)」のように番号だけの段落で、段落記号の先の文字を求めてエラーになるのを防ぐ）
                                        If k > .Characters.count Then Exit For
                                        If .Characters(k).Text Like "[ア-ン]" _
                                            Or .Characters(k).Text Like "[ｱ-ﾝ]" _
                                            Or .Characters(k).Text Like "[ａ-ｚ]" _
                                            Or .Characters(k).Text Like "[a-z]" Then
                                            
                                            ' さらにその次が空白であれば、見出しの区切りとみなす
                                            If .Characters(k + 1).Text = "　" Or .Characters(k + 1).Text = " " Then
                                                blnReturn = True
                                                j = k ' ループカウンタを進める
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    Exit For
                                End If
                            Next
                        
                        ' 2文字目がカタカナ/英字の場合
                        ElseIf .Characters(2).Text Like "[ア-ン]" _
                            Or .Characters(2).Text Like "[ｱ-ﾝ]" _
                            Or .Characters(2).Text Like "[ａ-ｚ]" _
                            Or .Characters(2).Text Like "[a-z]" Then
                            
                            ' 3文字目が閉じ括弧の場合
                            If .Characters(3).Text = "）" Or .Characters(3).Text = ")" Then
                                ' 4～5文字目の英字＋閉じ括弧パターンをチェック
                                For j = 4 To 5
                                    ' 段落の文字数を超えたらループを抜ける（「(ア)」だけの段落でエラーになるのを防ぐ）
                                    ' ※この確認がループの外にあると、Exit For が段落の繰り返しそのものを抜け、以降の段落が処理されない
                                    If j > .Characters.count Then Exit For
                                    If .Characters(j).Text Like "[ａ-ｚ]" _
                                        Or .Characters(j).Text Like "[a-z]" Then
                                        If .Characters(j + 1).Text = "）" Or .Characters(j + 1).Text = ")" Then
                                            blnReturn = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If
                    
                    ' 改行フラグが立っている場合（複合見出しの分離が必要な場合）
                    If blnReturn = True Then
                        ' 該当箇所の前に全角スペースと改行コード（vbCr）を挿入して分離する
                        .Characters(j).InsertBefore "　" & vbCr
                        ' フラグを戻して、再度同じ段落のチェックを繰り返す（分離された後半部分の処理のため）
                        blnReturn = False
                        GoTo Repeat
                    End If
                End With
            End With
ContinueLoop:
        Next par ' 次の段落へ
    Next sec ' 次のセクションへ
End Sub

'================================================================================
' マークダウンに変換
' 機能: インデントや記号を解析し、マークダウン記号(#, *)への変換と標準スタイル適用を行います。
'================================================================================
Private Sub ConvertToMarkdownSyntax()
    ' ステータスバーに進捗状況を表示する
    Application.StatusBar = "マークダウンに変換しています。"
    
    ' 変数の初期化
    Dim i As Long
    i = 0
    Dim totalPars As Long
    totalPars = ActiveDocument.Paragraphs.count
    
    Dim sec As Section
    Dim par As Paragraph
    
    ' フラグ: 見出し行を通過した後かどうか
    Dim blnAfterHeadding As Boolean
    blnAfterHeadding = False
    
    ' フラグ: 右揃え行（署名など）を通過した後かどうか
    Dim blnAfterAlignRight As Boolean
    blnAfterAlignRight = False
    
    ' 全セクションのループ
    For Each sec In ActiveDocument.Sections
        
        ' 2つ目以降のセクションの場合、フラグをリセットする
        If i > 0 Then
            blnAfterHeadding = False
            blnAfterAlignRight = False
        End If

        ' 全段落のループ
        For Each par In sec.Range.Paragraphs
            ' 段落番号カウント
            i = i + 1
            
            ' 50段落ごとに進捗表示を更新
            If i Mod 50 = 0 Then
                Dim intComp As Integer
                If totalPars > 0 Then intComp = Int(i / totalPars * 100)
                Application.StatusBar = "マークダウンに変換しています。　" & intComp & "%完了"
            End If
            
            ' 各段落の処理
            With par
                ' スタイルが「行間詰め」かつ設定で除外されていない場合はスキップ
                If Variables.blnNoParagraphSpace = False And .Style = "行間詰め" Then GoTo ContinueLoop
                        
                '--- マークダウン記号（## や * など）を決定するロジック ---
                Dim strMarks As String  ' 付与するマークダウン記号
                strMarks = ""
                Dim lngDelChr As Long   ' 記号付与後に削除する元の文字数（例: "1." を削除）
                lngDelChr = 0
                
                ' Wordの「見出し」スタイルが適用されているか、またはアウトラインレベルが設定されている場合
                ' ※ wdListOutlineNumbering: アウトライン番号付きリスト
                If Left(.Style, 3) = "見出し" Or _
                    .OutlineLevel <> 10 Or _
                    .Range.ListFormat.ListType = wdListOutlineNumbering Then
                    
                    ' リストの文字（番号など）を取得する
                    Dim strListString As String
                    strListString = .Range.ListFormat.ListString
                    
                    ' 見出し文字のパターンに応じて、適切なレベル（#の数）のマークダウンを生成する
                    ' RepeatHeadingMarks関数で "#" をレベル分繰り返す
                    If Left(strListString, 1) = "第" Then                    ' 例: 第1章 -> レベル1 (#)
                        strMarks = RepeatHeadingMarks(1)
                    ElseIf Left(strListString, 1) Like "#" Then              ' 例: 1. -> レベル2 (##)
                        strMarks = RepeatHeadingMarks(2)
                    ElseIf Mid(strListString, 2, 1) Like "#" Then            ' 例: (1) -> レベル3 (###)
                        strMarks = RepeatHeadingMarks(3)
                    ElseIf Left(strListString, 1) Like "[ア-ン]" Or Left(strListString, 1) Like "[ｱ-ﾝ]" Then ' 例: ア -> レベル4
                        strMarks = RepeatHeadingMarks(4)
                    ElseIf Mid(strListString, 2, 1) Like "[ア-ン]" Or Mid(strListString, 2, 1) Like "[ｱ-ﾝ]" Then ' 例: (ア) -> レベル5
                        strMarks = RepeatHeadingMarks(5)
                    ElseIf Left(strListString, 1) Like "[ａ-ｈ]" Or Left(strListString, 1) Like "[a-h]" Then ' 例: a -> レベル6
                        strMarks = RepeatHeadingMarks(6)
                    ElseIf Mid(strListString, 2, 1) Like "[ａ-ｈ]" Or Mid(strListString, 2, 1) Like "[a-h]" Then ' 例: (a) -> レベル7
                        strMarks = RepeatHeadingMarks(7)
                    Else
                        ' 上記以外はWordのアウトラインレベルに従う
                        strMarks = RepeatHeadingMarks(.OutlineLevel)
                    End If
                    
                    ' 本来のアウトライン機能を使っている場合は、文字削除は不要（スタイル解除で消えるため）
                    lngDelChr = 0
                    blnAfterHeadding = True
                
                ' 箇条書き（ブレット）の場合
                ElseIf .Range.ListFormat.ListType = wdListBullet Then
                    strMarks = strList ' 定数（例: *）をセット
                    lngDelChr = 0
                    blnAfterHeadding = True
                
                ' 段落番号の場合
                ElseIf .Range.ListFormat.ListType = wdListSimpleNumbering Then
                    strMarks = strNumList ' 定数（例: 1.）をセット
                    lngDelChr = 0
                    blnAfterHeadding = True
                    
                ' 手入力の記号（・、- など）で始まる場合
                ElseIf .Range.Characters(1).Text = "・" Or _
                    .Range.Characters(1).Text = "･" Or _
                    .Range.Characters(1).Text = "-" Or _
                    .Range.Characters(1).Text = "－" Then
                    
                    strMarks = strList ' マークダウンのリスト記号に変換
                    lngDelChr = 1      ' 元の記号1文字を削除
                    blnAfterHeadding = True
                
                ' 丸付き数字（①など）で始まり、2文字目が空白の場合
                ElseIf .Range.Characters(1).Text Like "[①-⑨]" Then
                    If .Range.Characters(2).Text = "　" Or .Range.Characters(2).Text = " " Then
                        strMarks = strNumList
                        lngDelChr = 2
                        blnAfterHeadding = True
                    End If
                    
                ' 「第」で始まり、2文字目が数字の手入力見出しの場合
                ElseIf .Range.Characters(1).Text = "第" Then
                    If .Range.Characters(2).Text Like "#" Then
                        Dim j As Long
                        ' 後続の空白を探して見出し範囲を特定
                        For j = 2 To .Range.Characters.count
                            With .Range.Characters(j)
                                If .Text = "　" Or .Text = " " Then
                                    ' 空白までの長さが3文字以内（例: 第1章_）なら見出しと判定
                                    If j - 1 - 1 <= 3 Then
                                        strMarks = RepeatHeadingMarks(1)
                                        lngDelChr = j ' 空白までを含めて削除対象とする
                                        blnAfterHeadding = True
                                    End If
                                    Exit For
                                End If
                            End With
                        Next
                    ' 「第 1」のように間に半角スペースがある場合
                    ' ※3文字目は Characters ではなく Mid で取得する（And は両辺を評価するため、
                    '   「第」だけの段落で Characters(3) を求めるとエラーになる。Mid は空文字列を返す）
                    ElseIf .Range.Characters(2).Text = " " And Mid(.Range.Text, 3, 1) Like "#" Then
                        For j = 3 To .Range.Characters.count
                            With .Range.Characters(j)
                                If .Text = "　" Or .Text = " " Then
                                    If j - 1 - 1 <= 3 Then
                                        strMarks = RepeatHeadingMarks(1)
                                        lngDelChr = j
                                        blnAfterHeadding = True
                                    End If
                                    Exit For
                                End If
                            End With
                        Next
                    End If
                    
                ' 数字で始まる手入力見出しの場合
                ElseIf .Range.Characters(1).Text Like "#" Then
                    With .Range
                        For j = 1 To .Characters.count
                            ' 空白が見つかった場合
                            If .Characters(j).Text = "　" Or .Characters(j).Text = " " Then
                                If j - 1 <= 3 Then
                                    strMarks = RepeatHeadingMarks(2) ' レベル2
                                    lngDelChr = j
                                    blnAfterHeadding = True
                                End If
                                Exit For
                            ' ピリオドが見つかった場合（例: 1. ）
                            ElseIf .Characters(j).Text = "." Or .Characters(j).Text = "．" Then
                                ' 小数点でないこと（次が数字でない）を確認
                                If Not .Characters(j + 1).Text Like "#" Then
                                    If j - 1 <= 3 Then
                                        strMarks = RepeatHeadingMarks(2)
                                        lngDelChr = j
                                        blnAfterHeadding = True
                                    End If
                                    Exit For
                                End If
                                Exit For
                            End If
                        Next
                    End With
                
                ' カタカナで始まる手入力見出しの場合（ア、イ...）
                ElseIf .Range.Characters(1).Text Like "[ア-ン]" Or .Range.Characters(1).Text Like "[ｱ-ﾝ]" Then
                    With .Range.Characters(2)
                        ' 2文字目が区切り文字なら見出しと判定
                        If .Text = "　" Or .Text = " " Or .Text = "." Or .Text = "．" Then
                            strMarks = RepeatHeadingMarks(4) ' レベル4
                            lngDelChr = 2
                            blnAfterHeadding = True
                        End If
                    End With
                
                ' 英字で始まる手入力見出しの場合（a, b...）
                ElseIf .Range.Characters(1).Text Like "[ａ-ｈ]" Or .Range.Characters(1).Text Like "[a-h]" Then
                    With .Range.Characters(2)
                        If .Text = "　" Or .Text = " " Or .Text = "." Or .Text = "．" Then
                            strMarks = RepeatHeadingMarks(6) ' レベル6
                            lngDelChr = 2
                            blnAfterHeadding = True
                        End If
                    End With
                
                ' ローマ数字（i, v, x）で始まる場合
                ElseIf .Range.Characters(1).Text = "i" Or .Range.Characters(1).Text = "v" Or .Range.Characters(1).Text = "x" Then
                    With .Range
                        For j = 2 To .Characters.count
                            With .Characters(j)
                                ' 区切り文字を探す
                                If .Text = "　" Or .Text = " " Or .Text = "．" Or .Text = "." Then
                                    If j - 1 <= 7 Then
                                        strMarks = RepeatHeadingMarks(8) ' レベル8
                                        lngDelChr = j
                                        blnAfterHeadding = True
                                    End If
                                    Exit For
                                End If
                                ' ローマ数字以外の文字が出たら中断
                                If .Text <> "i" And .Text <> "v" And .Text <> "x" Then
                                    Exit For
                                End If
                            End With
                        Next
                    End With
                
                ' 丸カッコ (1) 等で始まる手入力見出しの場合
                ElseIf .Range.Characters(1).Text = "（" Or .Range.Characters(1).Text = "(" Then
                    ' (数字) パターン
                    If .Range.Characters(2).Text Like "#" Then
                        For j = 2 To .Range.Characters.count
                            ' 閉じ括弧を探す
                            If .Range.Characters(j).Text = "）" Or .Range.Characters(j).Text = ")" Then
                                If j - 2 <= 3 Then
                                    strMarks = RepeatHeadingMarks(3) ' レベル3
                                    lngDelChr = j
                                    blnAfterHeadding = True
                                End If
                                Exit For
                            End If
                        Next
                    
                    ' (カタカナ) パターン
                    ElseIf .Range.Characters(2).Text Like "[ア-ン]" Or .Range.Characters(2).Text Like "[ｱ-ﾝ]" Then
                        If .Range.Characters(3).Text = "）" Or .Range.Characters(3).Text = ")" Then
                            strMarks = RepeatHeadingMarks(5) ' レベル5
                            lngDelChr = 3
                            blnAfterHeadding = True
                        End If
                    
                    ' (英字) パターン
                    ElseIf .Range.Characters(2).Text Like "[ａ-ｈ]" Or .Range.Characters(2).Text Like "[a-h]" Then
                        If .Range.Characters(3).Text = "）" Or .Range.Characters(3).Text = ")" Then
                            strMarks = RepeatHeadingMarks(7) ' レベル7
                            lngDelChr = 3
                            blnAfterHeadding = True
                        End If
                    
                    ' (ローマ数字) パターン
                    ElseIf .Range.Characters(2).Text = "i" Or .Range.Characters(2).Text = "v" Or .Range.Characters(2).Text = "x" Then
                        With .Range
                            For j = 3 To .Characters.count
                                With .Characters(j)
                                    If .Text = "）" Or .Text = ")" Then
                                        If j - 2 <= 7 Then
                                            strMarks = RepeatHeadingMarks(9) ' レベル9
                                            lngDelChr = j
                                            blnAfterHeadding = True
                                        End If
                                        Exit For
                                    End If
                                    If .Text <> "i" And .Text <> "v" And .Text <> "x" Then
                                        Exit For
                                    End If
                                End With
                            Next
                        End With
                    End If
                End If
                
                '--- 標準スタイルを適用し、書式をリセットする処理 ---
                ' 現在のスタイル名を取得
                Dim strStyle As String
                strStyle = .Style.NameLocal
                
                ' まだ最初の見出しが出てきていない場合
                If blnAfterHeadding = False Then
                    ' 設定「最初の見出しまでスキップ」がTrueなら処理しない
                    If blnSkipUntil = True Then GoTo ContinueLoop
                Else
                    ' 最初の見出し以降の処理
                    ' 設定「右揃え（署名など）が出たら以降の処理を停止」がTrueの場合
                    If blnStopAfter = True Then
                        ' 行間詰めスタイルでなく、かつ配置が右揃えであれば
                        If strStyle <> "行間詰め" And .Format.alignment = wdAlignParagraphRight Then
                            blnAfterAlignRight = True ' 停止フラグを立てる
                        End If
                    End If
                    ' 停止フラグが立っていれば処理をスキップ
                    If blnAfterAlignRight = True Then GoTo ContinueLoop
                End If
                
                ' スタイルを「標準」に設定
                .Style = "標準"
                ' フォント書式を標準設定にリセット
                .Range.Font.Reset
                ' 段落書式を標準設定にリセット
                .Range.ParagraphFormat.Reset
                
                '--- マークダウン記号の書き込み処理 ---
                ' 削除すべき文字数（元の擬似的な箇条書き番号など）がある場合
                If lngDelChr <> 0 Then
                    ' 後ろから順に削除（インデックスずれを防ぐためStep -1推奨だが、Range操作なのでループで削除）
                    For j = lngDelChr To 1 Step -1
                        .Range.Characters(j).Delete
                    Next
                End If
                
                ' 文頭に残った不要な空白やピリオドを掃除する
                With .Range.Characters
                    ' 文字数分ループ（動的に削除するため注意）
                    For j = 1 To .count
                        ' 先頭が全角スペースやピリオドの場合
                        If .First = "　" Or .First = "." Or .First = "．" Then
                            .First.Delete
                        ' 先頭が半角スペースの場合（BackSpaceトリック使用）
                        ElseIf .First = " " Then
                            .First.Select
                            Selection.MoveRight
                            Selection.TypeBackspace
                        Else
                            ' 不要文字がなくなったら終了
                            Exit For
                        End If
                    Next
                End With
                
                ' 生成したマークダウン記号を行頭に挿入する
                If strMarks <> "" Then
                    ' 設定「マークダウン記号の後に空白を入れる」がTrueなら空白を追加
                    If blnMarkdownSpace = True Then strMarks = strMarks + " "
                    ' 行頭に挿入
                    .Range.InsertBefore strMarks
                End If
            End With
ContinueLoop:
        Next par ' 次の段落へ
    Next sec ' 次のセクションへ
End Sub

'================================================================================
' 機能: 指定された見出しレベルの数だけ、見出し記号を繰り返した文字列を生成します。
'       ※このモジュール専用のPrivate関数として定義
'================================================================================
Private Function RepeatHeadingMarks(lngOutline As Long) As String
    Dim i As Long
    ' 指定されたレベル回数だけループ
    For i = 1 To lngOutline
        ' 見出し記号（strHead = "#"等と推測）を文字列に追加連結する
        RepeatHeadingMarks = RepeatHeadingMarks + strHead
    Next
End Function
