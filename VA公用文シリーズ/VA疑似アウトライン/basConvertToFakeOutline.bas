Option Explicit

'================================================================================
' モジュールレベル変数
'================================================================================
' 設定値を管理するクラス等のインスタンスを保持する変数
Private Variables As VariablesManager

'================================================================================
' ConvertToFakeOutline（メインプロシージャ）
' 機能: 現在開いているWord文書のマークダウン形式を解析し、
'       Wordのインデントや書式設定された見出し番号を持つ疑似アウトラインに変換
'       します。
' 引数: なし
'================================================================================
Sub ConvertToFakeOutline()
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

    ' ユーザーに処理開始の確認を行う（キャンセル時は処理終了）
    If MsgBox("疑似アウトラインへの変換を開始します。" & vbCr & _
        strUndoMessage & vbCr & vbCr & _
        "よろしいですか？", vbOKCancel, "VA書式修正") = vbCancel Then Exit Sub

    ' ステータスバーに準備中のメッセージを表示する
    Application.StatusBar = "疑似アウトラインへの変換を準備しています。"

    ' エラー発生時にエラー処理ラベルへジャンプするよう設定
    On Error GoTo HandleErr

    ' 元に戻す（Undo）の記録を開始し、変換全体を Ctrl+Z 1回で戻せるようにする
    Dim blnUndoRecordStarted As Boolean
    If blnUseUndoRecord Then
        ' Mac版で未対応の場合に備え、エラーは無視して処理を続行する
        On Error Resume Next
        Application.UndoRecord.StartCustomRecord "VA疑似アウトライン：疑似アウトライン"
        ' 記録を開始できたかどうかを記憶する（終了処理で記録を閉じるため）
        blnUndoRecordStarted = (Err.Number = 0)
        Err.Clear
        ' エラー処理ラベルへジャンプする設定に戻す
        On Error GoTo HandleErr
    End If

    ' 処理速度向上のため、画面描画を停止する
    Application.ScreenUpdating = False
    ' マウスカーソルを待機状態（砂時計など）にする
    Application.System.Cursor = wdCursorWait
    
    ' 設定に基づき、表内の段落に特定のスタイルを適用する
    If blnApplyTableStyle = True Then
        Dim objPreFormatter As clsPreFormatter
        Set objPreFormatter = New clsPreFormatter
        objPreFormatter.ApplyTableStyle "行間詰め", ActiveDocument
        Set objPreFormatter = Nothing
    End If
    
    ' ページ設定から1文字あたりの幅（アドバンス幅）を計算する
    ' 計算式: (ページ幅 - 左右余白) / 1行あたりの文字数
    With ActiveDocument.PageSetup
        Dim sngAdvWidth As Single
        sngAdvWidth = (.PageWidth - .LeftMargin - .RightMargin) / .CharsLine
    End With
    
    ' 設定に基づき、行頭空白をインデントに置換する
    If blnReplaceSpaceWithIndent = True Then
        Set objPreFormatter = New clsPreFormatter
        ' "行間詰め" スタイルを除外して実行
        objPreFormatter.ReplaceSpaceWithIndent sngAdvWidth, "行間詰め", ActiveDocument
        Set objPreFormatter = Nothing
    End If
    
    '--- スマート カット アンド ペーストの設定の退避 ---
    ' 有効なままだと、マークダウン記号の後の半角空白が Range.Delete で削除されないことがあるため、
    ' 現在の設定を退避してから無効にする（終了処理で必ず元に戻す）
    Dim blnSmartCutPaste As Boolean      ' 退避した設定値
    Dim blnSmartCutPasteSaved As Boolean ' 設定値を退避できたかどうか

    ' Mac版で未対応の場合に備え、取得・設定のエラーは無視する
    On Error Resume Next
    ' 現在の設定を退避する
    blnSmartCutPaste = Options.SmartCutPaste
    ' 退避できたかどうかを記録する
    blnSmartCutPasteSaved = (Err.Number = 0)
    ' スマート カット アンド ペーストを無効にする
    Options.SmartCutPaste = False
    Err.Clear
    ' エラー処理ラベルへジャンプする設定に戻す
    On Error GoTo HandleErr

    ' 複合見出し（改行のみで構成された見出し等）の情報を保持する配列変数
    Dim varRtnOnlyPar As Variant

    '--- メイン処理: マークダウンを解析し、疑似アウトライン（書式と番号）を適用 ---
    Call ConvertToFakeOutlineFormat(varRtnOnlyPar, sngAdvWidth)
    
    '--- 後処理: 記録された複合見出しの段落を調整 ---
    Call ProcessCompositeHeadings(varRtnOnlyPar, sngAdvWidth)
    
Cleanup:
    ' 終了処理ラベル
    ' 終了処理の途中でエラーが起きても、残りの終了処理を必ず行う
    On Error Resume Next
    ' 元に戻す（Undo）の記録を開始した場合のみ、記録を終了する
    If blnUndoRecordStarted Then Application.UndoRecord.EndCustomRecord
    ' スマート カット アンド ペーストの設定を、処理開始前の状態に戻す（正常終了・エラー終了の両方で必ず行う）
    If blnSmartCutPasteSaved Then Options.SmartCutPaste = blnSmartCutPaste
    ' 画面描画を再開する
    Application.ScreenUpdating = True
    ' ステータスバーをクリアする
    Application.StatusBar = ""
    ' マウスカーソルを通常に戻す
    Application.System.Cursor = wdCursorNormal
    ' クラスインスタンスを解放する
    Set Variables = Nothing
    ' プロシージャを抜ける
    Exit Sub

HandleErr:
    ' エラー処理ラベル
    ' エラーメッセージを表示する
    MsgBox "エラーが発生したため、処理を停止しました。", vbOKOnly + vbExclamation, "VA見出し変換"
    ' クリーンアップ処理へ移動して終了する
    Resume Cleanup
End Sub

'================================================================================
' 選択範囲の疑似アウトライン風書式をトグル（設定/解除）
' 機能: 選択範囲のフォント倍率が100%の場合は「疑似アウトライン風の全角1文字書式」
'       を適用し、それ以外（既に設定済み）の場合は書式を解除して復元します。
'       ※適用対象は「（」または「(」で始まる括弧付き文字のみに限定します。
' 引数: blnDummy (Boolean, Optional):
'       Wordの「マクロ」ダイアログボックスにこのマクロが表示されるのを
'       防ぐためのダミー引数です。
'================================================================================
Sub ToggleFakeOneChar(Optional blnDummy As Boolean)
    ' 変数管理クラスのインスタンス生成（設定値参照のため）
    Set Variables = New VariablesManager
    
    ' 選択範囲がない（カーソルのみ）場合は終了
    If Selection.Type = wdSelectionIP Then Exit Sub
    
    ' エラー処理
    On Error GoTo HandleErr

    ' 元に戻す（Undo）の記録を開始し、書式の設定・解除を Ctrl+Z 1回で戻せるようにする
    Dim blnUndoRecordStarted As Boolean
    If blnUseUndoRecord Then
        ' Mac版で未対応の場合に備え、エラーは無視して処理を続行する
        On Error Resume Next
        Application.UndoRecord.StartCustomRecord "VA疑似アウトライン：全角1文字"
        ' 記録を開始できたかどうかを記憶する（終了処理で記録を閉じるため）
        blnUndoRecordStarted = (Err.Number = 0)
        Err.Clear
        ' エラー処理に戻す
        On Error GoTo HandleErr
    End If

    ' 画面描画停止
    Application.ScreenUpdating = False

    ' 選択範囲（Range）を取得
    Dim rng As Range
    Set rng = Selection.Range
    
    ' ---------------------------------------------------------
    ' 状態判定と分岐
    ' Range.Font.Scaling が 100 であれば「未設定」とみなし適用処理を行う
    ' 9999999(wdUndefined: 混合) または 100以外 であれば「設定済み」とみなし復元処理を行う
    ' ---------------------------------------------------------
    If rng.Font.Scaling = 100 Then
        ' ==========================
        ' 書式適用処理（Apply）
        ' ==========================
        
        ' 選択範囲の文字列を取得し、前後の空白を除去
        Dim strNumber As String
        strNumber = Trim(rng.Text)
        
        ' --- 括弧で始まる場合のみ処理を実行する ---
        If Left(strNumber, 1) = "（" Or Left(strNumber, 1) = "(" Then
            
            ' ページ設定からアドバンス幅を計算（処理する場合のみ計算）
            Dim sngAdvWidth As Single
            With ActiveDocument.PageSetup
                sngAdvWidth = (.PageWidth - .LeftMargin - .RightMargin) / .CharsLine
            End With
            
            ' テキストをクリアして再構築
            rng.Text = ""
            
            Dim j As Long
            Dim targetVal As Double
            
            ' 全角括弧 "（" の場合
            If Left(strNumber, 1) = "（" Then
                strNumber = strNumber
                rng.Text = strNumber
                ' 括弧の倍率を縮小
                rng.Characters(1).Font.Scaling = FontMag(25, 1, sngAdvWidth)
                rng.Characters(Len(strNumber)).Font.Scaling = FontMag(25, 1, sngAdvWidth)
                
                ' 中身の文字倍率を調整
                For j = 2 To Len(strNumber) - 1
                    ' 文字が全角か半角かでターゲット値を変更
                    If LenB(StrConv(rng.Characters(j).Text, vbFromUnicode)) = 1 Then
                         targetVal = 100 ' 半角
                    Else
                         targetVal = 50  ' 全角
                    End If
                    rng.Characters(j).Font.Scaling = FontMag(targetVal, Len(strNumber) - 2, sngAdvWidth)
                Next
                
            ' 半角括弧 "(" の場合
            ElseIf Left(strNumber, 1) = "(" Then
                strNumber = strNumber
                rng.Text = strNumber
                ' 括弧の倍率を縮小
                rng.Characters(1).Font.Scaling = FontMag(50, 1, sngAdvWidth)
                rng.Characters(Len(strNumber)).Font.Scaling = FontMag(50, 1, sngAdvWidth)
                
                ' 中身の文字倍率を調整
                For j = 2 To Len(strNumber) - 1
                    ' 文字が全角か半角かでターゲット値を変更
                    If LenB(StrConv(rng.Characters(j).Text, vbFromUnicode)) = 1 Then
                         targetVal = 100 ' 半角
                    Else
                         targetVal = 50  ' 全角
                    End If
                    rng.Characters(j).Font.Scaling = FontMag(targetVal, Len(strNumber) - 2, sngAdvWidth)
                Next
            End If
            
            ' 最後にカーソルを変換後の文字列の後ろに移動（必要であればコメント解除）
            ' rng.Collapse Direction:=wdCollapseEnd
            ' rng.Select
            
        Else
            ' 括弧で始まらない場合（"第"や数字のみ）は、何もせずに終了する
        End If

    Else
        ' ==========================
        ' 書式復元処理（Restore）
        ' ==========================
        
        ' 1. フォント倍率を100%（標準）に戻す
        rng.Font.Scaling = 100
        
        ' 2. マクロが付与した余分な空白を除去する
        Dim strClean As String
        strClean = rng.Text
        
        ' 【重要】末尾に改行コードが含まれている場合、Trimが効かないため一時的に除去する
        Dim hasCr As Boolean
        hasCr = False
        If Right(strClean, 1) = vbCr Then
            hasCr = True
            strClean = Left(strClean, Len(strClean) - 1)
        End If
        
        ' 末尾の全角スペース（共通処理で付与されるもの）を削除
        If Right(strClean, 1) = "　" Then
            strClean = Left(strClean, Len(strClean) - 1)
        End If
        
        ' 前後の半角スペース（1文字パターンの調整等で付与されるもの）を削除
        strClean = Trim(strClean)
        
        ' 改行コードがあった場合は戻す
        If hasCr Then strClean = strClean & vbCr
        
        ' 文字列を書き戻す
        rng.Text = strClean
        
    End If

Cleanup:
    ' 終了処理の途中でエラーが起きても、残りの終了処理を必ず行う
    On Error Resume Next
    ' 元に戻す（Undo）の記録を開始した場合のみ、記録を終了する
    If blnUndoRecordStarted Then Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    Set Variables = Nothing
    Set rng = Nothing
    Exit Sub
    
HandleErr:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub

'================================================================================
' 疑似アウトラインに変換
' 機能: マークダウン記号の解析、インデント計算、見出し番号の生成と挿入を行う
' 引数: varRtnOnlyPar - 複合見出し情報を呼び出し元に返すための参照渡し変数
'================================================================================
Private Sub ConvertToFakeOutlineFormat(ByRef varRtnOnlyPar As Variant, sngAdvWidth As Single)
    ' ステータスバーに進捗状況を表示
    Application.StatusBar = "疑似アウトラインに変換しています。"
    
    ' インデント設定値を格納する配列を宣言（レベル0～9を想定）
    Dim dblHeadingLeftIndents(10) As Double       ' 見出し: 左インデント
    Dim dblHeadingFirstLineIndents(10) As Double  ' 見出し: 1行目のインデント（ぶら下げ等）
    Dim dblTextLeftIndents(10) As Double          ' 本文: 左インデント
    Dim dblTextFirstLineIndents(10) As Double     ' 本文: 1行目のインデント
    
    Dim x As Long
    ' Variablesオブジェクトから各レベルのインデント設定値を配列に読み込む
    For x = 0 To 9
        dblHeadingLeftIndents(x) = Val(Variables.varHeadingLeftIndents(x))
        dblHeadingFirstLineIndents(x) = Val(Variables.varHeadingFirstLineIndents(x))
        dblTextLeftIndents(x) = Val(Variables.varTextLeftIndents(x))
        dblTextFirstLineIndents(x) = Val(Variables.varTextFirstLineIndents(x))
    Next
    
    Dim i As Long           ' 段落の通し番号カウンタ
    Dim totalPars As Long   ' 文書全体の総段落数
    totalPars = ActiveDocument.Paragraphs.count
    
    Dim sec As Section      ' セクション用ループ変数
    Dim par As Paragraph    ' 段落用ループ変数
    
    Dim xFirst As Long      ' そのセクション内での「最初の見出しレベル」を記録する変数
    xFirst = 0
    
    Dim blnAfterHeadding As Boolean     ' 見出しが既に登場したかどうかのフラグ
    blnAfterHeadding = False
    
    Dim blnAfterAlignRight As Boolean   ' 右揃え段落が登場したかどうかのフラグ
    blnAfterAlignRight = False
    
    Dim varOrdinalNumber(1 To 10) As Variant    ' 見出しレベルごとの現在番号（序数）を保持する配列
    Erase varOrdinalNumber                      ' 配列を初期化
    
    '--- 全セクションに対するループ処理開始 ---
    For Each sec In ActiveDocument.Sections
        '--- セクションごとのリセット処理 ---
        ' 最初の段落以外の場合（セクションの変わり目）
        If i > 0 Then
            ' 最初の見出しレベルをリセット
            xFirst = 0
            
            ' 見出し通過フラグをリセット
            blnAfterHeadding = False
            
            ' 右揃え通過フラグをリセット
            blnAfterAlignRight = False
            
            ' 設定「セクション区切りで見出し番号を振り直す」がTrueなら番号カウンタをリセット
            If blnHeadingReset = True Then Erase varOrdinalNumber
        End If
        
        '--- セクション内の全段落に対するループ処理開始 ---
        For Each par In sec.Range.Paragraphs
            ' 段落番号をカウントアップ
            i = i + 1
            
            ' 50段落ごとにステータスバーの進捗率を更新（処理速度低下を防ぐため間引く）
            If i Mod 50 = 0 Then
                Dim intComp As Integer      ' 進捗率（%）
                If totalPars > 0 Then intComp = Int(i / totalPars * 100)
                Application.StatusBar = "疑似アウトラインに変換しています。　" & intComp & "%完了"
            End If
            
            ' 現在の段落に対する処理
            With par
                ' 現在の段落スタイル名を取得
                Dim strStyle As String
                strStyle = .Style.NameLocal
                
                ' 設定で「行間詰めを除外しない」がFalse、かつスタイルが「行間詰め」なら処理をスキップ
                If Variables.blnNoParagraphSpace = False And strStyle = "行間詰め" Then GoTo ContinueLoop
                
                ' 段落の文字列を取得
                Dim strtext As String
                strtext = .Range.Text
                
                ' 文字列の先頭がマークダウンの見出し記号（定義済み変数 strHead、通常は"#"）か判定
                If Left(strtext, 1) = strHead Then
                    blnAfterHeadding = True ' 見出しフラグを立てる
                End If
                
                '--- 見出しの前後の処理制御 ---
                If blnAfterHeadding = False Then
                    ' 最初の見出しが登場するまでスキップする設定の場合
                    If blnSkipUntil = True Then GoTo ContinueLoop
                Else
                    ' 「右揃え段落以降は処理を停止する」設定の場合の判定
                    If blnStopAfter = True Then
                        ' スタイルが「行間詰め」以外で、かつ右揃え設定の場合
                        If strStyle <> "行間詰め" And .Format.alignment = wdAlignParagraphRight Then
                            blnAfterAlignRight = True
                        End If
                    End If
                    ' 右揃え通過後のフラグが立っていればスキップ
                    If blnAfterAlignRight = True Then GoTo ContinueLoop
                End If
                
                '--- マークダウン解析と変換処理 ---
                
                ' 1. 見出し（#）の処理
                ' 先頭文字が見出し記号の場合
                If Left(strtext, 1) = strHead Then
                    ' レベル9から1まで逆順にチェックし、#の連続数（レベル）を特定する
                    For x = 9 To 1 Step -1
                        If Left(strtext, x) = String(x, strHead) Then
                            ' セクション内での最上位レベル（基準レベル）を記録
                            If xFirst = 0 Or xFirst > x Then xFirst = x
                            
                            blnAfterHeadding = True
                            
                            ' #記号を除いた文字数が2文字以下（改行含む）の場合の特殊処理
                            ' （空の見出しや、複合見出しの一部である可能性がある）
                            If .Range.Characters.count - x <= 2 Then
                                ' 文字数が極端に少ない、または#の直後にスペースがある場合
                                If .Range.Characters.count - x <= 1 Or Left(strtext, x + 1) = String(x, strHead) & " " Then
                                    ' 配列 varRtnOnlyPar を拡張して情報を格納する
                                    If IsEmpty(varRtnOnlyPar) = True Then
                                        ReDim varRtnOnlyPar(1 To 1)
                                    Else
                                        ReDim Preserve varRtnOnlyPar(1 To UBound(varRtnOnlyPar) + 1)
                                    End If
                                    
                                    ' 段落通し番号(i)と見出しレベル(x)をカンマ区切り文字列で記録
                                    varRtnOnlyPar(UBound(varRtnOnlyPar)) = i & "," & x
                                End If
                            End If
                            
                            Dim xCurr As Long         ' 現在の見出しレベル
                            
                            '--- 見出し番号のカウントアップ処理 ---
                            If x > xCurr Then
                                ' レベルが深くなった場合（親→子）、そのレベルを1にする
                                varOrdinalNumber(x) = 1
                            ElseIf x < xCurr Then
                                ' レベルが浅くなった場合（子→親）、そのレベルをカウントアップ
                                varOrdinalNumber(x) = varOrdinalNumber(x) + 1
                                ' 下位レベルのカウンタを全てリセット（1に戻す）
                                Dim xReset As Long
                                For xReset = x + 1 To 10
                                    varOrdinalNumber(xReset) = 1
                                Next
                            Else
                                ' 同一レベルの場合、単純にカウントアップ
                                varOrdinalNumber(x) = varOrdinalNumber(x) + 1
                            End If
                            
                            xCurr = x ' 現在レベルを更新
                            
                            '--- 見出し番号文字列の生成 ---
                            Dim strNumber As String
                            ' HeadingNumber関数: 書式設定文字列（例:第%1章）に番号を埋め込む
                            strNumber = HeadingNumber(Variables.varNumberFormats(x), _
                                Variables.varHeadingNumber(varOrdinalNumber(x), Variables.varNumberStyles(x)))
                             
                            ' 元のマークダウン記号（#）を削除する
                            Dim y As Long
                            For y = x To 1 Step -1
                                .Range.Characters(y).Delete
                            Next

                            ' マークダウン記号の後の空白を削除する（「第１　 見出し」とならないようにするため）
                            Call DeleteLeadingSpaces(.Range)

                             '--- 見出し番号の書式設定と挿入 ---
                             ' 設定「数字を全角1文字扱いにする」がTrueの場合
                             If Variables.blnOneFullWidth = True Then
                                ' 番号文字列のパターン（"第"で始まるか等）に応じた複雑な処理
                                ' フォントの倍率（Scaling）を調整して、見栄えを整えている
                                If Left(strNumber, 1) = "第" Then
                                    strNumber = Right(strNumber, Len(strNumber) - 1) ' "第"を除去
                                    
                                    ' 番号が1桁の場合
                                    If Len(strNumber) = 1 Then
                                        ' 半角数字かどうかで分岐
                                        If Len(Right(strNumber, 1)) = LenB(StrConv(Right(strNumber, 1), vbFromUnicode)) Then
                                            ' 半角の場合: "第 N 　" の形式にし、スペースの倍率を50%に縮小
                                            strNumber = "第 " & strNumber & " 　"
                                            
                                            .Range.InsertBefore strNumber
                                            .Range.Characters(2).Font.Scaling = FontMag(50, 1, sngAdvWidth)
                                            .Range.Characters(4).Font.Scaling = FontMag(50, 1, sngAdvWidth)
                                        Else
                                            ' 全角の場合
                                            strNumber = "第" & strNumber & "　"
                                            .Range.InsertBefore strNumber
                                        End If
                                    ' 番号が2桁の場合
                                    ElseIf Len(strNumber) = 2 Then
                                        strNumber = "第" & strNumber & "　"
                                        .Range.InsertBefore strNumber
                                        ' 数字部分のフォント倍率を調整して収める
                                        Dim j As Long
                                        For j = 2 To Len(strNumber) - 1
                                            If Variables.varNumberStyles(x) = 14 Or Variables.varNumberStyles(x) = 20 Then
                                                .Range.Characters(j).Font.Scaling = FontMag(100, Len(strNumber) - 2, sngAdvWidth)
                                            Else
                                                .Range.Characters(j).Font.Scaling = FontMag(200, Len(strNumber) - 2, sngAdvWidth)
                                            End If
                                        Next
                                    End If
                                ' 先頭が全角括弧 "（" の場合
                                ElseIf Left(strNumber, 1) = "（" Then
                                    strNumber = strNumber & "　"
                                    .Range.InsertBefore strNumber
                                    ' 括弧の倍率を縮小
                                    .Range.Characters(1).Font.Scaling = FontMag(25, 1, sngAdvWidth)
                                    .Range.Characters(Len(strNumber) - 1).Font.Scaling = FontMag(25, 1, sngAdvWidth)
                                    ' 中身の文字倍率を調整
                                    For j = 2 To Len(strNumber) - 2
                                        If Variables.varNumberStyles(x) = 14 Or Variables.varNumberStyles(x) = 20 Then
                                            .Range.Characters(j).Font.Scaling = FontMag(50, Len(strNumber) - 3, sngAdvWidth)
                                        Else
                                            .Range.Characters(j).Font.Scaling = FontMag(100, Len(strNumber) - 3, sngAdvWidth)
                                        End If
                                    Next
                                ' 先頭が半角括弧 "(" の場合
                                ElseIf Left(strNumber, 1) = "(" Then
                                     strNumber = strNumber & "　"
                                     .Range.InsertBefore strNumber
                                    ' 括弧の倍率を縮小
                                    .Range.Characters(1).Font.Scaling = FontMag(50, 1, sngAdvWidth)
                                    .Range.Characters(Len(strNumber) - 1).Font.Scaling = FontMag(50, 1, sngAdvWidth)
                                    ' 中身の文字倍率を調整
                                    For j = 2 To Len(strNumber) - 2
                                        If Variables.varNumberStyles(x) = 14 Or Variables.varNumberStyles(x) = 20 Then
                                            .Range.Characters(j).Font.Scaling = FontMag(50, Len(strNumber) - 3, sngAdvWidth)
                                        Else
                                            .Range.Characters(j).Font.Scaling = FontMag(100, Len(strNumber) - 3, sngAdvWidth)
                                        End If
                                    Next
                                ' その他の形式の場合
                                Else
                                    ' 1文字の場合
                                    If Len(strNumber) = 1 Then
                                        If Variables.varNumberStyles(x) = 14 Or Variables.varNumberStyles(x) = 20 Then
                                            strNumber = strNumber & "　"
                                            .Range.InsertBefore strNumber
                                        Else
                                            ' 前後に半角スペースを入れて調整
                                            strNumber = " " & strNumber & " " & "　"
                                            .Range.InsertBefore strNumber
                                            .Range.Characters(1).Font.Scaling = FontMag(50, 1, sngAdvWidth)
                                            .Range.Characters(3).Font.Scaling = FontMag(50, 1, sngAdvWidth)
                                        End If
                                    ' 複数文字の場合
                                    Else
                                        strNumber = strNumber & "　"
                                        .Range.InsertBefore strNumber
                                        ' 全体を縮小して枠に収める
                                        For j = 1 To Len(strNumber) - 1
                                            If Variables.varNumberStyles(x) = 14 Or Variables.varNumberStyles(x) = 20 Then
                                                .Range.Characters(j).Font.Scaling = FontMag(100, Len(strNumber) - 1, sngAdvWidth)
                                            Else
                                                .Range.Characters(j).Font.Scaling = FontMag(200, Len(strNumber) - 1, sngAdvWidth)
                                            End If
                                        Next
                                    End If
                                End If
                             Else
                                ' 設定「半角モード」の場合
                                ' 末尾が半角括弧閉じなら半角スペースを追加
                                If Right(strNumber, 1) = ")" Then
                                    strNumber = strNumber & " "
                                Else
                                    ' 特定の番号スタイルの場合は全角スペースを追加、そうでなければ全角スペース追加
                                    If Variables.varNumberStyles(x) = 4 Or Variables.varNumberStyles(x) = 2 Then
                                        strNumber = strNumber & "　 "
                                    Else
                                        strNumber = strNumber & "　"
                                    End If
                                End If
                                ' 文字列を挿入
                                .Range.InsertBefore strNumber
                             End If
                             
                             '--- インデントの設定 ---
                             ' 左インデント計算式: (基本インデント + レベル差分による追加分) * アドバンス幅
                             .LeftIndent = (dblHeadingLeftIndents(x) + lngLeftIndentsAdd * (x - xFirst)) * sngAdvWidth
                             ' 1行目のインデント（ぶら下げ）の設定
                             .FirstLineIndent = (dblHeadingFirstLineIndents(x)) * sngAdvWidth
                             ' このレベルの処理が終わったらループを抜ける
                             Exit For
                        End If
                    Next x
                    
                ' 2. 箇条書き（*）の処理
                ' 先頭文字が箇条書き記号の場合
                ElseIf Left(strtext, 1) = strList Then
                    ' マークダウン記号を削除
                    For y = 1 To 1 Step -1
                        .Range.Characters(y).Delete
                    Next
                    ' マークダウン記号の後の空白を削除する（「・ 項目」とならないようにするため）
                    Call DeleteLeadingSpaces(.Range)
                    ' 指定されたビュレット文字（記号）を挿入
                    .Range.InsertBefore strListBullFormat
                    ' インデントを設定（現在の見出しレベルに合わせて調整）
                    .LeftIndent = (dblHeadingLeftIndents(xCurr) + lngLeftIndentsAdd * (xCurr - xFirst) + lngListLeftIndent + 1) * sngAdvWidth
                    .FirstLineIndent = lngListFirstLineIndent * sngAdvWidth
                    
                ' 3. 段落番号（1.）の処理
                ' 先頭が数字で始まる場合
                ElseIf IsNumeric(Left(strtext, 1)) = True Then
                    ' 2文字目がドット "." の場合
                    If Mid(strtext, 2, 1) = "." Then
                        ' レベル10を箇条書き番号用として使用するための判定
                        If 10 <> x Then
                            varOrdinalNumber(10) = 1
                            x = 10
                        Else
                            varOrdinalNumber(10) = varOrdinalNumber(10) + 1
                        End If
                        
                        ' 段落番号文字列を生成
                        strNumber = HeadingNumber(strListNumbFormat, Variables.varHeadingNumber(varOrdinalNumber(10), 18))
                        
                        ' 元の "1." 等を削除
                        For y = 2 To 1 Step -1
                            .Range.Characters(y).Delete
                        Next
                        ' マークダウン記号の後の空白を削除する（「①　 項目」とならないようにするため）
                        Call DeleteLeadingSpaces(.Range)

                        ' 新しい番号を挿入
                        .Range.InsertBefore strNumber
                        ' インデントを設定
                        .LeftIndent = (dblHeadingLeftIndents(xCurr) + lngLeftIndentsAdd * (xCurr - xFirst) + lngListLeftIndent + 1) * sngAdvWidth
                        .FirstLineIndent = lngListFirstLineIndent * sngAdvWidth
                    End If
                    
                ' 4. 通常段落（本文）の処理
                ' マークダウン記号がない場合
                Else
                    ' インデントを設定（現在の見出しレベルに合わせて調整）
                    .LeftIndent = (dblTextLeftIndents(xCurr) + lngLeftIndentsAdd * (xCurr - xFirst)) * sngAdvWidth
                    .FirstLineIndent = dblTextFirstLineIndents(xCurr) * sngAdvWidth
                End If
            End With
ContinueLoop:
        Next par ' 次の段落へ
    Next sec ' 次のセクションへ
End Sub

'================================================================================
' 複合見出しを処理
' 機能: 記録された段落番号情報を元に、改行を削除して1つの見出しに結合し、インデントを調整する
' 引数: varRtnOnlyPar - 処理すべき段落情報（通し番号とレベル）の配列
'================================================================================
Private Sub ProcessCompositeHeadings(varRtnOnlyPar As Variant, sngAdvWidth As Single)
    ' ステータスバーに進捗を表示
    Application.StatusBar = "複合見出しを処理しています。"
    
    ' 処理対象の配列が空でない場合のみ実行
    If IsEmpty(varRtnOnlyPar) = True Then Exit Sub
        
    ' 左インデントの増加量が1（標準的なグリッド合わせ）でない場合は処理不可として終了
    If lngLeftIndentsAdd <> 1 Then
        MsgBox "左インデント増加量が1でないため、複合見出しは処理できません。"
        Exit Sub
    End If
    
    Dim k As Long
    ' Paragraphsコレクションオブジェクトを取得（インデックス参照用）
    Dim pars As Paragraphs
    Set pars = ActiveDocument.Paragraphs
    
    Dim lngIndentAdjust As Long ' インデント調整量
    lngIndentAdjust = 0
    
    ' 配列の末尾から先頭に向かって逆順処理（段落結合によるずれを防ぐため）
    For k = UBound(varRtnOnlyPar) To 1 Step -1
        Dim parIndex As Long
        ' カンマ区切りの文字列から段落インデックスを取得
        parIndex = CLng(Split(varRtnOnlyPar(k), ",")(0))
        
        ' 対象段落に対する処理
        With pars(parIndex)
            ' インデント調整量を加算
            lngIndentAdjust = lngIndentAdjust + lngLeftIndentsAdd
            
            ' インデントの再計算と適用
            If k = 1 Then
                ' 最後の要素（処理順では最後だが配列の先頭）の場合
                .LeftIndent = .LeftIndent + lngIndentAdjust * sngAdvWidth
                .FirstLineIndent = .FirstLineIndent - lngIndentAdjust * sngAdvWidth
                lngIndentAdjust = 0 ' 調整量をリセット
            Else
                ' 現在の段落と、その前の段落が連続していない場合（別のブロックの場合）
                If parIndex - 1 <> CLng(Split(varRtnOnlyPar(k - 1), ",")(0)) Then
                    .LeftIndent = .LeftIndent + lngIndentAdjust * sngAdvWidth
                    .FirstLineIndent = .FirstLineIndent - lngIndentAdjust * sngAdvWidth
                    lngIndentAdjust = 0 ' 調整量をリセット
                End If
            End If
            
            ' 改行の前に全角スペースがあれば削除する
            Dim lngRtnOnlyPar As Long
            lngRtnOnlyPar = Len(.Range.Text)
            If lngRtnOnlyPar > 1 Then
                If .Range.Characters(lngRtnOnlyPar - 1) = "　" Then
                    .Range.Characters(lngRtnOnlyPar - 1).Delete
                End If
            End If
            
            ' 末尾の改行コードを削除（これにより次の段落と結合される）
            lngRtnOnlyPar = Len(.Range.Text)
            If .Range.Characters.Last.Text = vbCr Then .Range.Characters.Last.Delete
            
        End With
    Next
    Set pars = Nothing ' オブジェクト開放
End Sub

'================================================================================
' 機能: 段落の先頭の空白（半角・全角）を削除する
'       マークダウン記号を削除した後に残る空白（「# 見出し」「- 項目」の「 」）が、
'       見出し番号や記号の後の全角空白と並ばないようにするため
'       ※半角空白は、スマート カット アンド ペーストが有効だと削除されないことがあるため、
'         呼び出し元（ConvertToFakeOutline）で無効にしてから呼び出す
' 引数: rng - 対象の段落の範囲
'       ※このモジュール専用のPrivateプロシージャとして定義
'================================================================================
Private Sub DeleteLeadingSpaces(rng As Range)
    Dim lngLenBefore As Long    ' 削除する前の段落の文字数

    ' 段落の1文字目が「全角スペース」または「半角スペース」の間、繰り返す
    Do While rng.Characters(1).Text = "　" Or rng.Characters(1).Text = " "
        ' 削除前の段落の文字数を記録する
        lngLenBefore = Len(rng.Text)

        ' 1文字目の空白を削除する
        rng.Characters(1).Delete

        ' 空白を削除できなかった場合は、無限ループを避けるため繰り返しを終了する
        If Len(rng.Text) >= lngLenBefore Then Exit Do
    Loop
End Sub

'================================================================================
' 機能: 番号書式テンプレート内のプレースホルダー（例: "%1"）を、実際の番号文字列に置換する
' 引数: strA - テンプレート文字列, strB - 挿入する番号文字列
'       ※このモジュール専用のPrivate関数として定義
'================================================================================
Private Function HeadingNumber(strA As Variant, strB As Variant) As String
    Dim i As Integer
    Dim strC As String
    Dim lenA As Integer
    
    lenA = Len(strA)
    i = 1
    strC = ""
    ' 文字列を1文字ずつ走査
    Do While i <= lenA
        ' "%"が見つかり、かつ末尾ではない場合
        If Mid(strA, i, 1) = "%" And i < lenA Then
            strC = strC & strB ' プレースホルダーの代わりに番号を結合
            i = i + 2          ' "%"と次の文字（数字）をスキップ
        Else
            strC = strC & Mid(strA, i, 1) ' 通常の文字はそのまま結合
            i = i + 1
        End If
    Loop
    HeadingNumber = strC
End Function

'================================================================================
' 機能: 指定した文字数の文字列を、目標幅に収めるためのフォント拡大・縮小率（%）を算出する
' 引数: dblTarget - 目標幅係数, lngLen - 文字列長, sngAdvWidth - アドバンス幅
'       ※このモジュール専用のPrivate関数として定義
'================================================================================
Private Function FontMag(dblTarget As Double, lngLen As Long, sngAdvWidth As Single) As Double
    Dim sngAdjust As Single
    ' アドバンス幅に応じた補正値を計算（フォントサイズ12ポイント基準の調整と思われる）
    sngAdjust = (sngAdvWidth - 12) * 4
    ' 基本倍率から補正値を引く
    ' 修正: 2.01倍にならないようIntで切り捨てを行う
    FontMag = Int((dblTarget / lngLen) - sngAdjust)
    ' 計算結果が0以下の場合は100%（等倍）とする
    If FontMag <= 0 Then
        FontMag = 100
    End If
End Function
