'変数の宣言を強制するオプションを設定します。これにより、未宣言の変数の使用を防ぎます。
Option Explicit

'-----------------------------------------------------------------------
' チェック実行プロシージャ
' 内容証明郵便の形式（字数・行数）をチェックするためのメイン処理です。
'-----------------------------------------------------------------------
Sub CheckCharsAndLines()
    '変数の宣言セクション
    Dim varMaxCharsAndLines As Variant      '最大文字数および行数設定値の配列を格納するバリアント型変数
    Dim intMaxChars As Integer              '許容される最大文字数を格納する整数型変数
    Dim intMaxLines As Integer              '許容される最大行数を格納する整数型変数
    Dim lngChars As Long                    '現在チェック中の行の文字数を格納する長整数型変数
    Dim lngLines As Long                    '現在チェック中のページの行数を格納する長整数型変数
    Dim lngLastPage As Long                 'ドキュメントの総ページ数を格納する長整数型変数
    Dim lngLastRect As Long                 '現在のページに含まれるRectangle（テキスト領域）の総数を格納する変数
    Dim lngLastLine As Long                 '現在のRectangleに含まれる行の総数を格納する変数
    Dim i As Long, j As Long, k As Long     'ループ処理用のカウンタ変数（i:ページ, j:Rectangle, k:行）
    Dim strText As String                   '取得した行内の文字列を一時的に格納する文字列型変数
    Dim blnDispRslt As Boolean              '行ごとの結果をメッセージボックスで表示するかどうかのフラグ変数
    Dim sec As Section                      'ヘッダー・フッター削除処理でループするためのセクションオブジェクト変数
    Dim rngLine As Range                    '行のRangeオブジェクト
    Dim blnHeaderFooterDeleted As Boolean   'ヘッダー・フッターを削除したかどうかのフラグ

    '以下のWebサイトの情報を参考にロジックを構築しています
    'https://www.relief.jp/docs/word-vba-select-page.html
    'https://www.relief.jp/docs/word-vba-get-text-each-line.html
    'Word 2013以降の仕様変更（Rectangleへのアクセス時の例外発生）に対応するため、
    'エラーハンドリングを用いて回避しています。
    '参考: https://support.microsoft.com/ja-jp/topic/word-2013-でオートメーションから特定の種類のオブジェクトの_
    '-rectangletypプロパティにアクセスすると例外が発生する-5ac4c453-496e-eaf9-1b13-4c9379ecec24

    'ユーザーに処理の開始を確認するメッセージボックスを表示します。
    If MsgBox("内容証明郵便のチェックを開始します。", vbOKCancel, "VA内容証明郵便") = vbCancel Then
        Exit Sub
    End If

    '処理中であることを視覚的に伝えるため、マウスポインタを待機状態にします。
    Application.System.Cursor = wdCursorWait
    'ステータスバーにメッセージを表示します。
    Application.StatusBar = "内容証明郵便のチェックを行っています..."

    '-----------------------------------------------------------------------
    ' 画面更新の停止（パフォーマンス向上）
    '-----------------------------------------------------------------------
    Application.ScreenUpdating = False

    '-----------------------------------------------------------------------
    ' ヘッダーおよびフッターの削除処理
    '-----------------------------------------------------------------------
    blnHeaderFooterDeleted = False 'フラグを初期化
    
    'ドキュメント内のすべてのセクションに対してループ処理を行います。
    For Each sec In ActiveDocument.Sections
        '各ヘッダー・フッターについて、内容があればフラグを立てて削除します
        Call DeleteAndCheck(sec.Headers(wdHeaderFooterPrimary).Range, blnHeaderFooterDeleted)
        Call DeleteAndCheck(sec.Footers(wdHeaderFooterPrimary).Range, blnHeaderFooterDeleted)
        Call DeleteAndCheck(sec.Headers(wdHeaderFooterFirstPage).Range, blnHeaderFooterDeleted)
        Call DeleteAndCheck(sec.Footers(wdHeaderFooterFirstPage).Range, blnHeaderFooterDeleted)
        Call DeleteAndCheck(sec.Headers(wdHeaderFooterEvenPages).Range, blnHeaderFooterDeleted)
        Call DeleteAndCheck(sec.Footers(wdHeaderFooterEvenPages).Range, blnHeaderFooterDeleted)
    Next sec

    'ヘッダーまたはフッターが削除された場合、ユーザーに通知します
    If blnHeaderFooterDeleted Then
        MsgBox "ヘッダーまたはフッターが設定されていたため、削除しました。", vbInformation, "VA内容証明郵便"
    End If

    '設定値を読み込むための処理です。
    varMaxCharsAndLines = Split("," & strMaxCharsAndLines, ",")
    intMaxChars = varMaxCharsAndLines(1)
    intMaxLines = varMaxCharsAndLines(2)
    
    '行ごとの詳細な結果表示（デバッグ用）は行わない設定にします。
    blnDispRslt = False

    'アクティブなウィンドウに対する操作を開始します。
    With ActiveWindow
        'ビューモードを「印刷レイアウト」表示に設定します（Rectangles取得に必須）。
        .View.Type = wdPrintView
            
        'アクティブなウィンドウ枠（Pane）に対する操作を開始します。
        With .ActivePane
            'ドキュメントの総ページ数を取得します。
            lngLastPage = .Pages.Count
            
            '1ページ目から最終ページまで、ページごとのループ処理を行います。
            For i = 1 To lngLastPage
                lngLines = 0 'ページ内の行数カウンタをリセット
                    
                '現在のページに含まれるRectangleの総数を取得します。
                lngLastRect = .Pages(i).Rectangles.Count
                
                'ページ内の各Rectangleに対してループ処理を行います。
                For j = 1 To lngLastRect
                    'エラーハンドリング（存在しないRectangleへのアクセス回避）
                    On Error GoTo HandleErr
                        lngLastLine = .Pages(i).Rectangles(j).Lines.Count
                    On Error GoTo 0
                   
                    'Rectangle内の各行に対してループ処理を行います。
                    For k = 1 To lngLastLine
                        '【リファクタリング】Selectを使用せず、直接Rangeオブジェクトを取得します。
                        'これにより処理速度が向上し、画面のちらつきを防止できます。
                        Set rngLine = .Pages(i).Rectangles(j).Lines(k).Range
                        strText = rngLine.Text
                        
                        '【リファクタリング】文字数計算ロジックを関数に分離しました。
                        lngChars = GetAdjustedCharCount(strText)
                        
                        '箇条書き等の段落番号が含まれる場合の補正処理
                        'Rectangle内の1行目であり、かつ箇条書き番号が存在する場合
                        If k = 1 And .Pages(i).Rectangles(j).Range.ListFormat.ListString <> "" Then
                            lngChars = lngChars + 1
                        End If
                        
                        '文字数が0より大きい場合（空行でない場合）、有効な行としてカウントします。
                        If lngChars > 0 Then
                            lngLines = lngLines + 1
                        End If
                    
                        'デバッグ用表示
                        If blnDispRslt = True Then
                            MsgBox "行数: " & lngLines & vbCr & "字数：" & lngChars, vbOKOnly, "VA内容証明郵便"
                        End If
                                            
                        '文字数チェック
                        If lngChars > intMaxChars Then
                            'エラー時は画面更新を戻してからメッセージを表示
                            Application.ScreenUpdating = True
                            Application.System.Cursor = wdCursorNormal
                            '該当箇所を選択してユーザーに場所を知らせる
                            rngLine.Select
                            MsgBox "字数が" & intMaxChars & "字を超えています。", vbCritical, "VA内容証明郵便"
                            Exit Sub
                        End If
                    Next
'エラー発生時のジャンプ先ラベル
Continue:
                Next
            
                'ページ行数チェック
                If lngLines > intMaxLines Then
                    Application.ScreenUpdating = True
                    Application.System.Cursor = wdCursorNormal
                    'ページの先頭付近を選択（概算位置）
                    .Pages(i).Rectangles(1).Range.Select
                    MsgBox "行数が" & intMaxLines & "行を超えています。", vbCritical, "VA内容証明郵便"
                    Exit Sub
                End If
            Next
        End With
    End With

    '選択範囲を解除し、カーソル位置のみの状態にします。
    Selection.Collapse

    '画面更新を再開します。
    Application.ScreenUpdating = True
    'マウスポインタを元に戻します。
    Application.System.Cursor = wdCursorNormal
    Application.StatusBar = ""

    '完了通知
    MsgBox "字数および行数のチェックを完了しました。", vbOKOnly, "VA内容証明郵便"

    Exit Sub

'エラーハンドリング用ラベル
HandleErr:
    Resume Continue

End Sub

'-----------------------------------------------------------------------
' 文字数計算用関数（リファクタリングにより分離）
' 文字列を受け取り、内容証明郵便のルールに基づいた有効文字数を返します。
'-----------------------------------------------------------------------
Private Function GetAdjustedCharCount(ByVal strTarget As String) As Long
    Dim lngLen As Long
    
    '基本の文字数を取得
    lngLen = Len(strTarget)
    
    '全角スペースを除外
    lngLen = lngLen - (Len(strTarget) - Len(Replace(strTarget, "　", "")))
    '半角スペースを除外
    lngLen = lngLen - (Len(strTarget) - Len(Replace(strTarget, " ", "")))
    
    '閉じ括弧 "）" を除外（簡易対応）
    lngLen = lngLen - (Len(strTarget) - Len(Replace(strTarget, "）", "")))
    '半角閉じ括弧 ")" を除外
    lngLen = lngLen - (Len(strTarget) - Len(Replace(strTarget, ")", "")))
    
    '制御文字（垂直タブ、改ページ、復帰）が含まれる場合の補正
    If InStr(strTarget, Chr(11)) > 0 Or _
       InStr(strTarget, Chr(12)) > 0 Or _
       InStr(strTarget, Chr(13)) > 0 Then
        lngLen = lngLen - 1
    End If
    
    GetAdjustedCharCount = lngLen
End Function

'-----------------------------------------------------------------------
' ヘッダー・フッター削除および判定用サブルーチン
' 対象のRangeにコンテンツが含まれている場合、フラグをTrueにして削除します。
'-----------------------------------------------------------------------
Private Sub DeleteAndCheck(ByVal rngTarget As Range, ByRef blnDeleted As Boolean)
    '文字列の長さが1より大きい（改行コードのみでない）、または図形等が含まれている場合
    If Len(rngTarget.Text) > 1 Or _
       rngTarget.ShapeRange.Count > 0 Or _
       rngTarget.InlineShapes.Count > 0 Then
        blnDeleted = True
    End If
    
    'コンテンツを削除
    rngTarget.Delete
End Sub

'-----------------------------------------------------------------------
' リボン用コールバックプロシージャ
'-----------------------------------------------------------------------
Sub OnBtnCheckCharsAndLines(control As IRibbonControl)
    Call CheckCharsAndLines
End Sub
