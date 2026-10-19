'================================================================================
' クラス名:     clsPreFormatter
' 概要:         文書の書式整形（表スタイル適用、インデント調整）を行うクラス
'               メイン処理の前に実行される「事前処理」を担当します。
' 作成日:       2025/11/28
'================================================================================
Option Explicit ' 変数の明示的な宣言を強制する（未宣言変数の使用を防ぐ）

'================================================================================
' 機能:         文書内のすべての表の段落に指定したスタイルを適用します。
' 引数:         targetStyleName (Required) - 適用するスタイル名（例："行間詰め"）
'               targetDoc (Optional)       - 対象のドキュメント。省略時はActiveDocument
'================================================================================
Public Sub ApplyTableStyle(ByVal targetStyleName As String, Optional ByVal targetDoc As Document = Nothing)
    ' 引数のドキュメントが指定されていない（Nothingの）場合
    If targetDoc Is Nothing Then
        ' アクティブなドキュメントを対象として設定する
        Set targetDoc = ActiveDocument
    End If
    
    ' エラーが発生した場合の処理（HandleErrラベルへジャンプ）を有効にする
    On Error GoTo HandleErr
    
    ' 対象ドキュメント内に表が1つも存在しない場合
    If targetDoc.Tables.count = 0 Then
        ' 処理を行わずにプロシージャを終了する
        Exit Sub
    End If
    
    '--- 画面更新の制御 ---
    ' 現在の画面更新設定（True/False）を一時変数に保存する
    Dim isScreenUpdating As Boolean
    isScreenUpdating = Application.ScreenUpdating
    
    ' 処理中の画面のちらつきを防ぐため、画面更新を停止する
    Application.ScreenUpdating = False
    
    '--- 変数宣言 ---
    Dim tbl As Table                    ' ループ用：個々の表オブジェクト
    Dim pars As Paragraphs              ' 表内の全段落コレクション
    Dim par As Paragraph                ' ループ用：個々の段落オブジェクト
    Dim fntBefore As Font               ' スタイル変更前のフォント情報を保持する変数
    Dim parAlign As WdParagraphAlignment ' スタイル変更前の段落配置を保持する変数
    
    ' ドキュメント内のすべての表に対してループ処理を行う
    For Each tbl In targetDoc.Tables
        ' 現在の表（tbl）に含まれるすべての段落を取得する
        Set pars = tbl.Range.Paragraphs
        
        ' 表内の各段落に対してループ処理を行う
        For Each par In pars
            ' 現在の段落のスタイルが、適用したいスタイル（targetStyleName）と異なる場合のみ処理する
            If par.Style.NameLocal <> targetStyleName Then
                
                '--- 書式の退避 ---
                ' 現在のフォント情報を複製（ディープコピー）して保存する
                ' ※単なる代入ではなくDuplicateを使うことで、参照ではなく実体のコピーを作成する
                Set fntBefore = par.Range.Font.Duplicate
                
                ' 現在の段落の配置（左揃え、中央揃えなど）を保存する
                parAlign = par.alignment
                
                '--- スタイルの適用 ---
                ' 指定されたスタイル名が存在しない場合のエラーを無視するため、一時的にエラートラップを変更
                On Error Resume Next
                
                ' 段落に指定されたスタイルを適用する
                par.Style = targetStyleName
                
                ' スタイル適用でエラーが発生しなかった場合（スタイルが存在し適用できた場合）
                If Err.Number = 0 Then
                    '--- 書式の復元 ---
                    ' 保存しておいたフォント情報を適用し直す（見た目を維持するため）
                    par.Range.Font = fntBefore
                    
                    ' 保存しておいた段落配置を適用し直す
                    par.alignment = parAlign
                End If
                
                ' エラー処理を元の設定（HandleErrへジャンプ）に戻す
                On Error GoTo HandleErr
                
                ' フォント情報保持用のオブジェクト変数を解放する
                Set fntBefore = Nothing
            End If
        Next par ' 次の段落へ
    Next tbl ' 次の表へ

Cleanup:
    '--- 終了処理 ---
    ' 画面更新の設定を、処理開始前の状態に戻す
    If isScreenUpdating Then Application.ScreenUpdating = True
    
    ' 使用したオブジェクト変数を解放する
    Set par = Nothing
    Set pars = Nothing
    Set tbl = Nothing
    
    ' プロシージャを正常終了する
    Exit Sub

HandleErr:
    '--- エラー処理 ---
    ' エラーが発生してもメッセージは表示せず、安全に終了処理を行う
    Resume Cleanup
End Sub

'================================================================================
' 機能:         文書内の全段落を対象に、行頭の不要なスペースをインデントに自動変換します。
' 引数:         sngAdvWidth (Required)     - 1文字あたりのポイント数（アドバンス幅）
'               excludeStyleName (Required)- 除外するスタイル名（例："行間詰め"）
'               targetDoc (Optional)       - 対象のドキュメント。省略時はActiveDocument
'================================================================================
Public Sub ReplaceSpaceWithIndent(ByVal sngAdvWidth As Single, ByVal excludeStyleName As String, Optional ByVal targetDoc As Document = Nothing)
    ' 引数のドキュメントが指定されていない（Nothingの）場合
    If targetDoc Is Nothing Then
        ' アクティブなドキュメントを対象として設定する
        Set targetDoc = ActiveDocument
    End If

    ' エラーが発生した場合の処理（HandleErrラベルへジャンプ）を有効にする
    On Error GoTo HandleErr

    '--- 画面更新の制御 ---
    ' 現在の画面更新設定（True/False）を一時変数に保存する
    Dim isScreenUpdating As Boolean
    isScreenUpdating = Application.ScreenUpdating
    
    ' 処理中の画面のちらつきを防ぐため、画面更新を停止する
    Application.ScreenUpdating = False

    '--- スマート カット アンド ペーストの設定の退避 ---
    ' 有効なままだと、Range.Delete で行頭の半角スペースが削除されないことがあるため、
    ' 現在の設定を退避してから無効にする（終了処理で必ず元に戻す）
    Dim blnSmartCutPaste As Boolean     ' 退避した設定値
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
    ' エラー処理をHandleErrラベルへのジャンプに戻す
    On Error GoTo HandleErr

    '--- 変数宣言 ---
    Dim pars As Paragraphs              ' ドキュメント内の全段落コレクション
    Dim par As Paragraph                ' ループ用：個々の段落オブジェクト
    Dim sngSpaceCnt As Single           ' インデント量（文字数換算）を計算するための変数
    Dim lngLenBefore As Long            ' スペースを削除する前の段落の文字数
    
    ' ドキュメント内の全段落を取得する
    Set pars = targetDoc.Paragraphs

    ' すべての段落に対してループ処理を行う
    For Each par In pars
        ' 段落内の文字数が1より大きい（改行マークだけでなく文字がある）場合のみ処理する
        If Len(par.Range.Text) > 1 Then
            
            ' 段落のスタイルが除外対象（excludeStyleName）と一致する場合
            If par.Style.NameLocal = excludeStyleName Then
                ' 処理をスキップして次の段落へ（Continueラベルへジャンプ）
                GoTo Continue
            End If

            ' 現在の段落（par）に対する操作を開始
            With par
                '--- 既存インデントの計算 ---
                ' アドバンス幅（sngAdvWidth）が正の値であることを確認する
                If sngAdvWidth > 0 Then
                    ' 現在設定されている「1行目のインデント」を文字数に換算して取得する
                    sngSpaceCnt = .FirstLineIndent / sngAdvWidth
                Else
                    ' アドバンス幅が0以下の場合は計算できないため、0とする（ゼロ除算防止）
                    sngSpaceCnt = 0
                End If

                '--- 行頭スペースの処理 ---
                ' 段落の1文字目が「全角スペース」または「半角スペース」以外になるまでループする
                Do Until .Range.Characters(1) <> "　" And .Range.Characters(1) <> " "
                    ' 削除前の段落の文字数を記録する
                    lngLenBefore = Len(.Range.Text)
                    
                    ' 1文字目が全角スペースの場合
                    If .Range.Characters(1) = "　" Then
                        ' インデント文字数カウンタに 1文字分 加算する
                        sngSpaceCnt = sngSpaceCnt + 1
                        
                        ' 全角スペースを削除する（Range操作なのでカーソルは動かない）
                        .Range.Characters(1).Delete
                    
                    ' 1文字目が半角スペースの場合
                    ElseIf .Range.Characters(1) = " " Then
                        ' インデント文字数カウンタに 0.5文字分 加算する
                        sngSpaceCnt = sngSpaceCnt + 0.5
                        
                        ' 半角スペースを削除する（スマート カット アンド ペーストを無効にしているため削除できる）
                        .Range.Characters(1).Delete
                    End If

                    ' スペースを削除できなかった場合は、無限ループを避けるため繰り返しを終了する
                    If Len(.Range.Text) >= lngLenBefore Then Exit Do
                Loop

                '--- インデントの適用 ---
                ' 計算した合計文字数（sngSpaceCnt）にアドバンス幅を掛けて、ポイント単位のインデントを設定する
                .FirstLineIndent = sngSpaceCnt * sngAdvWidth
            End With
        End If
        
        ' スキップ時のジャンプ先ラベル
Continue:
    Next par ' 次の段落へ

Cleanup:
    '--- 終了処理 ---
    ' スマート カット アンド ペーストの設定を、処理開始前の状態に戻す（正常終了・エラー終了の両方で必ず行う）
    If blnSmartCutPasteSaved Then
        On Error Resume Next
        Options.SmartCutPaste = blnSmartCutPaste
        On Error GoTo 0
    End If

    ' 画面更新の設定を、処理開始前の状態に戻す
    If isScreenUpdating Then Application.ScreenUpdating = True

    ' 使用したオブジェクト変数を解放する
    Set par = Nothing
    Set pars = Nothing
    
    ' プロシージャを正常終了する
    Exit Sub

HandleErr:
    '--- エラー処理 ---
    ' エラーが発生してもメッセージは表示せず、安全に終了処理を行う
    Resume Cleanup
End Sub