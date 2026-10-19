'//==================================================================================================
'// 機能: 疑似アウトラインとマークダウンの間の相互変換を制御します。
'//
'// 概要: このモジュールには、リボンからのコールバック、疑似アウトラインへの変換制御、
'//       マークダウンへの変換制御を行うプロシージャが含まれています。
'//==================================================================================================
Option Explicit

'================================================================================
' リボンコールバック (Ribbon Callbacks)
' 機能: リボンXML (CustomUI) からのボタンクリックイベントを受け取り、
'       対応する実処理プロシージャを呼び出します。
'================================================================================

' リボンの「疑似アウトライン変換」ボタンが押された時の処理
' onAction="rxbtnFakeOutline_click"
Sub rxbtnFakeOutline_click(control As IRibbonControl)
    Call ExecuteFakeOutline
End Sub

' リボンの「マークダウン変換」ボタンが押された時の処理
' onAction="rxbtnMarkdown_click"
Sub rxbtnMarkdown_click(control As IRibbonControl)
    Call ExecuteMarkdown
End Sub

' リボンの「設定」ボタンが押された時の処理
' onAction="rxbtnOption_click"
Sub rxbtnOption_click(control As IRibbonControl)
    Call ShowConversionOption
End Sub


'================================================================================
' 実処理プロシージャ (Actual Procedures)
' 機能: 実際の変換ロジックやフォーム表示の制御を行います。
'================================================================================

'================================================================================
' ShowConversionOption
' 機能: オプション設定用のユーザーフォーム「frmConvertOption」を表示します。
'================================================================================
Sub ShowConversionOption()
    ' 設定用フォームをモーダルで表示する
    frmConversionOption.Show
End Sub

'================================================================================
' ExecuteFakeOutline
' 機能: 選択範囲の有無を判定し、適切な疑似アウトライン変換処理を呼び出します。
'       - 選択範囲なし（カーソルのみ）: 文書全体を疑似アウトラインへ変換 (ConvertToFakeOutline)
'       - 選択範囲あり                : 選択部分を全角1文字書式へトグル (ToggleFakeOneChar)
'================================================================================
Sub ExecuteFakeOutline()
    
    ' Selection.Type を使用して、現在の選択状態を判定する
    ' wdSelectionIP は "Insertion Point"（挿入ポイント＝範囲選択なし）を意味します
    If Selection.Type = wdSelectionIP Then
        
        ' 選択範囲がない場合
        ' メインの疑似アウトライン変換プロシージャを呼び出す
        Call ConvertToFakeOutline
        
    Else
        
        ' 選択範囲がある場合（wdSelectionNormal 等）
        ' 全角1文字書式のトグル（設定/解除）プロシージャを呼び出す
        Call ToggleFakeOneChar
        
    End If

End Sub

'================================================================================
' ExecuteMarkdown
' 機能: マークダウンへの変換を呼び出します。
'================================================================================
Sub ExecuteMarkdown()
    ' メインのマークダウン変換プロシージャを呼び出す
    Call ConvertToMarkdown

End Sub