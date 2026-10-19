'================================================================================
' モジュール: frmDividendCalc
' 機能:      配当率計算ユーザーフォーム
' 概要:      ユーザーインターフェースを提供し、セルの選択、目標値の入力、
'            計算実行のトリガー、および結果の表示を制御します。
'================================================================================
Option Explicit

'================================================================================
' プロシージャ: btnSelectTotal_Click
' 機能:      合計値セルを選択します。
' 概要:      Application.InputBoxを使用してユーザーにセルを選択させ、
'            選択されたセルのアドレスを合計値テキストボックスに設定します。
'================================================================================
Private Sub btnSelectTotal_Click()
    '変数を設定する：選択されたセル範囲
    Dim rng As Range
    
    'エラー処理：キャンセル時にエラーとならないようにする
    On Error Resume Next
    
    'InputBoxを表示し、ユーザーにセルを選択させる（Type:=8 はセル参照）
    Set rng = Application.InputBox("合計値のセルを選択してください。", "セル選択", Type:=8)
    
    'エラー処理を解除
    On Error GoTo 0
    
    'セルが選択された場合（キャンセルされていない場合）
    If Not rng Is Nothing Then
        '選択されたセルのアドレスをテキストボックスに設定する
        Me.txtTotalCell.Value = rng.Address(External:=False) 'シート名は含めずアドレスのみ
    End If
End Sub

'================================================================================
' プロシージャ: btnSelectRate_Click
' 機能:      配当率セルを選択します。
' 概要:      Application.InputBoxを使用してユーザーにセルを選択させ、
'            選択されたセルのアドレスを配当率テキストボックスに設定します。
'================================================================================
Private Sub btnSelectRate_Click()
    '変数を設定する：選択されたセル範囲
    Dim rng As Range
    
    'エラー処理：キャンセル時にエラーとならないようにする
    On Error Resume Next
    
    'InputBoxを表示し、ユーザーにセルを選択させる（Type:=8 はセル参照）
    Set rng = Application.InputBox("配当率を入力するセルを選択してください。", "セル選択", Type:=8)
    
    'エラー処理を解除
    On Error GoTo 0
    
    'セルが選択された場合（キャンセルされていない場合）
    If Not rng Is Nothing Then
        '選択されたセルのアドレスをテキストボックスに設定する
        Me.txtRateCell.Value = rng.Address(External:=False)
    End If
End Sub

'================================================================================
' プロシージャ: txtTarget_AfterUpdate
' 機能:      目標金額のフォーマット
' 概要:      入力確定後に数値を3桁区切りのカンマ形式に変換します。
'            例: 1000000 -> 1,000,000
'================================================================================
Private Sub txtTarget_AfterUpdate()
    '入力値が数値であれば、カンマ区切り形式に整形する
    If IsNumeric(Me.txtTarget.Value) Then
        Me.txtTarget.Value = Format(Me.txtTarget.Value, "#,##0")
    End If
End Sub

'================================================================================
' プロシージャ: spnMaxDigits_Change
' 機能:      最大桁数の変更（スピンボタン）
' 概要:      スピンボタンの値をテキストボックスに反映させます。
'================================================================================
Private Sub spnMaxDigits_Change()
    'スピンボタンの値をテキストボックスに設定する
    Me.txtMaxDigits.Value = Me.spnMaxDigits.Value
End Sub

'================================================================================
' プロシージャ: btnCalc_Click
' 機能:      計算を実行します。
' 概要:      入力値の検証を行った後、標準モジュールのCalcDivRate関数を
'            呼び出して配当率を計算し、結果に応じたメッセージを表示します。
'================================================================================
Private Sub btnCalc_Click()
    '変数を設定する：合計配当金額のセル範囲オブジェクト
    Dim rngTotal As Range
    '変数を設定する：配当率のセル範囲オブジェクト
    Dim rngRate As Range
    '変数を設定する：合計配当金額の目標値（金額型）
    Dim lngTarget As Currency
    '変数を設定する：配分率の小数点以下の最大桁数
    Dim iMax As Long
    '変数を設定する：計算結果の配当率
    Dim dblResultRate As Double
    '変数を設定する：計算後の合計金額
    Dim curResultTotal As Currency
    
    'エラーが発生した場合にエラー処理ルーチンへ飛ばす
    On Error GoTo ErrorHandler

    '--- 入力チェック ---

    '合計値セルが空欄かチェック
    If Me.txtTotalCell.Value = "" Then
        MsgBox "合計値のセルを指定してください。", vbExclamation
        Exit Sub
    End If
    
    '配当率セルが空欄かチェック
    If Me.txtRateCell.Value = "" Then
        MsgBox "配当率のセルを指定してください。", vbExclamation
        Exit Sub
    End If

    '入力された合計値のセル番地が正しいか確認し、オブジェクトにセットする
    Set rngTotal = Range(Me.txtTotalCell.Value)
    
    '入力された配当率のセル番地が正しいか確認し、オブジェクトにセットする
    Set rngRate = Range(Me.txtRateCell.Value)
    
    '目標金額が数値で入力されているか確認する（カンマ付きでもIsNumericはTrueになります）
    If Not IsNumeric(Me.txtTarget.Value) Then
        MsgBox "目標金額には数値を入力してください。", vbExclamation
        Exit Sub
    End If
    
    'フォームから目標値を読み込む（CCurはカンマ付き文字列も数値に変換します）
    lngTarget = CCur(Me.txtTarget.Value)
    
    '最大桁数が数値で入力されているか確認する
    If Not IsNumeric(Me.txtMaxDigits.Value) Then
        MsgBox "最大桁数には数値を入力してください。", vbExclamation
        Exit Sub
    End If
    
    'フォームから小数点以下の最大桁数を読み込む
    iMax = CLng(Me.txtMaxDigits.Value)

    '最大桁数が設定されていない（0以下）の場合には終了する
    If iMax <= 0 Then
        MsgBox "最大桁数は1以上の整数を指定してください。"
        Exit Sub
    End If

    '--- 計算処理の呼び出し ---
    
    '標準モジュールのFunctionを呼び出して計算を実行する
    '関数名を変更しました: ExecuteDividendCalculation -> CalcDivRate
    dblResultRate = CalcDivRate(rngTotal, rngRate, lngTarget, iMax)
    
    '--- 結果の判定 ---
    
    '計算後の合計値を取得する
    curResultTotal = rngTotal.Value
    
    '計算結果が目標値と一致するか確認する
    If curResultTotal = lngTarget Then
        '成功メッセージを表示する
        MsgBox "解答が見つかりました。"
    Else
        '失敗メッセージを表示する
        MsgBox "小数点以下" & iMax & "桁までの間に解答が見つかりませんでした。"
        '（見つからなかった場合の配当率リセット処理はFunction内で行われています）
    End If
    
    'サブルーチンを終了する
    Exit Sub

'エラー処理ルーチン
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

'================================================================================
' プロシージャ: btnCancel_Click
' 機能:      フォームを閉じます。
' 概要:      ユーザーフォームをメモリからアンロードし、処理を終了します。
'================================================================================
Private Sub btnCancel_Click()
    'フォームをメモリからアンロードして閉じる
    Unload Me
End Sub