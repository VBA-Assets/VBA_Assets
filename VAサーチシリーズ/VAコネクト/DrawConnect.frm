Option Explicit
'=================================================================================
' モジュール名: DrawConnect
' 概要: コネクタ作成ツールのユーザーインターフェース（画面）を制御するモジュールです。
'       ボタンがクリックされた時の動作などを記述しています。
'=================================================================================

'---------------------------------------------------------------------------------
' プロシージャ名: CommandButton1_Click
' 概要: 実行ボタン（CommandButton1）がクリックされた時の処理です。
'       選択範囲を読み取り、設定値を取得して、コネクタ作成処理を呼び出します。
'---------------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    '始点と終点のセル範囲を格納する変数を宣言する
    Dim rngBegin, rngEnd
    
    'ユーザーが選択した設定（線種、方向、開始位置、角調整の有無）を格納する変数を宣言する
    Dim strType, strDirection, strStart, blnAlignElbow
    
    '現在選択されているオブジェクトが「セル範囲（Range）」でない場合はメッセージを表示して終了する
    '（図形などを選択した状態で実行するのを防ぐため）
    If TypeName(Selection) <> "Range" Then
        MsgBox "始点と終点のセルを選択してください"
        Exit Sub
    End If
    
    '選択されている範囲のアドレスをカンマ区切りで配列に分割して格納する
    '例："$A$1,$C$5" というアドレス文字列を ["$A$1", "$C$5"] という配列にする
    Dim varAddress As Variant
    varAddress = Split(Selection.Address, ",")
    
    '選択範囲が1箇所しかない（配列の要素数が0＝カンマがない）場合はメッセージを表示して終了する
    '（始点と終点の2箇所が選択されていないと線を引けないため）
    If UBound(varAddress) = 0 Then
        MsgBox "始点と終点のセルを選択してください"
        Exit Sub
    End If
    
    '配列の1つ目（インデックス0）を始点のセル範囲として設定する
    Set rngBegin = Range(varAddress(0))
    
    '配列の2つ目（インデックス1）を終点のセル範囲として設定する
    Set rngEnd = Range(varAddress(1))
    
    '-----------------------------------------------------
    ' フォーム上の選択状態を確認して設定変数に値をセットする
    '-----------------------------------------------------
    
    '線種の設定：OptionButton1が選択されている場合
    If OptionButton1.Value = True Then
        '「細線」を表す文字列をセット
        strType = "Normal"
    End If
    
    '線種の設定：OptionButton2が選択されている場合
    If OptionButton2.Value = True Then
        '「点線」を表す文字列をセット
        strType = "Dash"
    End If
    
    '線種の設定：OptionButton3が選択されている場合
    If OptionButton3.Value = True Then
        '「太線」を表す文字列をセット
        strType = "Heavy"
    End If
    
    '線種の設定：OptionButton4が選択されている場合
    If OptionButton4.Value = True Then
        '「二重線」を表す文字列をセット
        strType = "Double"
    End If
    
    '方向の設定：OptionButton5が選択されている場合
    If OptionButton5.Value = True Then
        '「右方向」を表す文字列をセット
        strDirection = "Right"
    End If
    
    '方向の設定：OptionButton6が選択されている場合
    If OptionButton6.Value = True Then
        '「下方向」を表す文字列をセット
        strDirection = "Down"
    End If
    
    '開始位置の設定：OptionButton7が選択されている場合
    If OptionButton7.Value = True Then
        '「外周」を表す文字列をセット
        strStart = "Edge"
    End If
    
    '開始位置の設定：OptionButton8が選択されている場合
    If OptionButton8.Value = True Then
        '「中央」を表す文字列をセット
        strStart = "Center"
    End If
    
    '角の位置調整：CheckBox1にチェックが入っている場合
    If CheckBox1.Value = True Then
        '「調整する」を表すTrueをセット
        blnAlignElbow = True
    End If
    
    'ConnectモジュールのAddConectorプロシージャを呼び出して、コネクタを作成する
    'ここまでに設定した変数（rngBegin, rngEnd, strType...）を引数として渡す
    Call Connect.AddConector(rngBegin, rngEnd, strType, strDirection, strStart, blnAlignElbow)
End Sub

'---------------------------------------------------------------------------------
' プロシージャ名: CommandButton2_Click
' 概要: 終了ボタン（CommandButton2）がクリックされた時の処理です。
'       フォームを閉じます。
'---------------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    '自分自身（Me = このユーザーフォーム）をメモリから解放して閉じる
    'ESCキーを押したときにフォームを閉じるためのボタン（非表示設定の場合あり）の処理です。
    Unload Me
End Sub
