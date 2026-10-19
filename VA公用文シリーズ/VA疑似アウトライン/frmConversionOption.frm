Option Explicit
Dim Variables As VariablesManager

'================================================================================
' 機能: ユーザーフォームが表示される際にコントロールの初期値を設定します。
'       保存されている設定値を読み込み、番号スタイルのオプションボタンと
'       「見出し番号を全角1文字に調節する」のチェックボックスに反映させます。
' 引数: なし
'================================================================================
Private Sub UserForm_Initialize()
    '変数オブジェクトを生成する
    Set Variables = New VariablesManager
    
    '「行間詰めスタイルも変換する」の初期値を設定する
    If Variables.blnNoParagraphSpace = True Then
        CheckBox1 = True
    Else
        CheckBox1 = False
    End If

    '「見出し番号を全角1文字に調節する」の初期値を設定する
    If Variables.blnOneFullWidth = True Then
        CheckBox11 = True
    Else
        CheckBox11 = False
    End If

    '変数オブジェクトを解放する
    Set Variables = Nothing
End Sub

'================================================================================
' 機能: 「OK」ボタンがクリックされた際の処理を実行します。
'       フォームで選択された設定内容をグローバル変数に格納すると共に、
'       設定ファイルに保存します。その後、フォームを閉じます。
' 引数: なし
'================================================================================
Private Sub CommandButton1_Click()
    '「OK」ボタンが押された場合に処理を行います。
       
    '変数オブジェクトを生成する
    Set Variables = New VariablesManager
    
    '「行間詰めスタイルも変換する」を設定する
    If CheckBox1.value = True Then
        Variables.strNoParagraphSpace = "Enable"
    Else
        Variables.strNoParagraphSpace = "Disable"
    End If
    
    '「見出し番号を全角1文字に調節する」を設定する
    If CheckBox11.value = True Then
        Variables.strOneFullWidth = "Enable"
    Else
        Variables.strOneFullWidth = "Disable"
    End If
    Dim varKeys(1 To 2) As Variant
    Dim varValues(1 To 2) As Variant
    
    '保存するキーと値を配列に設定
    varKeys(1) = "Convert No Paragraph Space"
    varValues(1) = Variables.strNoParagraphSpace
    
    varKeys(2) = "Adjust heading number"
    varValues(2) = Variables.strOneFullWidth
    
    'クラスのメソッドを呼び出して保存
    Dim settingsMgr As New SettingsManager
    settingsMgr.SaveSettingsToProperty varKeys, varValues
    Set settingsMgr = Nothing
    
    'フォームを閉じる
    Unload Me

    '変数オブジェクトを解放する
    Set Variables = Nothing
End Sub

'================================================================================
' 機能: 「キャンセル」ボタンがクリックされた際の処理を実行します。
'       設定を変更・保存せずにフォームを閉じます。
' 引数: なし
'================================================================================
Private Sub CommandButton2_Click()
    
    'フォームを閉じる
    Unload Me
End Sub