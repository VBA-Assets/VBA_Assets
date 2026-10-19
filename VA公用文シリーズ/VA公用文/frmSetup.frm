Option Explicit
'変数管理クラスを格納する変数を宣言する
Dim Variables As clsVariablesManager

'================================================================================
' イベント:     フォームが初期化されるとき（表示される直前）に実行されます。
' 機能:         フォーム上の各コントロールを、保存されている設定値に基づいて
'               初期表示します。
' 概要:         設定管理クラス（clsVariablesManager）を呼び出して前回保存された
'               設定を読み込みます。その設定内容に応じて、文書の種類、
'               ページ番号、行番号の各オプションボタンとチェックボックスの
'               初期状態を適切に設定します。
'================================================================================
Private Sub UserForm_Initialize()
    '変数オブジェクトを生成する
    Set Variables = New clsVariablesManager
    
    '文書の種類オプションボタンの初期値を設定する
    Select Case Variables.strDocumentType
        Case "Saiban"
            OptionButton1.value = True
        Case "Naiyou"
            OptionButton2.value = True
        Case "Kouyou"
            OptionButton3.value = True
        Case "Keiyaku1"
            OptionButton4.value = True
        Case "Keiyaku2"
            OptionButton5.value = True
        Case "Business"
            OptionButton6.value = True
        Case "Technical"
            OptionButton7.value = True
        Case Else
            OptionButton1.value = True
    End Select

    'ページ番号オプションボタンの初期値を設定する
    Select Case Variables.strPageNumber
        Case "TopRight"
            OptionButton11.value = True
        Case "BottomCenter"
            OptionButton12.value = True
        Case "BottomCenterDecorated"
            OptionButton13.value = True
        Case "BottomRight"
            OptionButton14.value = True
        Case "None"
            OptionButton15.value = True
        Case Else
            OptionButton15.value = True
    End Select

    '行番号チェックボックスの初期値を設定する
    Select Case Variables.strLineNumber
        Case "Show"
            CheckBox1.value = True
        Case Else
            CheckBox1.value = False
    End Select
    
    '変数オブジェクトを解放する
    Set Variables = Nothing
End Sub

'================================================================================
' イベント:     「OK」ボタン（CommandButton1）がクリックされたときに実行されます。
' 機能:         フォームで選択された設定内容を保存し、文書の書式設定を
'               開始します。
' 概要:         1. フォーム上の各コントロールの状態を読み取り、設定管理変数に
'                  格納します。
'               2. 格納した設定内容を、次回起動時のためにファイルへ保存します。
'               3. 文書全体の書式設定を行うメインプロシージャ（ApplyFormatSettings）を
'                  呼び出します。
'================================================================================
Private Sub CommandButton1_Click()
    '変数オブジェクトを生成する
    Set Variables = New clsVariablesManager
    
    '--- フォームの選択内容を変数に格納する ---
    '文書の種類を設定する
    If OptionButton1.value = True Then
        Variables.strDocumentType = "Saiban"
    ElseIf OptionButton2.value = True Then
        Variables.strDocumentType = "Naiyou"
    ElseIf OptionButton3.value = True Then
        Variables.strDocumentType = "Kouyou"
    ElseIf OptionButton4.value = True Then
        Variables.strDocumentType = "Keiyaku1"
    ElseIf OptionButton5.value = True Then
        Variables.strDocumentType = "Keiyaku2"
    ElseIf OptionButton6.value = True Then
        Variables.strDocumentType = "Business"
    ElseIf OptionButton7.value = True Then
        Variables.strDocumentType = "Technical"
    End If
    
    'ページ番号を設定する
    If OptionButton11.value = True Then
        Variables.strPageNumber = "TopRight"
    ElseIf OptionButton12.value = True Then
        Variables.strPageNumber = "BottomCenter"
    ElseIf OptionButton13.value = True Then
        Variables.strPageNumber = "BottomCenterDecorated"
    ElseIf OptionButton14.value = True Then
        Variables.strPageNumber = "BottomRight"
    ElseIf OptionButton15.value = True Then
        Variables.strPageNumber = "None"
    End If
    
    '行番号を設定する
    If CheckBox1.value = True Then
        Variables.strLineNumber = "Show"
    ElseIf CheckBox1.value = False Then
        Variables.strLineNumber = "Hide"
    End If
    
    'サンプル追加を設定する
    If CheckBox2.value = True Then
        Variables.strSample = "Add"
    ElseIf CheckBox2.value = False Then
        Variables.strSample = "DoNotAdd"
    End If
    
    '--- 設定をファイルに保存する ---
    Dim varKeys(1 To 4) As Variant
    Dim varValues(1 To 4) As Variant
    
    '保存するキーと値を配列に設定
    varKeys(1) = "DocumentType"
    varValues(1) = Variables.strDocumentType
    
    varKeys(2) = "PageNumber"
    varValues(2) = Variables.strPageNumber
    
    varKeys(3) = "LineNumber"
    varValues(3) = Variables.strLineNumber
    
    varKeys(4) = "Sample"
    varValues(4) = Variables.strSample
    
    'クラスのメソッドを呼び出して保存
    Dim settingsMgr As New clsSettingsManager
    settingsMgr.SaveSettingsToProperty varKeys, varValues
    Set settingsMgr = Nothing
    
    '変数オブジェクトを解放する
    Set Variables = Nothing
    
    '--- 3. 文書のセットアップを実行する ---
    Call basSetup.ApplyFormatSettings
    
End Sub

'================================================================================
' イベント:     「キャンセル」ボタン（CommandButton2）がクリックされたときに
'               実行されます。
' 機能:         すべての処理をキャンセルし、フォームを閉じます。
'================================================================================
Private Sub CommandButton2_Click()
    'フォームを閉じる
    Unload Me
End Sub

'================================================================================
' イベント:     文書の種類で「内容証明郵便」（OptionButton2）が選択された、
'               または選択解除されたときに実行されます。
' 機能:         「内容証明郵便」の選択状態に応じて、他の設定オプションの
'               有効／無効を動的に切り替えます。
' 概要:         内容証明郵便が選択された場合、その書式要件に従い、
'               ページ番号と行番号のオプションを無効化して値を固定します。
'               選択が解除された場合は、それらのオプションを再度有効に戻します。
'================================================================================
Private Sub OptionButton2_Change()
    '内容証明郵便の場合
    If OptionButton2.value = True Then
        'ページ番号オプションの表示を内容証明郵便用に切り替える
        OptionButton11.Enabled = False
        OptionButton12.Enabled = False
        OptionButton13.Enabled = False
        OptionButton14.Enabled = False
        OptionButton15.value = True
    Else
        'ページ番号オプションの表示を一般用に切り替える
        OptionButton11.Enabled = True
        OptionButton12.Enabled = True
        OptionButton13.Enabled = True
        OptionButton14.Enabled = True
    End If

    '内容証明郵便の場合
    If OptionButton2.value = True Then
        '行番号オプションの表示を内容証明郵便用に切り替える
        CheckBox1.value = False
        CheckBox1.Enabled = False
    Else
        '行番号オプションの表示を一般用に切り替える
        CheckBox1.Enabled = True
    End If
End Sub