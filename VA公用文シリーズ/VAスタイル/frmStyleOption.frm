'//==================================================================================================
'// 機能: ユーザー設定フォームのイベントを処理します。
'//
'// 概要: このモジュールは、ユーザーフォームの初期化、OKボタン、キャンセルボタンの
'//       クリックイベントを管理します。設定値はVariablesManagerクラスを通じて
'//       グローバルに管理され、SettingsManagerクラスを通じてファイルに保存されます。
'//==================================================================================================
Option Explicit

' 設定値を管理するためのクラスインスタンスをモジュールレベルで宣言
Dim Variables As VariablesManager

'//==================================================================================================
'// プロシージャ: UserForm_Initialize
'// 機能:       フォームがロードされたときにコントロールの初期値を設定します。
'//
'// 概要:       フォームが表示される直前に実行されます。
'//             VariablesManagerから現在の設定値を読み込み、各チェックボックスの
'//             オン/オフ状態に反映させます。
'//==================================================================================================
Private Sub UserForm_Initialize()
    '変数オブジェクトを生成する
    Set Variables = New VariablesManager
    
    '--- 「フォントをゴシックにする」の初期値を設定 ---
    If Variables.strChangeFontType = "Enable" Then
        CheckBox1 = True
    Else
        CheckBox1 = False
    End If
    
    '--- 「フォントを拡大する」の初期値を設定 ---
    If Variables.strChangeFontSize = "Enable" Then
        CheckBox2 = True
    Else
        CheckBox2 = False
    End If
    
    '--- 「段落前の間隔を1行開ける」の初期値を設定 ---
    If Variables.strLineUnitBeforeOne = "Enable" Then
        CheckBox11 = True
    Else
        CheckBox11 = False
    End If
    
    '--- 「次の段落と分離しない」の初期値を設定 ---
    If Variables.strKeepParagraphWithNext = "Enable" Then
        CheckBox12 = True
    Else
        CheckBox12 = False
    End If
    
    '--- 「段落を分割しない」の初期値を設定 ---
    If Variables.strKeepParagraphTogether = "Enable" Then
        CheckBox13 = True
    Else
        CheckBox13 = False
    End If

    '変数オブジェクトを解放する
    Set Variables = Nothing
End Sub

'//==================================================================================================
'// プロシージャ: CommandButton1_Click
'// 機能:       「OK」ボタンがクリックされたときの処理を実行します。
'//
'// 概要:       各チェックボックスの状態を読み取り、VariablesManagerの設定値を更新します。
'//             その後、SettingsManagerを使用して設定をファイルに保存し、フォームを閉じます。
'//==================================================================================================
Private Sub CommandButton1_Click()
    '変数オブジェクトを生成する
    Set Variables = New VariablesManager
    
    '--- 「フォントをゴシックにする」の設定を更新 ---
    If CheckBox1.value = True Then
        Variables.strChangeFontType = "Enable"
    Else
        Variables.strChangeFontType = "Disable"
    End If
    
    '--- 「フォントを拡大する」の設定を更新 ---
    If CheckBox2.value = True Then
        Variables.strChangeFontSize = "Enable"
    Else
        Variables.strChangeFontSize = "Disable"
    End If
    
    '--- 「段落前に間隔を1行空ける」の設定を更新 ---
    If CheckBox11.value = True Then
        Variables.strLineUnitBeforeOne = "Enable"
    Else
        Variables.strLineUnitBeforeOne = "Disable"
    End If
    
    '--- 「次の段落と分離しない」の設定を更新 ---
    If CheckBox12.value = True Then
        Variables.strKeepParagraphWithNext = "Enable"
    Else
        Variables.strKeepParagraphWithNext = "Disable"
    End If
    
    '--- 「段落を分割しない」の設定を更新 ---
    If CheckBox13.value = True Then
        Variables.strKeepParagraphTogether = "Enable"
    Else
        Variables.strKeepParagraphTogether = "Disable"
    End If
    
    '--- 設定をファイルに保存する準備 ---
    Dim varKeys(1 To 5) As Variant
    Dim varValues(1 To 5) As Variant
    
    ' 保存するキーと値を配列に設定
    varKeys(1) = "Change Font Type"
    varValues(1) = Variables.strChangeFontType
    
    varKeys(2) = "Change Font Size"
    varValues(2) = Variables.strChangeFontSize
    
    varKeys(3) = "Line Unit Before One"
    varValues(3) = Variables.strLineUnitBeforeOne
    
    varKeys(4) = "Paragraph Separation"
    varValues(4) = Variables.strKeepParagraphWithNext
    
    varKeys(5) = "Paragraph Split"
    varValues(5) = Variables.strKeepParagraphTogether
    
    ' SettingsManagerクラスのメソッドを呼び出して設定をファイルに保存
    Dim settingsMgr As New SettingsManager
    settingsMgr.SaveSettingsToProperty varKeys, varValues
    Set settingsMgr = Nothing
    
    '変数オブジェクトを解放する
    Set Variables = Nothing
    
    ' フォームを閉じる
    Unload Me
    
End Sub

'//==================================================================================================
'// プロシージャ: CommandButton2_Click
'// 機能:       「キャンセル」ボタンがクリックされたときの処理を実行します。
'//
'// 概要:       設定を保存せずにフォームを閉じます。
'//==================================================================================================
Private Sub CommandButton2_Click()
    ' フォームを閉じる
    Unload Me
End Sub
