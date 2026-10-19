'================================================================================
' クラスモジュール: VariabesManager
' 機能: 変数を管理する
'================================================================================
Option Explicit

Public strChangeFontType As String               'フォントをゴシックにする
Public strChangeFontSize As String               'フォントを拡大する
Public strKeepParagraphWithNext As String        '次の段落と分離しない
Public strKeepParagraphTogether As String        '段落を分割しない
Public strLineUnitBeforeOne As String            '段落前の間隔を1行空ける

Public blnChangeFontType As Boolean              'フォントをゴシックにする
Public blnChangeFontSize As Boolean              'フォントサイズを拡大する
Public blnKeepParagraphWithNext As Boolean       '次の段落と分離しない
Public blnKeepParagraphTogether As Boolean       '段落を分割しない
Public blnLineUnitBeforeOne As Boolean           '段落前の間隔を1行空ける

'================================================================================
' 機能: クラスがインスタンス化される際に自動的に実行されます。
'       SetVariablesプロシージャを呼び出し、すべてのプロパティと配列を
'       初期化します。
' 引数: なし
'================================================================================
Private Sub Class_Initialize()
    '変数の設定を自動的に行う
    SetVariables
End Sub

'================================================================================
' 機能: マクロの動作に必要なすべてのプロパティ（変数）を設定します。
'       まずプロパティをデフォルト値で初期化し、次いで設定ファイルから
'       ユーザー設定を読み込んで上書きします。
' 引数: なし
'================================================================================
Sub SetVariables()
    Dim varKeys() As Variant
    Dim varValues() As Variant
    Dim i As Long
    
    '変数をデフォルト値で初期化する
    strChangeFontType = "Enable"
    strChangeFontSize = "Disable"
    strKeepParagraphWithNext = "Enable"
    strKeepParagraphTogether = "Enable"
    
    'ファイルから設定を読み込む
    Dim settingsMgr As New SettingsManager
    settingsMgr.LoadSettingsFromProperty varKeys, varValues
    Set settingsMgr = Nothing
    
    '読み込んだ値を対応する変数に格納する
    On Error Resume Next
    If UBound(varKeys) > 0 Then
        For i = LBound(varKeys) To UBound(varKeys)
            Select Case varKeys(i)
                Case "Change Font Type": strChangeFontType = varValues(i)
                Case "Change Font Size": strChangeFontSize = varValues(i)
                Case "Line Unit Before One": strLineUnitBeforeOne = varValues(i)
                Case "Paragraph Separation": strKeepParagraphWithNext = varValues(i)
                Case "Paragraph Split": strKeepParagraphTogether = varValues(i)
            End Select
        Next i
    End If
    On Error GoTo 0
      
    '読み込んだ変数をブール値に変換する。
    If strChangeFontType = "Enable" Then
        blnChangeFontType = True
    Else
        blnChangeFontType = False
    End If
    
    If strChangeFontSize = "Enable" Then
        blnChangeFontSize = True
    Else
        blnChangeFontSize = False
    End If

    If strLineUnitBeforeOne = "Enable" Then
        blnLineUnitBeforeOne = True
    Else
        blnLineUnitBeforeOne = False
    End If
    
    If strKeepParagraphWithNext = "Enable" Then
        blnKeepParagraphWithNext = True
    Else
        blnKeepParagraphWithNext = False
    End If
    
    If strKeepParagraphTogether = "Enable" Then
        blnKeepParagraphTogether = True
    Else
        blnKeepParagraphTogether = False
    End If
End Sub