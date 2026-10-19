'================================================================================
' クラスモジュール: VariablesManager
' 機能: 変数を管理する
'================================================================================
Option Explicit

'配列
Public varNumberFormats As Variant            '番号書式設定値の配列
Public varNumberStyles As Variant             '番号スタイル設定値の配列
Public varHeadingLeftIndents As Variant       '見出しの左インデント設定値の配列
Public varHeadingFirstLineIndents As Variant  '見出しの最初の行のインデント設定値の配列
Public varTextLeftIndents As Variant          '本文の左インデント設定値の配列
Public varTextFirstLineIndents As Variant     '本文の最初の行のインデント設定値の配列
Public varHeadingNumber As Variant            '見出し番号の配列

'オプション
Public strNoParagraphSpace As String          '「行間詰めスタイルも変換する」のテキスト
Public blnNoParagraphSpace As Boolean         '「行間詰めスタイルも変換する」のコマンド

Public strOneFullWidth As String              '「見出し番号を全角1文字に調節する」のテキスト
Public blnOneFullWidth As Boolean             '「見出し番号を全角1文字に調節する」のコマンド


'================================================================================
' 機能: クラスがインスタンス化される際に自動的に実行されます。
'       SetVariablesプロシージャを呼び出し、すべてのプロパティと配列を
'       初期化します。
' 引数: なし
'================================================================================
Private Sub Class_Initialize()
    '変数の設定と配列の初期化を自動的に行う
    SetVariables
End Sub

'================================================================================
' 機能: マクロの動作に必要なすべてのプロパティ（変数）を設定します。
'       まずプロパティをデフォルト値で初期化し、次いで設定ファイルから
'       ユーザー設定を読み込んで上書きします。最終的に、選択された書式
'       プリセットに基づき、各種配列を生成します。
' 引数: なし
'================================================================================
Sub SetVariables()
    Dim varKeys() As Variant
    Dim varValues() As Variant
    Dim i As Long
    
    '変数をデフォルト値で初期化する
    blnOneFullWidth = False
    blnNoParagraphSpace = False
    
    'ファイルから設定を読み込む
    Dim settingsMgr As New SettingsManager
    settingsMgr.LoadSettingsFromProperty varKeys, varValues
    Set settingsMgr = Nothing
    
    '読み込んだ値を対応する変数に格納する
    On Error Resume Next
    If UBound(varKeys) > 0 Then
        For i = LBound(varKeys) To UBound(varKeys)
            Select Case varKeys(i)
                Case "Convert No Paragraph Space": strNoParagraphSpace = varValues(i)
                Case "Adjust heading number": strOneFullWidth = varValues(i)
            End Select
        Next i
    End If
    On Error GoTo 0
       
    '「行間詰めスタイルも変換する」に応じて変数を設定する。
    If strNoParagraphSpace = "Enable" Then
        blnNoParagraphSpace = True
    Else
        blnNoParagraphSpace = False
    End If
    
    '「見出し番号を全角1文字に調節する」に応じて変数を設定する。
    If strOneFullWidth = "Enable" Then
        blnOneFullWidth = True
    Else
        blnOneFullWidth = False
    End If
    
    '配列を生成する
    Call SetArrays
End Sub

'================================================================================
' 機能: 各種設定値を格納するための配列を生成・初期化します。
'       カンマ区切りの文字列で定義された書式設定を配列に変換します。
'       また、様々な番号スタイル（アラビア数字、ローマ数字、カタカナ等）
'       に対応する見出し番号の文字列を多次元配列に生成します。
' 引数: なし
'================================================================================
Private Sub SetArrays()
    '配列用変数に定数を読み込みます。
    '定数に配列が使用できないので、カンマ区切りで宣言された定数を配列変数に変換しています。
    varNumberFormats = ConvertToArray(strNumberFormats)
    varNumberStyles = ConvertToArray(strNumberStyles)
    varHeadingLeftIndents = ConvertToArray(strHeadingLeftIndents)
    varHeadingFirstLineIndents = ConvertToArray(strHeadingFirstLineIndents)
    varTextLeftIndents = ConvertToArray(strTextLeftIndents)
    varTextFirstLineIndents = ConvertToArray(strTextFirstLineIndents)
    
    '見出し番号の配列を生成する
    '全角アラビア数字：14、半角アラビア数字：0、全角カタカナ：20、半角カタカナ：12
    '英字（小文字）：4、ローマ数字（小文字）：2、丸囲み数字：18
    ReDim varHeadingNumber(1 To 99, 0 To 20)     '固定長の配列に変換する
    
    ' 全角カタカナと半角カタカナ
    Dim katakanaFull As String
    katakanaFull = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
    
    Dim katakanaHalf As String
    katakanaHalf = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ"
    
    ' 英字（小文字）
    Dim alphabet As String
    alphabet = "abcdefghijklmnopqrstuvwxyz"
    
    ' ローマ数字（小文字）
    Dim romanNumbers As Variant
    romanNumbers = Array("i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x", _
                         "xi", "xii", "xiii", "xiv", "xv", "xvi", "xvii", "xviii", "xix", "xx", _
                         "xxi", "xxii", "xxiii", "xxiv", "xxv", "xxvi", "xxvii", "xxviii", "xxix", "xxx", _
                         "xxxi", "xxxii", "xxxiii", "xxxiv", "xxxv", "xxxvi", "xxxvii", "xxxviii", "xxxix", "xl", _
                         "xli", "xlii", "xliii", "xliv", "xlv", "xlvi", "xlvii", "xlviii", "xlix", "l", _
                         "li", "lii", "liii", "liv", "lv", "lvi", "lvii", "lviii", "lix", "lx", _
                         "lxi", "lxii", "lxiii", "lxiv", "lxv", "lxvi", "lxvii", "lxviii", "lxix", "lxx", _
                         "lxxi", "lxxii", "lxxiii", "lxxiv", "lxxv", "lxxvi", "lxxvii", "lxxviii", "lxxix", "lxxx", _
                         "lxxxi", "lxxxii", "lxxxiii", "lxxxiv", "lxxxv", "lxxxvi", "lxxxvii", "lxxxviii", "lxxxix", "xc", _
                         "xci", "xcii", "xciii", "xciv", "xcv", "xcvi", "xcvii", "xcviii", "xcix")
    
    ' 丸囲み数字（20 まで）
    Dim encircledNumbers As Variant
    encircledNumbers = Array("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", _
                             "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳")
    
    ' 配列にデータを格納
    Dim i As Integer
    For i = 1 To 99
        varHeadingNumber(i, 14) = StrConv(i, vbWide) ' 全角アラビア数字
        varHeadingNumber(i, 0) = i                  ' 半角アラビア数字
        
        ' 全角カタカナ
        If i <= Len(katakanaFull) Then
            varHeadingNumber(i, 20) = Mid(katakanaFull, i, 1)
        Else
            varHeadingNumber(i, 20) = ""
        End If
        
        ' 半角カタカナ
        If i <= Len(katakanaHalf) Then
            varHeadingNumber(i, 12) = Mid(katakanaHalf, i, 1)
        Else
            varHeadingNumber(i, 12) = ""
        End If
        
        ' 英字（小文字）
        If i <= 26 Then
            varHeadingNumber(i, 4) = Mid(alphabet, i, 1)
        Else
            varHeadingNumber(i, 4) = ""
        End If
        
        ' ローマ数字（小文字）
        If i <= UBound(romanNumbers) + 1 Then
            varHeadingNumber(i, 2) = romanNumbers(i - 1)
        Else
            varHeadingNumber(i, 2) = ""
        End If
        
        ' 丸囲み数字（20 まで）
        If i <= 20 Then
            varHeadingNumber(i, 18) = encircledNumbers(i - 1)
        Else
            varHeadingNumber(i, 18) = ""
        End If
    Next i
    
End Sub

'================================================================================
' 機能: カンマ区切りの文字列を、Variant型の配列に変換します。
'       戻り値の配列は、インデックス1から使えるように調整されます。
' 引数: strArray - カンマで区切られた文字列 (ByVal)
' 戻り値: 分割された文字列を格納したVariant型の配列
'================================================================================
Function ConvertToArray(strArray As String) As Variant
    '文字列を配列変数に変換する。
    '「0」ではなく「1」から開始できるように、最初に空要素を追加しています。
    ConvertToArray = Split("," & strArray, ",")
End Function
