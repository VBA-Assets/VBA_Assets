Option Explicit

'================================================================================
' JSONパーサー & ジェネレーター
' 外部ライブラリ非依存版
'================================================================================

'--------------------------------------------------------------------------------
' 概要 : JSON文字列を解析し、DictionaryライクなCollectionのCollectionを返します
'        (配列はCollection、オブジェクトもCollectionとして扱います)
'--------------------------------------------------------------------------------
Public Function ParseJSON(ByVal jsonString As String) As Collection
    Dim index As Long
    index = 1
    jsonString = Trim(jsonString)
    
    ' 空チェック
    If Len(jsonString) = 0 Then
        Set ParseJSON = New Collection
        Exit Function
    End If
    
    ' 最初の文字で分岐
    Select Case Mid(jsonString, 1, 1)
        Case "{"
            Set ParseJSON = ParseObj(jsonString, index)
        Case "["
            Set ParseJSON = ParseArr(jsonString, index)
        Case Else
            ' エラーまたはプリミティブ単体だが、今回は想定外
            Set ParseJSON = New Collection
    End Select
End Function

'--------------------------------------------------------------------------------
' 内部関数: オブジェクト {} の解析
'--------------------------------------------------------------------------------
Private Function ParseObj(ByVal json As String, ByRef index As Long) As Collection
    Dim result As New Collection
    Dim key As String
    Dim char As String
    
    ' { をスキップ
    index = index + 1
    
    Do While index <= Len(json)
        SkipWhitespace json, index
        char = Mid(json, index, 1)
        
        If char = "}" Then
            index = index + 1
            Exit Do
        ElseIf char = "," Then
            index = index + 1
        ElseIf char = """" Then
            ' キーの取得
            key = ParseStr(json, index)
            
            SkipWhitespace json, index
            If Mid(json, index, 1) <> ":" Then Err.Raise 5001, , "Invalid JSON: Expected ':'"
            index = index + 1 ' : をスキップ
            
            ' 値の取得と格納
            ' CollectionをDictionary代わりにするため、Keyを指定してAddする
            ' キー重複はエラーになるがJSONとして正しければ重複はないはず
            On Error Resume Next
            result.Add ParseValue(json, index), key
            On Error GoTo 0
        Else
            Err.Raise 5001, , "Invalid JSON in Object at pos " & index
        End If
    Loop
    
    Set ParseObj = result
End Function

'--------------------------------------------------------------------------------
' 内部関数: 配列 [] の解析
'--------------------------------------------------------------------------------
Private Function ParseArr(ByVal json As String, ByRef index As Long) As Collection
    Dim result As New Collection
    Dim char As String
    
    ' [ をスキップ
    index = index + 1
    
    Do While index <= Len(json)
        SkipWhitespace json, index
        char = Mid(json, index, 1)
        
        If char = "]" Then
            index = index + 1
            Exit Do
        ElseIf char = "," Then
            index = index + 1
        Else
            result.Add ParseValue(json, index)
        End If
    Loop
    
    Set ParseArr = result
End Function

'--------------------------------------------------------------------------------
' 内部関数: 値の解析（再帰）
'--------------------------------------------------------------------------------
Private Function ParseValue(ByVal json As String, ByRef index As Long) As Variant
    SkipWhitespace json, index
    Dim char As String
    char = Mid(json, index, 1)
    
    Select Case char
        Case "{"
            Set ParseValue = ParseObj(json, index)
        Case "["
            Set ParseValue = ParseArr(json, index)
        Case """"
            ParseValue = ParseStr(json, index)
        Case "t", "f" ' true / false
            ParseValue = ParseBool(json, index)
        Case "n" ' null
            index = index + 4
            ParseValue = Null
        Case Else
            ' 数値とみなす
            ParseValue = ParseNum(json, index)
    End Select
End Function

'--------------------------------------------------------------------------------
' 内部関数: 文字列の解析（エスケープ処理対応）
'--------------------------------------------------------------------------------
Private Function ParseStr(ByVal json As String, ByRef index As Long) As String
    Dim sb As String
    Dim char As String
    
    ' " をスキップ
    index = index + 1
    
    Do While index <= Len(json)
        char = Mid(json, index, 1)
        
        If char = """" Then
            index = index + 1
            Exit Do
        ElseIf char = "\" Then
            index = index + 1
            char = Mid(json, index, 1)
            Select Case char
                Case """"
                    sb = sb & """"
                Case "\"
                    sb = sb & "\"
                Case "/"
                    sb = sb & "/"
                Case "b"
                    sb = sb & vbBack
                Case "f"
                    sb = sb & vbFormFeed
                Case "n"
                    sb = sb & vbLf ' VBAのエディタはvbLfまたはvbCrLfで改行する
                Case "r"
                    sb = sb & vbCr
                Case "t"
                    sb = sb & vbTab
                Case "u"
                    ' Unicodeエスケープ (\uXXXX) は簡易対応（飛ばすか、必要なら実装）
                    ' ここでは簡易的にHex変換を試みる
                    Dim hexCode As String
                    hexCode = Mid(json, index + 1, 4)
                    sb = sb & ChrW("&H" & hexCode)
                    index = index + 4
                Case Else
                    sb = sb & char
            End Select
            index = index + 1
        Else
            sb = sb & char
            index = index + 1
        End If
    Loop
    
    ParseStr = sb
End Function

'--------------------------------------------------------------------------------
' 内部関数: 数値の解析
'--------------------------------------------------------------------------------
Private Function ParseNum(ByVal json As String, ByRef index As Long) As Double
    Dim startPos As Long
    Dim endPos As Long
    Dim numStr As String
    
    startPos = index
    ' 数字、符号、ドット、Eなどが続く限り読み込む
    Do While index <= Len(json)
        Dim c As String
        c = Mid(json, index, 1)
        If InStr("0123456789+-.eE", c) > 0 Then
            index = index + 1
        Else
            Exit Do
        End If
    Loop
    
    numStr = Mid(json, startPos, index - startPos)
    ParseNum = Val(numStr)
End Function

'--------------------------------------------------------------------------------
' 内部関数: ブール値の解析
'--------------------------------------------------------------------------------
Private Function ParseBool(ByVal json As String, ByRef index As Long) As Boolean
    If Mid(json, index, 4) = "true" Then
        ParseBool = True
        index = index + 4
    ElseIf Mid(json, index, 5) = "false" Then
        ParseBool = False
        index = index + 5
    End If
End Function

'--------------------------------------------------------------------------------
' 内部関数: 空白スキップ
'--------------------------------------------------------------------------------
Private Sub SkipWhitespace(ByVal json As String, ByRef index As Long)
    Do While index <= Len(json)
        Select Case Mid(json, index, 1)
            Case " ", vbTab, vbCr, vbLf
                index = index + 1
            Case Else
                Exit Sub
        End Select
    Loop
End Sub

'================================================================================
' JSON生成 (エスケープ処理)
'================================================================================

Public Function EscapeString(ByVal str As String) As String
    Dim i As Long
    Dim char As String
    Dim sb As String
    
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        Select Case char
            Case """"
                sb = sb & "\"""
            Case "\"
                sb = sb & "\\"
            Case "/"
                sb = sb & "\/"
            Case vbBack
                sb = sb & "\b"
            Case vbFormFeed
                sb = sb & "\f"
            Case vbLf
                sb = sb & "\n"
            Case vbCr
                sb = sb & "\r"
            Case vbTab
                sb = sb & "\t"
            Case Else
                ' 制御文字等は必要に応じてUnicodeエスケープ
                ' 特定の漢字が消えるバグを修正するためIf AscW(char) < 32 Then Elseから修正
                If Not (AscW(char) And &HFFFF&) < 32 Then
                    sb = sb & char
                End If
        End Select
    Next i
    
    EscapeString = sb
End Function