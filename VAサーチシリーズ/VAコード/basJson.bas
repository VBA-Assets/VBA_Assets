Option Explicit

'================================================================================
' JSONパーサー（manifest.json 読み込み用）
' 外部ライブラリ非依存版
'================================================================================

'--------------------------------------------------------------------------------
' 概要 : manifest.json の "modules" 配列を解析し、Collectionを返します
'         各要素は {"name", "type", "file"} をキーとするCollectionです
'--------------------------------------------------------------------------------
Public Function ParseManifestModules(ByVal manifestJson As String) As Collection
    Dim result As New Collection
    Dim parsed As Collection
    Dim modulesArr As Collection
    Dim i As Long

    On Error GoTo ErrorHandler

    Set parsed = ParseJSON(manifestJson)
    If parsed Is Nothing Then GoTo ErrorHandler

    On Error Resume Next
    Set modulesArr = parsed("modules")
    On Error GoTo 0

    If modulesArr Is Nothing Then
        Set ParseManifestModules = result
        Exit Function
    End If

    For i = 1 To modulesArr.Count
        result.Add modulesArr(i)
    Next i

    Set ParseManifestModules = result
    Exit Function

ErrorHandler:
    Set ParseManifestModules = New Collection
End Function

'--------------------------------------------------------------------------------
' 概要 : JSON文字列を解析してCollectionを返します
'--------------------------------------------------------------------------------
Public Function ParseJSON(ByVal jsonString As String) As Collection
    Dim index As Long
    index = 1
    jsonString = Trim(jsonString)

    If Len(jsonString) = 0 Then
        Set ParseJSON = New Collection
        Exit Function
    End If

    Select Case Mid(jsonString, 1, 1)
        Case "{"
            Set ParseJSON = ParseObj(jsonString, index)
        Case "["
            Set ParseJSON = ParseArr(jsonString, index)
        Case Else
            Set ParseJSON = New Collection
    End Select
End Function

Private Function ParseObj(ByVal json As String, ByRef index As Long) As Collection
    Dim result As New Collection
    Dim key As String
    Dim char As String

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
            key = ParseStr(json, index)
            SkipWhitespace json, index
            If Mid(json, index, 1) <> ":" Then Err.Raise 5001, , "Invalid JSON: Expected ':'"
            index = index + 1
            On Error Resume Next
            result.Add ParseValue(json, index), key
            On Error GoTo 0
        Else
            Err.Raise 5001, , "Invalid JSON in Object at pos " & index
        End If
    Loop

    Set ParseObj = result
End Function

Private Function ParseArr(ByVal json As String, ByRef index As Long) As Collection
    Dim result As New Collection
    Dim char As String

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
        Case "t", "f"
            ParseValue = ParseBool(json, index)
        Case "n"
            index = index + 4
            ParseValue = Null
        Case Else
            ParseValue = ParseNum(json, index)
    End Select
End Function

Private Function ParseStr(ByVal json As String, ByRef index As Long) As String
    Dim sb As String
    Dim char As String

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
                Case """": sb = sb & """"
                Case "\": sb = sb & "\"
                Case "/": sb = sb & "/"
                Case "b": sb = sb & vbBack
                Case "f": sb = sb & vbFormFeed
                Case "n": sb = sb & vbLf
                Case "r": sb = sb & vbCr
                Case "t": sb = sb & vbTab
                Case "u"
                    Dim hexCode As String
                    hexCode = Mid(json, index + 1, 4)
                    sb = sb & ChrW("&H" & hexCode)
                    index = index + 4
                Case Else: sb = sb & char
            End Select
            index = index + 1
        Else
            sb = sb & char
            index = index + 1
        End If
    Loop

    ParseStr = sb
End Function

Private Function ParseNum(ByVal json As String, ByRef index As Long) As Double
    Dim startPos As Long
    startPos = index
    Do While index <= Len(json)
        Dim c As String
        c = Mid(json, index, 1)
        If InStr("0123456789+-.eE", c) > 0 Then
            index = index + 1
        Else
            Exit Do
        End If
    Loop
    ParseNum = Val(Mid(json, startPos, index - startPos))
End Function

Private Function ParseBool(ByVal json As String, ByRef index As Long) As Boolean
    If Mid(json, index, 4) = "true" Then
        ParseBool = True
        index = index + 4
    ElseIf Mid(json, index, 5) = "false" Then
        ParseBool = False
        index = index + 5
    End If
End Function

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
