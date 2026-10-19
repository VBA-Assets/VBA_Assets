Option Explicit

' モジュール概要: OSの差異を吸収し、ファイル操作やシステム情報取得を行うユーティリティ関数群
' Last Updated: 2026/02/12 10:25 (Ver416: Updated GetLocalPath with Mapping Constants)

' -----------------------------------------------------------------------------
' Windows API宣言 (Shiftキー状態判定用)
' -----------------------------------------------------------------------------
#If Mac Then
    ' Macの場合はAPIを使用しないため宣言不要
#Else
    ' Windows API宣言
    ' Excel 2010以降(VBA7)であれば、32bit/64bit問わず PtrSafe を付けるのが安全かつ確実
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    
    ' Shiftキーの仮想キーコード
    Private Const VK_SHIFT As Long = &H10
#End If

' -----------------------------------------------------------------------------
' 機能: パスのOS相互変換 (Windows <-> Mac)
' WindowsのドライブレターパスとMacのマウントパスを、cstSettingsの定数に基づいて変換する。
' -----------------------------------------------------------------------------
Public Function GetLocalPath(ByVal path As String) As String
    If path = "" Then Exit Function
    
    Dim result As String
    result = path
    
    #If Mac Then
        ' --- Mac環境でのパス解決 ---
        ' 入力がWindowsパス(O:\...)の可能性があるため、まず区切り文字を正規化
        result = Replace(result, "\", "/")
        
        ' マッピング定数を使用した変換 (優先)
        ' 例: "O:/..." -> "/Volumes/localhost/..."
        Dim winDrivePrefix As String
        winDrivePrefix = cstSettings.MAP_WIN_DRIVE & ":"
        
        ' 大文字小文字を区別せずに判定
        If UCase(Left(result, Len(winDrivePrefix))) = UCase(winDrivePrefix) Then
            ' ドライブレター部分を /Volumes/設定値 に置換
            result = "/Volumes/" & cstSettings.MAP_MAC_VOLUME & Mid(result, Len(winDrivePrefix) + 1)
        Else
            ' 既存の汎用ロジック (フォールバック: 1文字ドライブの場合のみ変換)
            If Mid(result, 2, 1) = ":" Then
                Dim driveLetter As String
                driveLetter = Left(result, 1)
                result = "/Volumes/" & driveLetter & Mid(result, 3)
            End If
        End If
        
    #Else
        ' --- Windows環境でのパス解決 ---
        ' 入力がMacパス(/Volumes/localhost/...)の可能性がある
        
        ' マッピング定数を使用した変換 (優先)
        ' 例: "/Volumes/localhost/..." -> "O:/..."
        Dim macVolPrefix As String
        macVolPrefix = "/Volumes/" & cstSettings.MAP_MAC_VOLUME
        
        If InStr(1, result, macVolPrefix, vbTextCompare) = 1 Then
            ' ボリュームパス部分を ドライブレター: に置換
            result = cstSettings.MAP_WIN_DRIVE & ":" & Mid(result, Len(macVolPrefix) + 1)
        Else
            ' 既存の汎用ロジック (フォールバック)
            If Left(result, 9) = "/Volumes/" Then
                Dim parts() As String
                parts = Split(result, "/")
                ' parts(0)="", parts(1)="Volumes", parts(2)="O" (ドライブ名)
                If UBound(parts) >= 2 Then
                    Dim driveName As String
                    driveName = parts(2)
                    ' ボリューム名が1文字の場合はドライブレターとみなす
                    If Len(driveName) = 1 Then
                        result = driveName & ":" & Mid(result, Len("/Volumes/" & driveName) + 1)
                    End If
                End If
            End If
        End If
        
        ' 最後に区切り文字をWindows用に正規化
        result = Replace(result, "/", "\")
    #End If
    
    GetLocalPath = result
End Function

' -----------------------------------------------------------------------------
' 機能: ネットワーク上のユーザー名を取得する(OS依存)
' Windows環境変数 "UserName" または Excelの Application.UserName を返す
' -----------------------------------------------------------------------------
Public Function GetNetworkUserName() As String
    On Error GoTo ErrorHandler
    Dim userName As String
    
    #If Mac Then
        ' Macの場合はApplication.UserNameを使用
        userName = Application.userName
    #Else
        ' Windowsの場合は環境変数を優先
        userName = Environ("UserName")
        If userName = "" Then userName = Application.userName
    #End If
    
    GetNetworkUserName = userName
    Exit Function
ErrorHandler:
    GetNetworkUserName = "Unknown"
End Function

' -----------------------------------------------------------------------------
' 機能: ファイル名として使用できない文字を置換する(サニタイズ)
' ユーザー名に特殊文字が含まれる場合の対策
' -----------------------------------------------------------------------------
Public Function SanitizeFileName(ByVal sName As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    ' ファイル名に使えない文字: \ / : * ? " < > |
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    Dim tempName As String
    tempName = sName
    
    For i = LBound(invalidChars) To UBound(invalidChars)
        tempName = Replace(tempName, invalidChars(i), "_")
    Next i
    
    SanitizeFileName = tempName
End Function

' -----------------------------------------------------------------------------
' 機能: パス結合(OS依存)
' -----------------------------------------------------------------------------
Public Function BuildPath(ByVal folderPath As String, ByVal fileName As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    
    If Right(folderPath, 1) = sep Then
        BuildPath = folderPath & fileName
    Else
        BuildPath = folderPath & sep & fileName
    End If
End Function

' -----------------------------------------------------------------------------
' 機能: ファイルの存在確認(OS依存)
' -----------------------------------------------------------------------------
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------
' 機能: ロックファイルの作成(内容付き)
' 2026/02/09: サイズ確保と情報記録のため内容を書き込むように変更
' -----------------------------------------------------------------------------
Public Sub CreateLockFile(ByVal filePath As String, ByVal content As String)
    On Error GoTo ErrorHandler
    Dim fn As Integer
    fn = FreeFile
    Open filePath For Output As #fn
    Print #fn, content
    Close #fn
    Exit Sub
ErrorHandler:
    ' エラー時は何もしない(呼び出し元でチェックする等は必要であれば実装)
End Sub

' -----------------------------------------------------------------------------
' 機能: テキストファイルをUTF-8で読み込む
' Windows/Mac両対応
' -----------------------------------------------------------------------------
Public Function ReadUTF8Content(ByVal path As String) As String
    If Not FileExists(path) Then Exit Function
    
    On Error GoTo ErrorHandler
    Dim result As String
    
    #If Mac Then
        Dim fn As Integer: fn = FreeFile
        Open path For Input As #fn
        Line Input #fn, result
        Close #fn
    #Else
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2
        stream.Charset = "UTF-8"
        stream.Open
        stream.LoadFromFile path
        stream.ReadText
        result = stream.ReadText
        stream.Close
    #End If
    
    ReadUTF8Content = Replace(Replace(result, vbCr, ""), vbLf, "")
    Exit Function
ErrorHandler:
    ReadUTF8Content = ""
End Function

' -----------------------------------------------------------------------------
' 機能: ユーザーブック側のプロシージャを実行する
' -----------------------------------------------------------------------------
Public Sub RunUserProcedure(ByVal procName As String, Optional ByVal TargetBook As Workbook = Nothing)
    On Error Resume Next
    If TargetBook Is Nothing Then
        Application.Run procName
    Else
        Application.Run "'" & TargetBook.Name & "'!" & procName
    End If
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' 機能: 指定されたシートが管理対象(設定済み)か判定
' -----------------------------------------------------------------------------
Public Function IsManagedSheet(ByVal sh As Worksheet) As Boolean
    IsManagedSheet = (basConfig.GetSheetSetting(sh, cstSettings.PROP_SUFFIX_SHEET) <> "")
End Function

' -----------------------------------------------------------------------------
' 機能: Shiftキーが現在押されているか判定する
' メンテナンスモード(アドイン機能停止)の検出に使用
' -----------------------------------------------------------------------------
Public Function IsShiftKeyPressed() As Boolean
    #If Mac Then
        ' MacではAPIによる判定が難しいため、常にFalse(検知不可)とする
        IsShiftKeyPressed = False
    #Else
        ' GetAsyncKeyStateの戻り値の最上位ビットを確認
        If (GetAsyncKeyState(VK_SHIFT) And &H8000) <> 0 Then
            IsShiftKeyPressed = True
        Else
            IsShiftKeyPressed = False
        End If
    #End If
End Function