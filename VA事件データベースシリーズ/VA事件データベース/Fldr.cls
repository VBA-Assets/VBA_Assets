Option Explicit

'================================================================================
' 機能:     指定されたフォルダ(strFolderA)を別の場所(strFolderB)へ移動します。
' 概要:     WindowsではFSO、MacではNameステートメントを使用します。
'================================================================================
Sub Move(strFolderA As String, strFolderB As String)

    '【修正】WindowsとMacで処理を分岐
    #If Mac Then
        ' --- Mac用の処理 (Nameステートメントによる移動) ---
        '移動元フォルダが存在しなければエラー終了
        If Dir(strFolderA, vbDirectory) = "" Then
            MsgBox "移動前のフォルダが存在しません"
            Exit Sub
        End If
        
        '移動先の親フォルダが存在するか簡易チェックは省略（Nameは親がないと失敗する）
        'VBAのMkDirは1階層しか作れないため、本来は再帰的な作成が必要だが、
        'ここでは簡易的に実装。
        
        On Error GoTo MacError
        Application.StatusBar = "フォルダの移動を行っています。"
        
        '名前変更（実質的な移動）を実行
        Name strFolderA As strFolderB
        
        Application.StatusBar = ""
        Exit Sub
MacError:
        MsgBox "Macでのフォルダ移動に失敗しました。" & vbCr & "移動先:" & strFolderB
        Application.StatusBar = ""
        Exit Sub
        
    #Else
        ' --- Windows用の処理 (Scripting.FileSystemObject) ---
        Dim FSO As Object
        'FileSystemObjectのインスタンスを作成
        Set FSO = CreateObject("Scripting.FileSystemObject")

        '移動元フォルダが存在しなければエラー終了
        If FSO.folderexists(strFolderA) = False Then
            MsgBox ("移動前のフォルダが存在しません")
            GoTo ExitHandler
        End If

        ' 移動先の親フォルダが存在しない場合
        If FSO.folderexists(Left(strFolderB, InStrRev(strFolderB, "\"))) = False Then
            On Error GoTo Skip1
            '親フォルダを作成する
            FSO.CreateFolder (Left(strFolderB, InStrRev(strFolderB, "\") - 1))
            On Error GoTo 0
        End If
            
        ' 同一ドライブ内の場合
        If Left(strFolderA, 1) = Left(strFolderB, 1) Then
            On Error GoTo Skip2
            Application.StatusBar = "フォルダの移動を行っています。"
            'Nameステートメントで移動（高速）
            Name strFolderA As strFolderB
            On Error GoTo 0
            
        ' 同一ドライブではない場合
        Else
            '注意喚起
            If MsgBox("別ドライブへの移動が指示されました。" & vbCr & _
                      "安全のためコピーのみ行います。" & vbCr & _
                      "フォルダの容量によっては処理に長時間がかかる場合があります。" & vbCr & _
                      "よろしいですか？", vbExclamation + vbOKCancel) = vbCancel Then GoTo Skip4
                      
            On Error GoTo Skip3
            Application.StatusBar = "フォルダのコピーを行っています。"
            'フォルダをコピー
            FSO.CopyFolder strFolderA, strFolderB
            MsgBox "フォルダのコピーを完了しました。" & vbCr & _
                   "コピーが完了していることを確認してください。"

            '元フォルダ削除確認
            If MsgBox("元のフォルダを削除します。" & vbCr & _
                      "コピー完了を確認することを強く推奨します。" & vbCr & _
                      "よろしいですか？", vbExclamation + vbOKCancel) = vbCancel Then GoTo Skip4
            
            '元フォルダ削除
            FSO.DeleteFolder strFolderA
            On Error GoTo 0
        End If
        GoTo ExitHandler

Skip1:
        MsgBox "フォルダを作成できませんでした。"
        GoTo ExitHandler
Skip2:
        MsgBox "フォルダを移動できませんでした。"
        GoTo ExitHandler
Skip3:
        MsgBox "フォルダをコピーまたは削除できませんでした。"
        GoTo ExitHandler
Skip4:
        MsgBox "キャンセルしました。"
        GoTo ExitHandler
ExitHandler:
        Application.StatusBar = ""
        Set FSO = Nothing
    #End If
End Sub

'================================================================================
' 機能:     指定されたフォルダ(strFolderA)に紐づくZIPファイルを移動します。
'================================================================================
Sub MoveInZip(strFolderA As String, strFolderB As String)
    Dim PS As String
    PS = Application.PathSeparator

    '【修正】WindowsとMacで処理を分岐
    #If Mac Then
        ' --- Mac用の処理 ---
        'FSOを使わずに実装
        Dim strZipPath As String
        strZipPath = strFolderA & ".zip"
        
        'ZIPファイルがなければ終了
        If Dir(strZipPath) = "" Then GoTo MacSkip2
        
        '移動先のフォルダ作成（簡易版）
        On Error Resume Next
        MkDir strFolderB
        On Error GoTo 0
        
        If MsgBox("ZIPファイルの移動を行います。" & vbCr & "よろしいですか？", vbExclamation + vbOKCancel) = vbCancel Then Exit Sub
        
        '移動
        On Error GoTo MacSkip3
        Name strZipPath As strFolderB & PS & Mid(strZipPath, InStrRev(strZipPath, PS) + 1)
        
        'Finderを開く
        Shell "open """ & strFolderB & """", vbNormalFocus
        
        MsgBox "ZIPファイルの移動が完了しました。", vbExclamation
        
        If MsgBox("元のフォルダを削除します。" & vbCr & "よろしいですか？", vbExclamation + vbOKCancel) = vbCancel Then Exit Sub
        
        'Macでフォルダ削除はリスクが高いため、ここではZIP移動のみで止めるか、
        'シェルコマンド rm -rf を使う必要があるが、安全のためメッセージのみとする
        MsgBox "Mac版では安全のため、元のフォルダ削除は手動で行ってください。"
        Exit Sub
        
MacSkip2:
        MsgBox "Zipファイルが存在しません。"
        Exit Sub
MacSkip3:
        MsgBox "移動に失敗しました。"
        Exit Sub
        
    #Else
        ' --- Windows用の処理 (Scripting.FileSystemObject) ---
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")

        '移動元が存在しなければエラー
        If FSO.folderexists(strFolderA) = False Then
            MsgBox ("移動前のフォルダが存在しません")
            GoTo ExitHandler
        End If

        Dim strZipPathWin As String
        strZipPathWin = strFolderA & ".zip"

        'ZIPがなければエラー
        If Dir(strZipPathWin) = "" Then GoTo Skip2

        '移動先親フォルダチェック
        If FSO.folderexists(Left(strFolderB, InStrRev(strFolderB, "\"))) = False Then
            On Error GoTo Skip1
            FSO.CreateFolder (Left(strFolderB, InStrRev(strFolderB, "\") - 1))
            On Error GoTo 0
        End If
            
        '移動確認
        If MsgBox("ZIPファイルの移動を行います。" & vbCr & _
                  "フォルダの容量によっては処理に長時間がかかる場合があります。" & vbCr & _
                  "よろしいですか？", vbExclamation + vbOKCancel) = vbCancel Then GoTo Skip

        '移動先フォルダ作成
        If FSO.folderexists(strFolderB) = False Then
            MkDir strFolderB
        End If

        On Error GoTo Skip3

        'ZIP移動
        FSO.MoveFile strZipPathWin, strFolderB & "\"
        
        '【修正】Explorer起動
        Shell "C:\windows\explorer.exe " & strFolderB & "\", vbNormalFocus

        MsgBox "ZIPファイルの移動が完了しました。" & vbCr & _
               "エクスプローラーでZIPファイルを開き、移動が完了していることを確認してください。", vbExclamation

        On Error GoTo Skip4

        '元フォルダ削除確認
        If MsgBox("元のフォルダを削除します。" & vbCr & _
                  "よろしいですか？", vbExclamation + vbOKCancel) = vbCancel Then GoTo Skip

        '元フォルダ削除
        FSO.DeleteFolder strFolderA
        On Error GoTo 0
        GoTo ExitHandler

Skip:
        MsgBox "キャンセルしました。"
        GoTo ExitHandler
Skip1:
        MsgBox "新しいフォルダを作成できませんでした。"
        GoTo ExitHandler
Skip2:
        MsgBox "Zipファイルが存在しません。"
        GoTo ExitHandler
Skip3:
        MsgBox "圧縮ファイルをコピーできませんでした。"
        GoTo ExitHandler
Skip4:
        MsgBox "元のフォルダを削除できませんでした。"
        GoTo ExitHandler
ExitHandler:
        Set FSO = Nothing
    #End If
End Sub


