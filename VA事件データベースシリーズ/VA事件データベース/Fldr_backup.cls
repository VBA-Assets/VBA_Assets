
Option Explicit

'================================================================================
' モジュール名: Fldr
' 概要:         ファイルシステム操作をカプセル化するクラスモジュールです。
'               フォルダの移動やZIPファイルの操作など、Windows/Mac両環境で
'               動作するコードを提供します。
' 更新日時:     2026/02/01 05:37
'================================================================================

'================================================================================
' プロシージャ: Move
' 機能:         フォルダの移動・リネーム
' 概要:         指定されたフォルダ(strFolderA)を別の場所(strFolderB)へ
'               移動または名称変更します。Nameステートメントを使用し、
'               エラー時にはメッセージを表示します。
'================================================================================
Sub Move(strFolderA As String, strFolderB As String)
    'フォルダの存在確認
    If Dir(strFolderA, vbDirectory) = "" Then
        MsgBox "移動前のフォルダが存在しません: " & strFolderA, vbExclamation, "事件データベース"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    'Nameステートメントは多くのケース（同一ドライブ）で移動として機能する
    'WindowsとMac両方で利用可能な標準ステートメント
    Name strFolderA As strFolderB
    
    Exit Sub

ErrorHandler:
    'Nameステートメントでの移動に失敗した場合（別ドライブなど）
    MsgBox "フォルダの移動に失敗しました。" & vbCrLf & _
           "エラー: " & Err.Description & vbCrLf & _
           "手動で移動を行ってください。", vbExclamation, "事件データベース"
End Sub

'================================================================================
' プロシージャ: MoveInZip
' 機能:         ZIPファイルの移動
' 概要:         指定フォルダに対応するZIPファイルを移動先の構成に合わせて
'               移動します。移動先フォルダが存在しない場合は作成を試みます。
'================================================================================
Sub MoveInZip(strFolderA As String, strFolderB As String)
    Dim PS As String
    PS = Application.PathSeparator
    
    'ZIPファイルのパスを構築
    Dim strZipPath As String
    strZipPath = strFolderA & ".zip"
    
    'ZIPファイルが存在しない場合は終了
    If Dir(strZipPath) = "" Then Exit Sub
    
    On Error Resume Next
    
    '移動先フォルダ階層の作成を試みる（簡易的なMkDir）
    '※多階層の作成が必要な場合は再帰処理が必要だが、ここでは親フォルダが存在する前提か1階層のみ作成
    If Dir(strFolderB, vbDirectory) = "" Then
        MkDir strFolderB
    End If
    
    On Error GoTo 0
    
    '移動先ファイルパス
    Dim strZipDest As String
    'strFolderBの中に配置するのか、strFolderBと同列に配置するのか文脈によるが
    '元のコードではMac版は `strFolderB & PS & ファイル名` としている
    
    'ファイル名部分を取得
    Dim fileName As String
    fileName = Mid(strZipPath, InStrRev(strZipPath, PS) + 1)
    
    '移動先パス
    strZipDest = strFolderB & PS & fileName
    
    '移動実行
    On Error GoTo MoveError
    Name strZipPath As strZipDest
    
    '成功メッセージ
    MsgBox "ZIPファイルを移動しました。" & vbCrLf & _
           "移動元: " & strZipPath & vbCrLf & _
           "移動先: " & strZipDest, vbInformation, "事件データベース"
    Exit Sub

MoveError:
    MsgBox "ZIPファイルの移動に失敗しました（Mac/Win共通エラー）。" & vbCrLf & _
           Err.Description, vbExclamation, "事件データベース"
End Sub


