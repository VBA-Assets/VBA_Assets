'更新日時: 2026/10/18
'概要: アプリケーションの初期設定・連携およびバックアップ処理モジュール
Option Explicit
Public Evnt As New Evnt         'イベント制御用クラスの宣言

'--------------------------------------------------
' 概要 : 外部連携等のロード後に呼び出され、電話メモのソート、書式設定、および日次バックアップを実行する
' 引数 : dummy (Byte) - オプションダミー引数
' 戻り値 : なし
'--------------------------------------------------
Public Sub VABridge_AfterLoad(Optional ByVal dummy As Byte = 0)
    'エラーが発生した場合、エラー処理ルーチンへジャンプします
    On Error GoTo ErrHandler

    'イベントの発生や画面更新を停止し、処理速度を向上させます
    Evnt.Disable
    
    '電話メモのテーブルを番号順に並び替え、書式（色や表示形式）を設定します
    'FormatTableにダミー引数がついたが、Optionalなのでそのまま呼び出し可能
    sht11.FormatTable
   
    '日次バックアップ処理を実行します（データがある場合のみ実行されます）
    PerformDailyBackup
    
    '停止していたイベントや画面更新を再開させます
    Evnt.Enable
    '処理を正常に終了します
    Exit Sub

ErrHandler:
    'エラーが発生した場合、メッセージボックスを表示します
    MsgBox "初期化処理中にエラーが発生しました。", vbCritical, "電話メモ"
    'エラー時でも操作不能にならないよう、イベントや画面更新を再開させます
    Evnt.Enable
End Sub

'--------------------------------------------------
' 概要 : 設定されたバックアップフォルダへ、現在のブックのコピーを日付付きファイル名で保存する
' 引数 : dummy (Byte) - オプションダミー引数
' 戻り値 : なし
'--------------------------------------------------
Sub PerformDailyBackup(Optional ByVal dummy As Byte = 0)
    'バックアップフォルダのパスを格納する変数を宣言します
    Dim backupFolder As String
    'バックアップファイルのファイル名を格納する変数を宣言します
    Dim backupFileName As String
    '保存先のフルパスを格納する変数を宣言します
    Dim targetPath As String
    'パス区切り文字を格納する変数を宣言します
    Dim PS As String
    '元のファイル名（拡張子なし）を格納する変数を宣言します
    Dim baseName As String
    '新規作成したブックを操作するための変数を宣言します
    Dim wbNew As Workbook
    '電話メモのテーブルオブジェクトを操作するための変数を宣言します
    Dim tbl As ListObject

    '電話メモ（sht11）にある「電話メモ」テーブルをオブジェクト変数にセットします
    Set tbl = sht11.ListObjects("電話メモ")
    
    'テーブルにデータ行が存在しない（0件）場合は、バックアップ処理を行わずに終了します
    If tbl.ListRows.Count = 0 Then Exit Sub

    'OSごとのパス区切り文字（Windowsは￥、Macは/）を取得します
    PS = Application.PathSeparator
    
    '設定値の読み込みエラーを無視するため、エラー処理を一時的に無効化します
    On Error Resume Next
    '設定シートからバックアップ保存先のフォルダパスを取得します
    backupFolder = shtOpt.Range("バックアップフォルダ").Value
    'エラー処理の無効化を解除し、通常のエラー検知に戻します
    On Error GoTo 0
    
    'バックアップフォルダが設定されていない場合は、処理を終了します
    If backupFolder = "" Then Exit Sub
    
    'これ以降の処理でエラーが発生した場合、エラー処理ルーチンへ移動するように設定します
    On Error GoTo ErrHandler
    
    '指定されたバックアップ先フォルダが存在しない場合は、処理を終了します
    If Dir(backupFolder, vbDirectory) = "" Then Exit Sub
    
    '現在のブック名から拡張子を除いたベース名を取得します
    baseName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "V") - 1)
    '今日の日付とベース名を組み合わせて、保存するファイル名を作成します（拡張子はxlsx）
    backupFileName = Format(Date, "yyyymmdd") & "_" & baseName & ".xlsx"
    
    'フォルダパスの末尾にパス区切り文字がない場合を確認します
    If Right(backupFolder, 1) <> PS Then
        '区切り文字を付加して、フォルダパスとファイル名を結合します
        targetPath = backupFolder & PS & backupFileName
    'フォルダパスの末尾に既にパス区切り文字がある場合です
    Else
        'そのままフォルダパスとファイル名を結合します
        targetPath = backupFolder & backupFileName
    'パス結合の条件分岐を終了します
    End If
    
    'その日のバックアップファイルがまだ存在しない場合のみ、以下の処理を実行します
    If Dir(targetPath) = "" Then
        '保存時の確認メッセージ等を非表示にするため、警告を無効にします
        Application.DisplayAlerts = False
        
        '現在のブックの全シートをコピーして、新規ブックを作成します
        ThisWorkbook.Sheets.Copy
        '作成された新規ブックを変数に格納します
        Set wbNew = ActiveWorkbook
        
        '新規ブックを指定されたパスに、マクロなしのExcelブック形式(.xlsx)で保存します
        wbNew.SaveAs Filename:=targetPath, FileFormat:=xlOpenXMLWorkbook
        '保存したバックアップ用ブックを閉じます
        wbNew.Close SaveChanges:=False
        
        '無効にしていた警告メッセージを再度有効にします
        Application.DisplayAlerts = True
    'バックアップ作成処理の条件分岐を終了します
    End If
    
    '正常に完了した場合、ここでプロシージャを抜けます
    Exit Sub

ErrHandler:
    'エラーの内容をメッセージボックスでユーザーに通知します
    MsgBox "自動バックアップの作成中にエラーが発生しました。" & vbCrLf & _
           "保存先: " & backupFolder & vbCrLf & _
           "詳細: " & Err.Description, vbExclamation, "電話メモ"
    'エラー終了時でも警告機能が元に戻るように設定します
    Application.DisplayAlerts = True
End Sub

