Option Explicit

'================================================================================
' モジュール名: Setup
' 概要:         アプリケーション全体の初期化と共通機能を提供する標準モジュールです。
'               グローバル変数の宣言、日次バックアップ、文字列操作関数等を
'               実装しています。
' 更新日時:     2026/02/06 23:55
'================================================================================

'--- モジュールレベル（グローバル）変数の宣言 ---
Public lngActCase As Long '現在処理中の事件番号
Public Evnt As New Evnt   'イベント制御クラス
Public Fldr As New Fldr   'フォルダ操作クラス

'================================================================================
' 関数:         FncSplit
' 機能:         安全な文字列分割関数
' 概要:         指定された区切り文字で文字列を分割し、指定番目（1始まり）の
'               要素を返します。エラー時は空文字を返します。
'================================================================================
Function FncSplit(StrSplitData As String, StrSplitChar As String, StrDataNumber As Integer) As String
    'エラーが発生した場合、エラー処理ルーチンへジャンプします
    On Error GoTo catchError
    
    '指定番目の要素を返すため、配列のインデックスを調整して値を返します（配列は0始まりなので -1）
    FncSplit = Split(StrSplitData, StrSplitChar)(StrDataNumber - 1)
    '正常に処理が完了したら関数を抜けます
    Exit Function

catchError:
    'インデックスエラー等が発生した場合は、空文字を返します
    FncSplit = ""
End Function

'================================================================================
' プロシージャ: VABridge_AfterLoad
' 機能:         ロード完了後の初期化処理
' 概要:         外部連携等のロード後に呼び出され、事件シートのソート、
'               書式設定、および日次バックアップを実行します。
'================================================================================
Public Sub VABridge_AfterLoad(Optional ByVal dummy As Byte = 0)
    'エラーが発生した場合、エラー処理ルーチンへジャンプします
    On Error GoTo ErrHandler

    'イベントの発生や画面更新を停止し、処理速度を向上させます
    Evnt.Disable
    
    '事件シートのテーブルデータを指定されたキーで並び替えます
    'SortTableにダミー引数がついたが、Optionalなのでそのまま呼び出し可能
    sht11.SortTable
    
    '事件シートの各行の書式（色や表示形式）を設定します
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
    MsgBox "初期化処理中にエラーが発生しました。", vbCritical, "事件データベース"
    'エラー時でも操作不能にならないよう、イベントや画面更新を再開させます
    Evnt.Enable
    
End Sub

'================================================================================
' プロシージャ: PerformDailyBackup
' 機能:         日次自動バックアップの実行
' 概要:         設定されたバックアップフォルダへ、現在のブックのコピーを
'               日付付きファイル名（yyyymmdd_BookName.xlsx）で保存します。
'               その日最初の起動時のみ実行されます。
'================================================================================
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
    '事件シートのテーブルオブジェクトを操作するための変数を宣言します
    Dim tbl As ListObject

    '事件シート（sht11）にある「事件」テーブルをオブジェクト変数にセットします
    Set tbl = sht11.ListObjects("事件")
    
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
        wbNew.SaveAs fileName:=targetPath, FileFormat:=xlOpenXMLWorkbook
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
           "詳細: " & Err.Description, vbExclamation, "事件データベース"
    'エラー終了時でも警告機能が元に戻るように設定します
    Application.DisplayAlerts = True
End Sub

