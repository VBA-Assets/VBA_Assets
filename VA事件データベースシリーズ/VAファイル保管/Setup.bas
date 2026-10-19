Option Explicit

'================================================================================
' モジュール名: Setup
' 概要:           プロジェクト全体で使用するグローバル変数の宣言と、
'                 共通の初期化処理やバックアップ処理を行う標準モジュールです。
' 更新日時:       2026/02/08 12:00
'================================================================================

'// プロジェクト全体で使用するグローバル変数とオブジェクト変数を宣言

Public strActCase As String              ' 現在選択中の事件呼名を保持する変数
Public brnMoveFile As Boolean            ' ファイル移動処理が正常完了したかを示すフラグ
Public brnMakeShortcutsCopies As Boolean ' ショートカット・コピー作成が完了したかを示すフラグ
Public Evnt As New Evnt                  ' Evntクラス（イベント制御）のインスタンス

'================================================================================
' プロシージャ: VABridge_AfterLoad
' 概要:           外部連携等のロード完了後に実行される初期化処理（現在はソートと位置合わせ）
'                 マクロメニューに表示させないため、ダミー引数(Optional)を設定しています。
'================================================================================
Public Sub VABridge_AfterLoad(Optional ByVal dummy As Byte = 0)
    ' エラーが発生した場合の処理先を指定します
    On Error GoTo ErrHandler
    
    ' 【修正】SortAndFormatTable呼び出し前にイベントを停止する
    ' これにより、SortAndFormatTable内でのScreenUpdating操作が不要になります。
    Evnt.Disable

    ' オブジェクト変数を使用してシート操作を行います
    Dim wsHokan As Worksheet
    Set wsHokan = Sht1Hokan
    
    ' 「保管」シートをアクティブにします
    wsHokan.Activate
    
    ' 「保管」テーブルのソートと書式設定を実行します
    ' Worksheet型の変数経由では独自メソッドを呼べないため、
    ' オブジェクト名Sht1Hokanを直接指定して呼び出します。
    Call Sht1Hokan.SortAndFormatTable

    ' 処理対象（保管日時が空欄の最初の行）を探索します
    Dim r As Long
    Dim tbl As ListObject
    Set tbl = wsHokan.ListObjects("保管") ' テーブルオブジェクトを取得
    
    ' テーブルのデータをループします
    For r = 1 To tbl.ListRows.Count
        ' 保管日時列が空欄かどうか判定します
        If tbl.ListColumns("保管日時").DataBodyRange(r).Value = "" Then Exit For
    Next

    ' 未処理行が見つかった場合
    If r <= tbl.ListRows.Count Then
        ' その行の事件呼名セルを選択します（ユーザーに位置を示すため）
        tbl.ListColumns("事件呼名").DataBodyRange(r).Select
        ' 画面をその行までスクロールします
        ActiveWindow.ScrollRow = tbl.ListColumns("事件呼名").DataBodyRange(r).Row
    End If

    ' バックアップ処理を実行します
    PerformDailyBackup
    
    ' イベント等を再開します
    Evnt.Enable
    
    ' 処理を終了します
    Exit Sub

ErrHandler:
    ' エラーメッセージを表示します
    MsgBox "初期化処理中にエラーが発生しました。", vbCritical, "ファイル保管"
    ' イベント等を確実に再開させます
    Evnt.Enable
    
End Sub


'================================================================================
' プロシージャ: PerformDailyBackup
' 概要:           現在のブックのバックアップを作成します。
'                 マクロメニューに表示させないため、ダミー引数(Optional)を設定しています。
'================================================================================
Sub PerformDailyBackup(Optional ByVal dummy As Byte = 0)
    ' バックアップフォルダのパスを格納する変数を宣言します
    Dim backupFolder As String
    ' バックアップファイルのファイル名を格納する変数を宣言します
    Dim backupFileName As String
    ' 保存先のフルパスを格納する変数を宣言します
    Dim targetPath As String
    ' パス区切り文字を格納する変数を宣言します
    Dim PS As String
    ' 元のファイル名（拡張子なし）を格納する変数を宣言します
    Dim baseName As String
    ' 新規作成したブックを操作するための変数を宣言します
    Dim wbNew As Workbook
    
    '保管シートのテーブルオブジェクトを操作するための変数を宣言します
    Dim tbl As ListObject

    '保管シートにある「保管」テーブルをオブジェクト変数にセットします
    Set tbl = Sht1Hokan.ListObjects("保管")
    
    'テーブルにデータ行が存在しない（0件）場合は、バックアップ処理を行わずに終了します
    If tbl.ListRows.Count = 0 Then Exit Sub
    
    ' OSごとのパス区切り文字（Windowsは￥、Macは/）を取得します
    PS = Application.PathSeparator
    
    ' 設定値の読み込みエラーを無視するため、エラー処理を一時的に無効化します
    On Error Resume Next
    ' 設定シートからバックアップ保存先のフォルダパスを取得します
    backupFolder = ShtOption.Range("バックアップフォルダ").Value
    ' エラー処理の無効化を解除し、通常のエラー検知に戻します
    On Error GoTo 0
    
    ' バックアップフォルダが設定されていない場合は、処理を終了します
    If backupFolder = "" Then Exit Sub
    
    ' これ以降の処理でエラーが発生した場合、エラー処理ルーチンへ移動するように設定します
    On Error GoTo ErrHandler
    
    ' 指定されたバックアップ先フォルダが存在しない場合は、処理を終了します
    If Dir(backupFolder, vbDirectory) = "" Then Exit Sub
    
    ' 現在のブック名から拡張子を除いたベース名を取得します
    ' "V"の位置を基準にしているため、ファイル名規則に依存しています
    baseName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "V") - 1)
    ' 今日の日付とベース名を組み合わせて、保存するファイル名を作成します（拡張子はxlsx）
    backupFileName = Format(Date, "yyyymmdd") & "_" & baseName & ".xlsx"
    
    ' フォルダパスの末尾にパス区切り文字がない場合を確認します
    If Right(backupFolder, 1) <> PS Then
        ' 区切り文字を付加して、フォルダパスとファイル名を結合します
        targetPath = backupFolder & PS & backupFileName
    ' フォルダパスの末尾に既にパス区切り文字がある場合です
    Else
        ' そのままフォルダパスとファイル名を結合します
        targetPath = backupFolder & backupFileName
    ' パス結合の条件分岐を終了します
    End If
    
    ' その日のバックアップファイルがまだ存在しない場合のみ、以下の処理を実行します
    If Dir(targetPath) = "" Then
        ' 保存時の確認メッセージ等を表示しないようにするため、警告を無効にします
        Application.DisplayAlerts = False
        
        ' 現在のブックの全シートをコピーして、新規ブックを作成します
        ThisWorkbook.Sheets.Copy
        ' 作成された新規ブックを変数に格納します
        Set wbNew = ActiveWorkbook
        
        ' 新規ブックを指定されたパスに、マクロなしのExcelブック形式(.xlsx)で保存します
        wbNew.SaveAs Filename:=targetPath, FileFormat:=xlOpenXMLWorkbook
        ' 保存したバックアップ用ブックを閉じます
        wbNew.Close SaveChanges:=False
        
        ' 無効にしていた警告メッセージを再度有効にします
        Application.DisplayAlerts = True
    ' バックアップ作成処理の条件分岐を終了します
    End If
    
    ' 正常に完了した場合、ここでプロシージャを抜けます
    Exit Sub

ErrHandler:
    ' エラーの内容をメッセージボックスでユーザーに通知します
    MsgBox "自動バックアップの作成中にエラーが発生しました。" & vbCrLf & _
           "保存先: " & backupFolder & vbCrLf & _
           "詳細: " & Err.Description, vbExclamation, "ファイル保管"
    ' エラー終了時でも警告機能が元に戻るように設定します
    Application.DisplayAlerts = True
End Sub

