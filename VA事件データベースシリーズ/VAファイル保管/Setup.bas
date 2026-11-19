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


'--------------------------------------------------------------------------------
' 概要：設定したバックアップフォルダに、このブックの写し（マクロなしの .xlsx）を
'       「yyyymmdd_ファイル保管.xlsx」の名前で作る。その日の写しが既にあれば作らない（1日に1回）。
'       データがないとき、フォルダが未設定か存在しないときは、何もしない。
' 引数：blnDummy - ［マクロ］ダイアログに出さないためのダミー（使わない）
' 戻り値：なし
' 備考：ブックの名前は、配布のときは版つき（ファイル保管VerNNN.xlsm）、インストールの後は
'       版なし（ファイル保管.xlsm）になる。どちらでも同じ写しの名前になるよう、GetBackupBaseName で求める。
'--------------------------------------------------------------------------------
Sub PerformDailyBackup(Optional ByVal blnDummy As Boolean = False)
    Dim strBackupFolder As String   'バックアップフォルダ
    Dim strTargetPath As String     '写しのフルパス
    Dim strPS As String             'パスの区切り文字（Windows は「\」、Mac は「/」）
    Dim blnAlerts As Boolean        '処理の前の DisplayAlerts の値（後処理で戻す）
    Dim wbNew As Workbook           '写しのブック
    Dim tbl As ListObject           'データのテーブル

    '--- 前処理 ---
    blnAlerts = Application.DisplayAlerts
    On Error GoTo HandleErr

    'データがないときは、写しを作らない
    Set tbl = Sht1Hokan.ListObjects("保管")
    If tbl.ListRows.Count = 0 Then GoTo Cleanup

    'オプションシートのバックアップフォルダを読む（名前の定義がないときは、空のままにする）
    On Error Resume Next
    strBackupFolder = CStr(ShtOption.Range("バックアップフォルダ").Value)
    On Error GoTo HandleErr

    'フォルダが未設定か、存在しないときは、何もしない
    If strBackupFolder = "" Then GoTo Cleanup
    If Dir(strBackupFolder, vbDirectory) = "" Then GoTo Cleanup

    '--- メイン処理 ---
    '写しのフルパスを作る（例：<フォルダ>\20261119_ファイル保管.xlsx）
    strPS = Application.PathSeparator
    If Right(strBackupFolder, 1) <> strPS Then strBackupFolder = strBackupFolder & strPS
    strTargetPath = strBackupFolder & Format(Date, "yyyymmdd") & "_" & GetBackupBaseName(ThisWorkbook.Name) & ".xlsx"

    'その日の写しが既にあれば、作らない
    If Dir(strTargetPath) <> "" Then GoTo Cleanup

    '全シートを新しいブックに写し、マクロなしのブック（.xlsx）として保存して閉じる
    '（シートのコードを保存できない旨の確認を出さないため、DisplayAlerts を止める）
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets.Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs Filename:=strTargetPath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing

Cleanup:
    '--- 後処理 ---
    On Error Resume Next
    '途中で失敗したときは、写しのブックを保存せずに閉じる
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    Application.DisplayAlerts = blnAlerts
    Exit Sub

HandleErr:
    MsgBox "自動バックアップの作成中にエラーが発生しました。" & vbCrLf & _
           "保存先：" & strBackupFolder & vbCrLf & _
           "エラー番号：" & Err.Number & vbCrLf & _
           "内容：" & Err.Description, vbExclamation, "VAファイル保管"
    Resume Cleanup
End Sub

'--------------------------------------------------------------------------------
' 概要：ブックのファイル名から、バックアップの写しの名前に使う部分を返す。
'       拡張子を除き、末尾が版（「Ver」と英数字3文字）なら、それも除く。
'       例：「ファイル保管VerNNN.xlsm」→「ファイル保管」、「ファイル保管.xlsm」→「ファイル保管」
' 引数：strFileName - ブックのファイル名
' 戻り値：写しの名前に使う部分
'--------------------------------------------------------------------------------
Private Function GetBackupBaseName(ByVal strFileName As String) As String
    Dim strBase As String   '拡張子を除いた名前
    Dim lngDot As Long      '最後の「.」の位置

    '拡張子（最後の「.」から後ろ）を除く
    lngDot = InStrRev(strFileName, ".")
    If lngDot > 1 Then
        strBase = Left(strFileName, lngDot - 1)
    Else
        strBase = strFileName
    End If

    '末尾の版（例：Ver403、VerA05）を除く
    If strBase Like "*Ver[0-9A-Z][0-9A-Z][0-9A-Z]" Then strBase = Left(strBase, Len(strBase) - 6)

    GetBackupBaseName = strBase
End Function

