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

'--------------------------------------------------------------------------------
' 概要：設定したバックアップフォルダに、このブックの写し（マクロなしの .xlsx）を
'       「yyyymmdd_事件データベース.xlsx」の名前で作る。その日の写しが既にあれば作らない（1日に1回）。
'       データがないとき、フォルダが未設定か存在しないときは、何もしない。
' 引数：blnDummy - ［マクロ］ダイアログに出さないためのダミー（使わない）
' 戻り値：なし
' 備考：ブックの名前は、配布のときは版つき（事件データベースVerNNN.xlsm）、インストールの後は
'       版なし（事件データベース.xlsm）になる。どちらでも同じ写しの名前になるよう、GetBackupBaseName で求める。
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
    Set tbl = sht11.ListObjects("事件")
    If tbl.ListRows.Count = 0 Then GoTo Cleanup

    'オプションシートのバックアップフォルダを読む（名前の定義がないときは、空のままにする）
    On Error Resume Next
    strBackupFolder = CStr(shtOpt.Range("バックアップフォルダ").Value)
    On Error GoTo HandleErr

    'フォルダが未設定か、存在しないときは、何もしない
    If strBackupFolder = "" Then GoTo Cleanup
    If Dir(strBackupFolder, vbDirectory) = "" Then GoTo Cleanup

    '--- メイン処理 ---
    '写しのフルパスを作る（例：<フォルダ>\20261119_事件データベース.xlsx）
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
    wbNew.SaveAs fileName:=strTargetPath, FileFormat:=xlOpenXMLWorkbook
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
           "内容：" & Err.Description, vbExclamation, "VA事件データベース"
    Resume Cleanup
End Sub

'--------------------------------------------------------------------------------
' 概要：ブックのファイル名から、バックアップの写しの名前に使う部分を返す。
'       拡張子を除き、末尾が版（「Ver」と英数字3文字）なら、それも除く。
'       例：「事件データベースVerNNN.xlsm」→「事件データベース」、「事件データベース.xlsm」→「事件データベース」
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

