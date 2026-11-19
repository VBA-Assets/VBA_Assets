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

'--------------------------------------------------------------------------------
' 概要：設定したバックアップフォルダに、このブックの写し（マクロなしの .xlsx）を
'       「yyyymmdd_電話メモ.xlsx」の名前で作る。その日の写しが既にあれば作らない（1日に1回）。
'       データがないとき、フォルダが未設定か存在しないときは、何もしない。
' 引数：blnDummy - ［マクロ］ダイアログに出さないためのダミー（使わない）
' 戻り値：なし
' 備考：ブックの名前は、配布のときは版つき（電話メモVerNNN.xlsm）、インストールの後は
'       版なし（電話メモ.xlsm）になる。どちらでも同じ写しの名前になるよう、GetBackupBaseName で求める。
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
    Set tbl = sht11.ListObjects("電話メモ")
    If tbl.ListRows.Count = 0 Then GoTo Cleanup

    'オプションシートのバックアップフォルダを読む（名前の定義がないときは、空のままにする）
    On Error Resume Next
    strBackupFolder = CStr(shtOpt.Range("バックアップフォルダ").Value)
    On Error GoTo HandleErr

    'フォルダが未設定か、存在しないときは、何もしない
    If strBackupFolder = "" Then GoTo Cleanup
    If Dir(strBackupFolder, vbDirectory) = "" Then GoTo Cleanup

    '--- メイン処理 ---
    '写しのフルパスを作る（例：<フォルダ>\20261119_電話メモ.xlsx）
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
           "内容：" & Err.Description, vbExclamation, "VA電話メモ"
    Resume Cleanup
End Sub

'--------------------------------------------------------------------------------
' 概要：ブックのファイル名から、バックアップの写しの名前に使う部分を返す。
'       拡張子を除き、末尾が版（「Ver」と英数字3文字）なら、それも除く。
'       例：「電話メモVerNNN.xlsm」→「電話メモ」、「電話メモ.xlsm」→「電話メモ」
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

