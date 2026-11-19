'更新日時: 2026/02/22 16:44
'概要: プロジェクト全体で使用するグローバル変数の宣言と、共通の初期化処理やバックアップ処理を行う
' 変数の宣言を強制する
Option Explicit

' 対象事件番号用のパブリック長整数型変数を宣言する
Public lngActCase As Long
' Evntクラスのインスタンスを生成してパブリック変数に代入する
Public Evnt As New Evnt

'--------------------------------------------------
' 概要: 外部連携等のロード完了後に実行される初期化処理
' 引数: dummy (Byte)
' 戻り値: なし
'--------------------------------------------------
' VAブリッジのロード完了後イベントプロシージャ
Public Sub VABridge_AfterLoad(Optional ByVal dummy As Byte = 0)
    '--- 前処理 ---
    ' エラー発生時はHandleErrラベルへジャンプする
    On Error GoTo HandleErr
    
    ' イベント処理および画面更新を一時的に無効化する
    Evnt.Disable

    ' Worksheet型変数wsHokanを宣言する
    Dim wsHokan As Worksheet
    ' 保管シートのオブジェクトをwsHokanにセットする
    Set wsHokan = ShtHokan
    
    ' 保管シートをアクティブにする
    wsHokan.Activate
    ' 保管シートのテーブルフォーマット処理を呼び出す
    Call ShtHokan.FormatTable

    ' 行番号用の長整数型変数を宣言する
    Dim r As Long
    ' ListObject型変数を宣言する
    Dim tbl As ListObject
    ' 保管シート内の「保管」テーブルをtblにセットする
    Set tbl = wsHokan.ListObjects("保管")
    
    ' テーブルの行数分だけループ処理を行う
    For r = 1 To tbl.ListRows.Count
        ' 保管日時が空の行を見つけたらループを抜ける
        If tbl.ListColumns("保管日時").DataBodyRange(r).Value = "" Then Exit For
    ' 次の行へ進む
    Next

    ' テーブル内に空欄の行が存在した場合
    If r <= tbl.ListRows.Count Then
        ' 該当行の種別セルを選択する
        tbl.ListColumns("種別").DataBodyRange(r).Select
        ' 選択した行が画面内に表示されるようにスクロールする
        ActiveWindow.ScrollRow = tbl.ListColumns("種別").DataBodyRange(r).Row
    ' 条件分岐を終了する
    End If

    ' 日次バックアップ処理を呼び出す
    Call PerformDailyBackup

Cleanup:
    '--- 後処理 ---
    ' イベント処理および画面更新を、停止する前の値に戻す（エラーのときも通る）
    On Error Resume Next
    Evnt.Enable
    Exit Sub

HandleErr:
    ' エラーの番号と内容を表示して、後処理へ進む
    MsgBox "データを読み込んだ後の処理でエラーが発生しました。" & vbCrLf & _
           "エラー番号：" & Err.Number & vbCrLf & _
           "内容：" & Err.Description, vbExclamation, strAppTitle
    Resume Cleanup
End Sub

'--------------------------------------------------------------------------------
' 概要：設定したバックアップフォルダに、このブックの写し（マクロなしの .xlsx）を
'       「yyyymmdd_領収書等ファイル保管.xlsx」の名前で作る。その日の写しが既にあれば作らない（1日に1回）。
'       データがないとき、フォルダが未設定か存在しないときは、何もしない。
' 引数：blnDummy - ［マクロ］ダイアログに出さないためのダミー（使わない）
' 戻り値：なし
' 備考：ブックの名前は、配布のときは版つき（領収書等ファイル保管VerNNN.xlsm）、インストールの後は
'       版なし（領収書等ファイル保管.xlsm）になる。どちらでも同じ写しの名前になるよう、GetBackupBaseName で求める。
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
    Set tbl = ShtHokan.ListObjects("保管")
    If tbl.ListRows.Count = 0 Then GoTo Cleanup

    'オプションシートのバックアップフォルダを読む（名前の定義がないときは、空のままにする）
    On Error Resume Next
    strBackupFolder = CStr(ShtOption.Range("バックアップフォルダ").Value)
    On Error GoTo HandleErr

    'フォルダが未設定か、存在しないときは、何もしない
    If strBackupFolder = "" Then GoTo Cleanup
    If Dir(strBackupFolder, vbDirectory) = "" Then GoTo Cleanup

    '--- メイン処理 ---
    '写しのフルパスを作る（例：<フォルダ>\20261119_領収書等ファイル保管.xlsx）
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
           "内容：" & Err.Description, vbExclamation, strAppTitle
    Resume Cleanup
End Sub

'--------------------------------------------------------------------------------
' 概要：ブックのファイル名から、バックアップの写しの名前に使う部分を返す。
'       拡張子を除き、末尾が版（「Ver」と英数字3文字）なら、それも除く。
'       例：「領収書等ファイル保管VerNNN.xlsm」→「領収書等ファイル保管」、「領収書等ファイル保管.xlsm」→「領収書等ファイル保管」
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

