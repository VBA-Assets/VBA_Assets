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
    ' エラー発生時はErrHandlerラベルへジャンプする
    On Error GoTo ErrHandler
    
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
    ' イベント処理および画面更新を有効化して元に戻す
    Evnt.Enable
    ' エラーがなければプロシージャを終了する
    Exit Sub

' エラー処理用のラベル
ErrHandler:
    ' エラー内容をメッセージボックスで表示する
    MsgBox "初期化処理中にエラーが発生しました。", vbCritical, "ファイル保管"
    ' イベント処理および画面更新を有効化して元に戻す
    Evnt.Enable
' プロシージャを終了する
End Sub

'--------------------------------------------------
' 概要: 現在のブックのバックアップを作成します。
' 引数: dummy (Byte)
' 戻り値: なし
'--------------------------------------------------
' バックアップ作成プロシージャ
Sub PerformDailyBackup(Optional ByVal dummy As Byte = 0)
    ' バックアップフォルダパス格納用の文字列変数を宣言する
    Dim backupFolder As String
    ' バックアップファイル名格納用の文字列変数を宣言する
    Dim backupFileName As String
    ' 保存先フルパス格納用の文字列変数を宣言する
    Dim targetPath As String
    ' パスセパレータ格納用の文字列変数を宣言する
    Dim PS As String
    ' ベースファイル名格納用の文字列変数を宣言する
    Dim baseName As String
    ' 新規ブック格納用のオブジェクト変数を宣言する
    Dim wbNew As Workbook
    ' テーブル情報格納用のオブジェクト変数を宣言する
    Dim tbl As ListObject

    ' 「保管」テーブルをtblにセットする
    Set tbl = ShtHokan.ListObjects("保管")
    ' テーブルにデータが存在しない場合はバックアップを作成せず終了する
    If tbl.ListRows.Count = 0 Then Exit Sub
    
    ' OS標準のパスセパレータ（\など）を取得する
    PS = Application.PathSeparator
    
    ' エラーが発生しても次の行の処理を続行する
    On Error Resume Next
    ' オプションシートからバックアップフォルダパスを取得する
    backupFolder = ShtOption.Range("バックアップフォルダ").Value
    ' エラー処理を初期状態に戻す
    On Error GoTo 0
    
    ' バックアップフォルダが指定されていなければ処理を終了する
    If backupFolder = "" Then Exit Sub
    
    ' 以降でエラーが発生した場合はErrHandlerラベルへ飛ぶ
    On Error GoTo ErrHandler
    ' バックアップ先フォルダが存在しない場合は処理を終了する
    If Dir(backupFolder, vbDirectory) = "" Then Exit Sub
    
    ' 現在のブック名から「V」より前の部分をベース名として抽出する
    baseName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "V") - 1)
    ' 日付とベース名を結合し、バックアップファイル名を生成する
    backupFileName = Format(Date, "yyyymmdd") & "_" & baseName & ".xlsx"
    
    ' フォルダパスの末尾がパスセパレータでない場合
    If Right(backupFolder, 1) <> PS Then
        ' パスセパレータを挟んでフルパスを結合する
        targetPath = backupFolder & PS & backupFileName
    ' フォルダパスの末尾がパスセパレータである場合
    Else
        ' そのままファイル名と結合する
        targetPath = backupFolder & backupFileName
    ' 条件分岐を終了する
    End If
    
    ' 同日のバックアップファイルがまだ存在しない場合
    If Dir(targetPath) = "" Then
        ' 確認メッセージの表示をオフにする
        Application.DisplayAlerts = False
        ' 現在のブックのすべてのシートを新規ブックにコピーする
        ThisWorkbook.Sheets.Copy
        ' コピーして新しく開かれたブックをwbNewにセットする
        Set wbNew = ActiveWorkbook
        
        ' 新規ブックを標準のxlsx形式でバックアップフォルダに保存する
        wbNew.SaveAs Filename:=targetPath, FileFormat:=xlOpenXMLWorkbook
        ' 新規ブックを保存せずに（保存済みのためそのまま）閉じる
        wbNew.Close SaveChanges:=False
        
        ' 確認メッセージの表示をオンに戻す
        Application.DisplayAlerts = True
    ' 条件分岐を終了する
    End If
    ' プロシージャを正常終了する
    Exit Sub

' エラー処理用のラベル
ErrHandler:
    ' エラーの詳細状況を表示する
    MsgBox "自動バックアップの作成中にエラーが発生しました。" & vbCrLf & _
           "保存先: " & backupFolder & vbCrLf & _
           "詳細状況: " & Err.Description, vbExclamation, "ファイル保管"
    ' 確認メッセージの表示をオンに戻す
    Application.DisplayAlerts = True
' プロシージャを終了する
End Sub

