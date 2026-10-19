' 変数宣言を強制する
Option Explicit

' クラス概要: ブックの保存・開く・閉じる動作を監視し、排他制御と同期を制御する。
' 更新日時: 2026/04/04 12:44

' アプリケーション全体のイベントをフックするための変数を宣言する
Public WithEvents App As Application

' 処理の目的: ブックが開かれた時の初期化処理
' 引数: Wb - 開かれた対象のブック
' 戻り値: なし
' ブックが開かれた際に実行されるイベントプロシージャ
Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    ' 開かれたブックがアドイン自身であれば処理を終了する
    If Wb.IsAddin Then Exit Sub
    
    ' Shiftキーが押されながら開かれたか判定する
    If basUtility.IsShiftKeyPressed() Then
        ' メンテナンスモードとして対象ブックを登録する
        basMain.RegisterMaintenanceBook Wb
        ' イベント処理を終了する
        Exit Sub
    ' 判定の終了
    End If
    
    ' 対象ブックに管理対象シートが存在するか判定する
    If HasManagedSheets(Wb) Then
        ' エラーが発生しても次の行へ進むよう設定する
        On Error Resume Next
        ' クラウドの自動保存機能を強制的に無効化する
        Wb.AutoSaveOn = False
        ' エラーハンドリングを標準状態に戻す
        On Error GoTo 0
        
        ' 遅延実行のためにグローバル変数に対象ブックを保持する
        Set basMain.gTargetBook = Wb
        
        ' 1秒後に読み込み処理を遅延実行するようスケジュールを登録する
        Application.OnTime Now + TimeValue("00:00:01"), "DelayedStartup"
    ' 判定の終了
    End If
' プロシージャの終了
End Sub

' 処理の目的: ブック保存時のブロック処理
' 引数: Wb - 保存対象のブック, SaveAsUI - 名前を付けて保存ダイアログ表示フラグ, Cancel - 保存キャンセルフラグ
' 戻り値: なし
' ブックが保存される直前に実行されるイベントプロシージャ
Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' アドインまたは管理対象外のブックであれば処理を終了する
    If Wb.IsAddin Or Not HasManagedSheets(Wb) Then Exit Sub
    
    ' メンテナンスモードまたはデータが存在しない場合は処理を終了する
    If basMain.IsMaintenanceBook(Wb) Or Not HasDataInManagedSheets(Wb) Then Exit Sub
    
    ' セキュリティ確保のためExcel標準の保存動作をキャンセルする
    Cancel = True
    
    ' ユーザーに保存不可のメッセージを表示する
    MsgBox "セキュリティ確保のため、このファイルの保存は禁止されています。" & vbCrLf & "マスタファイルへのデータ反映は、リボンの「データ書き込み」を使用してください。", vbExclamation, cstSettings.MSG_TITLE
' プロシージャの終了
End Sub

' 処理の目的: ブックを閉じる際の終了確認とクリーンアップ
' 引数: Wb - 閉じる対象のブック, Cancel - 閉じるキャンセルフラグ
' 戻り値: なし
' ブックが閉じられる直前に実行されるイベントプロシージャ
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    ' アドインまたは管理対象外のブックであれば処理を終了する
    If Wb.IsAddin Or Not HasManagedSheets(Wb) Then Exit Sub
    
    ' メンテナンスモードまたはデータなしの場合を判定する
    If basMain.IsMaintenanceBook(Wb) Or Not HasDataInManagedSheets(Wb) Then
        ' メンテナンスモードの登録を解除する
        If basMain.IsMaintenanceBook(Wb) Then basMain.UnregisterMaintenanceBook Wb
        ' 特殊な終了処理を行わず処理を終了する
        Exit Sub
    ' 判定の終了
    End If
    
    ' マスタファイルのパスを格納する変数を宣言する
    Dim strMasterPath As String
    ' 現在のOS環境に合わせてマスタパスを取得し変換する
    strMasterPath = basUtility.GetLocalPath(basConfig.GetBookSetting(cstSettings.PROP_MASTER_PATH, Wb))
    
    ' 対象ブックにおいて自身の有効なロック(編集モード)が存在するか判定する
    If basSync.hasMyLock(Wb) Then
        ' 対象ブックに未保存の変更があるか判定する
        If Not Wb.Saved Then
            ' ユーザーの選択結果を格納する変数を宣言する
            Dim intAns As VbMsgBoxResult
            ' 保存確認のメッセージを表示し結果を取得する
            intAns = MsgBox("マスタファイルにデータを書き込みますか？", vbYesNoCancel + vbQuestion, cstSettings.MSG_TITLE)
            
            ' ユーザーの選択結果に応じて処理を分岐する
            Select Case intAns
                ' 「はい」が選択された場合
                Case vbYes
                    ' マスタへの書き込み処理を実行し、失敗したかを判定する
                    If Not basSync.writeToMaster(Wb, True) Then
                        ' 書き込みに失敗した場合は閉じる処理をキャンセルする
                        Cancel = True
                        ' プロシージャを終了する
                        Exit Sub
                    ' 書き込み判定の終了
                    End If
                    
                ' 「いいえ」が選択された場合
                Case vbNo
                    ' 保存せずに終了するためマスタのロックを解放する
                    basSync.releaseLock strMasterPath
                    
                ' 「キャンセル」が選択された場合
                Case vbCancel
                    ' 閉じる処理自体をキャンセルする
                    Cancel = True
                    ' プロシージャを終了する
                    Exit Sub
            ' 分岐処理の終了
            End Select
        ' 未保存の変更がない場合の処理
        Else
            ' 変更がない場合でも安全のためにロックを解放する
            basSync.releaseLock strMasterPath
        ' 未保存判定の終了
        End If
    ' 自身の有効なロックが存在しない(閲覧モード)の場合の処理
    Else
        ' 念のため残存している自身の過去のロック解放を試みる
        basSync.releaseLock strMasterPath
    ' ロック状態判定の終了
    End If
    
    ' 開いている可視状態のブック数をカウントする変数を宣言する
    Dim lngVisibleWbCount As Long
    ' ループ処理用のブック変数を宣言する
    Dim objLoopWb As Workbook
    ' 開かれているすべてのブックをループ処理する
    For Each objLoopWb In Application.Workbooks
        ' アドイン自身ではないブックか判定する
        If Not objLoopWb.IsAddin Then
            ' ブックにウィンドウが存在するか判定する
            If objLoopWb.Windows.Count > 0 Then
                ' ウィンドウが可視状態であるか判定する
                If objLoopWb.Windows(1).Visible Then
                    ' 可視状態のブック数を1つカウントアップする
                    lngVisibleWbCount = lngVisibleWbCount + 1
                ' 可視状態判定の終了
                End If
            ' ウィンドウ存在判定の終了
            End If
        ' アドイン判定の終了
        End If
    ' 次のブックへループを進める
    Next objLoopWb
    
    ' 可視状態のブックが最後の1つであるか判定する
    If lngVisibleWbCount <= 1 Then
        ' イベントの連鎖を防ぐためイベント発生を一時無効化する
        Application.EnableEvents = False
        ' 確認ダイアログの表示を一時無効化する
        Application.DisplayAlerts = False
        ' 保存済み状態と偽装して保存確認ダイアログを防ぐ
        Wb.Saved = True
        
        ' Excelアプリケーション全体を強制的に終了させる
        Application.Quit
    ' 他に開いているブックが存在する場合の処理
    Else
        ' 標準の閉じる処理を一旦キャンセルして独自処理を行う
        Cancel = True
        
        ' イベントの発生を一時無効化する
        Application.EnableEvents = False
        ' 確認ダイアログの表示を一時無効化する
        Application.DisplayAlerts = False
        ' 保存済み状態と偽装する
        Wb.Saved = True
        
        ' 対象のブックのみを保存せずに強制終了させる
        Wb.Close SaveChanges:=False
        
        ' 無効化した確認ダイアログの設定を元に戻す
        Application.DisplayAlerts = True
        ' 無効化したイベント発生の設定を元に戻す
        Application.EnableEvents = True
    ' 終了分岐判定の終了
    End If
' プロシージャの終了
End Sub

' 処理の目的: ブック内の管理対象シートの有無を判定する
' 引数: Wb - 判定対象のブック
' 戻り値: 管理対象シートが存在すればTrue
' 管理対象シートの存在を判定するプライベート関数
Private Function HasManagedSheets(Wb As Workbook) As Boolean
    ' ループ処理用のワークシート変数を宣言する
    Dim wsLoop As Worksheet
    ' ブック内の全シートをループ処理する
    For Each wsLoop In Wb.Worksheets
        ' 現在のシートが管理対象であるか判定する
        If basUtility.IsManagedSheet(wsLoop) Then
            ' 管理対象であれば戻り値をTrueに設定する
            HasManagedSheets = True
            ' 存在が確認できたため関数を終了する
            Exit Function
        ' 判定の終了
        End If
    ' 次のシートへループを進める
    Next wsLoop
    ' 管理対象シートが一つもないため戻り値をFalseに設定する
    HasManagedSheets = False
' 関数の終了
End Function

' 処理の目的: ブック内の管理対象テーブルにデータが存在するか判定する
' 引数: Wb - 判定対象のブック
' 戻り値: データが存在すればTrue
' データの存在を判定するプライベート関数
Private Function HasDataInManagedSheets(Wb As Workbook) As Boolean
    ' ループ処理用のワークシート変数を宣言する
    Dim wsLoop As Worksheet
    ' テーブルを操作するためのリストオブジェクト変数を宣言する
    Dim objTbl As ListObject
    
    ' 戻り値のデフォルトをFalse(データなし)に初期化する
    HasDataInManagedSheets = False
    
    ' ブック内の全シートをループ処理する
    For Each wsLoop In Wb.Worksheets
        ' 現在のシートが管理対象であるか判定する
        If basUtility.IsManagedSheet(wsLoop) Then
            ' テーブル取得時のエラーを一時的に無視する
            On Error Resume Next
            ' シート内の最初のテーブルオブジェクトを取得する
            Set objTbl = wsLoop.ListObjects(1)
            ' エラーハンドリングを標準状態に戻す
            On Error GoTo 0
            
            ' テーブルオブジェクトが正常に取得できたか判定する
            If Not objTbl Is Nothing Then
                ' テーブルのデータ行が存在するか判定する
                If Not objTbl.DataBodyRange Is Nothing Then
                    ' データが存在するため戻り値をTrueに設定する
                    HasDataInManagedSheets = True
                    ' データが存在するため直ちに関数を終了する
                    Exit Function
                ' データ行存在判定の終了
                End If
            ' テーブル存在判定の終了
            End If
        ' 管理対象判定の終了
        End If
    ' 次のシートへループを進める
    Next wsLoop
' 関数の終了
End Function
