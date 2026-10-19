' 変数宣言を強制する
Option Explicit

' モジュール概要: マスタファイルとの読み込み・書き込み・強力な排他制御を行うコアロジック
' 更新日時: 2026/06/07 13:20

' ロックファイルの拡張子を定数として定義する
Private Const EXT_LOCK_SUFFIX As String = ".lock"

' -----------------------------------------------------------------------------
' プロシージャ概要: 自身のロックファイルが有効か判定する(ステートレス判定)
' 引数: TargetBook = 判定対象のブック
' 戻り値: 有効なロックが存在すればTrue、それ以外はFalse
' -----------------------------------------------------------------------------
' 自身のロックファイルを判定するパブリック関数
Public Function hasMyLock(ByVal TargetBook As Workbook) As Boolean
    ' エラー発生時はErrorHandlerラベルへジャンプするよう設定する
    On Error GoTo ErrorHandler
    
    ' マスタパスを格納する変数を宣言する
    Dim strMasterPath As String
    ' 対象ブックの設定からマスタパスを取得しローカルパスに変換して代入する
    strMasterPath = basUtility.GetLocalPath(basConfig.GetBookSetting(cstSettings.PROP_MASTER_PATH, TargetBook))
    
    ' マスタパスが空文字かどうか判定する
    If strMasterPath = "" Then
        ' 空の場合は戻り値をFalseに設定する
        hasMyLock = False
        ' 関数を抜ける
        Exit Function
    ' 空文字判定の終了
    End If
    
    ' フォルダパスを格納する変数を宣言する
    Dim strFolderPath As String
    ' マスタパスからフォルダ部分を抽出して代入する
    strFolderPath = Left(strMasterPath, InStrRev(strMasterPath, Application.PathSeparator))
    
    ' マスタファイル名を格納する変数を宣言する
    Dim strMasterFileName As String
    ' マスタパスからファイル名部分を抽出して代入する
    strMasterFileName = Mid(strMasterPath, InStrRev(strMasterPath, Application.PathSeparator) + 1)
    
    ' 自身のユーザー名を格納する変数を宣言する
    Dim strMyUser As String
    ' ネットワークユーザー名を取得しファイル名として使えるようサニタイズして代入する
    strMyUser = basUtility.SanitizeFileName(basUtility.GetNetworkUserName())
    
    ' 自身のロックファイルのフルパスを格納する変数を宣言する
    Dim strMyLockPath As String
    ' フォルダパスとロックファイル名を結合して代入する
    strMyLockPath = basUtility.BuildPath(strFolderPath, strMasterFileName & "." & strMyUser & EXT_LOCK_SUFFIX)
    
    ' ロックファイルが存在するかどうか判定する
    If basUtility.FileExists(strMyLockPath) Then
        ' マスタファイルの更新日時を格納する変数を宣言する
        Dim dtmMasterTime As Date
        ' 一時的にエラーを無視する設定に変更する
        On Error Resume Next
        ' マスタファイルの更新日時を取得して代入する
        dtmMasterTime = FileDateTime(strMasterPath)
        ' エラー処理をErrorHandlerへ戻す設定に変更する
        On Error GoTo ErrorHandler
        
        ' ロックファイルの更新日時を格納する変数を宣言する
        Dim dtmLockTime As Date
        ' ロックファイルの更新日時を取得して代入する
        dtmLockTime = FileDateTime(strMyLockPath)
        
        ' ロックファイルがマスタファイルより古いかどうか判定する
        If dtmLockTime < dtmMasterTime Then
            ' 古い場合は無効とみなし戻り値をFalseに設定する
            hasMyLock = False
            ' 関数を抜ける
            Exit Function
        ' 古いかどうかの判定終了
        End If
        
        ' 全ての条件をクリアした場合は有効とみなし戻り値をTrueに設定する
        hasMyLock = True
    ' ロックファイルが存在しない場合の処理
    Else
        ' 存在しない場合は戻り値をFalseに設定する
        hasMyLock = False
    ' ファイル存在判定の終了
    End If
    
    ' 正常終了のため関数を抜ける
    Exit Function
    
' エラー発生時のジャンプ先ラベル
ErrorHandler:
    ' エラー時は安全のため戻り値をFalseに設定する
    hasMyLock = False
' 関数の終了
End Function

' -----------------------------------------------------------------------------
' プロシージャ概要: マスタからのデータ読み込み処理
' 引数: TargetBook = 処理対象のブック, isSilent = メッセージ表示抑制フラグ
' 戻り値: 成功ならTrue、失敗・中止ならFalse
' -----------------------------------------------------------------------------
' マスタからの読み込み処理を定義するパブリック関数
Public Function loadFromMaster(ByVal TargetBook As Workbook, ByVal isSilent As Boolean) As Boolean
    ' 対象ブックのウィンドウが存在しないか判定する
    If TargetBook.Windows.Count = 0 Then
        ' 戻り値をFalse(失敗)に設定する
        loadFromMaster = False
        ' 関数を終了して抜ける
        Exit Function
    ' ウィンドウ判定の終了
    End If

    ' マスタファイルのパスを格納する変数を宣言する
    Dim strMasterPath As String
    ' 設定からマスタパスを取得しローカル環境に合わせて変換する
    strMasterPath = basUtility.GetLocalPath(basConfig.GetBookSetting(cstSettings.PROP_MASTER_PATH, TargetBook))
    
    ' マスタパスが未設定または存在しないか判定する
    If strMasterPath = "" Or Not basUtility.FileExists(strMasterPath) Then
        ' エラーメッセージを表示する
        If Not isSilent Then MsgBox "マスタファイルが見つかりません。" & vbCrLf & "パス: " & strMasterPath, vbCritical, cstSettings.MSG_TITLE
        ' 関数を終了して抜ける
        Exit Function
    ' 存在判定の終了
    End If

    ' 編集モード中に未書き込みの変更がある場合の警告処理
    If Not isSilent Then
        If hasMyLock(TargetBook) And Not TargetBook.Saved Then
            ' ユーザーの選択結果を格納する変数を宣言する
            Dim intReloadAns As VbMsgBoxResult
            ' 書き込み確認のメッセージを表示し結果を取得する
            intReloadAns = MsgBox("編集中のデータがマスタへ書き込まれていません。" & vbCrLf & "再読み込みする前にマスタファイルへデータを書き込みますか？", vbYesNoCancel + vbQuestion, cstSettings.MSG_TITLE)

            ' ユーザーの選択結果に応じて処理を分岐する
            Select Case intReloadAns
                ' 「はい」が選択された場合: 先に書き込みを実行してから読み込みへ続く
                Case vbYes
                    If Not writeToMaster(TargetBook, True) Then
                        ' 書き込みに失敗した場合は読み込みを中止する
                        loadFromMaster = False
                        Exit Function
                    End If
                ' 「いいえ」が選択された場合: 変更を破棄してロックを解放してから読み込みへ続く
                Case vbNo
                    releaseLock strMasterPath
                ' 「キャンセル」が選択された場合: 読み込み操作自体を中止する
                Case vbCancel
                    loadFromMaster = False
                    Exit Function
            ' 分岐処理の終了
            End Select
        ' 未書き込み変更判定の終了
        End If
    ' サイレント判定の終了
    End If

    ' 他のユーザーによってロックされているか判定するフラグ変数を宣言する
    Dim blnIsLockedByOther As Boolean
    ' ロックを取得しているユーザー名を格納する変数を宣言する
    Dim strLockUser As String
    ' ロックが取得された日時を格納する変数を宣言する
    Dim dtmLockTime As Date
    ' ユーザーが編集モードを要求したかを示すフラグ変数を宣言する
    Dim blnRequestEdit As Boolean
    
    ' マスタファイルの現在のロック状態を確認し結果を変数に代入する
    blnIsLockedByOther = isLocked(strMasterPath, strLockUser, dtmLockTime)
    
    ' 他のユーザーがロック中であるか判定する
    If blnIsLockedByOther Then
        ' 編集モードの要求を強制的にFalse(閲覧のみ)に設定する
        blnRequestEdit = False
        ' サイレントモードではないか判定する
        If Not isSilent Then
            ' ユーザーの確認結果を格納する変数を宣言する
            Dim intConfirmAns As VbMsgBoxResult
            ' 閲覧モードで開く旨を通知し結果を取得する
            intConfirmAns = MsgBox("現在、ユーザー [" & strLockUser & "] が編集中です。" & vbCrLf & "閲覧モード(読み取り専用)でデータを読み込みます。", vbOKCancel + vbExclamation, cstSettings.MSG_TITLE)
            ' ユーザーがキャンセルを選択したか判定する
            If intConfirmAns = vbCancel Then
                ' 処理を中止するため戻り値をFalseに設定する
                loadFromMaster = False
                ' 関数を終了する
                Exit Function
            ' キャンセル判定の終了
            End If
        ' サイレント判定の終了
        End If
    ' ロックされておらずサイレントモードであるか判定する
    ElseIf isSilent Then
        ' 自動実行時は安全のため強制的に閲覧モード(False)とする
        blnRequestEdit = False
    ' ロックされておらずユーザー操作による実行である場合の処理
    Else
        ' モード選択フォームのインスタンス用変数を宣言する
        Dim objFrm As frmModeSelect
        ' フォームの新しいインスタンスを生成する
        Set objFrm = New frmModeSelect
        ' フォームに対象ブックの参照を渡す
        Set objFrm.TargetBook = TargetBook
        ' モーダルダイアログとしてフォームを画面に表示する
        objFrm.Show vbModal
        ' フォームでキャンセル操作が行われたか判定する
        If objFrm.IsCancelled Then
            ' フォームをメモリから破棄する
            Unload objFrm
            ' オブジェクト変数を解放する
            Set objFrm = Nothing
            ' 処理を中止するため戻り値をFalseに設定する
            loadFromMaster = False
            
            ' ローカルテーブルにデータが存在するか判定するフラグ変数を宣言する
            Dim hasLocalData As Boolean
            ' フラグをFalseで初期化する
            hasLocalData = False
            
            ' データ存在チェック用のシート変数を宣言する
            Dim wsCheck As Worksheet
            ' データ存在チェック用のテーブル変数を宣言する
            Dim objTblCheck As ListObject
            
            ' 対象ブックの全シートをループ処理する
            For Each wsCheck In TargetBook.Worksheets
                ' 管理対象シートであるか判定する
                If basUtility.IsManagedSheet(wsCheck) Then
                    ' オブジェクト変数を初期化する
                    Set objTblCheck = Nothing
                    ' テーブル取得エラーを一時的に無視する
                    On Error Resume Next
                    ' シートの1つ目のテーブルオブジェクトを取得する
                    Set objTblCheck = wsCheck.ListObjects(1)
                    ' エラー処理の設定を元に戻す
                    On Error GoTo 0
                    
                    ' テーブルオブジェクトが正常に取得できたか判定する
                    If Not objTblCheck Is Nothing Then
                        ' テーブルのデータ範囲が存在するか判定する
                        If Not objTblCheck.DataBodyRange Is Nothing Then
                            ' データが存在するためフラグをTrueに設定する
                            hasLocalData = True
                            ' 1つでもデータがあれば十分なためループを抜ける
                            Exit For
                        ' データ存在判定の終了
                        End If
                    ' テーブル存在判定の終了
                    End If
                ' 管理対象シート判定の終了
                End If
            ' 次のシートへループを進める
            Next wsCheck
            
            ' ローカルテーブルにデータが存在しないか判定する
            If Not hasLocalData Then
                ' 開いている可視状態のブック数をカウントする変数を宣言する
                Dim lngVisibleWbCount As Long
                ' ループ用のブック変数を宣言する
                Dim objLoopWb As Workbook
                
                ' カウント変数を初期化する
                lngVisibleWbCount = 0
                
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
                            ' 可視状態の判定を終了する
                            End If
                        ' ウィンドウ存在の判定を終了する
                        End If
                    ' アドイン判定を終了する
                    End If
                ' 次のブックへ処理を進める
                Next objLoopWb
                
                ' イベント発生を一時無効化する
                Application.EnableEvents = False
                ' 確認ダイアログの表示を一時無効化する
                Application.DisplayAlerts = False
                ' 保存済み状態と偽装して保存確認ダイアログを防ぐ
                TargetBook.Saved = True
                
                ' 可視状態のブックが最後の1つであるか判定する
                If lngVisibleWbCount <= 1 Then
                    ' Excelアプリケーション全体を強制的に終了させる
                    Application.Quit
                ' ほかに開いているブックがある場合の処理
                Else
                    ' 対象のブックのみを保存せずに強制終了させる
                    TargetBook.Close SaveChanges:=False
                    
                    ' 無効化したダイアログ表示設定を元に戻す
                    Application.DisplayAlerts = True
                    ' 無効化したイベント発生設定を元に戻す
                    Application.EnableEvents = True
                ' 可視状態のブック数の判定を終了する
                End If
            ' データ非存在判定の終了
            End If
            
            ' 関数を終了する
            Exit Function
        ' フォームで有効な選択が行われた場合の処理
        Else
            ' ユーザーが選択したモードを取得する
            blnRequestEdit = objFrm.ResultEditMode
        ' キャンセル判定の終了
        End If
        ' フォームをメモリから破棄する
        Unload objFrm
        ' オブジェクト変数を解放する
        Set objFrm = Nothing
    ' モード決定処理の終了
    End If
    
    ' 編集モードが要求されたか判定する
    If blnRequestEdit Then
        ' ロック取得結果を格納する一時変数を宣言する
        Dim blnAcquired As Boolean
        ' 排他制御用のロック取得処理を実行する
        blnAcquired = acquireLock(strMasterPath)
    ' 閲覧モードが選択された場合の処理
    Else
        ' 閲覧モードの場合は念のため自身の過去のロックを解放する
        releaseLock strMasterPath
    ' モード要求判定の終了
    End If

    ' 画面の描画更新を停止して処理速度を向上させる
    Application.ScreenUpdating = False
    ' イベント発生を抑止して余計な処理を防ぐ
    Application.EnableEvents = False
    ' 確認メッセージ等の表示を抑止する
    Application.DisplayAlerts = False
    
    ' 実行処理開始時にステータスバーを表示する
    Application.StatusBar = "マスタファイルのデータを読み込み中です..."

    ' エラーが発生した場合のジャンプ先を指定する
    On Error GoTo ErrorHandler

    ' マスタブック操作用の変数を宣言する
    Dim wbMaster As Workbook
    ' シート操作用の変数を宣言する
    Dim wsLoop As Worksheet
    ' 処理前の選択シートを記憶するための変数を宣言する
    Dim objOriginalSh As Object
    ' 処理前の選択セル位置を記憶するための変数を宣言する
    Dim strOriginalCellAddr As String
    
    ' 現在アクティブなシートとセルの位置記憶時にエラーを無視する
    On Error Resume Next
    ' 現在アクティブなシートオブジェクトを記憶する
    Set objOriginalSh = ActiveSheet
    ' 選択範囲がセルの場合はアドレスを記憶する
    If TypeName(Selection) = "Range" Then strOriginalCellAddr = ActiveCell.Address
    ' エラー処理を標準のハンドラに戻す
    On Error GoTo ErrorHandler

    ' マスタファイルをリンク更新なしの読み取り専用で開く
    Set wbMaster = Workbooks.Open(strMasterPath, ReadOnly:=True, UpdateLinks:=False)
    
    ' 対象ブックの全シートに対してループ処理を行う
    For Each wsLoop In TargetBook.Worksheets
        ' 管理対象のシートであるかチェックする
        If basUtility.IsManagedSheet(wsLoop) Then
            ' データ書き込みのためシート保護を一時的に解除する
            wsLoop.Unprotect
            
            ' マスタ側の対応するシート名を格納する変数を宣言する
            Dim strMShName As String
            ' シート設定からマスタシート名を取得する
            strMShName = basConfig.GetSheetSetting(wsLoop, cstSettings.PROP_SUFFIX_SHEET)
            
            ' ローカルのテーブルオブジェクト変数を宣言する
            Dim objTblLocal As ListObject
            ' マスタのテーブルオブジェクト変数を宣言する
            Dim objTblMaster As ListObject
            
            ' テーブル取得時のエラーを一時的に無視する
            On Error Resume Next
            ' ローカルシートの1つ目のテーブルを取得する
            Set objTblLocal = wsLoop.ListObjects(1)
            ' マスタシートの1つ目のテーブルを取得する
            Set objTblMaster = wbMaster.Worksheets(strMShName).ListObjects(1)
            ' エラー処理を復帰させる
            On Error GoTo ErrorHandler
            
            ' 両方のテーブルが正常に取得できた場合のみ処理を行う
            If Not objTblLocal Is Nothing And Not objTblMaster Is Nothing Then
                ' ヘッダー列構造が一致しているか確認する
                If checkHeaderConsistency(objTblLocal, objTblMaster) Then
                    ' データが残留していた場合の事故を防ぐためローカルデータを削除する
                    If Not objTblLocal.DataBodyRange Is Nothing Then objTblLocal.DataBodyRange.Delete
                    
                    ' マスタテーブルにデータが存在するか判定する
                    If Not objTblMaster.DataBodyRange Is Nothing Then
                        ' マスタのデータ部分をコピーする
                        objTblMaster.DataBodyRange.Copy
                        
                        ' クリップボードの安定化のためOS側に処理を一時譲る
                        DoEvents
                        
                        ' ローカルテーブルのヘッダーの1行下に貼り付ける
                        objTblLocal.HeaderRowRange.Offset(1, 0).PasteSpecial xlPasteAll
                        ' コピー状態を解除する
                        Application.CutCopyMode = False
                    ' マスタデータ存在判定の終了
                    End If
                ' ヘッダーが不一致の場合の処理
                Else
                     ' 警告を表示する
                     MsgBox "シート [" & wsLoop.Name & "] のヘッダー構造がマスタと一致しません。同期をスキップしました。", vbExclamation, cstSettings.MSG_TITLE
                ' ヘッダー整合性判定の終了
                End If
            ' テーブル存在判定の終了
            End If
        ' 管理対象判定の終了
        End If
    ' 次のシートへループを進める
    Next wsLoop
    
    ' 読み込み完了時点のマスタ更新日時を最終同期日時として保存する
    basConfig.SaveBookSetting cstSettings.PROP_LAST_SYNC, CStr(FileDateTime(strMasterPath)), TargetBook
    
    ' コピーモードを確実に解除する
    Application.CutCopyMode = False
    ' マスタファイルを保存せずに閉じる
    wbMaster.Close SaveChanges:=False
    
    ' 元々選択していたシートが存在するか判定する
    If Not objOriginalSh Is Nothing Then
        ' 復元時のエラーを一時的に無視する
        On Error Resume Next
        ' 元のシートをアクティブにする
        objOriginalSh.Activate
        ' 元のセル位置を選択状態にする
        If strOriginalCellAddr <> "" Then objOriginalSh.Range(strOriginalCellAddr).Select
        ' エラー処理を解除する
        On Error GoTo 0
    ' 選択復元処理の終了
    End If
    
    ' ユーザー定義の読み込み後処理を実行する
    basUtility.RunUserProcedure "VABridge_AfterLoad", TargetBook
    
    ' ブックを保存済み状態にする
    TargetBook.Saved = True
    
    ' 現在のモード状態を格納する変数を宣言する
    Dim blnIsNowEditMode As Boolean
    ' 自身の有効なロックがあるか確認し結果を代入する
    blnIsNowEditMode = hasMyLock(TargetBook)
    ' 編集モードか閲覧モードかに応じてシートのデザインを変更する
    changeDesignMode TargetBook, blnIsNowEditMode
    
    ' メッセージボックス表示前にステータスバーを消去する
    Application.StatusBar = False

    ' サイレントモードではないか判定する
    If Not isSilent Then
        ' 編集モード状態であるか判定する
        If blnIsNowEditMode Then
            ' 編集モード完了のメッセージを表示する
            MsgBox "編集モードで読み込みが完了しました。" & vbCrLf & "マスタファイルをロックしています。", vbInformation, cstSettings.MSG_TITLE
        ' 閲覧モード状態の場合の処理
        Else
            ' 閲覧モード完了のメッセージを表示する
            MsgBox "閲覧モードで読み込みが完了しました。", vbInformation, cstSettings.MSG_TITLE
        ' モード判定の終了
        End If
    ' サイレント判定の終了
    End If
    
    ' 戻り値をTrue(成功)に設定する
    loadFromMaster = True
    ' 終了処理ラベルへジャンプする
    GoTo Finalize

' エラー発生時の処理ラベル
ErrorHandler:
    ' マスタファイルが開いているか判定する
    If Not wbMaster Is Nothing Then
        ' コピー状態を解除する
        Application.CutCopyMode = False
        ' マスタファイルを強制的に閉じる
        wbMaster.Close False
    ' 開閉判定の終了
    End If
    
    ' 自身の有効なロックが存在するか判定する
    If hasMyLock(TargetBook) Then
        ' 安全のためにロックを解放する
        releaseLock strMasterPath
        ' デザインも閲覧モードに戻す
        changeDesignMode TargetBook, False
    ' 有効ロック判定の終了
    End If
    
    ' エラー時もステータスバーを消去する
    Application.StatusBar = False

    ' エラー内容をユーザーに通知する
    MsgBox "読み込み中にエラーが発生しました: " & Err.Description, vbCritical, cstSettings.MSG_TITLE
    ' 戻り値をFalse(失敗)に設定する
    loadFromMaster = False
    
' 共通の終了処理ラベル
Finalize:
    ' ステータスバーを念のため既定に戻す
    Application.StatusBar = False

    ' コピー状態を解除する
    Application.CutCopyMode = False
    ' 画面の描画更新を再開する
    Application.ScreenUpdating = True
    ' イベント発生設定を元に戻す
    Application.EnableEvents = True
    ' 確認ダイアログの表示設定を元に戻す
    Application.DisplayAlerts = True
' 関数の終了
End Function

' -----------------------------------------------------------------------------
' プロシージャ概要: マスタへの書き込み処理
' 引数: TargetBook = 処理対象のブック, isSilent = メッセージ表示抑制フラグ
' 戻り値: 成功ならTrue、失敗ならFalse
' -----------------------------------------------------------------------------
' マスタへの書き込み処理を定義するパブリック関数
Public Function writeToMaster(ByVal TargetBook As Workbook, ByVal isSilent As Boolean) As Boolean
    ' 自身の有効なロックがない場合(閲覧モードの場合)か判定する
    If Not hasMyLock(TargetBook) Then
        ' 書き込み不可のメッセージを表示する
        MsgBox "現在は閲覧モードのため、書き込みできません。", vbInformation, cstSettings.MSG_TITLE
        ' 関数を終了する
        Exit Function
    ' ロック判定の終了
    End If
    
    ' Mac環境であるか判定するコンパイラディレクティブ
    #If Mac Then
        ' 非対応メッセージを表示する
        MsgBox "Mac環境での書き込みはサポートされていません。" & vbCrLf & "Windows環境で実行してください。", vbExclamation, cstSettings.MSG_TITLE
        ' 関数を終了する
        Exit Function
    ' コンパイラディレクティブの終了
    #End If
    
    ' マスタパスを格納する変数を宣言する
    Dim strMasterPath As String
    ' マスタブックを操作する変数を宣言する
    Dim wbMaster As Workbook
    ' シート操作用の変数を宣言する
    Dim wsLoop As Worksheet
    ' 最終同期日時を格納する変数を宣言する
    Dim dtmLastSyncTime As Date
    ' 現在のマスタ更新日時を格納する変数を宣言する
    Dim dtmCurrentMasterTime As Date
    
    ' OS環境に合わせて解決済みのマスタパスを取得する
    strMasterPath = basUtility.GetLocalPath(basConfig.GetBookSetting(cstSettings.PROP_MASTER_PATH, TargetBook))
    
    ' 楽観的ロックチェックのための日時取得時にエラーを一時的に無視する
    On Error Resume Next
    ' 保存しておいた最終同期日時を取得する
    dtmLastSyncTime = CDate(basConfig.GetBookSetting(cstSettings.PROP_LAST_SYNC, TargetBook))
    ' 現在のマスタファイルの更新日時を取得する
    dtmCurrentMasterTime = FileDateTime(strMasterPath)
    ' エラーハンドリングを標準状態に戻す
    On Error GoTo 0
    
    ' マスタファイルが読み込み後に更新されているか判定する
    If dtmCurrentMasterTime > dtmLastSyncTime Then
        ' 競合ガイドメッセージを表示する
        MsgBox cstSettings.MSG_CONFLICT_GUIDE, vbCritical, cstSettings.MSG_TITLE
        ' 終了処理ラベルへジャンプする
        GoTo Finalize
    ' 競合判定の終了
    End If
    
    ' ユーザー定義の書き込み前処理を実行する
    basUtility.RunUserProcedure "VABridge_BeforeSave", TargetBook
    
    ' 実行処理開始前にステータスバーを表示する
    Application.StatusBar = "マスタファイルへデータを書き込み中です..."
    
    ' 画面描画を停止する
    Application.ScreenUpdating = False
    ' イベント発生を停止する
    Application.EnableEvents = False
    ' 確認メッセージを停止する
    Application.DisplayAlerts = False
    
    ' エラー発生時の処理先を設定する
    On Error GoTo ErrorHandler
    
    ' リトライ回数をカウントする変数を宣言する
    Dim lngRetryCount As Long
    ' オープン成功フラグを宣言しFalseで初期化する
    Dim blnIsOpenSuccess As Boolean: blnIsOpenSuccess = False
    
    ' 最大3回のリトライ処理ループを開始する
    For lngRetryCount = 1 To 3
        ' マスタファイルを書き込み可モードで開く
        Set wbMaster = Workbooks.Open(strMasterPath, ReadOnly:=False)
        
        ' 開いたマスタファイルが読み取り専用状態か判定する
        If wbMaster.ReadOnly Then
            ' 読み取り専用なら直ちに閉じる
            wbMaster.Close SaveChanges:=False
            ' ブックオブジェクトを解放する
            Set wbMaster = Nothing
            
            ' OSの排他制御遅延対策として1秒待機する
            Application.Wait Now + TimeValue("00:00:01")
        ' 書き込みモードで開けた場合の処理
        Else
            ' 成功フラグをTrueに設定する
            blnIsOpenSuccess = True
            ' リトライループを抜ける
            Exit For
        ' 読み取り専用判定の終了
        End If
    ' 次のリトライへ進む
    Next lngRetryCount
    
    ' 3回試行しても成功しなかったか判定する
    If Not blnIsOpenSuccess Then
        ' エラーメッセージを表示する
        MsgBox "マスタファイルを書き込みモードで開けませんでした。" & vbCrLf & "他のユーザーが開いているか、ネットワーク遅延の可能性があります。", vbCritical, cstSettings.MSG_TITLE
        ' 終了処理ラベルへジャンプする
        GoTo Finalize
    ' オープン成功判定の終了
    End If
    
    ' 対象ブックの全シートに対してループ処理を開始する
    For Each wsLoop In TargetBook.Worksheets
        ' 管理対象シートか判定する
        If basUtility.IsManagedSheet(wsLoop) Then
            ' マスタシート名を格納する変数を宣言する
            Dim strMShName As String
            ' ローカルテーブルオブジェクトを宣言する
            Dim objTblLocal As ListObject
            ' マスタテーブルオブジェクトを宣言する
            Dim objTblMaster As ListObject
            
            ' 設定からマスタシート名を取得する
            strMShName = basConfig.GetSheetSetting(wsLoop, cstSettings.PROP_SUFFIX_SHEET)
            
            ' ローカルテーブルを取得する
            Set objTblLocal = wsLoop.ListObjects(1)
            ' マスタテーブル取得時にエラーを無視する
            On Error Resume Next
            ' マスタテーブルの取得を試みる
            Set objTblMaster = wbMaster.Worksheets(strMShName).ListObjects(1)
            ' エラーハンドラを復帰させる
            On Error GoTo ErrorHandler
            
            ' マスタテーブルが存在するか判定する
            If Not objTblMaster Is Nothing Then
                ' ヘッダー整合性を再確認する
                If checkHeaderConsistency(objTblLocal, objTblMaster) Then
                    ' マスタ側の既存データを全削除する
                    If Not objTblMaster.DataBodyRange Is Nothing Then objTblMaster.DataBodyRange.Delete
                    
                    ' ローカルにデータが存在するか判定する
                    If Not objTblLocal.DataBodyRange Is Nothing Then
                        ' ローカルデータをコピーする
                        objTblLocal.DataBodyRange.Copy
                        
                        ' クリップボード安定化のため待機する
                        DoEvents
                        
                        ' マスタへ値のみ貼り付けを行う
                        objTblMaster.HeaderRowRange.Offset(1, 0).PasteSpecial xlPasteValues
                        
                        ' クリップボードをクリアする
                        Application.CutCopyMode = False
                    ' ローカルデータ存在判定の終了
                    End If
                ' ヘッダー整合性判定の終了
                End If
            ' マスタテーブル存在判定の終了
            End If
        ' 管理対象シート判定の終了
        End If
    ' 次のシートへ進む
    Next wsLoop
    
    ' 念のため再度クリップボードをクリアする
    Application.CutCopyMode = False
    
    ' 保存処理のエラーを独自に捕捉するため一時的にエラーを無視する
    On Error Resume Next
    ' マスタファイルを上書き保存する
    wbMaster.Save
    
    ' 保存時にエラーが発生したか判定する
    If Err.Number <> 0 Then
        ' エラー説明文を格納する変数を宣言する
        Dim strErrDesc As String
        ' エラー内容を取得する
        strErrDesc = Err.Description
        ' エラーハンドラを標準に復帰させる
        On Error GoTo ErrorHandler
        
        ' エラーメッセージを表示する
        MsgBox "マスタファイルの保存に失敗しました。" & vbCrLf & strErrDesc, vbCritical, cstSettings.MSG_TITLE
        ' 保存せず強制的に閉じる
        wbMaster.Close SaveChanges:=False
        ' 戻り値をFalse(失敗)に設定する
        writeToMaster = False
        ' 終了処理ラベルへジャンプする
        GoTo Finalize
    ' 保存エラー判定の終了
    End If
    
    ' エラーハンドラを標準に復帰させる
    On Error GoTo ErrorHandler
    
    ' 保存に成功したので閉じる
    wbMaster.Close SaveChanges:=False
    
    ' 書き込み完了に伴いロックを解放する
    releaseLock strMasterPath
    
    ' デザインを閲覧モードに戻す
    changeDesignMode TargetBook, False
    ' ローカルブックを保存済み状態にする
    TargetBook.Saved = True
    
    ' メッセージボックス表示前にステータスバーを消去する
    Application.StatusBar = False

    ' 完了メッセージを表示する
    If Not isSilent Then MsgBox "書き込みが完了しました。閲覧モードに戻ります。", vbInformation, cstSettings.MSG_TITLE
    ' 戻り値をTrue(成功)に設定する
    writeToMaster = True
    ' 終了処理ラベルへジャンプする
    GoTo Finalize

' エラー発生時の処理ラベル
ErrorHandler:
    ' マスタが開いているか判定する
    If Not wbMaster Is Nothing Then
        ' クリップボードをクリアする
        Application.CutCopyMode = False
        ' マスタを強制的に閉じる
        wbMaster.Close False
    ' マスタオープン判定の終了
    End If
    
    ' エラー時もステータスバーを消去する
    Application.StatusBar = False
    
    ' エラー内容を通知する
    MsgBox "書き込みに失敗しました: " & Err.Description, vbCritical, cstSettings.MSG_TITLE
    ' 戻り値をFalse(失敗)に設定する
    writeToMaster = False
    
' 共通の終了処理ラベル
Finalize:
    ' ステータスバーを念のため既定に戻す
    Application.StatusBar = False

    ' クリップボードをクリアする
    Application.CutCopyMode = False
    ' 画面描画を再開する
    Application.ScreenUpdating = True
    ' イベント発生を再開する
    Application.EnableEvents = True
    ' 確認ダイアログを再開する
    Application.DisplayAlerts = True
' 関数の終了
End Function

' -----------------------------------------------------------------------------
' プロシージャ概要: ロック状態の判定処理
' 引数: strMasterPath = マスタファイルのパス, strOutUser = ロック主の名前(出力), dtmOutTime = ロック日時(出力)
' 戻り値: ロックされているならTrue, されていないならFalse
' -----------------------------------------------------------------------------
' マスタファイルのロック状態を判定するパブリック関数
Public Function isLocked(ByVal strMasterPath As String, ByRef strOutUser As String, ByRef dtmOutTime As Date) As Boolean
    ' ロックファイルの検索パターンを格納する変数を宣言する
    Dim strLockPattern As String
    ' 検索でヒットしたファイル名を格納する変数を宣言する
    Dim strFileName As String
    ' マスタファイルのフォルダパスを格納する変数を宣言する
    Dim strFolderPath As String
    ' ロックファイルから抽出したユーザー名を格納する変数を宣言する
    Dim strFileUser As String
    ' ロックファイルの更新日時を格納する変数を宣言する
    Dim dtmFileTime As Date
    ' 自身のユーザー名を格納する変数を宣言する
    Dim strMyUser As String
    
    ' 自身のユーザー名を取得しファイル名形式にサニタイズする
    strMyUser = basUtility.SanitizeFileName(basUtility.GetNetworkUserName())
    
    ' マスタパスからフォルダ部分を抽出する
    strFolderPath = Left(strMasterPath, InStrRev(strMasterPath, Application.PathSeparator))
    
    ' マスタファイル名を格納する変数を宣言する
    Dim strMasterFileName As String
    ' マスタパスからファイル名部分を抽出する
    strMasterFileName = Mid(strMasterPath, InStrRev(strMasterPath, Application.PathSeparator) + 1)
    
    ' 検索用パターン文字列を作成する
    strLockPattern = strMasterFileName & ".*" & EXT_LOCK_SUFFIX
    ' フォルダ内でパターンに一致する最初のファイルを検索する
    strFileName = Dir(basUtility.BuildPath(strFolderPath, strLockPattern))
    
    ' 一致するファイルが見つかる間ループを継続する
    Do While strFileName <> ""
        ' ロックファイルのフルパスを格納する変数を宣言する
        Dim strFullLockPath As String
        ' フォルダパスとファイル名を結合する
        strFullLockPath = basUtility.BuildPath(strFolderPath, strFileName)
        ' 該当ロックファイルの更新日時を取得する
        dtmFileTime = FileDateTime(strFullLockPath)
        
        ' ユーザー名抽出用のプレフィックス文字列を作成する
        Dim strPrefix As String: strPrefix = strMasterFileName & "."
        
        ' ファイル名が所定の形式に合致するかチェックする
        If (Left(strFileName, Len(strPrefix)) = strPrefix) And (Right(strFileName, Len(EXT_LOCK_SUFFIX)) = EXT_LOCK_SUFFIX) Then
            ' ユーザー名部分の長さを格納する変数を宣言する
            Dim lngUserLen As Long
            ' プレフィックスと拡張子を除いた文字数を算出する
            lngUserLen = Len(strFileName) - Len(strPrefix) - Len(EXT_LOCK_SUFFIX)
            
            ' ユーザー名が1文字以上存在するか判定する
            If lngUserLen > 0 Then
                ' ファイル名からユーザー名部分を抽出する
                strFileUser = Mid(strFileName, Len(strPrefix) + 1, lngUserLen)
                
                ' ロックファイルがマスタファイル以上新しいか判定する
                If FileDateTime(strMasterPath) <= dtmFileTime Then
                    ' ロックが自身以外のユーザーによるものか判定する
                    If strFileUser <> strMyUser Then
                        ' 出力用引数にユーザー名をセットする
                        strOutUser = strFileUser
                        ' 出力用引数にロック日時をセットする
                        dtmOutTime = dtmFileTime
                        ' ロックされていると判定して戻り値をTrueにする
                        isLocked = True
                        ' 条件を満たしたため関数を終了する
                        Exit Function
                    ' 他ユーザー判定の終了
                    End If
                ' 日時新旧判定の終了
                End If
            ' ユーザー名長判定の終了
            End If
        ' ファイル形式判定の終了
        End If
        ' 次に一致するファイルを検索する
        strFileName = Dir()
    ' ループの先頭に戻る
    Loop
    ' 有効な他人のロックが見つからなかったため戻り値をFalseにする
    isLocked = False
' 関数の終了
End Function

' -----------------------------------------------------------------------------
' プロシージャ概要: 無効な古いロックファイルをクリーンアップする処理
' 引数: strMasterPath = マスタファイルのパス
' 戻り値: なし
' -----------------------------------------------------------------------------
' 古いロックファイルを削除するプライベートプロシージャ
Private Sub cleanUpStaleLocks(ByVal strMasterPath As String)
    ' ロックファイルの検索パターンを格納する変数を宣言する
    Dim strLockPattern As String
    ' 検索でヒットしたファイル名を格納する変数を宣言する
    Dim strFileName As String
    ' マスタファイルのフォルダパスを格納する変数を宣言する
    Dim strFolderPath As String
    ' ロックファイルのフルパスを格納する変数を宣言する
    Dim strFullLockPath As String
    ' ロックファイルの更新日時を格納する変数を宣言する
    Dim dtmFileTime As Date
    ' マスタファイルの更新日時を格納する変数を宣言する
    Dim dtmMasterTime As Date

    ' 日時取得時のエラーを一時的に無視する
    On Error Resume Next
    ' マスタファイルの更新日時を取得する
    dtmMasterTime = FileDateTime(strMasterPath)
    ' エラーハンドラを標準に復帰させる
    On Error GoTo 0

    ' マスタパスからフォルダ部分を抽出する
    strFolderPath = Left(strMasterPath, InStrRev(strMasterPath, Application.PathSeparator))
    ' マスタファイル名を格納する変数を宣言する
    Dim strMasterFileName As String
    ' マスタパスからファイル名部分を抽出する
    strMasterFileName = Mid(strMasterPath, InStrRev(strMasterPath, Application.PathSeparator) + 1)
    
    ' 検索用パターン文字列を作成する
    strLockPattern = strMasterFileName & ".*" & EXT_LOCK_SUFFIX
    ' パターンに一致する最初のファイルを検索する
    strFileName = Dir(basUtility.BuildPath(strFolderPath, strLockPattern))
    
    ' ファイルが見つかる間ループを継続する
    Do While strFileName <> ""
        ' ファイルのフルパスを作成する
        strFullLockPath = basUtility.BuildPath(strFolderPath, strFileName)
        ' ファイル情報取得時のエラーを一時的に無視する
        On Error Resume Next
        ' ファイルの更新日時を取得する
        dtmFileTime = FileDateTime(strFullLockPath)
        
        ' 削除すべきか判定するフラグ変数をFalseで初期化する
        Dim blnShouldDelete As Boolean: blnShouldDelete = False
        
        ' マスタファイルより古いロックファイルか判定する
        If dtmFileTime < dtmMasterTime Then
            ' 古い場合は削除対象フラグをTrueにする
            blnShouldDelete = True
        ' 削除判定の終了
        End If
        
        ' 削除対象フラグがTrueであればファイルを削除する
        If blnShouldDelete Then Kill strFullLockPath
        ' エラーハンドラを標準に復帰させる
        On Error GoTo 0
        ' 次の対象ファイルを検索する
        strFileName = Dir()
    ' ループの先頭に戻る
    Loop
' プロシージャの終了
End Sub

' -----------------------------------------------------------------------------
' プロシージャ概要: ロック取得処理
' 引数: strMasterPath = マスタファイルのパス
' 戻り値: 取得成功ならTrue, 失敗ならFalse
' -----------------------------------------------------------------------------
' 新たにロックを取得するパブリック関数
Public Function acquireLock(ByVal strMasterPath As String) As Boolean
    ' 既存ロックユーザー名受け取り用の変数を宣言する
    Dim strOutUser As String
    ' 既存ロック日時受け取り用の変数を宣言する
    Dim dtmOutTime As Date
    ' 作成する自身のロックファイルパスを格納する変数を宣言する
    Dim strMyLockPath As String
    ' 自身のユーザー名を格納する変数を宣言する
    Dim strMyUser As String
    
    ' まず古い無効なロックファイルを全て削除してクリーンアップする
    cleanUpStaleLocks strMasterPath
    
    ' 現在有効な他人のロックが存在するか判定する
    If isLocked(strMasterPath, strOutUser, dtmOutTime) Then
        ' エラーメッセージを表示する
        MsgBox "既に他のユーザー (" & strOutUser & ") が編集を開始しています。", vbCritical, cstSettings.MSG_TITLE
        ' ロック取得失敗として戻り値をFalseに設定する
        acquireLock = False
        ' 関数を終了する
        Exit Function
    ' 他人ロック存在判定の終了
    End If
    
    ' 自身のユーザー名を取得しファイル名形式に変換する
    strMyUser = basUtility.SanitizeFileName(basUtility.GetNetworkUserName())
    ' 自身のロックファイルのフルパスを作成する
    strMyLockPath = strMasterPath & "." & strMyUser & EXT_LOCK_SUFFIX
    
    ' ロックファイルに書き込む内容を格納する変数を宣言する
    Dim strLockContent As String
    ' ロックファイルの内容文字列を構築する
    strLockContent = "LOCKED_BY=" & Environ("USERNAME") & vbCrLf & "MACHINE=" & Environ("COMPUTERNAME") & vbCrLf & "TIMESTAMP=" & Now
    
    ' 指定したパスと内容でロックファイルを実際に作成する
    basUtility.CreateLockFile strMyLockPath, strLockContent
    
    ' 自身が作成した直後に他人の有効なロックが存在しないか再度確認する
    If isLocked(strMasterPath, strOutUser, dtmOutTime) Then
        ' 自身のロック削除時のエラーを無視する
        On Error Resume Next
        ' 競合負けと判断し自身の作成したロックファイルを削除する
        Kill strMyLockPath
        ' エラーハンドラを標準に戻す
        On Error GoTo 0
        ' 競合エラーメッセージを表示する
        MsgBox "競合が発生しました。他のユーザー (" & strOutUser & ") が僅かに早くロックを取得しました。", vbExclamation, cstSettings.MSG_TITLE
        ' ロック取得失敗として戻り値をFalseに設定する
        acquireLock = False
    ' 競合が発生しなかった場合の処理
    Else
        ' ロック取得成功として戻り値をTrueに設定する
        acquireLock = True
    ' 競合判定の終了
    End If
' 関数の終了
End Function

' -----------------------------------------------------------------------------
' プロシージャ概要: ロック解除処理
' 引数: strMasterPath = マスタファイルのパス
' 戻り値: なし
' -----------------------------------------------------------------------------
' 自身のロックを解除するパブリックプロシージャ
Public Sub releaseLock(ByVal strMasterPath As String)
    ' 自身のユーザー名を格納する変数を宣言する
    Dim strMyUser As String
    ' ロックファイルの検索パターンを格納する変数を宣言する
    Dim strLockPattern As String
    ' 検索されたファイル名を格納する変数を宣言する
    Dim strFileName As String
    ' マスタのフォルダパスを格納する変数を宣言する
    Dim strFolderPath As String
    
    ' 自身のユーザー名を取得しファイル名形式に変換する
    strMyUser = basUtility.SanitizeFileName(basUtility.GetNetworkUserName())
    ' マスタパスからフォルダ部分を抽出する
    strFolderPath = Left(strMasterPath, InStrRev(strMasterPath, Application.PathSeparator))
    
    ' マスタファイル名を格納する変数を宣言する
    Dim strMasterFileName As String
    ' マスタパスからファイル名部分を抽出する
    strMasterFileName = Mid(strMasterPath, InStrRev(strMasterPath, Application.PathSeparator) + 1)
    
    ' 自身のロックファイルのみに一致するパターンを作成する
    strLockPattern = strMasterFileName & "." & strMyUser & EXT_LOCK_SUFFIX
    
    ' 指定フォルダ内で自身のロックファイルを検索する
    strFileName = Dir(basUtility.BuildPath(strFolderPath, strLockPattern))
    ' 自身のロックファイルが見つかる間ループを継続する
    Do While strFileName <> ""
        ' 削除時のエラーを一時的に無視する
        On Error Resume Next
        ' 見つかったロックファイルを削除する
        Kill basUtility.BuildPath(strFolderPath, strFileName)
        ' エラーハンドラを標準に戻す
        On Error GoTo 0
        ' 同名の他のファイルがないか再度検索する
        strFileName = Dir()
    ' ループの先頭に戻る
    Loop
' プロシージャの終了
End Sub

' -----------------------------------------------------------------------------
' プロシージャ概要: デザインモード変更処理
' 引数: Wb = 対象ブック, blnToEdit = 編集モードにするかどうかのフラグ
' 戻り値: なし
' -----------------------------------------------------------------------------
' シートのデザインや保護状態を変更するパブリックプロシージャ
Public Sub changeDesignMode(ByVal Wb As Workbook, ByVal blnToEdit As Boolean)
    ' ループ処理用のシート変数を宣言する
    Dim wsLoop As Worksheet
    ' テーブル操作用のオブジェクト変数を宣言する
    Dim objTbl As ListObject
    
    ' 対象ブック内の全シートをループ処理する
    For Each wsLoop In Wb.Worksheets
        ' 現在のシートが管理対象か判定する
        If basUtility.IsManagedSheet(wsLoop) Then
            ' シートの1つ目のテーブルオブジェクトを取得する
            Set objTbl = wsLoop.ListObjects(1)
            
            ' 編集モードへの変更要求か判定する
            If blnToEdit Then
                ' データ編集を可能にするためシートの保護を解除する
                wsLoop.Unprotect
                ' 現在のテーブルスタイルが編集用スタイルではないか判定する
                If objTbl.TableStyle <> cstSettings.STYLE_EDIT_MODE Then
                    ' 復元用に現在のスタイル名をシート設定に保存する
                    basConfig.SaveSheetSetting wsLoop, cstSettings.PROP_SUFFIX_STYLE, objTbl.TableStyle
                ' スタイル判定の終了
                End If
                ' 存在しないスタイル名エラー等を防ぐためエラーを無視する
                On Error Resume Next
                ' テーブルのスタイルを目立つ編集用スタイルに変更する
                objTbl.TableStyle = cstSettings.STYLE_EDIT_MODE
                ' エラーハンドラを標準に戻す
                On Error GoTo 0
            
            ' 閲覧モードへの変更要求の場合の処理
            Else
                ' 保存しておいた元のスタイル名を格納する変数を宣言する
                Dim strOldStyle As String
                ' シート設定から元のスタイル名を取得する
                strOldStyle = basConfig.GetSheetSetting(wsLoop, cstSettings.PROP_SUFFIX_STYLE)
                ' 元のスタイル名が存在するか判定する
                If strOldStyle <> "" Then
                    ' スタイル適用エラーを防ぐためエラーを無視する
                    On Error Resume Next
                    ' テーブルスタイルを元の状態に戻す
                    objTbl.TableStyle = strOldStyle
                    ' エラーハンドラを標準に戻す
                    On Error GoTo 0
                ' スタイル存在判定の終了
                End If
                
                ' ユーザーの誤編集を防ぐためシートを保護しつつ書式やソート等は許可する
                wsLoop.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True
            ' モード判定の終了
            End If
        ' 管理対象判定の終了
        End If
    ' 次のシートへループを進める
    Next wsLoop
' プロシージャの終了
End Sub

' -----------------------------------------------------------------------------
' プロシージャ概要: ヘッダー整合性検証処理
' 引数: objLocalTbl = ローカルテーブル, objMasterTbl = マスタテーブル
' 戻り値: 完全一致ならTrue, 異なればFalse
' -----------------------------------------------------------------------------
' ローカルとマスタのヘッダー構造が一致しているか確認するプライベート関数
Private Function checkHeaderConsistency(objLocalTbl As ListObject, objMasterTbl As ListObject) As Boolean
    ' 両テーブルのヘッダー列数が異なるか判定する
    If objLocalTbl.HeaderRowRange.Columns.Count <> objMasterTbl.HeaderRowRange.Columns.Count Then
        ' 異なる場合は不一致として戻り値をFalseにする
        checkHeaderConsistency = False
        ' 関数を終了する
        Exit Function
    ' 列数判定の終了
    End If
    
    ' ループ処理用の列番号変数を宣言する
    Dim i As Long
    ' 1列目から最終列までループ処理を行う
    For i = 1 To objLocalTbl.HeaderRowRange.Columns.Count
        ' 対象列のヘッダー文字列が異なるか判定する
        If objLocalTbl.HeaderRowRange.Cells(1, i).value <> objMasterTbl.HeaderRowRange.Cells(1, i).value Then
            ' 1つでも異なれば不一致として戻り値をFalseにする
            checkHeaderConsistency = False
            ' 関数を直ちに終了する
            Exit Function
        ' ヘッダー文字列判定の終了
        End If
    ' 次の列へループを進める
    Next i
    ' 全ての列が一致したため戻り値をTrueにする
    checkHeaderConsistency = True
' 関数の終了
End Function
