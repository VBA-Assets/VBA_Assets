Option Explicit

' モジュール概要: アドインのエントリポイントおよびリボンUIの管理
' Last Updated: 2026/02/07 15:00 (Ver402: Added Maintenance Mode state management)

' アプリケーションイベントクラスのインスタンスを保持する変数
Public gAppEvents As clsAppEvents

' 【新規追加】遅延実行時に対象ブックを一時的に保持する変数
Public gTargetBook As Workbook

' 【新規追加】メンテナンスモード（Shift起動）中のブックを保持するコレクション
' キー: ObjPtr(Workbook)の文字列, 値: Workbookオブジェクト
Public gMaintenanceBooks As Collection

' -----------------------------------------------------------------------------
' 初期化・終了処理
' ブックが開かれたときに自動実行される
' -----------------------------------------------------------------------------
Sub Auto_Open()
    ' アプリケーションレベルのイベント監視を開始するためにクラスをインスタンス化
    Set gAppEvents = New clsAppEvents
    ' Excelアプリケーション本体をイベント監視対象としてセット
    Set gAppEvents.App = Application
    
    ' メンテナンスモード管理用コレクションの初期化
    Set gMaintenanceBooks = New Collection
End Sub

Sub Auto_Close()
    ' ブックを閉じる際にイベント監視用オブジェクトを破棄
    Set gAppEvents = Nothing
    ' グローバル変数の解放
    Set gTargetBook = Nothing
    Set gMaintenanceBooks = Nothing
End Sub

' -----------------------------------------------------------------------------
' 【新規追加】メンテナンスモード管理用ヘルパー関数群
' -----------------------------------------------------------------------------

' ブックをメンテナンスモードとして登録する
Public Sub RegisterMaintenanceBook(Wb As Workbook)
    If gMaintenanceBooks Is Nothing Then Set gMaintenanceBooks = New Collection
    
    On Error Resume Next
    ' キーにオブジェクトポインタを使用して一意に識別する
    gMaintenanceBooks.Add Wb, CStr(ObjPtr(Wb))
    On Error GoTo 0
End Sub

' ブックがメンテナンスモードかどうか判定する
Public Function IsMaintenanceBook(Wb As Workbook) As Boolean
    If gMaintenanceBooks Is Nothing Then
        IsMaintenanceBook = False
        Exit Function
    End If
    
    On Error Resume Next
    Dim tmp As Workbook
    Set tmp = gMaintenanceBooks(CStr(ObjPtr(Wb)))
    IsMaintenanceBook = Not (tmp Is Nothing)
    On Error GoTo 0
End Function

' ブックのメンテナンスモード登録を解除する
Public Sub UnregisterMaintenanceBook(Wb As Workbook)
    If gMaintenanceBooks Is Nothing Then Exit Sub
    
    On Error Resume Next
    gMaintenanceBooks.Remove CStr(ObjPtr(Wb))
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' 【新規追加】遅延起動処理（クラッシュ対策）
' Workbook_OpenからApplication.OnTimeで1秒後に呼び出される
' -----------------------------------------------------------------------------
Public Sub DelayedStartup()
    ' 対象ブック変数が空でないか確認
    If Not gTargetBook Is Nothing Then
        ' エラー対策: ブックが既に閉じられている場合は無視する
        On Error Resume Next
        
        ' 読み込み処理を実行する
        ' 第1引数: True = 読み込み処理
        ' 第2引数: gTargetBook = 対象ブック
        ' 第3引数: False = サイレントモードOFF（これによりモード選択フォームが表示される）
        ProcessAllSheets True, gTargetBook, False
        
        ' 処理が終わったら変数をクリアする
        Set gTargetBook = Nothing
        On Error GoTo 0
    End If
End Sub

' -----------------------------------------------------------------------------
' リボンコールバック
' リボンUIからの操作を受け取るプロシージャ群
' -----------------------------------------------------------------------------

' RibbonXMLのロード時に呼ばれる初期化処理用
Sub OnRibbonLoad(ribbon As IRibbonUI)
    ' 必要に応じてリボンUIの状態を更新する処理を（現在は特になし）
End Sub

' リボンの「データ読み込み」ボタンが押された時の処理
Sub OnLoadPressed(control As IRibbonControl)
    ' 全シートを対象に読み込み処理を実行（True = 読み込み）
    ' ※ボタン操作時はサイレントモードOFF（isSilent=False）
    ProcessAllSheets True
End Sub

' リボンの「データ書き込み」ボタンが押された時の処理
Sub OnWritePressed(control As IRibbonControl)
    ' 全シートを対象に書き込み処理を実行（False = 書き込み）
    ProcessAllSheets False
End Sub

' リボンの「マスタ設定」ボタンが押された時の処理
Sub OnSettingsPressed(control As IRibbonControl)
    ' 設定フォームを表示する
    frmSettings.Show
End Sub

' -----------------------------------------------------------------------------
' 共通プロシージャ: 処理の振り分け
' 読み込みまたは書き込み処理のエントリーポイント
' isSilent: 自動実行時などメッセージを抑止する場合はTrueを指定する（読み込み時のみ有効）
' -----------------------------------------------------------------------------
Public Sub ProcessAllSheets(ByVal isLoad As Boolean, Optional ByVal Wb As Workbook = Nothing, Optional ByVal isSilent As Boolean = False)
    ' 対象ブックが指定されていない場合はアクティブブックを使用
    If Wb Is Nothing Then Set Wb = ActiveWorkbook
    
    ' ブック内に管理対象シートが1つでもあるかを確認する
    Dim sh As Worksheet, hasManaged As Boolean
    hasManaged = False
    For Each sh In Wb.Worksheets
        ' 管理対象シートかどうかを判定
        If basUtility.IsManagedSheet(sh) Then
            hasManaged = True
            Exit For
        End If
    Next sh
    
    ' 管理対象シートが一つもない場合の処理
    If Not hasManaged Then
        ' 読み込み操作でない場合（ユーザー操作時）はメッセージを表示
        ' ※サイレントモード時はメッセージを出さない
        If Not isLoad And Not isSilent Then MsgBox "管理対象シートがありません。", vbInformation, cstSettings.MSG_TITLE
        Exit Sub
    End If

    ' ブック単位で一括同期処理を実行する
    If isLoad Then
        ' マスタからの読み込みを実行
        ' 引数で受け取ったisSilentモードを適用する
        basSync.loadFromMaster Wb, isSilent
    Else
        ' マスタへの書き込みを実行
        basSync.writeToMaster Wb, False
    End If
End Sub