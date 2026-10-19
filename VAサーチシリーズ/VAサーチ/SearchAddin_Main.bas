Option Explicit

' --- グローバル変数 ---
' 検索ロジッククラスのインスタンスを保持
Public g_SearchLogic As New SearchLogic

' --- シート保護の設定 ---
' 一時的に解除する前の保護の設定を控え、保護し直すときに同じ設定に戻すために使います
Private Type typProtection
    blnDrawingObjects As Boolean
    blnScenarios As Boolean
    blnAllowFormattingCells As Boolean
    blnAllowFormattingColumns As Boolean
    blnAllowFormattingRows As Boolean
    blnAllowInsertingColumns As Boolean
    blnAllowInsertingRows As Boolean
    blnAllowInsertingHyperlinks As Boolean
    blnAllowDeletingColumns As Boolean
    blnAllowDeletingRows As Boolean
    blnAllowSorting As Boolean
    blnAllowFiltering As Boolean
    blnAllowUsingPivotTables As Boolean
End Type

'================================================================================
' プロシージャ: ShowSearchForm
' 機能:      検索フォームを表示します。
' 概要:      アクティブセルがデータ範囲内にあるか確認し、
'            検索フォームをモードレスで表示します。
'            ※チェック時にシートが保護されているとエラーになる場合があるため一時解除します。
'            ※パスワードで保護されたシートでは、保護の解除を案内してフォームを表示しません。
'================================================================================
Sub ShowSearchForm()
    Dim ws As Worksheet
    Dim isProtected As Boolean
    Dim checkResult As Boolean
    Dim prt As typProtection

    ' データ範囲の簡易チェック
    ' (厳密なチェックは実行時に行うが、明らかに空白の場合は警告する)
    If TypeName(Selection) <> "Range" Then Exit Sub

    Set ws = ActiveCell.Parent
    isProtected = ws.ProtectContents

    ' 保護されている場合は一時的に解除（読み取りチェックのため）
    ' パスワードで保護されている場合は、案内を表示して終了
    If isProtected Then
        If Not UnprotectTemporarily(ws, prt) Then Exit Sub
    End If

    ' チェック実行
    ' エラーが発生しても保護再設定まで進めるようにエラーハンドリングを追加
    checkResult = False
    On Error Resume Next
    If ActiveCell.Value = "" And ActiveCell.ListObject Is Nothing And ActiveCell.CurrentRegion.Cells.Count = 1 Then
        checkResult = True
    End If
    On Error GoTo 0

    ' 保護を再設定（元々保護されていた場合のみ）
    If isProtected Then Call ProtectAgain(ws, prt)

    ' チェック結果による分岐
    If checkResult Then
        MsgBox "データリスト内のセルを選択してから実行してください。", vbInformation, "VAサーチ"
        Exit Sub
    End If
    
    ' フォームを表示（モードレス）
    SearchForm.Show vbModeless
End Sub

'================================================================================
' プロシージャ: ShowSearchFormFromRibbon
' 機能:      リボンから呼び出されるコールバック関数
' 概要:      リボンのボタンクリックイベント（onAction）から呼び出されます。
'            引数 control は IRibbonControl 型ですが、参照設定の手間を省くため
'            Object 型、または遅延バインディングで動作するようにしています。
'================================================================================
Sub ShowSearchFormFromRibbon(control As Object)
    ' 既存の表示プロシージャを呼び出し
    Call ShowSearchForm
End Sub

'================================================================================
' プロシージャ: ExecuteSearch
' 機能:      検索を実行します（フォームから呼ばれます）。
' 概要:      アクティブなデータ範囲（テーブルまたは通常範囲）を特定し、
'            指定された条件で検索し、ヒットしない行を非表示にします。
'            ※シート保護時は一時解除して実行します（パスワードで保護されている場合は実行しません）。
'================================================================================
Sub ExecuteSearch(strText As String, strMode As String)
    Dim rngTargetBody As Range
    Dim ws As Worksheet
    Dim isProtected As Boolean
    Dim prt As typProtection

    Application.ScreenUpdating = False

    ' ワークシートの取得と保護状態の確認
    If TypeName(Selection) = "Range" Then
        Set ws = ActiveCell.Parent
        isProtected = ws.ProtectContents
    Else
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 保護されている場合は一時的に解除
    ' パスワードで保護されている場合は、案内を表示して終了
    If isProtected Then
        If Not UnprotectTemporarily(ws, prt) Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If

    ' 1. データ範囲（ヘッダーを除く本体）を取得
    Set rngTargetBody = GetTargetDataBody(ActiveCell)
    
    If rngTargetBody Is Nothing Then
        MsgBox "検索対象のデータが見つかりません。" & vbCrLf & "リスト内のセルを選択してください。", vbExclamation
        ' 保護を戻して終了
        If isProtected Then Call ProtectAgain(ws, prt)
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 2. 既存のフィルタと非表示状態をリセット（すべて表示）
    Call ResetFilterAndHidden(rngTargetBody)
    
    ' 3. 検索ロジックの適用（非表示処理も含む）
    Dim visibleCount As Long
    visibleCount = g_SearchLogic.ApplySearchAndHide(rngTargetBody, strText, strMode)
    
    ' 4. 先頭行へスクロール
    Application.GoTo rngTargetBody.Cells(1, 1), True
    
    ' 5. 結果確認
    If visibleCount = 0 Then
        MsgBox "該当するデータはありませんでした。", vbInformation, "検索結果"
    End If

    ' 保護を再設定（元々保護されていた場合のみ）
    If isProtected Then Call ProtectAgain(ws, prt)
    
    Application.ScreenUpdating = True
End Sub

'================================================================================
' プロシージャ: ClearSearch
' 機能:      検索を解除します。
' 概要:      行の非表示状態を解除し、すべて表示します。
'            ※シート保護時は一時解除して実行します（パスワードで保護されている場合は実行しません）。
'================================================================================
Sub ClearSearch()
    Dim rngTargetBody As Range
    Dim ws As Worksheet
    Dim isProtected As Boolean
    Dim prt As typProtection

    Application.ScreenUpdating = False

    ' ワークシートの取得と保護状態の確認
    If TypeName(Selection) = "Range" Then
        Set ws = ActiveCell.Parent
        isProtected = ws.ProtectContents
    Else
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 保護されている場合は一時的に解除
    ' パスワードで保護されている場合は、案内を表示して終了
    If isProtected Then
        If Not UnprotectTemporarily(ws, prt) Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If

    ' データ範囲を取得
    Set rngTargetBody = GetTargetDataBody(ActiveCell)
    
    If Not rngTargetBody Is Nothing Then
        ' リセット実行
        Call ResetFilterAndHidden(rngTargetBody)
        ' 選択位置へスクロール
        Application.GoTo ActiveCell, True
    End If

    ' 保護を再設定（元々保護されていた場合のみ）
    If isProtected Then Call ProtectAgain(ws, prt)

    Application.ScreenUpdating = True
End Sub

'================================================================================
' 関数:      UnprotectTemporarily
' 機能:      シートの保護を一時的に解除します。
' 概要:      保護の設定（並べ替え・フィルターの許可など）を prt に控えてから、
'            パスワードなしで解除します。パスワードで保護されている場合は解除できないため、
'            保護の解除を案内するメッセージを表示して False を返します。
'            （パスワードを省略して解除すると、パスワードの入力を求められ、
'              誤ったときに実行時エラーになるため、空のパスワードを指定して解除を試みます）
' 引数:      ws  - 対象のワークシート
'            prt - 解除前の保護の設定（戻り値として設定されます）
' 戻り値:    解除できた場合は True
'================================================================================
Private Function UnprotectTemporarily(ws As Worksheet, prt As typProtection) As Boolean
    ' 解除前の保護の設定を控える
    prt.blnDrawingObjects = ws.ProtectDrawingObjects
    prt.blnScenarios = ws.ProtectScenarios
    With ws.Protection
        prt.blnAllowFormattingCells = .AllowFormattingCells
        prt.blnAllowFormattingColumns = .AllowFormattingColumns
        prt.blnAllowFormattingRows = .AllowFormattingRows
        prt.blnAllowInsertingColumns = .AllowInsertingColumns
        prt.blnAllowInsertingRows = .AllowInsertingRows
        prt.blnAllowInsertingHyperlinks = .AllowInsertingHyperlinks
        prt.blnAllowDeletingColumns = .AllowDeletingColumns
        prt.blnAllowDeletingRows = .AllowDeletingRows
        prt.blnAllowSorting = .AllowSorting
        prt.blnAllowFiltering = .AllowFiltering
        prt.blnAllowUsingPivotTables = .AllowUsingPivotTables
    End With

    ' 空のパスワードで解除を試みる（パスワードで保護されている場合はエラーになる）
    On Error Resume Next
    ws.Unprotect Password:=""
    On Error GoTo 0

    ' 解除できなかった場合は、保護の解除を案内する
    If ws.ProtectContents Then
        Application.ScreenUpdating = True
        MsgBox "このシートはパスワードで保護されているため、検索できません。" & vbCrLf & _
               "シートの保護を解除してから実行してください。", vbExclamation, "VAサーチ"
        UnprotectTemporarily = False
    Else
        UnprotectTemporarily = True
    End If
End Function

'================================================================================
' プロシージャ: ProtectAgain
' 機能:      一時的に解除したシートを、元の保護の設定で保護し直します。
' 引数:      ws  - 対象のワークシート
'            prt - UnprotectTemporarily で控えた保護の設定
'================================================================================
Private Sub ProtectAgain(ws As Worksheet, prt As typProtection)
    ws.Protect DrawingObjects:=prt.blnDrawingObjects, Contents:=True, Scenarios:=prt.blnScenarios, _
        AllowFormattingCells:=prt.blnAllowFormattingCells, _
        AllowFormattingColumns:=prt.blnAllowFormattingColumns, _
        AllowFormattingRows:=prt.blnAllowFormattingRows, _
        AllowInsertingColumns:=prt.blnAllowInsertingColumns, _
        AllowInsertingRows:=prt.blnAllowInsertingRows, _
        AllowInsertingHyperlinks:=prt.blnAllowInsertingHyperlinks, _
        AllowDeletingColumns:=prt.blnAllowDeletingColumns, _
        AllowDeletingRows:=prt.blnAllowDeletingRows, _
        AllowSorting:=prt.blnAllowSorting, _
        AllowFiltering:=prt.blnAllowFiltering, _
        AllowUsingPivotTables:=prt.blnAllowUsingPivotTables
End Sub

'================================================================================
' 関数:      GetTargetDataBody
' 機能:      検索対象となる「データ部分（ヘッダー除く）」のRangeを取得します。
' 概要:      テーブルがあればそのDataBodyRange、なければCurrentRegionの2行目以降を返します。
'================================================================================
Private Function GetTargetDataBody(targetCell As Range) As Range
    Dim lo As ListObject
    Dim rngRegion As Range
    
    ' 1. テーブル(ListObject)の判定
    On Error Resume Next
    Set lo = targetCell.ListObject
    On Error GoTo 0
    
    If Not lo Is Nothing Then
        ' テーブルの場合、DataBodyRangeを返す（データがない場合はNothing）
        Set GetTargetDataBody = lo.DataBodyRange
        Exit Function
    End If
    
    ' 2. 通常範囲(CurrentRegion)の判定
    Set rngRegion = targetCell.CurrentRegion
    
    ' 範囲が単一セル、または1行しかない（ヘッダーのみ）場合は対象外
    If rngRegion.Rows.Count <= 1 Then
        Set GetTargetDataBody = Nothing
        Exit Function
    End If
    
    ' 1行目をヘッダーとみなし、2行目以降をデータ範囲として返す
    Set GetTargetDataBody = rngRegion.Offset(1, 0).Resize(rngRegion.Rows.Count - 1, rngRegion.Columns.Count)
End Function

'================================================================================
' プロシージャ: ResetFilterAndHidden
' 機能:      オートフィルタの解除と行の再表示を行います。
'================================================================================
Private Sub ResetFilterAndHidden(rngBody As Range)
    Dim ws As Worksheet
    Set ws = rngBody.Parent
    
    ' オートフィルタの絞り込み解除
    If ws.FilterMode Then
        ws.ShowAllData
    End If
    
    ' 行の非表示解除
    rngBody.EntireRow.Hidden = False
End Sub

'================================================================================
' テストデータ生成用マクロ
' 機能: 検索機能のテスト用に、異字体、全角半角、日付などが混在した表を作成します。
'================================================================================
'Sub GenerateTestData()
'    Dim ws As Worksheet
'    Dim rngTable As Range
'
'    ' 新しいシートを追加
'    Set ws = Worksheets.Add
'    ws.Name = "テストデータ_" & Format(Now, "mmdd_hhmm")
'
'    ' ヘッダー作成
'    ws.Range("A1:F1").Value = Array("ID", "氏名", "フリガナ", "部署", "日付", "備考・商品名")
'
'    ' テストデータ配列
'    ' ID, 氏名, フリガナ, 部署, 日付, 備考
'    Dim data As Variant
'    data = Array( _
'        Array(1, "高橋 太郎", "タカハシ タロウ", "営業部", "2025/1/1", "iPhone 15 (半角)"), _
'        Array(2, "髙橋 次郎", "タカハシ ジロウ", "営業部", "2025/1/1", "ｉＰｈｏｎｅ １５ (全角)"), _
'        Array(3, "渡辺 花子", "ワタナベ ハナコ", "総務部", "2025/3/3", "サーバー (長音あり)"), _
'        Array(4, "渡邉 花子", "ワタナベ ハナコ", "総務部", "R7.3.3", "サーバ (長音なし)"), _
'        Array(5, "中島 健一", "ナカジマ ケンイチ", "開発部", "2024/12/31", "Canon Camera"), _
'        Array(6, "中嶋 健一", "ナカシマ ケンイチ", "開発部", "2024-12-31", "キヤノン (大文字ヤ)"), _
'        Array(7, "鈴木 一郎", "スズキ イチロウ", "人事部", "2025/5/5", "キャノン (小文字ャ)"), _
'        Array(8, "齋藤 学", "サイトウ マナブ", "経理部", "2025/1/10", "MacBook"), _
'        Array(9, "斎藤 學", "サイトウ マナブ", "経理部", "2025/1/10", "ＭａｃＢｏｏｋ"), _
'        Array(10, "山崎 浩", "ヤマザキ ヒロシ", "広報部", "2025/4/1", "リンゴ (カタカナ)"), _
'        Array(11, "山﨑 浩", "ヤマサキ ヒロシ", "広報部", "2025/4/1", "りんご (ひらがな)"), _
'        Array(12, "佐藤 栄作", "サトウ エイサク", "製造部", "2025/6/20", "ダイヤモンド"), _
'        Array(13, "佐藤 榮作", "サトウ エイサク", "製造部", "2025/6/20", "ダイヤ (部分一致確認)"), _
'        Array(14, "大沢 樹生", "オオサワ ミキオ", "営業部", "2025/7/7", "ヴァイオリン"), _
'        Array(15, "大澤 樹生", "オオサワ ミキオ", "営業部", "2025/7/7", "バイオリン (濁点変化)") _
'    )
'
'    ' データ書き込み
'    Dim i As Long
'    For i = 0 To UBound(data)
'        ws.Cells(i + 2, 1).Resize(1, 6).Value = data(i)
'    Next i
'
'    ' 見た目の調整
'    ws.Range("A1").CurrentRegion.Columns.AutoFit
'
'    ' テーブルとして書式設定
'    ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes).Name = "Table_Test" & Format(Now, "mmss")
'
'    ' A1セルを選択しておく
'    ws.Range("A1").Select
'
'    MsgBox "テストデータを作成しました。" & vbCrLf & "「VAサーチ」を実行して動作を確認してください。", vbInformation
'End Sub
