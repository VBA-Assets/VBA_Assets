Option Explicit

'================================================================================
' モジュール名: sht11 (事件シート)
' 概要:         事件データベースのメインシート操作を制御するクラスモジュールです。
'               事件の追加、編集、並び替え、書式設定、入力規則の制御など、
'               主要なユーザーインターフェース機能を提供します。
' 更新日時:     2026/03/11 11:30
'================================================================================

'================================================================================
' イベント:     Worksheet_Activate
' 機能:         シートアクティブ時の初期化
' 概要:         シートが表示された際に、データの並び替え、書式設定を行い、
'               直前に編集していた事件番号の行へスクロールします。
'================================================================================
Private Sub Worksheet_Activate()
    Evnt.Disable
    
    'シートを並び替える
    Call SortTable
    
    'シートの書式を設定する
    Call FormatTable
    
    Evnt.Enable
    
    '直前に選択されていた事件番号（lngActCase）が記録されていなければ処理を終了する
    If lngActCase = 0 Then Exit Sub

    '変数を宣言
    Dim r As Long

    'テーブル内の事件番号列に対象の番号が存在するか確認し、行番号を取得する
    If WorksheetFunction.CountIf(Range("事件[事件番号]"), lngActCase) > 0 Then
        r = WorksheetFunction.Match(lngActCase, Range("事件[事件番号]"), 0)
    End If

    '該当する行が見つからなかった場合（r=0）は終了する
    If r = 0 Then Exit Sub

    '対象の行の「事件呼名」セルを選択する
    Range("事件[事件呼名]")(r).Select

    'ウィンドウをスクロールして、選択した行が表示されるようにする
    ActiveWindow.ScrollRow = r
End Sub

'================================================================================
' イベント:     Worksheet_Change
' 機能:         セル変更時の自動処理
' 概要:         「現況」変更時の色変更やフォルダ移動、「データフォルダ」入力時の
'               存在チェックなど、値の変更に応じたアクションを実行します。
'================================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    ' シートが保護されている場合は、値の変更を伴う処理を行わず終了する
    If Me.ProtectContents = True Then Exit Sub

    '変更されたセルの行番号と列番号を、テーブル内の相対位置に変換するための変数を宣言する
    Dim r As Long, s As Long
    
    'テーブル範囲外のエラー回避
    On Error Resume Next
    
    '行番号を計算する（シート上の行番号 - テーブル開始行 + 1）
    r = Target.Row - Range("事件").Row + 1
    '列番号を計算する（シート上の列番号 - テーブル開始列 + 1）
    s = Target.Column - Range("事件").Column + 1
    
    On Error GoTo 0

    '変更箇所がテーブルの範囲外（行または列）であれば処理を終了する
    If r < 1 Or Range("事件").Rows.Count < r Then Exit Sub
    If s < 1 Or Range("事件").Columns.Count < s Then Exit Sub

    '処理中の再帰的なイベント発生を防ぐため、イベントを停止する
    Evnt.Disable

    'ヘッダー名を取得して処理を分岐
    Dim headerName As String
    headerName = Range("事件").ListObject.HeaderRowRange(s).Value

    '変更された列が「現況」列である場合の処理
    If headerName = "現況" Then
        '現況の値に応じて処理を分岐する
        Select Case Range("事件[現況]")(r).Value
            Case "E:相談不可"
                '行のフォント色を「相談不可」用の色（アクセント2）に変更する
                Range("事件").Rows(r).Font.ThemeColor = xlThemeColorAccent2
                
            Case "D:相談のみ"
                '行のフォント色を「相談のみ」用の色（アクセント6）に変更する
                Range("事件").Rows(r).Font.ThemeColor = xlThemeColorAccent6
                
            Case "C:終了"
                '行のフォント色を「終了」用の色（アクセント3）に変更する
                Range("事件").Rows(r).Font.ThemeColor = xlThemeColorAccent3
                
                '終了年月日がまだ入力されていない場合
                If Range("事件[終了年月日]")(r).Value = "" Then
                    '今日の日付を自動入力するかユーザーに確認する
                    If MsgBox("終了年月日に今日の日付を入力します。" & vbCr & _
                        "よろしいですか？", vbYesNo, "事件データベース") = vbYes Then _
                        Range("事件[終了年月日]")(r).Value = Date
                End If
                
                '終了案件としてデータフォルダを移動する処理を呼び出す
                Call MoveFolder(r)
                
            Case "B:精算管理"
                '行のフォント色を自動（黒）に戻す
                Range("事件").Rows(r).Font.ColorIndex = xlAutomatic
                '精算管理案件としてデータフォルダを移動する処理を呼び出す
                Call MoveFolder(r)
                
            Case Else
                'それ以外の場合、フォント色を自動（黒）に戻す
                Range("事件").Rows(r).Font.ColorIndex = xlAutomatic
                'データフォルダが正しく存在するか確認する
                Call CheckFolder(r)
        End Select
    End If

    '変更された列が「データフォルダ」列である場合の処理
    If headerName = "データフォルダ" Then
        '入力されたデータフォルダが存在するか確認する
        Call CheckFolder(r)
    End If

    'イベントを再開する
    Evnt.Enable
End Sub

'================================================================================
' イベント:     Worksheet_SelectionChange
' 機能:         選択セル変更時の処理
' 概要:         選択された行の事件番号をグローバル変数に記録し、
'               シート間移動時の位置保持に使用します。
'================================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    '選択されたセルの行番号をテーブル内の行番号に変換する
    Dim r As Long
    
    'エラー回避（テーブル外選択時など）
    On Error Resume Next
    r = Target.Row - Range("事件").Row + 1
    On Error GoTo 0

    '選択範囲がテーブルの外であれば処理を終了する
    If r < 1 Or Range("事件").Rows.Count < r Then Exit Sub

    '現在アクティブな事件番号をグローバル変数 lngActCase に記録する
    '（他のシートから戻ってきたときに同じ行を選択するため）
    lngActCase = Range("事件[事件番号]")(r).Value
End Sub

'================================================================================
' イベント:     Worksheet_BeforeDoubleClick
' 機能:         ダブルクリック時のアクション実行
' 概要:         クリックされた列に応じて、新規事件の追加、フォルダを開く、
'               TeamsURLを開く等のコマンドを実行します。
'================================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'ダブルクリックされた行番号と列番号をテーブル相対位置に変換する
    Dim r As Long, s As Long
    
    On Error Resume Next
    r = Target.Row - Range("事件").Row + 1
    s = Target.Column - Range("事件").Column + 1
    On Error GoTo 0

    'クリック位置がテーブルの範囲外であれば処理を終了する
    If r < 1 Or Range("事件").Rows.Count < r Then Exit Sub
    If s < 1 Or Range("事件").Columns.Count < s Then Exit Sub

    'デフォルトのダブルクリック動作（セル編集モード移行など）をキャンセルする
    Cancel = True

    'イベントを停止する
    Evnt.Disable

    'ダブルクリックされた列の見出し（項目名）に応じて処理を分岐する
    Dim headerName As String
    headerName = Range("事件").ListObject.HeaderRowRange(s).Value
    
    Select Case headerName
        Case "事件番号"
            'マウスポインタを標準に戻す
            Application.Cursor = xlNormal
            '事件を追加する
            InsertCase
            
            '行の入力規則（ドロップダウンリスト）などを再設定する
            Call FormatRow(r)
        
        Case "データフォルダ"
            'データフォルダを開く処理
            If Range("事件[データフォルダ]")(r).Value <> "" Then
                On Error Resume Next
                'フォルダが存在するか確認して開く (Mac対応のため条件分岐の可能性を考慮、DirはMacでも動作)
                If Dir(Range("事件[データフォルダ]")(r).Value, vbDirectory) <> "" Then
                    ThisWorkbook.FollowHyperlink Range("事件[データフォルダ]")(r).Value
                Else
                    MsgBox "データフォルダが見つかりません。", vbExclamation, "事件データベース"
                End If
                On Error GoTo 0
            End If
            
        Case "TeamsURL"
            'TeamsのURLを開く処理
            If Range("事件[TeamsURL]")(r).Value <> "" Then
                On Error Resume Next
                ThisWorkbook.FollowHyperlink Range("事件[TeamsURL]")(r).Value
                If Err.Number <> 0 Then
                    MsgBox "URLを開けませんでした。" & vbCr & Err.Description, vbExclamation, "事件データベース"
                End If
                On Error GoTo 0
            End If
    End Select

    'イベントを再開する
    Evnt.Enable
End Sub

'================================================================================
' プロシージャ: InsertCase
' 機能:         新規事件データの追加
' 概要:         新規依頼人または既存依頼人の追加案件として、新しい事件番号を
'               自動採番し、テーブルに行を追加します。
'================================================================================
Private Sub InsertCase()
    ' シートが保護されている場合は新規追加をブロックする
    If Me.ProtectContents = True Then
        MsgBox "シートが保護されているため、新規事件の追加はできません。" & vbCrLf & _
               "シートの保護を解除してから実行してください。", vbInformation, "事件データベース"
        Exit Sub
    End If

    '現在選択されている行番号を取得する
    Dim r As Long
    
    If Selection.Count > 1 Then
        r = Selection.Cells(1).Row - Range("事件").Row + 1
    Else
        r = Selection.Row - Range("事件").Row + 1
    End If
    
    '範囲外なら最終行とする
    If r < 1 Or r > Range("事件").Rows.Count Then
        r = Range("事件").Rows.Count
    End If

    '新規の依頼人か、既存の依頼人の追加案件かをユーザーに確認する
    Select Case MsgBox("追加するのは、新規の依頼人ですか？", vbYesNoCancel, "事件データベース")
        Case vbYes
            '-- 新規依頼人の場合 --
            Dim rTmp As Long
            Dim lngTmpCase As Long
            lngTmpCase = 0
            
            'すべての行を走査して最大値を見つける
            For rTmp = 1 To Range("事件").Rows.Count
                If Range("事件[事件番号]")(rTmp).Value > lngTmpCase Then
                    lngTmpCase = Range("事件[事件番号]")(rTmp).Value
                End If
            Next
            
            '現在行に既にデータがある場合は、行を挿入してスペースを作る
            If Range("事件[事件番号]")(r).Value <> "" Then
                Range("事件").Rows(r + 1).EntireRow.Insert
                r = r + 1
            End If
                
            '新しい事件番号を採番する（最大値の百の位を繰り上げ、下2桁を01にする）
            Range("事件[事件番号]")(r).Value = Int(lngTmpCase / 100) * 100 + 101
            '入力しやすくするため、次の列を選択する
            Range("事件")(r, 2).Select
        
        Case vbNo
            '-- 既存依頼人の場合 --
            Dim currentClientID As Long
            currentClientID = Int(Range("事件[事件番号]")(r).Value / 100)
            
            lngTmpCase = Range("事件[事件番号]")(r).Value
            
            '同一依頼人（事件番号の上位桁が同じ）の中で最大の番号を探す
            For rTmp = 1 To Range("事件").Rows.Count
                If Int(Range("事件[事件番号]")(rTmp).Value / 100) = currentClientID And _
                    lngTmpCase < Range("事件[事件番号]")(rTmp).Value Then
                    lngTmpCase = Range("事件[事件番号]")(rTmp).Value
                End If
            Next
                            
            '現在行にデータがある場合は行を挿入する
            If Range("事件[事件番号]")(r).Value <> "" Then
                Range("事件").Rows(r + 1).EntireRow.Insert
                r = r + 1
            End If
                
            '同一依頼人の次の枝番（最大値 + 1）を採番して入力する
            Range("事件[事件番号]")(r).Value = lngTmpCase + 1
            '入力しやすくするため、次の列を選択する
            Range("事件")(r, 2).Select
            
        Case vbCancel
            Exit Sub
    End Select

    '新しい行の入力規則や書式を設定する
    '※列全体に入力規則が設定されていれば、テーブル機能により自動的に継承されます
    Call FormatRow(r)
End Sub

'================================================================================
' プロシージャ: SortTable
' 機能:         テーブルデータの並び替え
' 概要:         現況、受任区分、事件区分、事件番号の優先順位でテーブルを
'               昇順に並び替えます。
'================================================================================
Sub SortTable(Optional ByVal dummy As Byte = 0)
    ' シートが保護されている場合は処理を行わず終了する
    If Me.ProtectContents = True Then Exit Sub

    'ステータスバーに処理中であることを表示する
    Application.StatusBar = "「事件」テーブルの並び替えを行っています。"

    '「事件」テーブルのソート設定を行う
    With ListObjects("事件").Sort
        With .SortFields
            .Clear
            '第1キー：現況（昇順）
            .Add Key:=Range("事件[現況]"), SortOn:=xlSortOnValues, Order:=xlAscending
            '第2キー：受任区分（昇順）
            .Add Key:=Range("事件[受任区分]"), SortOn:=xlSortOnValues, Order:=xlAscending
            '第3キー：事件区分（昇順）
            .Add Key:=Range("事件[事件区分]"), SortOn:=xlSortOnValues, Order:=xlAscending
            '第4キー：事件番号（昇順）
            .Add Key:=Range("事件[事件番号]"), SortOn:=xlSortOnValues, Order:=xlAscending
        End With
        
        .Header = xlYes
        .Orientation = xlTopToBottom
        .SortMethod = xlStroke
        .Apply
    End With
    
    'ステータスバーをクリアする
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: FormatTable
' 機能:         テーブル全体の書式設定
' 概要:         各行の「現況」に応じたフォント色の一括設定、日付形式の設定、
'               および入力規則の動的再設定を行います。
'               配列処理と一括適用により処理速度を最適化しています。
'================================================================================
Sub FormatTable(Optional ByVal dummy As Byte = 0)
    ' シートが保護されている場合は処理を行わず終了する
    If Me.ProtectContents = True Then Exit Sub

    Application.StatusBar = "「事件」テーブルの書式を設定しています。"

    Dim tbl As ListObject
    Dim statusValues As Variant
    Dim r As Long, startRow As Long, lastRow As Long
    Dim currentStatus As String, lastStatus As String

    'テーブル変数の設定
    Set tbl = Me.ListObjects("事件")
    
    'テーブルにデータがない場合は処理を終了
    If tbl.ListRows.Count = 0 Then
        Application.StatusBar = ""
        Exit Sub
    End If
    
    '描画停止（念のため）
    Application.ScreenUpdating = False
    
    '1. 全体のフォント色をデフォルト（自動/黒）に一括リセット
    '   これにより、条件に合致しない行を個別に処理する必要がなくなります
    tbl.DataBodyRange.Font.ColorIndex = xlAutomatic
    
    '2. 「現況」列の値を配列として一括取得（セルへの繰り返しアクセスを回避）
    statusValues = tbl.ListColumns("現況").DataBodyRange.Value
    
    lastRow = UBound(statusValues, 1)
    startRow = 1
    lastStatus = statusValues(1, 1)
    
    '3. 配列を走査して連続するステータスごとにまとめて書式適用
    For r = 2 To lastRow
        currentStatus = statusValues(r, 1)
        
        'ステータスが変わったタイミングで、前の区間を一括処理
        If currentStatus <> lastStatus Then
            Call ApplyStatusColor(tbl, startRow, r - 1, lastStatus)
            startRow = r
            lastStatus = currentStatus
        End If
    Next r
    
    '最後の区間を処理
    Call ApplyStatusColor(tbl, startRow, lastRow, lastStatus)

    '4. セル内の折り返し表示を有効にする
    tbl.DataBodyRange.WrapText = True

    '5. 特定の列の表示形式とフォント設定（列全体への一括操作なので高速）
    'TeamsURLはフォントサイズを小さく
    tbl.ListColumns("TeamsURL").DataBodyRange.Font.Size = 6
    '日付形式を和暦に設定
    tbl.ListColumns("生年月日").DataBodyRange.NumberFormatLocal = "[$-ja-JP]ge.m.d"
    tbl.ListColumns("相談年月日").DataBodyRange.NumberFormatLocal = "[$-ja-JP]ge.m.d"
    tbl.ListColumns("終了年月日").DataBodyRange.NumberFormatLocal = "[$-ja-JP]ge.m.d"
    
    '--- 入力規則の一括設定 ---
    'INDIRECT関数を使用して動的にリストを参照します
    Call SetColumnValidation

    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: ApplyStatusColor
' 機能:         範囲への色適用（FormatTableヘルパー）
' 概要:         指定された行範囲に対して、ステータスに応じたフォント色を設定します。
'================================================================================
Private Sub ApplyStatusColor(tbl As ListObject, startRow As Long, endRow As Long, status As String)
    '対象外のステータスなら何もしない（既にデフォルト色になっているため）
    If status <> "C:終了" And status <> "D:相談のみ" And status <> "E:相談不可" Then Exit Sub
    
    Dim targetRange As Range
    '行範囲を特定（Rangeオブジェクトの生成）
    Set targetRange = tbl.DataBodyRange.Rows(startRow & ":" & endRow)
    
    Select Case status
        Case "C:終了"
            targetRange.Font.ThemeColor = xlThemeColorAccent3
        Case "D:相談のみ"
            targetRange.Font.ThemeColor = xlThemeColorAccent6
        Case "E:相談不可"
            targetRange.Font.ThemeColor = xlThemeColorAccent2
    End Select
End Sub

'================================================================================
' プロシージャ: SetColumnValidation
' 機能:         入力規則の動的設定
' 概要:         テーブルの各列に対し、INDIRECT関数を用いてオプションシートの
'               リスト範囲を動的に参照する入力規則を設定します。
'================================================================================
Private Sub SetColumnValidation()
    Dim tbl As ListObject
    Set tbl = Me.ListObjects("事件")

    'エラー回避（列名が存在しない場合など）
    On Error Resume Next
    
    '相談弁護士 -> 担当[担当区分]
    ApplyValidation tbl.ListColumns("相談弁護士"), "担当区分[担当区分]"
    
    '受任区分 -> 受任区分[受任区分]
    ApplyValidation tbl.ListColumns("受任区分"), "受任区分[受任区分]"
    
    '事件区分 -> 事件区分[事件区分]
    ApplyValidation tbl.ListColumns("事件区分"), "事件区分[事件区分]"
    
    '担当区分 -> 担当区分[担当区分]
    ApplyValidation tbl.ListColumns("担当区分"), "担当区分[担当区分]"
    
    '依頼人立場 -> 依頼人立場[依頼人立場]
    ApplyValidation tbl.ListColumns("依頼人立場"), "依頼人立場[依頼人立場]"

    On Error GoTo 0
End Sub

'================================================================================
' プロシージャ: ApplyValidation
' 機能:         入力規則設定の共通処理（ヘルパー）
' 概要:         指定された列に対して、INDIRECT関数を使用した入力規則リストを
'               適用します。
'================================================================================
Private Sub ApplyValidation(lc As ListColumn, referenceFormula As String)
    With lc.DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=INDIRECT(""" & referenceFormula & """)"
    End With
End Sub

'================================================================================
' プロシージャ: FormatRow
' 機能:         特定行の書式設定
' 概要:         指定行の「現況」フィールドの値に基づき、行全体のフォント色を
'               変更します。
'================================================================================
Private Sub FormatRow(r)
    ' シートが保護されている場合は処理を行わず終了する
    If Me.ProtectContents = True Then Exit Sub

    '現況に応じたフォント色の設定
    Dim strStatus As String
    strStatus = Range("事件[現況]")(r).Value

    With Range("事件").Rows(r)
        Select Case strStatus
            Case "E:相談不可": .Font.ThemeColor = xlThemeColorAccent2
            Case "D:相談のみ": .Font.ThemeColor = xlThemeColorAccent6
            Case "C:終了":     .Font.ThemeColor = xlThemeColorAccent3
            Case Else:         .Font.ColorIndex = xlAutomatic
        End Select
    End With
End Sub

'================================================================================
' プロシージャ: CheckFolders
' 機能:         全データのフォルダ存在チェック
' 概要:         すべてのレコードについて、データフォルダが存在するか確認し、
'               存在しない場合はフォント色を赤に変更します。
'================================================================================
Sub CheckFolders(Optional ByVal dummy As Byte = 0)
    Application.StatusBar = "データフォルダのチェックを行っています。"
    Me.Unprotect

    Dim r As Long
    For r = 1 To Range("事件").Rows.Count
        'データフォルダが入力されている場合のみチェック
        If Range("事件[データフォルダ]")(r).Value <> "" Then
            With Range("事件[データフォルダ]")(r).Font
                On Error Resume Next
                'Dir関数はMacでも動作する
                If Dir(Range("事件[データフォルダ]")(r).Value, vbDirectory) = "" Then
                    .ThemeColor = xlThemeColorAccent2 '存在しない場合: 赤
                Else
                    .ColorIndex = xlAutomatic         '存在する場合: 黒
                End If
                On Error GoTo 0
            End With
        End If
    Next

    If shtOpt.Range("事件モード") = "入力" Then Me.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    
    '完了メッセージ
    MsgBox "データフォルダのチェックを完了しました。", vbOKOnly, "事件データベース"
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: CheckFolder
' 機能:         単一データのフォルダ存在チェック
' 概要:         指定行のデータフォルダが存在するか確認し、存在しない場合は
'               警告メッセージを表示します。
'================================================================================
Private Sub CheckFolder(r As Long)
    Application.StatusBar = "データフォルダのチェックを行っています。"

    If Range("事件[データフォルダ]")(r).Value <> "" Then
        On Error Resume Next
        If Dir(Range("事件[データフォルダ]")(r).Value, vbDirectory) = "" Then
            'エラーメッセージ
            MsgBox "データフォルダの値が不正か、フォルダが存在しません。", vbCritical, "事件データベース"
        End If
        On Error GoTo 0
    End If

    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: CheckNumbers
' 機能:         事件番号の整合性チェック
' 概要:         テーブル内で事件番号の重複や空白がないか検証し、異常があれば
'               警告を表示します。
'================================================================================
Sub CheckNumbers(Optional ByVal dummy As Byte = 0)
    Application.StatusBar = "事件番号の重複および空白をチェックしています。"
    Me.Unprotect

    Dim r As Long
    For r = 1 To Range("事件").Rows.Count
        '重複または0（空白）をチェック
        If WorksheetFunction.CountIf(Range("事件[事件番号]"), Range("事件[事件番号]")(r).Value) > 1 Or _
            Range("事件[事件番号]")(r).Value = 0 Then
            
            'エラーメッセージ
            MsgBox r & "行目の「事件番号」に重複または空白のデータがあります。", vbCritical, "事件データベース"
            GoTo Cleanup
        End If
    Next

    '完了メッセージ
    MsgBox "事件番号の重複および空白のチェックを完了しました。", vbOKOnly, "事件データベース"

Cleanup:
    If shtOpt.Range("事件モード") = "入力" Then Me.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: MoveFolder
' 機能:         フォルダの移動・整理
' 概要:         事件の状態（終了、精算管理等）に応じて、データフォルダを
'               所定のアーカイブ場所へ移動・ZIP化します。
'================================================================================
Private Sub MoveFolder(r As Long)
    '移動元のフォルダパスを取得
    Dim strFolderA As String
    strFolderA = Range("事件[データフォルダ]")(r).Value

    If strFolderA = "" Then Exit Sub
    
    '移動元が存在するか確認
    If Dir(strFolderA, vbDirectory) = "" Then
        MsgBox "移動元フォルダが見つかりません: " & strFolderA, vbExclamation, "事件データベース"
        Exit Sub
    End If

    '終了年を取得（2017以前はまとめる）
    Dim strYear As Variant
    strYear = Year(Range("事件[終了年月日]")(r))
    If Val(strYear) <= 2017 Then strYear = "-2017"

    'フォルダシート（shtFldr）を参照して移動先ルールを決定
    Dim strFolderB As String
    Dim PS As String
    PS = Application.PathSeparator
    
    With shtFldr
        Dim t As Long
        For t = 1 To .Range("フォルダ").Rows.Count
            '条件一致（現況、受任区分、事件区分）を確認
            If .Range("フォルダ[現況]")(t).Value = Range("事件[現況]")(r).Value And _
                .Range("フォルダ[受任区分]")(t).Value = Range("事件[受任区分]")(r).Value And _
                .Range("フォルダ[事件区分]")(t).Value = Range("事件[事件区分]")(r).Value Then
                
                '移動先のベースパスを取得
                strFolderB = .Range("フォルダ[フォルダ]")(t).Value
            
                'プレースホルダ "yyyy" の置換処理（パス区切り文字に配慮）j
                'Mac/Windows両対応のため Application.PathSeparator で検索
                If InStr(strFolderB, PS & "yyyy") <> 0 Then
                    strFolderB = Left(strFolderB, InStrRev(strFolderB, PS))
                    strFolderB = strFolderB & strYear
                'Windowsのバックスラッシュがハードコードされていた場合のフォールバック
                ElseIf InStr(strFolderB, "yyyy") <> 0 Then
                     strFolderB = Replace(strFolderB, "yyyy", PS & strYear)
                End If
                
                '末尾に移動元のフォルダ名（依頼者名など）を付加
                'InStrRevで最後の区切り文字を探すが、Macの場合は "/" になる
                strFolderB = strFolderB & Mid(strFolderA, InStrRev(strFolderA, PS))
                Exit For
            End If
        Next
    End With
        
    If strFolderB = "" Then
        MsgBox "移動先のフォルダルールが見つかりませんでした。", vbExclamation, "事件データベース"
        Exit Sub
    End If

    If strFolderA = strFolderB Then Exit Sub

    '移動確認
    If MsgBox("フォルダを以下の場所に移動します。" & vbCr & vbCr & _
        "移動元：" & strFolderA & vbCr & _
        "移動先：" & strFolderB & vbCr & vbCr & _
        "よろしいですか？", vbOKCancel, "事件データベース") = vbCancel Then Exit Sub

    'ZIP移動処理の呼び出し
    Call Fldr.MoveInZip(strFolderA, strFolderB)

    '移動成功確認
    On Error Resume Next
    If Dir(strFolderB, vbDirectory) <> "" Or Dir(strFolderB & ".zip") <> "" Then
        'テーブル上のパスを更新
        Range("事件[データフォルダ]")(r).Value = strFolderB
        
        'scanフォルダの作成（必要に応じて）
        Dim strTemp As String
        strTemp = Mid(strFolderB, InStrRev(strFolderB, PS) + 1)
        If InStr(strTemp, "：") <> 0 Then strTemp = Mid(strTemp, InStr(strTemp, "：") + 1)
        
        '移動先の中にscanフォルダを作成
        'Dirチェックをしてフォルダが存在する場合のみ作成
        If Dir(strFolderB, vbDirectory) <> "" Then
             MkDir strFolderB & PS & "0000" & strTemp & "scan"
             MsgBox "終了事件フォルダにscanフォルダを作成しました。", vbOKOnly, "事件データベース"
        End If
    End If
    On Error GoTo 0
End Sub

'================================================================================
' プロシージャ: WriteNameToCSV
' 機能:         氏名データのエクスポート
' 概要:         事件データから氏名情報を抽出し、外部連携用のCSVファイルに
'               書き出します。
'================================================================================
Private Sub WriteNameToCSV()
    Application.StatusBar = "氏名データを書き込んでいます。"
    On Error GoTo HdlErr

    Dim strCSVFile As String
    strCSVFile = shtOpt.Range("名前データファイル名").Value
    
    If strCSVFile = "" Then Exit Sub

    Dim lines() As String
    Dim lineCount As Long
    lineCount = 0
    ReDim lines(1000)
    
    Dim r As Long
    Dim i As Integer, n As Integer
    Dim strName As String, strTargetCol As String
    Dim intCategory As Integer, intBracket As Integer

    '全データを走査して配列に格納
    For r = 1 To Range("事件").Rows.Count
        '除外条件: 終了、相談不可、古い相談のみ案件
        If Range("事件[現況]")(r).Value = "C:終了" Or Range("事件[現況]")(r).Value = "E:相談不可" Then GoTo ContinueLoop
        If Range("事件[現況]")(r).Value = "C:相談のみ" And _
            Int(Range("事件[事件番号]")(r).Value / 100000) < (Year(Date) Mod 100) - 1 Then GoTo ContinueLoop

        Dim lngCase As String
        lngCase = Left(Range("事件[事件番号]")(r).Value, 5)

        '3種類の名前列について処理
        For i = 1 To 3
            Select Case i
                Case 1: strTargetCol = "依頼者氏名": intCategory = 1
                Case 2: strTargetCol = "関係者":     intCategory = 2
                Case 3: strTargetCol = "相手方氏名": intCategory = 3
            End Select
            
            n = 1
            strName = "dummy"
            
            '改行区切りの名前を分解
            Do Until strName = ""
                strName = FncSplit(Range("事件[" & strTargetCol & "]")(r).Value, Chr(10), n)
                
                If strName <> "" Then
                    '括弧書き除去
                    intBracket = InStr(Replace(strName, "(", "（"), "（")
                    If intBracket > 1 Then strName = Left(strName, intBracket - 1)

                    '配列拡張
                    If lineCount > UBound(lines) Then ReDim Preserve lines(UBound(lines) + 1000)
                    
                    lines(lineCount) = lngCase & vbTab & intCategory & vbTab & strName
                    lineCount = lineCount + 1
                End If
                n = n + 1
            Loop
        Next i
ContinueLoop:
    Next r

    'ファイル保存（Mac対応のためFSO廃止、標準I/O使用）
    If lineCount > 0 Then
        ReDim Preserve lines(lineCount - 1)
        
        Dim intFile As Integer
        intFile = FreeFile
        
        '標準のOpenステートメントで書き込み（Mac互換）
        Open strCSVFile For Output As #intFile
        Print #intFile, Join(lines, vbCrLf)
        Close #intFile
    End If

    Application.StatusBar = ""
    Exit Sub
HdlErr:
    'エラーメッセージ
    MsgBox "名前データの書き込みに失敗しました。" & vbCrLf & Err.Description, vbExclamation, "事件データベース"
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: WriteFolderToCSV
' 機能:         フォルダデータのエクスポート
' 概要:         事件データからフォルダパス情報を抽出し、外部連携用のCSVファイルに
'               書き出します。
'================================================================================
Private Sub WriteFolderToCSV()
    Application.StatusBar = "フォルダデータを書き込んでいます。"
    On Error GoTo HdlErr

    Dim strCSVFile As String
    strCSVFile = shtOpt.Range("フォルダデータファイル名").Value
    
    If strCSVFile = "" Then Exit Sub

    Dim lines() As String
    Dim lineCount As Long
    ReDim lines(Range("事件").Rows.Count)
    lineCount = 0

    Dim r As Long
    '全データを走査して配列に格納
    For r = 1 To Range("事件").Rows.Count
        'スキップ条件
        If Range("事件[データフォルダ]")(r).Value = "" Then GoTo ContinueLoop
        If Range("事件[現況]")(r).Value = "D:相談のみ" Then GoTo ContinueLoop
        
        lines(lineCount) = Range("事件[事件番号]")(r).Value & vbTab & _
                           Range("事件[事件呼名]")(r).Value & vbTab & _
                           Range("事件[担当区分]")(r).Value & vbTab & _
                           Range("事件[データフォルダ]")(r).Value
        lineCount = lineCount + 1
ContinueLoop:
    Next r

    'ファイル保存（Mac対応のためFSO廃止、標準I/O使用）
    If lineCount > 0 Then
        ReDim Preserve lines(lineCount - 1)
        
        Dim intFile As Integer
        intFile = FreeFile
        
        '標準のOpenステートメントで書き込み
        Open strCSVFile For Output As #intFile
        Print #intFile, Join(lines, vbCrLf)
        Close #intFile
    End If

    Application.StatusBar = ""
    Exit Sub
HdlErr:
    'エラーメッセージ
    MsgBox "フォルダデータの書き込みに失敗しました。" & vbCrLf & Err.Description, vbExclamation, "事件データベース"
    Application.StatusBar = ""
End Sub


