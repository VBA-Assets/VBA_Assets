Option Explicit

' =========================================================
' 変更履歴
' 2026/03/18    Ver100    新規作成
' 2026/03/19    Ver101    行の途中で改ページしないように設定
' 2026/03/21    Ver110    標目の結合・分割を実行前に選択するように修正
' 2026/05/16    Ver120    結合・分割を「作成年月日」および「作成者」をダブルクリックしても可能に変更
'                         作成年月日の書式設定の起動は停止
' =========================================================

' 処理中かどうかを判定するためのモジュールレベル変数（多重実行防止用）
Private isProcessing As Boolean

' ---------------------------------------------------------
' 共通関数：セル内のテキストからセル区切り文字を削除して取得する
' ---------------------------------------------------------
Private Function GetCellText(ByVal rng As Range) As String
    ' セル末尾の不要な制御文字（Chr(13) と Chr(7)）を消去し、前後の空白も削除して返す
    GetCellText = Trim(Replace(rng.Text, Chr(13) & Chr(7), ""))
End Function

' ---------------------------------------------------------
' 原告・被告マークをトグル移動する
'「原告・被告」の文字列をダブルクリックすると起動します。
' ---------------------------------------------------------
Sub TogglePartyCircle()
    ' 既に別の処理が実行中の場合は、何もせずに終了する（連打対策）
    If isProcessing Then Exit Sub
    isProcessing = True
    System.Cursor = wdCursorWait
    On Error GoTo ErrorHandler

    Dim doc As Document
    Dim shp As Shape
    Dim moveDistance As Single
    
    ' 【設定】移動距離を指定します（単位：ポイント）
    moveDistance = 50
    Set doc = ActiveDocument
    
    ' 1. 名前で図形を取得（エラーを一時的に無視して安全に取得）
    On Error Resume Next
    Set shp = doc.Shapes("MarkerCircle")
    On Error GoTo ErrorHandler ' エラーハンドリングを通常に戻す
    
    If shp Is Nothing Then
        MsgBox "対象の図形（MarkerCircle）が見つかりません。", vbExclamation
        GoTo ExitHandler
    End If

    ' 2. 移動処理（図形の代替テキストに現在の状態を記録して判定）
    If shp.AlternativeText <> "Hikoku" Then
        ' 初回実行時、または「原告」状態のときは「被告」へ移動（右へプラス）
        shp.Left = shp.Left + moveDistance
        shp.AlternativeText = "Hikoku" ' 状態を保存
    Else
        ' 「被告」状態のときは「原告」へ戻る（左へマイナス）
        shp.Left = shp.Left - moveDistance
        shp.AlternativeText = "Genkoku" ' 状態を保存
    End If

' === 正常終了時・中断時の共通出口 ===
ExitHandler:
    System.Cursor = wdCursorNormal
    isProcessing = False
    Exit Sub

' === エラー発生時の処理 ===
ErrorHandler:
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & "エラー内容: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

' ---------------------------------------------------------
' 標目の結合・分割を選択する（ユーザーフォームの呼び出し）
'「標目」・「作成年月日」・「作成者」の文字列をダブルクリックすると起動します。
' ---------------------------------------------------------
Sub ToggleMergeByHeader()
    If isProcessing Then Exit Sub
    
    ' カーソルが表の中になければ終了処理へ飛ぶ
    If Not Selection.Information(wdWithInTable) Then Exit Sub
    
    ' 名称を変更したユーザーフォームを表示する
    frmMergeSplit.Show
End Sub

' ---------------------------------------------------------
' 実際の結合・分割処理（frmMergeSplitから呼び出される）
' 引数 actionType: "Merge" または "Split"
' ---------------------------------------------------------
Public Sub ExecuteTableFormat(ByVal actionType As String)
    isProcessing = True
    System.Cursor = wdCursorWait
    On Error GoTo ErrorHandler

    Dim tbl As Table
    Dim rw As Row
    Dim i As Long
    Dim c As Integer
    Dim textCol3 As String
    Dim textCol4 As String
    
    Set tbl = Selection.Tables(1)
    Application.ScreenUpdating = False
    
    ' 2行目から表の最終行まで順番に処理を繰り返す
    For i = 2 To tbl.Rows.Count
        Set rw = tbl.Rows(i)
        
        If actionType = "Split" Then
            ' ---------------------------------------------------------
            ' 【分割処理】すでに結合されている行（セル数が4）を対象とする
            ' ---------------------------------------------------------
            If rw.Cells.Count = 4 Then
                rw.Cells(2).Split NumRows:=1, NumColumns:=3
                
                For c = 1 To 6
                    rw.Cells(c).PreferredWidthType = tbl.Rows(1).Cells(c).PreferredWidthType
                    rw.Cells(c).PreferredWidth = tbl.Rows(1).Cells(c).PreferredWidth
                    On Error Resume Next
                    rw.Cells(c).Width = tbl.Rows(1).Cells(c).Width
                    On Error GoTo ErrorHandler
                Next c
            End If
            
        ElseIf actionType = "Merge" Then
            ' ---------------------------------------------------------
            ' 【結合処理】結合されていない標準行（セル数が6）を対象とする
            ' ---------------------------------------------------------
            If rw.Cells.Count = 6 Then
                ' 3列目および4列目のテキストを取得する
                textCol3 = GetCellText(rw.Cells(3).Range)
                textCol4 = GetCellText(rw.Cells(4).Range)
                
                ' 3列目および4列目のいずれかに文字列が入力されている場合は、処理を行わずにスキップする
                If textCol3 <> "" Or textCol4 <> "" Then
                    GoTo NextRow
                End If
                
                ' 結合処理の実行
                rw.Cells(3).Range.Text = ""
                rw.Cells(4).Range.Text = ""
                rw.Cells(2).Merge MergeTo:=rw.Cells(4)
            End If
        End If
        
NextRow:
    Next i
    
    ' 見出しセルが選択された状態になっているため、選択範囲を先頭に潰して解除する
    Selection.Collapse Direction:=wdCollapseStart

' === 正常終了時・中断時の共通出口 ===
ExitHandler:
    Application.ScreenUpdating = True
    System.Cursor = wdCursorNormal
    isProcessing = False
    Exit Sub

' === エラー発生時の処理 ===
ErrorHandler:
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & "エラー内容: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

' ---------------------------------------------------------
' 作成年月日の書式を設定する（表全体）
' ---------------------------------------------------------
Sub FormatDatesByHeader()
    If isProcessing Then Exit Sub
    isProcessing = True
    System.Cursor = wdCursorWait
    On Error GoTo ErrorHandler

    Dim tbl As Table
    Dim rw As Row
    Dim i As Long
    Dim textCol3 As String
    Dim dateVal As Date
    
    ' カーソルが表の中になければ終了処理へ飛ぶ
    If Not Selection.Information(wdWithInTable) Then GoTo ExitHandler
    
    Set tbl = Selection.Tables(1)
    Application.ScreenUpdating = False
    
    ' 2行目から最終行まで順番に処理を繰り返す
    For i = 2 To tbl.Rows.Count
        Set rw = tbl.Rows(i)
        
        ' 結合されていない標準行（セル数が6個の行）のみを処理対象とする
        If rw.Cells.Count = 6 Then
            ' 共通関数を使って3列目のテキストをきれいに取得する
            textCol3 = GetCellText(rw.Cells(3).Range)
            
            ' 3列目に何かしらの文字が入力されている場合
            If textCol3 <> "" Then
                ' その文字列が「日付」として認識できるデータであるかを判定する
                If IsDate(textCol3) Then
                    dateVal = CDate(textCol3)
                    rw.Cells(3).Range.Text = Format(dateVal, "ge.m.d")
                End If
            End If
        End If
    Next i
    
    ' 選択状態を解除してカーソルを戻す
    Selection.Collapse Direction:=wdCollapseStart

' === 正常終了時・中断時の共通出口 ===
ExitHandler:
    Application.ScreenUpdating = True
    System.Cursor = wdCursorNormal
    isProcessing = False
    Exit Sub

' === エラー発生時の処理 ===
ErrorHandler:
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & "エラー内容: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub