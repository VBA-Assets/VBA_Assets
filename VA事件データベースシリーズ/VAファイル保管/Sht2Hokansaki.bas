Option Explicit

'================================================================================
' モジュール名: Sht2Hokansaki
' 概要:         「保管先」シートのイベントやデータ操作を管理するクラスです。
' 更新日時:     2026/01/31 18:23
'================================================================================

' ADODB.Stream用定数 (ReadCSVで使用)
Private Const adReadLine = -2

'================================================================================
' イベント:  Worksheet_Activate
' 機能:      ワークシートがアクティブになったときの処理
' 概要:      「保管」シートで選択されていた事件呼名と同じ事件呼名をこのシートで探し、
'            選択状態にします。ユーザーの視線移動を補助します。
'================================================================================
Private Sub Worksheet_Activate()
    Dim r As Long

    ' 対象事件呼名がなければ終了
    If strActCase = "" Then Exit Sub
    
    ' オブジェクト変数を使用して操作
    Dim rngCaseList As Range
    Set rngCaseList = Range("保管先[事件呼名]")
    
    ' 事件呼名が存在するか確認
    If WorksheetFunction.CountIf(rngCaseList, strActCase) > 0 Then
        ' 行番号を取得して選択
        r = WorksheetFunction.Match(strActCase, rngCaseList, 0)
        Range("保管先")(r, 1).Select
    End If
End Sub

'================================================================================
' イベント:  Worksheet_BeforeDoubleClick
' 機能:      ワークシートがダブルクリックされたときの処理
' 概要:      選択した「保管先」の事件呼名を、「保管」シートの現在行に転記します。
'            オブジェクト変数を使用し、明示的なセル操作を行っています。
'================================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim r As Long, s As Long
    Dim t As Long
    Dim rngTable As Range
    
    Set rngTable = Range("保管先")
    
    ' クリック位置の相対座標を取得
    With rngTable
        r = Target.Row - .Row + 1
        s = Target.Column - .Column + 1
        ' 範囲外なら終了
        If r = 0 Or r > .Rows.Count Or s > .Columns.Count Then Exit Sub
    End With

    ' イベントの発生を停止する
    Evnt.Disable

    ' 事件呼名の列がダブルクリックされた場合
    If s = Range("保管先[事件呼名]").Column - Range("保管先").Column + 1 Then
        Dim wsHokan As Worksheet
        Set wsHokan = Sht1Hokan
        
        With wsHokan
            .Activate
            ' ActiveCellを使用せざるを得ないコンテキスト（直前の操作位置依存）のため維持するが、
            ' テーブル範囲基準で位置を特定する
            t = ActiveCell.Row - .Range("保管").Row + 1
            
            ' 事件呼名を入力する
            .Range("保管[事件呼名]")(t).Value = rngTable(r, s).Value
            
            ' 事件呼名にフリガナを設定する
            .Range("保管[事件呼名]")(t).Phonetic.Text = Application.GetPhonetic(.Range("保管[事件呼名]")(t).Value)
        
            ' 次の列（内容）を選択する
            .Range("保管[内容]")(t).Select
        End With
        
        Cancel = True
    End If

    ' イベントの発生を再開する
    Evnt.Enable
End Sub


'================================================================================
' プロシージャ: InputScan
' 機能:         [作業用] スキャンフォルダのパスを生成し、保管フォルダ列に入力します。
' 概要:         「データフォルダ」列の値を基に、命名規則に従ったスキャンフォルダの
'               フルパスを組み立て、「保管フォルダ」列を更新します。
'================================================================================
Sub InputScan()
    Dim r As Long
    Dim strFolder As String

    For r = 1 To Range("保管先").Rows.Count
        strFolder = Range("保管先[データフォルダ]")(r).Value
        ' 特定のフォルダパターン（：を含み、scanを含まない）の場合の処理
        If InStr(strFolder, "：") <> 0 And InStr(strFolder, "scan") = 0 Then
            Range("保管先[保管フォルダ]")(r).Value = strFolder & "\" & _
                Replace(Mid(strFolder, InStr(strFolder, "：") + 1), "\", "") & "scan"
        Else
            Range("保管先[保管フォルダ]")(r).Value = strFolder
        End If
    Next
End Sub

'================================================================================
' プロシージャ: UpdateTable
' 機能:         [作業用] 「保管」シートのデータを「保管先」シートに転記・更新します。
' 概要:         「保管」シートの全データをループし、まだ「保管先」シートに存在しない
'               事件呼名のデータを、「保管先」シートの末尾に追加します。
'================================================================================
Sub UpdateTable()
    Dim r As Long, t As Long
    Dim wsHokan As Worksheet
    Set wsHokan = Sht1Hokan
    
    With wsHokan
        For t = 1 To .Range("保管").Rows.Count
            ' 未登録の事件呼名の場合
            If WorksheetFunction.CountIf(Range("保管先[事件呼名]"), .Range("保管[事件呼名]")(t).Value) = 0 Then
                r = Range("保管先").Rows.Count + 1
            
                ' データを転記する
                Range("保管先[事件呼名]")(r).Value = .Range("保管[事件呼名]")(t).Value
                Range("保管先[保管フォルダ]")(r).Value = .Range("保管[保管フォルダ名]")(t).Value
                Range("保管先[担当区分]")(r).Value = .Range("保管[複製作成先]")(t).Value
            End If
        Next
    End With
End Sub

