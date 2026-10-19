Option Explicit
'================================================================================
' フォーム: frmTokenType
' 機能: 仮名化する語の種別を選ぶ。種別のオプションボタンを選んで［仮名化］を押す（または種別をダブルクリックする）と、
'       その種別を選んで閉じる。
'       見た目は、シリーズのほかのアドインのフォームにそろえ、Word 標準の部品（Label・Frame・OptionButton・
'       CommandButton）だけを使う（2026-10-17 班長指示）。部品は班長が VBE のデザイン画面に置く
'       （名前と配置は、作業フォルダの frmTokenType_部品一覧.md）。
'       ［仮名化］（cmdOK）は Default、［キャンセル］は Cancel を True にしてあるので、Enter で仮名化、Esc でキャンセルになる。
' 使い方: strType = frm.SelectType(strWord, strDefaultType)（キャンセルした場合は空の文字列が返る）
'================================================================================

Private strSelectedType As String ' 選んだ種別（キャンセルした場合は空）
Private lngInitialType As Long ' 最初に選んでおく種別のオプションボタンの番号（表示したときに焦点を移す）

Private Const lngTypeCount As Long = 5 ' 種別のオプションボタンの数（デザイン画面に置いた数）

'================================================================================
' 機能:         フォームを表示して、種別を選ばせます。
' 引数:         strWord - 仮名化する語（フォームに表示する）
'               strDefaultType - 最初に選んでおく種別（前回選んだ種別。空の場合は最初の種別）
' 戻り値:       選んだ種別（キャンセルした場合は空の文字列）
'================================================================================
Public Function SelectType(ByVal strWord As String, ByVal strDefaultType As String) As String
    Dim varTypes As Variant ' 種別の一覧
    Dim blnSelected As Boolean ' 最初に選んでおく種別を選べたかどうか
    Dim i As Long

    '種別のオプションボタンの文字を、定数の種別の一覧から設定する（種別がボタンの数より少ない場合は隠す）
    varTypes = Split(strTokenTypes, ",")
    blnSelected = False
    lngInitialType = 1
    For i = 1 To lngTypeCount
        With Me.Controls("optType" & i)
            If i - 1 <= UBound(varTypes) Then
                .Caption = varTypes(i - 1)
                '前回選んだ種別を選んでおく
                If varTypes(i - 1) = strDefaultType Then
                    .Value = True
                    blnSelected = True
                    lngInitialType = i
                End If
            Else
                .Visible = False
            End If
        End With
    Next i
    '前回選んだ種別がない場合は、最初の種別を選んでおく
    If Not blnSelected Then optType1.Value = True

    '仮名化する語を表示し、語の長さに合わせて、下の部品とフォームの高さを調整する
    Call ShowWord(strWord)

    '選んだ種別を初期化して、フォームを表示する（閉じるまで待つ）
    strSelectedType = ""
    Me.Show vbModal

    '選んだ種別を返す
    SelectType = strSelectedType
End Function

'================================================================================
' 機能:         仮名化する語を表示し、語のラベルが高くなった分だけ、下の部品とフォームを広げます。
' 引数:         strWord - 仮名化する語
'================================================================================
Private Sub ShowWord(ByVal strWord As String)
    Dim sngOldHeight As Single ' 語のラベルの、もとの高さ
    Dim sngOldWidth As Single ' 語のラベルの、もとの幅
    Dim sngDelta As Single ' 高さの差
    Dim ctl As Object ' 部品

    With lblWord
        sngOldHeight = .Height
        sngOldWidth = .Width
        .Caption = strWord
        '幅を固定したまま折り返し、高さを文字に合わせる
        .WordWrap = True
        .AutoSize = True
        .AutoSize = False
        .Width = sngOldWidth
        '1行の高さより低くしない
        If .Height < sngOldHeight Then .Height = sngOldHeight
        sngDelta = .Height - sngOldHeight
    End With
    If sngDelta <= 0 Then Exit Sub

    '語のラベルより下にある部品を下げる（フレームの中の部品は、フレームと一緒に動くので動かさない）
    For Each ctl In Me.Controls
        If ctl.Parent Is Me Then
            If ctl.Top > lblWord.Top Then ctl.Top = ctl.Top + sngDelta
        End If
    Next ctl
    'フォームを高くする
    Me.Height = Me.Height + sngDelta
End Sub

'================================================================================
' 機能:         選んでいる種別を記録して、フォームを閉じます（呼び出し元で Unload する）。
'================================================================================
Private Sub DecideType()
    Dim i As Long

    For i = 1 To lngTypeCount
        If Me.Controls("optType" & i).Value = True Then
            strSelectedType = Me.Controls("optType" & i).Caption
            Exit For
        End If
    Next i
    Me.Hide
End Sub

'================================================================================
' イベント
'================================================================================

' フォームが表示されたとき、選んでおいた種別に焦点を移す（矢印キーで種別を変えるとき、そこから動くようにする）
Private Sub UserForm_Activate()
    On Error Resume Next
    Me.Controls("optType" & lngInitialType).SetFocus
    On Error GoTo 0
End Sub

' ［仮名化］（cmdOK）が押されたとき（Enter キーでも押される）
Private Sub cmdOK_Click()
    Call DecideType
End Sub

' ［キャンセル］が押されたとき（Esc キーでも押される）
Private Sub cmdCancel_Click()
    strSelectedType = ""
    Me.Hide
End Sub

' 閉じるボタン（×）が押されたとき、フォームを破棄せずに隠す（呼び出し元で選んだ種別を読み取ってから Unload するため）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Call cmdCancel_Click
    End If
End Sub

' 種別をダブルクリックしたとき、その種別で仮名化する
Private Sub optType1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    optType1.Value = True
    Call DecideType
End Sub
Private Sub optType2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    optType2.Value = True
    Call DecideType
End Sub
Private Sub optType3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    optType3.Value = True
    Call DecideType
End Sub
Private Sub optType4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    optType4.Value = True
    Call DecideType
End Sub
Private Sub optType5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    optType5.Value = True
    Call DecideType
End Sub
