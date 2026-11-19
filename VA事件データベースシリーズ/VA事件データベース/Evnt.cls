Option Explicit

'================================================================================
' モジュール名: Evnt
' 概要:         マクロの処理の間、Excel の画面更新・イベント・自動計算を止め、
'               処理が終わったら、止める前の値に戻すクラスモジュールです。
' 備考:         VA事件データベースシリーズの4本（事件データベース・ファイル保管・
'               領収書等ファイル保管・電話メモ）が、同じものを写して持ちます。
'               1つを直したら、ほかの3本も同じに直します（標準 03「ツールどうしの関係」）。
'               利用者が自動計算を手動にしていても、処理の後に手動のまま戻すため、
'               止める前の値を控えておきます（2026-11-19。それまでは自動などに決め打ちで戻していた）。
'================================================================================

'止める前の値（Disable で控え、Enable で戻す）
Private blnSaved As Boolean             '控えた値があるか（Disable の後、Enable の前は True）
Private blnEverSaved As Boolean         '一度でも値を控えたか
Private blnScreenUpdating As Boolean    '画面の更新
Private blnEnableEvents As Boolean      'イベント
Private lngCalculation As Long          '計算の方法（自動・手動）
Private lngCursor As Long               'マウスポインタ

'--------------------------------------------------------------------------------
' 概要：画面の更新・イベント・自動計算を止め、マウスポインタを待機の形にする。
'       止める前の値を控える（Enable の前に Disable がもう一度呼ばれたときは、最初に控えた値を残す）。
' 引数：なし
' 戻り値：なし
'--------------------------------------------------------------------------------
Sub Disable()
    With Application
        '止める前の値を控える（2回目の Disable で、止めた後の値を控えないようにする）
        If Not blnSaved Then
            blnScreenUpdating = .ScreenUpdating
            blnEnableEvents = .EnableEvents
            lngCalculation = .Calculation
            lngCursor = .Cursor
            blnSaved = True
            blnEverSaved = True
        End If

        .ScreenUpdating = False
        .Cursor = xlWait
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Sub

'--------------------------------------------------------------------------------
' 概要：Disable で止めたものを、止める前の値に戻す。
'       一度も値を控えていないとき（Disable の前にエラーで止まったときなど）は、
'       Excel の既定の値（自動計算・イベントあり・画面の更新あり）に戻す。
' 引数：なし
' 戻り値：なし
'--------------------------------------------------------------------------------
Sub Enable()
    With Application
        If blnEverSaved Then
            .Calculation = lngCalculation
            .EnableEvents = blnEnableEvents
            .Cursor = lngCursor
            .ScreenUpdating = blnScreenUpdating
        Else
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
            .Cursor = xlDefault
            .ScreenUpdating = True
        End If
    End With
    blnSaved = False
End Sub
