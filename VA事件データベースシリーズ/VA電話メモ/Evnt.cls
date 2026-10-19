'更新日時: 2026/02/23 20:39
'概要: イベント・画面描画・自動計算の抑止および再開を制御するクラス
Option Explicit

'--------------------------------------------------
' 概要 : イベント抑止と速度向上のための設定を行う
' 引数 : なし
' 戻り値 : なし
'--------------------------------------------------
Sub Disable()
    'イベント抑止と速度向上のための設定を行います
    With Application
        .ScreenUpdating = False '画面描画を停止
        .Cursor = xlWait 'マウスポインタをウエイトカーソル
        .EnableEvents = False 'イベントを抑止
        .Calculation = xlCalculationManual '計算を手動に
    End With
End Sub

'--------------------------------------------------
' 概要 : イベント抑止と速度向上のための設定を解除する
' 引数 : なし
' 戻り値 : なし
'--------------------------------------------------
Sub Enable()
    'イベント抑止と速度向上のための設定を解除します
    With Application
        .Calculation = xlCalculationAutomatic '計算を自動に
        .EnableEvents = True 'イベントを開始
        .Cursor = xlNormal 'マウスポインタを標準カーソル
        .ScreenUpdating = True '画面描画を開始
    End With
End Sub