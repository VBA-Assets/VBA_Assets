Option Explicit

'================================================================================
' モジュール名: shtOpt (設定シート)
' 概要:         アプリケーションの各種設定を管理するクラスモジュールです。
'               設定値の変更検知や、メンテナンス用コマンドの実行トリガーを
'               制御します。
' 更新日時:     2026/02/04 06:40
'================================================================================

'================================================================================
' イベント:     Worksheet_BeforeDoubleClick
' 機能:         設定シート上のコマンド実行
' 概要:         セルをダブルクリックすることで、フォルダ確認や番号チェックなどの
'               メンテナンス機能を実行します。
'================================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Evnt.Disable
    Cancel = True 'デフォルト動作（編集モード）をキャンセル

    'クリックされたセル（名前付き範囲）に応じて処理分岐
    Select Case Target.Address
        Case Range("データフォルダ確認").Address
            '全件のフォルダ存在チェック実行
            'CheckFoldersにダミー引数がついたが、Optionalなのでそのまま呼び出し可能
            Call sht11.CheckFolders

        Case Range("事件番号確認").Address
            '事件番号の重複・空白チェック実行
            'CheckNumbersにダミー引数がついたが、Optionalなのでそのまま呼び出し可能
            Call sht11.CheckNumbers
    End Select

    Evnt.Enable
End Sub
