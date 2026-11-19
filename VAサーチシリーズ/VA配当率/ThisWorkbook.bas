'VA配当率
'https://vba-assets.net/automatically-calculate-the-dividend-rate/
'Copyright (c) 2026 VBAアセット
'MIT License

'配当表で、配当金の合計が目標の金額になる配当率を、自動で計算するアドインです。

'更新履歴
'2026/01/03  Ver100  作成（配当表（配当率計算ブック）から計算機能を抽出）
'2026/01/23  Ver110  リボンから起動できるように修正
'2026/11/19  Ver111  計算の途中でエラーが起きたときも、画面の更新を元に戻し、番号と内容を表示するよう修正
'                    メッセージの題名とアイコンを揃え、定数のモジュール（cstSettings）を追加
'                    VBA プロジェクトの名前を VB_Dividened から VA_Dividend に変更
'                    冒頭の著作権の表示を改め、MIT License を明記
'                    画面の題名を「VA配当率」から「配当率計算」に変更
