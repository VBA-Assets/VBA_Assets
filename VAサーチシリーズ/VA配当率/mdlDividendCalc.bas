'================================================================================
' モジュール: mdlDividendCalc
' 機能:      配当率計算ロジックおよびフォーム管理
' 概要:      配当率計算フォームの表示機能と、
'            目標金額から配当率を算出するメインロジックを提供します。
'================================================================================
Option Explicit

'================================================================================
' プロシージャ: ShowDividendCalcForm
' 機能:      配当率計算フォームを表示します。
' 概要:      ユーザーフォーム「frmDividendCalc」のインスタンスを生成し、
'            Excel操作が可能なモードレス状態で表示します。
'================================================================================
Sub ShowDividendCalcForm()
    'フォームのインスタンス用変数を宣言する
    Dim frm As frmDividendCalc
    
    '新しいフォームのインスタンスを生成する
    Set frm = New frmDividendCalc
    
    'フォームをモードレス（Excel操作可能な状態）で表示する
    '※モードレスにすることで、フォームを開いたままセルの確認が可能になります
    frm.Show vbModeless
    
End Sub

'================================================================================
' プロシージャ: OnRibbonBtnClick
' 機能:      リボンボタンからのコールバック
' 概要:      リボン上のボタンがクリックされたときに呼び出されます。
'            Custom UI Editorなどで設定するonAction属性に対応します。
' 引数:      control (IRibbonControl) - クリックされたリボンコントロールの情報
'================================================================================
Sub OnRibbonBtnClick(control As IRibbonControl)
    '既存のフォーム表示プロシージャを呼び出す
    ShowDividendCalcForm
End Sub

'================================================================================
' プロシージャ: CalcDivRate
' 機能:      目標配当金額になる配当率を計算します。
' 概要:      指定されたセル範囲と目標金額に基づき、小数点以下の数値を
'            下限・上限を絞り込みながら探索し、最適な配当率を返します。
'            ※この関数はマクロから呼び出して使用します（UDF使用不可）。
'================================================================================
Public Function CalcDivRate( _
    ByRef rngTotal As Range, _
    ByRef rngRate As Range, _
    ByVal lngTarget As Currency, _
    ByVal iMax As Long _
) As Double
    '変数を設定する：合計配当金額の計算値（金額型）
    Dim lngResult As Currency
    '変数を設定する：配分率の小数点以下の桁数
    Dim i As Long
    '変数を設定する：配分率の小数点以下の数値
    Dim j As Double
    '変数を設定する：配分率の小数点以下の数値の下限値
    Dim jLower As Double
    '変数を設定する：配分率の小数点以下の数値の上限値
    Dim jUpper As Double
    
    '画面の更新を停止する（処理速度向上のため）
    Application.ScreenUpdating = False
    
    'エラーが発生した場合にエラー処理ルーチンへ飛ばす
    On Error GoTo ErrorHandler

    '配分率の小数点以下の数値の下限値と上限値の初期値を設定する
    'この下限値と上限値は、次のFor～Next構文で計算の桁数が増えるたびに更新されるようになっています。

    'ここがこのマクロのポイントです。
    'この下限値と上限値を設定しないで計算をすると、最初は0.0～1.0、次の桁の際には0.00～１.00の間のすべての数値について、
    '目標値と計算値が合致するかどうかを確認しなければならず、膨大な時間がかかってしまいます。
    '（私のＰＣでは5～6桁が限界でした。）
    'このため、ある数値で計算を行った結果が目標値を下回ったのか、目標値を上回ったのかに応じて、
    '下限値と上限値を設定するようにしています。
    '例えば、0.07の次の桁を計算する場合は、この設定を行わなければ、0.000～1.000のすべての値について計算しなければなりませんが、
    'この設定をおこなえば、0.070～0.080を計算するだけで済みます。
    
    jLower = 0
    jUpper = 1

    '小数点以下1桁から最大桁数まで計算を繰り返す
    'For～Next構文のネスト（入れ子）になっています。
    '細部の説明は省略しますが、For～Next構文の練習問題のようなマクロですので、ぜひ、理解しておくことをお勧めします。
    '細部については、こちらを参照してください。
    'https://www.sejuku.net/blog/30720
    For i = 1 To iMax
        '配分率の小数点以下の数値の下限値を1桁繰り上げる（10倍する）
        jLower = jLower * 10
        '配分率の小数点以下の数値の上限値を1桁繰り上げる（10倍する）
        jUpper = jUpper * 10
        
        '小数点以下の数値の下限値から上限値まで1刻みで計算を繰り返す
        For j = jLower To jUpper Step 1
                   
            'シートに配分率の計算結果を出力する（指定された配当率セルへ入力）
            rngRate.Value = j / (10 ^ i)
            
            'シートから合計配当金額の計算値を取得する（指定された合計値セルから取得）
            lngResult = rngTotal.Value
            
            '合計配当金額の計算値が目標値よりも小さい場合は、その小数点以下の数値を下限値とする
            If lngResult < lngTarget Then
                '下限値を現在のjに更新する
                jLower = j
            'If文の終了
            End If
            
            '合計配当金額の計算値が目標値よりも大きい場合は、上限値を設定してループを抜ける
            If lngResult > lngTarget Then
                '上限値を現在のjに更新する
                jUpper = j
                'ループを抜ける
                Exit For
            'If文の終了
            End If
            
            '合計配当金額の計算値が目標値と等しい場合は、繰り返しを終了する
            If lngResult = lngTarget Then
                'ループを抜ける
                Exit For
            'If文の終了
            End If
        '次のjへ
        Next j
    
        '合計配当金額の計算値が目標値と等しい場合は、外側のループも終了する
        If lngResult = lngTarget Then
            'ループを抜ける
            Exit For
        'If文の終了
        End If
    '次の桁数へ
    Next i
    
    '画面の更新を再開する
    Application.ScreenUpdating = True
       
    '計算結果が目標値と一致するか確認する
    If lngResult <> lngTarget Then
        '見つからなかった場合、配当率を0に戻す
        rngRate.Value = 0
    'If文の終了
    End If
    
    '計算後の配当率を関数の戻り値として返す
    CalcDivRate = rngRate.Value
    
    Exit Function

ErrorHandler:
    '画面の更新を再開する
    Application.ScreenUpdating = True
    'エラー内容を表示する
    MsgBox "計算処理中にエラーが発生しました: " & Err.Description, vbCritical
    'エラー時は0を返す
    CalcDivRate = 0
End Function
