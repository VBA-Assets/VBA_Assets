Option Explicit
'=================================================================================
' モジュール名: Connect
' 概要: コネクタ（結合線）を描画するためのメイン処理を行うモジュールです。
'       ユーザーフォームの表示や、実際に図形を作成・加工する機能を含みます。
'=================================================================================

'---------------------------------------------------------------------------------
' プロシージャ名: ShowDrawConnect
' 概要: 設定画面（ユーザーフォーム）を表示します。
'---------------------------------------------------------------------------------
Sub ShowDrawConnect()
    'DrawConnectという名前のユーザーフォームを画面に表示する
    DrawConnect.Show
End Sub

'---------------------------------------------------------------------------------
' プロシージャ名: ShowDrawConnectRibbon
' 概要: リボン上のボタンがクリックされた時に呼び出されるコールバック関数です。
' 引数: control (IRibbonControl) - クリックされたボタンの情報
'---------------------------------------------------------------------------------
Sub ShowDrawConnectRibbon(control As IRibbonControl)
    '既存のフォーム表示プロシージャを呼び出す
    ShowDrawConnect
End Sub

'---------------------------------------------------------------------------------
' プロシージャ名: AddConector
' 概要: 指定された条件に基づいて、シート上にコネクタ（線）を追加します。
' 引数:
'   rngBegin      : 始点となるセル範囲
'   rngEnd        : 終点となるセル範囲
'   strType       : 線の種類（Normal, Dash, Heavy, Double）
'   strDirection  : 線の進行方向（Right, Down）
'   strStart      : 始点の位置（Edge=端, Center=中央）
'   blnAlignElbow : カギ線の折れ曲がり位置を調整するかどうか（True/False）
'---------------------------------------------------------------------------------
Sub AddConector(rngBegin, rngEnd, strType, strDirection, strStart, blnAlignElbow)
    
    '始点および終点のX座標（横）とY座標（縦）を格納する変数を宣言する
    Dim BeginX, BeginY, EndX, EndY
    
    '図形の幅と高さを格納する変数を宣言する
    Dim XLength, YLength
    
    '図形の退避距離（回転処理時の位置ズレ防止用）を格納する変数を宣言する
    Dim EscapeLength
    
    '作成した図形オブジェクトを操作するために格納する変数を宣言する
    Dim objShape

    '---------------------------------------------------------
    ' 1. 座標の計算と図形の追加処理
    '---------------------------------------------------------
    
    '進行方向が「右方向」の場合の処理
    If strDirection = "Right" Then
        
        '始点の座標を設定するための処理（Withを使ってrngBeginの記述を省略）
        With rngBegin
            '始点が「外周（セルの右端）」指定の場合
            If strStart = "Edge" Then
                'セルの左端座標(.Left)に幅(.Width)を足して、右端の座標を算出する
                BeginX = .Left + .Width
            
            '始点が「中央」指定の場合
            ElseIf strStart = "Center" Then
                'セルの左端座標(.Left)に幅の半分(.Width / 2)を足して、中央の座標を算出する
                BeginX = .Left + .Width / 2
            End If
            
            'Y座標（縦）は常にセルの高さの中央にする（上端 + 高さの半分）
            BeginY = .Top + .Height / 2
        End With
        
        '終点の座標を設定するための処理
        With rngEnd
            'X座標はセルの左端にする
            EndX = .Left
            'Y座標はセルの高さの中央にする
            EndY = .Top + .Height / 2
        End With
    
        '始点と終点のY座標が同じ（水平）かどうかを判定する
        If BeginY = EndY Then
            '水平の場合は、直線（msoConnectorStraight）を追加する
            Set objShape = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, BeginX, BeginY, EndX, EndY)
            
        '水平でない（高さが違う）場合
        Else
            'カギ線の描画に必要な幅（XLength）と高さ（YLength）を計算する
            XLength = EndX - BeginX
            YLength = EndY - BeginY
            
            'カギ線（msoConnectorElbow）を追加する
            Set objShape = ActiveSheet.Shapes.AddConnector(msoConnectorElbow, BeginX, BeginY, EndX, EndY)
    
            '「カギ線の角の位置を合わせる」が有効になっている場合の処理
            If blnAlignElbow = True Then
                '幅が0ではなく、かつ幅が終点セルの1つ左のセルの幅より大きい場合
                '（近すぎる場合に調整すると形が崩れるのを防ぐチェック）
                If XLength <> 0 And XLength > rngEnd.Cells(1, 1).Offset(0, -1).Width Then
                    'カギ線の折れ曲がり位置（黄色いハンドル）を調整して、終点の手前の列で曲がるようにする
                    'Adjustments.Item(1) は比率(0.0～1.0)で指定するため、計算で求める
                    objShape.Adjustments.Item(1) = (XLength - rngEnd.Cells(1, 1).Offset(0, -1).Width) / XLength
                End If
            End If
        End If
    
    '進行方向が「下方向」の場合の処理
    '※右方向と同じように単純に追加すると、カギ線の形状が意図した通りにならないため、回転処理を行います。
    ElseIf strDirection = "Down" Then
        
        '始点の座標を設定する
        With rngBegin
            'X座標はセルの幅の中央にする
            BeginX = .Left + .Width / 2
            
            '始点が「外周（セルの下端）」指定の場合
            If strStart = "Edge" Then
                'セルの上端(.Top)に高さ(.Height)を足して、下端の座標を算出する
                BeginY = .Top + .Height
            
            '始点が「中央」指定の場合
            ElseIf strStart = "Center" Then
                'セルの上端(.Top)に高さの半分(.Height / 2)を足して、中央の座標を算出する
                BeginY = .Top + .Height / 2
            End If
        End With
        
        '終点の座標を設定する
        With rngEnd
            'X座標はセルの幅の中央にする
            EndX = .Left + .Width / 2
            'Y座標はセルの上端にする
            EndY = .Top
        End With
        
        '始点と終点のX座標が同じ（垂直）かどうかを判定する
        If BeginX = EndX Then
            '垂直の場合は、直線（msoConnectorStraight）を追加する
            Set objShape = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, BeginX, BeginY, EndX, EndY)
        
        '垂直でない（左右にズレている）場合
        Else
            'カギ線の幅と高さを計算する
            XLength = EndX - BeginX
            YLength = EndY - BeginY
            
            '図形の退避距離を計算する（回転させたときに画面外にはみ出してエラーになるのを防ぐため）
            EscapeLength = Abs(EndX - BeginX) + Abs(EndY - BeginY)
    
            '一度、高さと幅を入れ替えた状態でカギ線を作成する
            'このとき、退避距離(EscapeLength)分だけ座標をずらして作成する
            Set objShape = ActiveSheet.Shapes.AddConnector(msoConnectorElbow, _
                BeginX + EscapeLength, BeginY + EscapeLength, _
                BeginX + YLength + EscapeLength, BeginY + XLength + EscapeLength)
            
            '作成したカギ線を上下反転させる
            objShape.Flip msoFlipVertical
            
            'カギ線を時計回りに90度回転させる
            objShape.IncrementRotation 90
            
            'カギ線の位置を本来あるべき場所に移動させる
            '回転によって生じたズレと、退避距離分のズレを補正する計算
            objShape.IncrementLeft -YLength / 2 + XLength / 2 - EscapeLength
            objShape.IncrementTop -XLength / 2 + YLength / 2 - EscapeLength
            
            '「カギ線の角の位置を合わせる」が有効になっている場合の処理
            If blnAlignElbow = True Then
                '高さが0ではなく、かつ高さが終点セルの1つ上のセルの高さより大きい場合
                If YLength <> 0 And YLength > rngEnd.Cells(1, 1).Offset(-1, 0).Height Then
                   'カギ線の折れ曲がり位置を調整して、終点の上の行で曲がるようにする
                   objShape.Adjustments.Item(1) = (YLength - rngEnd.Cells(1, 1).Offset(-1, 0).Height) / YLength
                End If
            End If
        End If
    End If
    
    '---------------------------------------------------------
    ' 2. 線のスタイル（見た目）の設定
    '---------------------------------------------------------
    
    '追加した図形の線（Line）プロパティを操作する
    With objShape.Line
        '線の色をテーマカラーの黒（テキスト1）に設定する
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        
        '線の種類が「点線 (Dash)」の場合
        If strType = "Dash" Then
            '太さを設定値（Settingモジュールの定数）にする
            .Weight = sngWeight
            '線のスタイルを点線にする
            .DashStyle = msoLineDash
            
        '線の種類が「太線 (Heavy)」の場合
        ElseIf strType = "Heavy" Then
            '太さを太めの設定値にする
            .Weight = sngHeavyWeight
            
        '線の種類が「二重線 (Double)」の場合
        ElseIf strType = "Double" Then
            '線のスタイルを二重線にする
            .Style = msoLineThinThin
            '太さを太めの設定値にする（細すぎると二重線に見えないため）
            .Weight = sngHeavyWeight
            
        'それ以外（通常の細線）の場合
        Else
            '太さを通常の設定値にする
            .Weight = sngWeight
        End If
    End With
        
    '追加した図形を選択状態にする（ユーザーがすぐに調整できるようにするため）
    objShape.Select

End Sub
