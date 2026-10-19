'更新日時: 2026/10/18
'概要: 「電話メモ」シートイベントおよび書式設定等の制御
Option Explicit

'--------------------------------------------------
' 概要 : ワークシートがアクティブになった場合に処理を行う
' 引数 : なし
' 戻り値 : なし
'--------------------------------------------------
Private Sub Worksheet_Activate()
    '「電話メモ」テーブルを対象
    With Range("電話メモ")
        '編集行の塗りつぶしを解除する
        With .Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With
End Sub

'--------------------------------------------------
' 概要 : ワークシートがダブルクリックされた場合に処理を行う
' 引数 : Target (Range) ダブルクリックされたセル, Cancel (Boolean) 既定アクションのキャンセル
' 戻り値 : なし
'--------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'ターゲットのテーブル行・列番号を計算する
    'ターゲットの行列番号は、それぞれTarget.RowとTarget.columnで求められますが、それは
    'シート上の行列番号であって、テーブル上の行列番号ではありません。このため、それぞ
    'れの値をRange(“メモ”)の位置を元に修正しています。
    '「テーブルの上や左にテーブル以外の行を作らない」ということを守れるのであれば、見
    '出し行の1行分がずれていると認識するだけでも大丈夫です。とにかくシートの行列番号と
    'テーブルの行列番号は、違うということだけはしっかり認識しておかないと、ずれたデー
    'タを参照してしまう可能性があります。
    Dim r As Long
    Dim s As Long
    Dim lngNo As Long   '新しい行の番号
    r = Target.Row - Range("電話メモ").Row + 1          'テーブル行番号
    s = Target.Column - Range("電話メモ").Column + 1    'テーブル列番号

    'テーブル行・列番号が入力範囲外の場合は終了する
    'ダブルクリックしたセル（ターゲット）が入力可能な範囲（テーブルのすぐ下の行を含む）
    'にあるかどうかをIf構文で判断します。
    'if 構文の細部については、こちらのサイトをご覧ください。
    'http://excelvba.pc-users.net/fol6/6_1.html
    If r < 1 Or Range("電話メモ").Rows.Count + 1 < r Then Exit Sub
    If s < 1 Or Range("電話メモ").Columns.Count < s Then Exit Sub

    'テーブルのすぐ下の行は、「番号」の場合だけ処理する（ほかの列は、ふつうのセルの編集にする）
    If r = Range("電話メモ").Rows.Count + 1 And Range("電話メモ")(0, s).Value <> "番号" Then Exit Sub
            
    'セルの編集をキャンセルする
    'ダブルクリックの本来の機能であるセルの編集をキャンセルするため、変数CancelにTrueを
    '代入します。これにより、ダブルクリックした後、編集状態にならないまま、マクロを終了
    'させることができます。
    Cancel = True

    'イベントを停止する
    Evnt.Disable

    '「電話メモ」テーブルを対象
    With Range("電話メモ")
        'リストを解除する
        .Validation.Delete
        
        '塗りつぶしを解除する
        With .Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With

    'ターゲット列の値に応じて処理を行う
    'ターゲットの列の値に応じて、処理を変える必要がありますので、Select Caseを用いて分
    '岐処理を行います。
    'Select Caseについて、細部を確認したい方は、こちらをご覧ください。
    'http://excelvba.pc-users.net/fol6/6_2.html
    'ターゲットの列は、範囲（メモ）の行番号0で列番号がターゲットの列番号のセルの値で特
    '定することができます。「Select Case s」として、ターゲットの列番号に応じて、処理を
    '分岐させても良いのですが､将来､列を追加した場合､列番号で指示しているとマクロが誤作
    '動することになります。フィールド名で指定しておけば、列が追加になっても正常に作動
    'しますし、もし、そのフィールド名自体が変更になった場合も、エラーが生じるだけで、
    '誤作動はしません。
    '特定のフィールドの特定の行番号のセルの値は、構造化参照を使って「Range(“電話メモ
    '[番号]”)(r).value」で求めることができます。
    'テーブルを使わなくても、列に範囲名をつけてやれば、同じように列が追加になっても正
    '常に作動するマクロが作れますが、その範囲名を管理するのは、かなり面倒です。テーブ
    'ルなら、それを自動でやってくれます。
    Select Case Range("電話メモ")(0, s).Value
        '「番号」の場合
        Case "番号"
            '新しい行の場合（テーブルのすぐ下の行、またはテーブルの中の空白行）
            If r = Range("電話メモ").Rows.Count + 1 Or _
                WorksheetFunction.CountA(Range("電話メモ").Rows(r)) = 0 Then
                'テーブルのすぐ下の行の場合は、テーブルの最終行が空白ならその行を使い、
                '空白でなければテーブルに行を追加する
                'VBAでテーブルのすぐ下のセルに値を書き込んでも、テーブルは広がらないためです。
                If r = Range("電話メモ").Rows.Count + 1 Then
                    If WorksheetFunction.CountA(Range("電話メモ").Rows(r - 1)) = 0 Then
                        r = r - 1
                    Else
                        ListObjects("電話メモ").ListRows.Add AlwaysInsert:=True
                    End If
                End If
                
                '新しい行を選択する
                Range("電話メモ[番号]")(r).Select
                
                '番号を入力する
                'すべての番号の最大値の次の番号を入力します。
                lngNo = WorksheetFunction.Max(Range("電話メモ[番号]")) + 1
                Range("電話メモ[番号]")(r).Value = lngNo
            
                '入力者名を入力する
                Call InputUserName(r)
                
                '現在の日時を入力する
                Range("電話メモ[日時]")(r).Value = Now
            
                '書式を設定する
                'セルの高さを調節するために必要です。
                Call FormatTable
                
                '並べ替えで移った新しい行の行番号を把握する
                r = WorksheetFunction.Match(lngNo, Range("電話メモ[番号]"), 0)
                
                '相手先を選択する
                Range("電話メモ[相手先]")(r).Select
                
                '画面をスクロールする
                ActiveWindow.ScrollRow = r + 1
            End If
            
        '「日時」の場合
        Case "日時"
            '現在の日時を入力する
            Range("電話メモ[日時]")(r).Value = Now
            
        Case "メール送信日時"
            'メールの送信を行う
            ''SendMailプロシージャに変数rを値渡しでテーブル列番号を渡します。
            Call SendMail(r)
    End Select

    '「電話メモ」テーブルを対象
    With Range("電話メモ")
        'ターゲット行を対象
        With .Rows(r)
            'リストを設定する
            Call FormatRow(r)
        End With
    End With

Skip:
    'イベントを再開する
    Evnt.Enable
End Sub

'--------------------------------------------------
' 概要 : テーブルの書式・ソート設定を行う
' 引数 : dummy (Byte) - マクロの一覧に表示しないためのダミー引数
' 戻り値 : なし
'--------------------------------------------------
Public Sub FormatTable(Optional ByVal dummy As Byte = 0)
    'ステータスバーを表示する
    Application.StatusBar = "「電話メモ」テーブルの書式を設定しています。"

    '並び替えを行う
    With ListObjects("電話メモ").Sort
        .SortFields.Clear

        .SortFields.Add2 _
            Key:=Range("電話メモ[番号]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'シート全体を対象
    With Cells
        'セル内の文字列を折り返して全体を表示する
        .WrapText = True
        
        'セルの高さを調節する
        .Rows.AutoFit
    End With

    '日時の書式を設定する
    Range("電話メモ[日時]").NumberFormatLocal = "yyyy/m/d h:mm"
    Range("電話メモ[メール送信日時]").NumberFormatLocal = "yyyy/m/d h:mm"

    Application.StatusBar = ""
End Sub

'--------------------------------------------------
' 概要 : 指定行に入力規則（ドロップダウンリスト）を設定する
' 引数 : r (Variant) - 行番号
' 戻り値 : なし
'--------------------------------------------------
Private Sub FormatRow(r)
    'リスト設定用配列を定義する
    Dim varLists(1 To 6, 1 To 2)

    '「オプション」シートを対象
    With shtOpt
        Set varLists(1, 1) = Range("電話メモ[入力者]")(r)
        Set varLists(1, 2) = .Range("所員別設定[所員名]")
        Set varLists(2, 1) = Range("電話メモ[用件]")(r)
        Set varLists(2, 2) = .Range("用件")
        Set varLists(3, 1) = Range("電話メモ[対応]")(r)
        Set varLists(3, 2) = .Range("対応")
        Set varLists(4, 1) = Range("電話メモ[宛先]")(r)
        Set varLists(4, 2) = .Range("所員別設定[所員名]")
        Set varLists(5, 1) = Range("電話メモ[確認者]")(r)
        Set varLists(5, 2) = .Range("所員別設定[所員名]")
        Set varLists(6, 1) = Range("電話メモ[処理]")(r)
        Set varLists(6, 2) = .Range("処理")
    End With

    Dim x As Integer
    Dim strFormula As String

    'リストを設定する
    For x = 1 To 6
        '設定項目が1件だけの場合、TransposeとJoinがエラーになるため分岐
        If varLists(x, 2).Cells.Count = 1 Then
            strFormula = varLists(x, 2).Value
        Else
            strFormula = Join(WorksheetFunction.Transpose(varLists(x, 2)), ",")
        End If

        With varLists(x, 1).Validation
            .Delete
            .Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:=strFormula
        End With
    Next x
End Sub

'--------------------------------------------------
' 概要 : Windowsのログイン名から入力者名を自動取得して設定する
' 引数 : r (Variant) - 行番号
' 戻り値 : なし
'--------------------------------------------------
Private Sub InputUserName(r)
    'エラー処理を行う
    On Error GoTo Skip

    'WshNetworkオブジェクトを利用可能にする
    ''参考
    ''https://www.moug.net/tech/exvba/0150087.html
    Dim objUser
    Set objUser = CreateObject("WScript.Network")
       
    '入力者名を入力する
    With shtOpt
        Range("電話メモ[入力者]")(r).Value = _
            WorksheetFunction.XLookup(objUser.UserName, .Range("所員別設定[ユーザー名（Windows）]"), .Range("所員別設定[所員名]"))
    End With

    'オブジェクト変数を解放する
    Set objUser = Nothing
Skip:
End Sub
