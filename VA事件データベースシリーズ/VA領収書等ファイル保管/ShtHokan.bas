Option Explicit
'モジュールレベルの変数を宣言する
Dim r As Long, s As Long                 '保管を行うテーブル行・列番号
'Dim brnMoveFile As Boolean               '保管完了フラグ値

'================================================================================
' イベント:  ワークシートがダブルクリックされたとき
' 機能:      ダブルクリックされたセルに応じて、特定の処理を実行します。
' 概要:      ダブルクリックされた列（フィールド）に応じて、データ削除、フォルダや
'            ファイルのオープン、保管フォルダ名/ファイル名の自動入力、ファイル移動
'            などの多様なアクションを実行します。処理後、必要に応じてデータベース
'            に内容を反映します。
'================================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'セルをダブルクリックした際の処理を行います。

    Dim rTmp As Long
    'ターゲットのテーブル行・列番号を計算する
    'ターゲットの行列番号は、それぞれTarget.RowとTarget.columnで求められますが、それはシート上の行列番号であって、
    'テーブル上の行列番号ではありません。
    'それぞれの値をRange(“保管”)の位置を元に修正しています。

    r = Target.Row - Range("保管").Row + 1          'テーブル行番号
    s = Target.Column - Range("保管").Column + 1    'テーブル列番号

    'テーブル行・列番号が入力範囲外の場合は終了する
    If r < 1 Or Range("保管").Rows.Count + 1 < r Then Exit Sub
    If s < 1 Or Range("保管").Columns.Count < s Then Exit Sub

    'イベントの発生を停止する
    Call Evnt.Disable

    'ターゲットの列は、Range(“保管”)の行番号0で列番号がターゲットの列番号のセルの値で特定することができます。
    'この「Range(“保管”)(0, s).Value」で「保管」テーブルのs列のフィールド名が取得できるというのが、
    'テーブルを使うことによる大きなメリットのひとつです。
    '「Select Case s」として、ターゲットの列番号に応じて、処理を分岐させる方法もありますが、将来、列を追加した
    '場合には、マクロが誤作動する可能性があります。
    'フィールド名で指定しておけば、列が追加になっても正常に作動しますし、
    'もし、そのフィールド名自体が変更になった場合も、エラーが生じるだけで、誤作動はしません。

    '特定のフィールドの特定の行番号のセルの値は、「Range(“保管[番号]”).value」のように構造化参照を使って求めることができます。
    'これもテーブルを使うことによる大きなメリットです。
    'テーブルを使わなくても、列に範囲名をつけてやれば、同じように列が追加になっても正常に作動するマクロにすることができますが、
    'その範囲名を管理するのは、かなり面倒です。
    'テーブルなら、それを自動でやってくれます。

    'ファイルやフォルダを規定のアプリケーションで開くのには、セルにハイパーリンクを設定してもよいのですが、
    'ファイルのサイズが大きくなってしまうため、FollowHyperlinkメソッドを使用しています。
    'この方法については、こちらのサイトの記事を参考にさせていただきました。
    '「VBAでファイルを既定のアプリで開く方法」
    'https://excel-ubara.com/excelvba4/EXCEL295.html

    Select Case Range("保管")(0, s).Value
        Case "番号"
            'シートが保護されている場合は処理を行わない
            If Me.ProtectContents Then GoTo Skip
            
            '新しい行の場合
            If r = Range("保管").Rows.Count + 1 Then
                '保管対象ファイルを読み込む
                If MsgBox("読込フォルダ内のファイル名を読み込みます。", vbOKCancel) = vbOK Then
                    'ファイル名を読み込む
                    Call ReadFileNames
                End If
            End If

            '処理を終了する
            GoTo Skip
        Case "読込フォルダ名"
            '読込フォルダ名が入力されている場合のみ処理を行う
            If Range("保管[読込フォルダ名]")(r).Value <> "" Then
                '読込フォルダが存在する場合のみ処理を行う
                If Dir(Range("保管[読込フォルダ名]")(r).Value, vbDirectory) <> "" Then
                    '読込フォルダを開く
                    On Error Resume Next
                    ThisWorkbook.FollowHyperlink Range("保管[読込フォルダ名]")(r).Value
                    On Error GoTo 0
                Else
                    MsgBox "フォルダが見つかりませんでした。"
                End If
            End If

            '処理を終了する
            GoTo Skip
        Case "読込ファイル名"
            '読込フォルダ名が入力されている場合のみ処理を行う
            If Range("保管[読込フォルダ名]")(r).Value <> "" Then

                '読込ファイル名が入力されている場合のみ処理を行う
                If Range("保管[読込ファイル名]")(r).Value <> "" Then
                    '読込ファイルが存在する場合のみ処理を行う
                    If Dir(Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value) <> "" Then
                        '読込ファイルを開く
                        On Error Resume Next
                        ThisWorkbook.FollowHyperlink _
                            Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value
                        On Error GoTo 0
                    Else
                        MsgBox "ファイルが見つかりませんでした。"
                    End If
                End If
            End If

            '処理を終了する
            GoTo Skip
        Case "保管フォルダ名"
            '保管フォルダ名の入力の有無に応じて処理を行う
            If Range("保管[保管フォルダ名]")(r).Value <> "" Then
                '保管フォルダが存在する場合のみ処理を行う
                If Dir(Range("保管[保管フォルダ名]")(r).Value, vbDirectory) <> "" Then
                    '保管フォルダを開く
                    On Error Resume Next
                    ThisWorkbook.FollowHyperlink Range("保管[保管フォルダ名]")(r).Value
                    On Error GoTo 0
                Else
                    MsgBox "フォルダが見つかりませんでした。"
                End If

                '処理を終了する
                GoTo Skip
            Else
                'シートが保護されている場合は処理を行わない
                If Me.ProtectContents Then GoTo Skip
                
                '保管フォルダ名を入力する
                With ShtOption
                    If WorksheetFunction.CountIf(.Range("保管フォルダ[種別]"), Range("保管[種別]")(r).Value) > 0 Then
                        rTmp = WorksheetFunction.Match(Range("保管[種別]")(r).Value, .Range("保管フォルダ[種別]"), 0)
                    End If

                    Range("保管[保管フォルダ名]")(r).Value = .Range("保管フォルダ[保管フォルダ]")(rTmp).Value
                End With
            End If

        Case "保管ファイル名"
            '読込フォルダ名が入力されている場合のみ処理を行う
            If Range("保管[読込フォルダ名]")(r).Value <> "" Then

                '保管ファイル名の入力の有無に応じて処理を行う
                If Range("保管[保管ファイル名]")(r).Value <> "" Then
                    '保管ファイルが存在する場合のみ処理を行う
                    If Dir(Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value) <> "" Then
                        '保管ファイルを開く
                        On Error Resume Next
                        ThisWorkbook.FollowHyperlink _
                            Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value
                        On Error GoTo 0
                    Else
                        MsgBox "ファイルが見つかりませんでした。"
                    End If

                    '処理を終了する
                    GoTo Skip
                Else
                    'シートが保護されている場合は処理を行わない
                    If Me.ProtectContents Then GoTo Skip
                    
                    '読み込みファイル名に拡張子が記載されている場合
                    If InStr(Range("保管[読込ファイル名]")(r).Value, ".") > 0 Then
                        '枝番の値に応じて処理を行う
                        If Range("保管[枝番]")(r).Value = 0 Then
                            '枝番を付加しない保管ファイル名を入力する
                            Range("保管[保管ファイル名]")(r).Value = Range("保管[日時]")(r).Value & _
                                "_" & Range("保管[相手先]")(r).Value & _
                                "_" & Range("保管[金額]")(r).Value & _
                                "." & GetExtension(Range("保管[読込ファイル名]")(r).Value)
                        Else
                            '枝番を付加した保管ファイル名を入力する
                            Range("保管[保管ファイル名]")(r).Value = Range("保管[日時]")(r).Value & _
                                "_" & Range("保管[相手先]")(r).Value & _
                                "_" & Range("保管[金額]")(r).Value & _
                                "_" & Range("保管[枝番]")(r).Value & _
                                "." & GetExtension(Range("保管[読込ファイル名]")(r).Value)
                        End If
                    End If
                End If
            End If

        Case "保管日時"
            'ファイルを移動し、保管日時を入力する

            'シートが保護されている場合は処理を行わない
            If Me.ProtectContents Then GoTo Skip
            
            '保管フォルダ名または保管ファイル名が入力されていない場合は終了する
            If Range("保管[保管フォルダ名]")(r).Value = "" Or _
                Range("保管[保管ファイル名]")(r).Value = "" Then GoTo Skip

            '保管日時の入力状態に応じて処理を行う
            If Range("保管[保管日時]")(r).Value = "" Then
                'ファイルをリネームして移動する
                Call MoveFile
            Else
                'ファイルをリネームして元のフォルダに移動する
                Call RstrFile
            End If

            'テーブルの書式を設定する
            Call FormatTable
    End Select

Skip:
    Cancel = True

    'イベントの発生を再開する
    Call Evnt.Enable
End Sub

'================================================================================
' イベント:  セルの値が変更されたとき
' 機能:      変更されたセルの内容に応じて、関連するフィールドを自動更新します。
' 概要:      「種別」が変更された場合は「保管フォルダ名」を、日時、相手先、金額、
'            枝番が変更された場合は「枝番」と「保管ファイル名」を再計算して
'            自動入力します。処理後、データベースフォルダが設定されていれば
'            TSVファイルにレコードを書き込みます。
'================================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    'セルの値が変更された場合の処理を行います。

    Dim t As Long, rTmp As Long
    Dim strContent As String, lngFileCnt As Long

    'ターゲットのテーブル行・列番号を計算する
    'このプロシージャにおいても、まず、ターゲットの行列番号をテーブルの行列番号に変換し、
    'テーブルの列見出しに応じて、処理を行っています。
    '「テーブルの上や左にテーブル以外の行を作らない」ということを守れるのであれば、
    '「見出し行の1行分がずれている」という考え方でも問題ありません。
    'とにかくシートの行列番号とテーブルの行列番号は､違うということだけはしっかり認識しておかないと､
    'ずれたデータを参照してしまう可能性があるので気を付けてください｡

    r = Target.Row - Range("保管").Row + 1          'テーブル行番号
    s = Target.Column - Range("保管").Column + 1    'テーブル列番号

    'テーブル行・列番号が入力範囲外の場合は終了する
    If r < 1 Or Range("保管").Rows.Count < r Then Exit Sub
    If s < 1 Or Range("保管").Columns.Count < s Then Exit Sub

    'イベントの発生を停止する
    Call Evnt.Disable

    Select Case Range("保管")(0, s).Value
        Case "種別"
            '保管フォルダ名を入力する
            With ShtOption
                If WorksheetFunction.CountIf(.Range("保管フォルダ[種別]"), Range("保管[種別]")(r).Value) > 0 Then
                    rTmp = WorksheetFunction.Match(Range("保管[種別]")(r).Value, .Range("保管フォルダ[種別]"), 0)
                End If

                Range("保管[保管フォルダ名]")(r).Value = .Range("保管フォルダ[保管フォルダ]")(rTmp).Value
            End With

        Case "日時", "相手先", "金額"
            '枝番を入力する
            '重複するファイルの数をカウントする
            strContent = Range("保管[日時]")(r).Value & _
                    Range("保管[相手先]")(r).Value & _
                    Range("保管[金額]")(r).Value
            For rTmp = 1 To r
                If Range("保管[日時]")(rTmp).Value & _
                    Range("保管[相手先]")(rTmp).Value & _
                    Range("保管[金額]")(rTmp).Value = strContent Then
                lngFileCnt = lngFileCnt + 1
                End If
            Next rTmp

            'オプションの枝番の設定が「常に付加する」以外の場合はファイル数の開始を0に変更する
            If ShtOption.Range("枝番").Value <> "常に付加する" Then lngFileCnt = lngFileCnt - 1

            'ファイル数を枝番に入力する
            Range("保管[枝番]")(r).Value = lngFileCnt

            '読み込みファイル名に拡張子が記載されている場合
            If InStr(Range("保管[読込ファイル名]")(r).Value, ".") > 0 Then
                '枝番の値に応じて処理を行う
                If Range("保管[枝番]")(r).Value = 0 Then
                    '枝番を付加しない保管ファイル名を入力する
                    Range("保管[保管ファイル名]")(r).Value = Range("保管[日時]")(r).Value & _
                        "_" & Range("保管[相手先]")(r).Value & _
                        "_" & Range("保管[金額]")(r).Value & _
                        "." & GetExtension(Range("保管[読込ファイル名]")(r).Value)
                Else
                    '枝番を付加した保管ファイル名を入力する
                    Range("保管[保管ファイル名]")(r).Value = Range("保管[日時]")(r).Value & _
                        "_" & Range("保管[相手先]")(r).Value & _
                        "_" & Range("保管[金額]")(r).Value & _
                        "_" & Range("保管[枝番]")(r).Value & _
                        "." & GetExtension(Range("保管[読込ファイル名]")(r).Value)
                End If
            End If
        Case "枝番"
            '読み込みファイル名に拡張子が記載されている場合
            If InStr(Range("保管[読込ファイル名]")(r).Value, ".") > 0 Then
                '枝番の値に応じて処理を行う
                If Range("保管[枝番]")(r).Value = 0 Then
                    '枝番を付加しない保管ファイル名を入力する
                    Range("保管[保管ファイル名]")(r).Value = Range("保管[日時]")(r).Value & _
                        "_" & Range("保管[相手先]")(r).Value & _
                        "_" & Range("保管[金額]")(r).Value & _
                        "." & GetExtension(Range("保管[読込ファイル名]")(r).Value)
                Else
                    '枝番を付加した保管ファイル名を入力する
                    Range("保管[保管ファイル名]")(r).Value = Range("保管[日時]")(r).Value & _
                        "_" & Range("保管[相手先]")(r).Value & _
                        "_" & Range("保管[金額]")(r).Value & _
                        "_" & Range("保管[枝番]")(r).Value & _
                        "." & GetExtension(Range("保管[読込ファイル名]")(r).Value)
                End If
            End If
    End Select

    'イベントの発生を再開する
    Call Evnt.Enable
End Sub

'================================================================================
' 機能:      各フォルダのファイルの一覧を取得し、テーブルに追加します。
' 概要:      オプションシートの「読込フォルダ」に設定されたフォルダから、Dir関数を
'            使ってファイル名を取得します。ファイル名から「日時」「相手先」
'            「金額」を抽出し、既存のレコードとの重複をチェックして「枝番」を
'            自動計算し、テーブルの新しい行にデータを追加します。
'================================================================================
Sub ReadFileNames()
    '各フォルダのファイルの一覧を取得します。
    'Dir関数を使う方法とFileSystemObjectを使う方法があります。
    'Dir関数を使う方法は、「引数を省略して再度実行すると次のファイルが読みだされる」というのが何となく裏技っぽくって
    '好きになれないのですが、私の環境の場合はFileSystemObjectを使うよりもかなり高速に処理ができましたので、
    'こちらを採用しています。

    'ファイルの一覧を取得する方法の細部については､こちらをご覧ください｡
    'http://officetanaka.net/excel/vba/file/file07.htm

    Dim x As Integer, v As Long, rTmp As Long
    Dim strFolder As String, strFolders As Variant, strFile As String
    Dim lngReadCnt As Long
    Dim strContent As String, lngFileCnt As Long

    Application.StatusBar = "保管するファイルのファイル名を取得しています。"

    '入力するテーブル行を取得する
    If Range("保管")(1, 1).Value = "" Then
        r = 1
    Else
        r = Range("保管").Rows.Count + 1
    End If

    '読み込みフォルダを取得する
    strFolders = Split(ShtOption.Range("読込フォルダ").Value, vbLf)

    For x = 0 To UBound(strFolders)
        '読み込みフォルダを選択する
        strFolder = Trim(Replace(strFolders(x), vbCr, ""))

        '空の行（最後の改行など）は飛ばす
        If strFolder = "" Then GoTo NextFolder

        Application.StatusBar = strFolder & "からファイル名を読み込んでいます。"

        '読み込みフォルダの存在を確認する
        If Dir(strFolder, vbDirectory) = "" Then
            MsgBox "読み込みファイルのフォルダが見つかりません。", vbExclamation
            Exit Sub
        End If

        'エラー処理を行う
        On Error GoTo errInvalidFileName

        '読み込みファイルフォルダからファイル名を読み込む
        strFile = Dir(strFolder & "\" & "*.*")
        Do While strFile <> ""
            '枝番を常に付加する場合は、保管済みのファイルは、処理しない。
            '読込フォルダと保管フォルダを同一にできるようにするためです。
            If ShtOption.Range("枝番").Value = "常に付加する" Then
                If WorksheetFunction.CountIf(Range("保管[保管ファイル名]"), strFile) > 0 Then GoTo Cont
            End If

            '同じ名前のファイルが既に読み込まれている場合は、処理しない
            If WorksheetFunction.CountIf(Range("保管[読込ファイル名]"), strFile) > 0 Then _
                GoTo Cont

            '番号を入力する
            Range("保管[番号]")(r).Value = WorksheetFunction.Max(Range("保管[番号]")) + 1

            '読込フォルダ名を入力する
            Range("保管[読込フォルダ名]")(r).Value = strFolder

            '読込ファイル名を入力する
            Range("保管[読込ファイル名]")(r).Value = strFile

            '種別にデフォルト値を入力する
            Range("保管[種別]")(r).Value = ShtOption.Range("保管フォルダ[種別]")(1).Value

            '日時を入力する
            Range("保管[日時]")(r).Value = Left(Range("保管[読込ファイル名]")(r).Value, 8)

            '相手先を入力する
            If InStr(Range("保管[読込ファイル名]")(r).Value, "_") > 0 And _
                InStr(Range("保管[読込ファイル名]")(r).Value, ".") > 0 Then
                    Range("保管[相手先]")(r).Value = Split(GetBaseName(Range("保管[読込ファイル名]")(r).Value), "_")(1)
            End If

            '金額を入力する
            If UBound(Split(Range("保管[読込ファイル名]")(r).Value, "_")) > 1 And _
                InStr(Range("保管[読込ファイル名]")(r).Value, ".") > 0 Then
                    Range("保管[金額]")(r).Value = Split(GetBaseName(Range("保管[読込ファイル名]")(r).Value), "_")(2)
            End If

            '枝番を入力する
            'ファイル数をリセットする
            lngFileCnt = 0

            '重複するファイルの数をカウントする
            strContent = Range("保管[日時]")(r).Value & _
                    Range("保管[相手先]")(r).Value & _
                    Range("保管[金額]")(r).Value
            For rTmp = 1 To r
                If Range("保管[日時]")(rTmp).Value & _
                    Range("保管[相手先]")(rTmp).Value & _
                    Range("保管[金額]")(rTmp).Value = strContent Then
                lngFileCnt = lngFileCnt + 1
                End If
            Next rTmp

            'オプションの枝番の設定が「常に付加する」以外の場合はファイル数の開始を0に変更する
            If ShtOption.Range("枝番").Value <> "常に付加する" Then lngFileCnt = lngFileCnt - 1

            'ファイル数を枝番に入力する
            Range("保管[枝番]")(r).Value = lngFileCnt


            '保管フォルダ名を入力する
            With ShtOption
                If WorksheetFunction.CountIf(.Range("保管フォルダ[種別]"), Range("保管[種別]")(r).Value) > 0 Then
                    rTmp = WorksheetFunction.Match(Range("保管[種別]")(r).Value, .Range("保管フォルダ[種別]"), 0)
                End If

                Range("保管[保管フォルダ名]")(r).Value = .Range("保管フォルダ[保管フォルダ]")(rTmp).Value
            End With

            '読み込みファイル名に拡張子が記載されている場合
            If InStr(Range("保管[読込ファイル名]")(r).Value, ".") > 0 Then
                '枝番の値に応じて処理を行う
                If Range("保管[枝番]")(r).Value = 0 Then
                    '枝番を付加しない保管ファイル名を入力する
                    Range("保管[保管ファイル名]")(r).Value = Range("保管[日時]")(r).Value & _
                        "_" & Range("保管[相手先]")(r).Value & _
                        "_" & Range("保管[金額]")(r).Value & _
                        "." & GetExtension(Range("保管[読込ファイル名]")(r).Value)
                Else
                    '枝番を付加した保管ファイル名を入力する
                    Range("保管[保管ファイル名]")(r).Value = Range("保管[日時]")(r).Value & _
                        "_" & Range("保管[相手先]")(r).Value & _
                        "_" & Range("保管[金額]")(r).Value & _
                        "_" & Range("保管[枝番]")(r).Value & _
                        "." & GetExtension(Range("保管[読込ファイル名]")(r).Value)
                End If
            End If
            
            '読み込みファイル数をカウントする
            lngReadCnt = lngReadCnt + 1

            '次のテーブル行に移動する
            r = r + 1
Cont:
            '次のファイルを読み込む
            strFile = Dir()
        Loop

        'エラー処理を解除する
        On Error GoTo 0
NextFolder:
    Next

    '書式を設定する
    Call FormatTable

    '読み込み結果のメッセージを表示
    If lngReadCnt = 0 Then
        MsgBox "読み込まれていないファイルが見つかりませんでした。"
    Else
        MsgBox lngReadCnt & "個のファイルを読み込みました。"
    End If

    Application.StatusBar = ""

    Exit Sub

errInvalidFileName:
    MsgBox "ファイルの読み込みが異常終了しました。" & vbCrLf & _
        "読み込みフォルダ内に不正なファイル名が存在する可能性があります。", vbExclamation

    Application.StatusBar = ""

End Sub

'================================================================================
' 機能:      読み込みフォルダから保管フォルダへファイルを移動します。
' 概要:      Nameステートメントを使用して、現在の行の「読込フォルダ名・ファイル名」
'            から「保管フォルダ名・ファイル名」へファイルをリネームして移動します。
'            移動に成功した場合、「保管日時」を現在時刻で更新します。
'================================================================================
Sub MoveFile()
    'ファイルの移動を行います。
    'Nameステートメントを使ってフォルダの名前を変えてやれば、移動したことと同じになります。

    'ファイルの移動など､ファイルの操作全般については､こちらを参考にしてください｡
    'http://officetanaka.net/excel/vba/tips/tips91.htm

    Dim strReadedFile As String, strStoredFile As String
    Dim strErrMsg As String

    'コピー元とコピー先を確認する
    Application.StatusBar = "ファイルのコピー元と移動先を確認しています"

    'コピー元ファイルが存在することを確認し、コピー元ファイルを取得する
    If Dir(Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value) <> "" Then
        strReadedFile = Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value
    Else
        strErrMsg = Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value & "が見つかりません。"
        GoTo HdlPrelimErr
    End If

    '保管フォルダ名が存在することを確認し、移動後ファイルを設定する
    If Dir(Range("保管[保管フォルダ名]")(r).Value, vbDirectory) <> "" Then
        strStoredFile = Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value
    Else
        strErrMsg = Range("保管[保管フォルダ名]")(r).Value & "が見つかりません。"
        GoTo HdlPrelimErr
    End If

    'ファイルを保管フォルダに移動する。
    Application.StatusBar = "ファイルを保管フォルダに移動しています。"

    '同じ名前のファイルの有無に応じて処理を行う
    If Dir(strStoredFile) = "" Then
        'ファイル名を変更して保管フォルダに移動する
        On Error GoTo HdlWriteErr
        Name strReadedFile As strStoredFile
        On Error GoTo 0

        '保管日時を記録する
        Range("保管[保管日時]")(r).Value = Now

    Else
        strErrMsg = strStoredFile & "と同じ名前のファイルが既にあります。"
        GoTo HdlPrelimErr
    End If

    Application.StatusBar = ""
    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & "このファイルの移動を中止しました。", vbExclamation
    Application.StatusBar = ""
    Exit Sub

HdlWriteErr:
    MsgBox "このファイルを移動できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイルが別のアプロケーションで開かれている。" & vbCr & _
        "・ファイル名に使うことのできない文字が含まれている。" _
        , vbExclamation
    Application.StatusBar = ""
    Exit Sub

End Sub

'================================================================================
' 機能:      保管フォルダから元の読み込みフォルダへファイルを戻します。
' 概要:      MoveFileプロシージャの逆の処理です。Nameステートメントを使用して、
'            現在の行の「保管フォルダ名・ファイル名」から「読込フォルダ名・ファイル名」
'            へファイルをリネームして移動します。成功した場合、「保管日時」、
'            「保管フォルダ名」、「保管ファイル名」をクリアします。
'================================================================================
Sub RstrFile()
    '一旦移動したファイルを元に戻します。
    '「MoveFile」プロシージャを逆にしただけですが、この機能があるのとないのとでは、使い勝手が全然違います。

    Dim strReadedFile As String, strStoredFile As String
    Dim strErrMsg As String

    'コピー元とコピー先を確認する
    Application.StatusBar = "ファイルのコピー元と移動先を確認しています"

    '保管ファイルが存在することを確認し、保管ファイル名を取得する
    If Dir(Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value) <> "" Then
        strReadedFile = Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value
    Else
        strErrMsg = Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value & "が見つかりません。"
        GoTo HdlPrelimErr
    End If

    '読込フォルダ名が存在することを確認し、読込ファイルを取得する
    If Dir(Range("保管[読込フォルダ名]")(r).Value, vbDirectory) <> "" Then
        strStoredFile = Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value
    Else
        strErrMsg = Range("保管[読込フォルダ名]")(r).Value & "が見つかりません。"
        GoTo HdlPrelimErr
    End If

    '読込フォルダに移動する
    Application.StatusBar = "ファイルを読込フォルダに移動しています。"

    '同じ名前のファイルの有無に応じて処理を行う
    If Dir(strStoredFile) = "" Then
        'ファイル名を変更して読込フォルダに移動する
        On Error GoTo HdlWriteErr
        Name strReadedFile As strStoredFile
        On Error GoTo 0

        '保管日時を削除する
        Range("保管[保管日時]")(r).Value = ""

    Else
        '処理を中止する
        strErrMsg = strStoredFile & "と同じ名前のファイルが既にあります。"
        GoTo HdlPrelimErr
    End If

    '保管フォルダ名、保管ファイル名および複製作成先を削除する
    Range("保管[保管フォルダ名]")(r).Value = ""
    Range("保管[保管ファイル名]")(r).Value = ""

    Application.StatusBar = ""

    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & "このファイルの元のフォルダへの移動を中止しました。", vbExclamation
    Application.StatusBar = ""
    Exit Sub

HdlWriteErr:
    MsgBox "このファイルを元のフォルダに移動できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイルが別のアプロケーションで開かれている。" & vbCr & _
        "・ファイル名に使うことのできない文字が含まれている。" _
        , vbExclamation
    Application.StatusBar = ""
    Exit Sub

End Sub

'================================================================================
' 機能:      テーブルの表示形式（フォーマット）を設定します。
' 概要:      「保管日時」「番号」「日時」の表示形式を設定し、
'            「種別」列にオプションシートから取得した値を基にドロップダウンリスト
'            を設定します。「保管日時」が未入力の行は自動色、入力済みの行は
'            テーマカラーを設定します。
'================================================================================
Sub FormatTable()
    'テーブルのフォーマットを設定します。

    Dim varItems As Variant
    Dim rTmp As Long

    '日時の表示形式を設定する
    Range("保管[保管日時]").NumberFormatLocal = "yyyy/m/d h:mm"

    'リストを設定する
    ''リストアイテムが変更された場合などに対応するため、すべてのリストを再設定するようにしています。
    ''細部は、https://valmore.work/vba-create-dropdown-list/を参照してください。
    varItems = WorksheetFunction.Transpose(ShtOption.Range("保管フォルダ[種別]"))
    With Range("保管[種別]").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Join(varItems, ",")
    End With

    '文字色を設定する
    With Range("保管")
        For rTmp = 1 To .Rows.Count
            With .Rows(rTmp)
                Select Case Range("保管[保管日時]")(rTmp).Value
                    Case ""
                        .Font.ColorIndex = xlAutomatic
                    Case Else
                        .Font.ThemeColor = xlThemeColorAccent3
                End Select
            End With
        Next
    End With

    'ワークシート全体を対象
    With Cells
        'セル内の文字列を折り返して全体を表示する
        .WrapText = True

        'セルの高さを調節する
        .Rows.AutoFit
    End With

    '表示形式を設定する
    Range("保管[番号]").NumberFormatLocal = "0"
    Range("保管[日時]").NumberFormatLocal = "0"
    Range("保管[保管日時]").NumberFormatLocal = "yyyy/m/d h:mm"
End Sub

'================================================================================
' 機能:      ファイル名の拡張子を返します。
' 概要:      最後の「.」より後ろを返します（「.」が2つ以上あるファイル名に対応）。
'================================================================================
Private Function GetExtension(ByVal strFileName As String) As String
    GetExtension = Mid(strFileName, InStrRev(strFileName, ".") + 1)
End Function

'================================================================================
' 機能:      ファイル名から拡張子を除いた部分を返します。
' 概要:      最後の「.」より前を返します。「.」がない場合は、ファイル名をそのまま返します。
'================================================================================
Private Function GetBaseName(ByVal strFileName As String) As String
    If InStrRev(strFileName, ".") > 0 Then
        GetBaseName = Left(strFileName, InStrRev(strFileName, ".") - 1)
    Else
        GetBaseName = strFileName
    End If
End Function
