'更新日時: 2026/10/18
'概要: Outlookを利用した電話メモ内容のメール送信処理モジュール
Option Explicit

'--------------------------------------------------
' 概要 : OLEオートメーションを利用して指定行の電話メモをメール送信する
' 引数 : r (Variant) - テーブルの対象行番号
' 戻り値 : なし
'--------------------------------------------------
Sub SendMail(r)
    'SendMailプロシージャでは、OLEオートメーションと呼ばれる機能を設定し、Outlookを使
    'ってメールを送信する準備を整えます。
    'テーブル行番号を引数として受け取ります。

    'エラー処理を設定する
    On Error GoTo HdlErr

    '送信済みの場合は、もう一度送信するかを確かめ、「いいえ」の場合は終了する
    If Range("電話メモ[メール送信日時]")(r).Value <> "" Then
        If MsgBox("既に送信しています（" & _
                Format(Range("電話メモ[メール送信日時]")(r).Value, "yyyy/m/d h:mm") & "）。" & vbCrLf & _
                vbCrLf & _
                "もう一度送信しますか？", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    'Outlookオブジェクトの実行時バインディングを行う
    'OLEオートメーションを利用する方法には、事前バインディングと実行時バインディングの
    '2つがありますが、電話メモには、実行時バインディングを用いています。理由は、事前
    'バインディングの場合は、使われるPCすべてでVBEを使った設定が必要となるからです。
    '実行時バインディングの細部については、VBAエキスパート公式テキストのExcelVBAスタン
    'ダード（2009年版）p185を参照してください。
    Dim objOutlook
    Set objOutlook = CreateObject("Outlook.Application")

    'MailItemオブジェクトを作成する
    'メールを操作するオブジェクトであるMailItemを作成します。OutLookの場合は、これ以外
    'に予定、タスクおよび連絡先のオブジェクトの作成が可能です。
    Dim objMail
    Set objMail = objOutlook.CreateItem(0)

    'メール宛先を把握する
    '「電話メモ」テーブルに入力されている「宛先」の値を読み取ります。
    Dim strMailTo
    strMailTo = Range("電話メモ[宛先]")(r).Value

    'メールアドレスを把握する
    'シートの特定には、シート名である「オプション」ではなく、コードネームである
    '「Sheet2」を用います。コードネームを使用することにより、ワークシートの名前が変更
    'になっても、マクロを変更する必要がなくなります。
    With shtOpt
        '「所員別設定」テーブルから、その所員のメールアドレスを取得します。
        Dim strMailAddress
        strMailAddress = WorksheetFunction.XLookup(strMailTo, .Range("所員別設定[所員名]"), .Range("所員別設定[メールアドレス]"))
    End With

    'メールアドレスがなかった場合は終了する
    'テーブルにメールアドレスが入力されていない場合は、メッセージを表示して終了します。
    If strMailAddress = "" Then
        MsgBox "メールアドレスが設定されていません｡ "
        Exit Sub
    End If
       
    'メール送信を確認し、キャンセルされた場合は終了する。
    If MsgBox("この行の内容を次の宛先にメールで送信します。" & vbCrLf & _
            vbCrLf & _
            strMailTo & "(" & strMailAddress & ")" & vbCrLf & _
            vbCrLf & _
            "よろしいですか？", vbOKCancel) = vbCancel Then
        MsgBox "メールの送信を中止しました｡ "
        Exit Sub
    End If
        
    'メール本文を作成する
    '「メモ」テーブルのターゲット行(r)のデータを使って、メール本文を作成します。
    '「日付」と「時間」は、「日時」のデータからそれぞれFormat関数を用いて表示します。
    Dim strMailBody
    strMailBody = _
        "番　号：" & Range("電話メモ[番号]")(r).Value & vbCrLf & _
        "日　付：" & Format(Range("電話メモ[日時]")(r).Value, "yyyy/m/d") & vbCrLf & _
        "時　間：" & Format(Range("電話メモ[日時]")(r).Value, "h:mm") & vbCrLf & _
        "相手先：" & Range("電話メモ[相手先]")(r).Value & vbCrLf & _
        "電　話：" & Range("電話メモ[電話番号]")(r).Value & vbCrLf & _
        "宛　先：" & Range("電話メモ[宛先]")(r).Value & vbCrLf & _
        "用　件：" & Range("電話メモ[用件]")(r).Value & vbCrLf & _
        "対　応：" & Range("電話メモ[対応]")(r).Value & vbCrLf & _
        "内　容：" & Range("電話メモ[内容]")(r).Value & vbCrLf & _
        "入力者：" & Range("電話メモ[入力者]")(r).Value & vbCrLf
                
    'メール情報を設定する
    objMail.To = strMailAddress       'メール宛先
    objMail.Subject = "電話メモ"    'メール題名
    objMail.Body = strMailBody    'メール本文

    'メールを送信する
    objMail.Send

    'オブジェクトを解放する
    'オブジェクトの解放は、通常は、プロシージャの終了と同時に自動的に行われるので不要
    'なのですが、何らかの理由で解放されなかった場合に備えて、強制的に開放を行っておき
    'ます。
    Set objOutlook = Nothing
    Set objMail = Nothing

    'メール送信完了時の処理を行う
    Range("電話メモ[メール送信日時]")(r).Value = Now                    '送信日時を記録
    MsgBox strMailTo & "へのメール送信が完了しました。"

    'エラー処理を解除する
    On Error GoTo 0
    Exit Sub

HdlErr:
    'エラーメッセージを表示する
    MsgBox "メールが送信できませんでした。", vbExclamation
End Sub

