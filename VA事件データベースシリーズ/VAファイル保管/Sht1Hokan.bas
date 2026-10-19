Option Explicit

'================================================================================
' モジュール名: Sht1Hokan
' 概要:         「保管」シートに対する操作を制御するクラスモジュールです。
'               ファイル名の読み込み、移動、リネーム、ショートカット作成等の
'               メイン機能を実装しています。
' 更新日時:     2026/02/01 01:41
'================================================================================

'================================================================================
' イベント:  Worksheet_BeforeDoubleClick
' 機能:      ワークシートがダブルクリックされたときの処理
' 概要:      ダブルクリックされたセル（列）に応じて、ファイル読込や移動、
'            フォルダオープンなどのアクションを実行します。
'================================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' ターゲット行/列番号をテーブル行/列番号に変換するための変数を宣言
    Dim r As Long, s As Long
    ' テーブル範囲を取得するための変数を宣言
    Dim rngTable As Range
    Set rngTable = Range("保管")
    
    ' クリック位置の相対行・列番号を算出
    r = Target.Row - rngTable.Row + 1
    s = Target.Column - rngTable.Column + 1

    ' テーブル行・列番号が入力範囲外の場合は終了する
    If r < 1 Or rngTable.Rows.Count + 1 < r Then Exit Sub
    If s < 1 Or rngTable.Columns.Count < s Then Exit Sub

    ' イベントの発生を停止する（無限ループ防止と高速化）
    Evnt.Disable

    ' ターゲットのテーブル列見出しに応じて処理を分岐
    Select Case rngTable(0, s).Value
        Case "番号"
            'シートが保護されている場合は処理を行わない
            If Me.ProtectContents Then GoTo Skip
            
            ' 新しい行（データ末尾の次の行）の場合
            If r = rngTable.Rows.Count + 1 Then
                ' 保管対象ファイルを読み込むか確認
                If MsgBox("読込フォルダ内のファイル名を読み込みます。", vbOKCancel) = vbOK Then
                    ' ファイル名を読み込むプロシージャを呼び出し
                    Call ReadFileNames
                End If
            End If
            ' 処理を終了するためのラベルへジャンプ
            GoTo Skip
            
        Case "読込フォルダ名"
            ' 読込フォルダ名が入力されている場合のみ処理を行う
            If Range("保管[読込フォルダ名]")(r).Value <> "" Then
                ' 読込フォルダが存在する場合のみ処理を行う
                On Error Resume Next
                If Dir(Range("保管[読込フォルダ名]")(r).Value, vbDirectory) <> "" Then
                    ' 読込フォルダをエクスプローラー等で開く
                    ThisWorkbook.FollowHyperlink Range("保管[読込フォルダ名]")(r).Value
                Else
                    MsgBox "フォルダが見つかりませんでした。"
                End If
                On Error GoTo 0
            End If
            ' 処理を終了する
            GoTo Skip

        Case "読込ファイル名"
            ' 読込フォルダ名が入力されている場合のみ処理を行う
            If Range("保管[読込フォルダ名]")(r).Value <> "" Then
                ' 読込ファイル名が入力されている場合のみ処理を行う
                If Range("保管[読込ファイル名]")(r).Value <> "" Then
                    ' 読込ファイルが存在する場合のみ処理を行う
                    On Error Resume Next
                    Dim strFullPath As String
                    strFullPath = Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value
                    If Dir(strFullPath) <> "" Then
                        ' 読込ファイルを開く
                        ThisWorkbook.FollowHyperlink strFullPath
                    Else
                        MsgBox "ファイルが見つかりませんでした。"
                    End If
                    On Error GoTo 0
                End If
            End If
            ' 処理を終了する
            GoTo Skip

        Case "日時"
            'シートが保護されている場合は処理を行わない
            If Me.ProtectContents Then GoTo Skip
            
            ' ファイル名が"20"で始まっている（YYYY...形式）場合
            If Left(Range("保管[読込ファイル名]")(r).Value, 2) = 20 And _
                IsNumeric(Left(Range("保管[読込ファイル名]")(r).Value, 12)) = True Then
                ' 日時データをファイル名から読み込んで入力
                Range("保管[日時]")(r).Value = Left(Range("保管[読込ファイル名]")(r).Value, 12)
            ' ファイル名が日付形式でない場合
            Else
                ' 読み込みファイルのタイムスタンプを取得して入力する
                Dim strReadedFile As String
                Dim datToday As Date
                strReadedFile = Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value
                ' ファイルが存在する場合のみ
                If Dir(strReadedFile) <> "" Then
                    datToday = FileDateTime(strReadedFile)
                    ' 日時列にフォーマットして入力
                    Range("保管[日時]")(r).Value = Format(datToday, "yyyymmddhhnn")
                End If
            End If

        Case "事件呼名"
            'シートが保護されている場合は処理を行わない
            If Me.ProtectContents Then GoTo Skip
           
           ' 事件番号に応じて事件呼名を入力する処理
            With Sht2Hokansaki
                ' 事件呼名が未入力の場合
                If Range("保管[事件呼名]")(r).Value = "" Then
                    ' 「保管先」シートをアクティブにして選択させる
                    .Activate
                End If
            End With
            ' 処理を終了する
            GoTo Skip

        Case "保管フォルダ名"
            '保管フォルダ名が入力されている場合
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
            '保管フォルダ名が入力されていない場合
            Else
                'シートが保護されている場合は処理を行わない
                If Me.ProtectContents Then GoTo Skip
            
                '保管フォルダ名をダイアログで入力する
                With Application.FileDialog(msoFileDialogFolderPicker)
                    .InitialFileName = ShtOption.Range("保管フォルダ")
                    If .Show = True Then
                        Range("保管[保管フォルダ名]")(r).Value = .SelectedItems(1)
                    End If
                End With
            End If

        Case "保管ファイル名"
            ' 保管フォルダ名が入力されている場合
            If Range("保管[保管フォルダ名]")(r).Value <> "" Then
                ' 保管ファイル名が入力済みの場合
                If Range("保管[保管ファイル名]")(r).Value <> "" Then
                    ' 保管ファイルが存在する場合のみ処理を行う
                    If Dir(Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value) <> "" Then
                        ' 保管ファイルを開く
                        On Error Resume Next
                        ThisWorkbook.FollowHyperlink _
                            Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value
                        On Error GoTo 0
                    Else
                        MsgBox "ファイルが見つかりませんでした。"
                    End If
                ' 保管ファイル名が入力未了の場合
                Else
                    'シートが保護されている場合は処理を行わない
                    If Me.ProtectContents Then GoTo Skip
                    
                    ' 保管ファイル名を自動生成して入力する
                    Call InputStoredFile(r)
                    ' 保管ファイル名が入力できた場合のみ
                    If Range("保管[保管ファイル名]")(r).Value = "" Then
                        ' 次の列（複製作成先）を選択する
                        Range("保管[複製作成先]")(r).Select
                    End If
                End If
            End If

        Case "複製作成先"
            'シートが保護されている場合は処理を行わない
            If Me.ProtectContents Then GoTo Skip
            
            ' 事件呼名が入力されている場合
            If Range("保管[事件呼名]")(r).Value <> "" Then
               ' 複製作成先を自動入力する
                Call InputSCFor(r)
            End If

        Case "保管日時"
            'シートが保護されている場合は処理を行わない
            If Me.ProtectContents Then GoTo Skip
            
            ' 「保管フォルダ名」と「保管ファイル名」の両方が入力されている場合
            If Range("保管[保管フォルダ名]")(r).Value <> "" _
                And Range("保管[保管ファイル名]")(r).Value <> "" Then
                ' 保管日時が入力されていない場合の処理（新規保管）を行う
                If Range("保管[保管日時]")(r).Value = "" Then
                    ' ファイルをリネームして移動する
                    Call MoveFile(r)
                    
                    ' 移動が成功し、かつ複製作成先がある場合
                    If brnMoveFile = True And Range("保管[複製作成先]")(r).Value <> "" Then
                         ' ショートカットとコピーを作成する
                         Call MakeShortcuts(r)
                         Call MakeCopies(r)
                    Else
                         ' ショートカット作成不要または移動失敗の場合、フラグをTrueにして続行可能にする
                         brnMakeShortcutsCopies = True
                    End If
                    
                    ' 内容に「職務上請求」が含まれている場合は専用コピーを作成する
                    If InStr(Range("保管[内容]")(r).Value, "職務上請求") <> 0 Then Call MakeCopyOfSpecialFile(r, "職務上請求")
                    
                    ' 内容に「弁護士会照会」が含まれている場合は専用コピーを作成する
                    If InStr(Range("保管[内容]")(r).Value, "弁護士会照会") <> 0 Then Call MakeCopyOfSpecialFile(r, "弁護士会照会")
                    
                    ' 移動成功時、保管日時を現在日時で記録する
                    If brnMoveFile = True Then
                        Range("保管[保管日時]")(r).Value = Now
                    End If
                
                ' 保管日時が既に入力済みの場合（保管の取り消し・復元）
                Else
                    ' ファイルをリネームして元のフォルダに移動（復元）する
                    Call RstrFile(r)
                            
                    ' ショートカットおよびコピーを作成していた場合は削除する
                    If Range("保管[複製作成先]")(r).Value <> "" Then
                         Call DeleteShortcuts(r)
                         Call DeleteCopies(r)
                    Else
                         brnMakeShortcutsCopies = True
                    End If
                    
                    ' 内容に「職務上請求」が含まれている場合は専用コピーを削除する
                    If InStr(Range("保管[内容]")(r).Value, "職務上請求") <> 0 Then Call DeleteCopyOfSpecialFile(r, "職務上請求")
                    
                    ' 内容に「弁護士会照会」が含まれている場合は専用コピーを削除する
                    If InStr(Range("保管[内容]")(r).Value, "弁護士会照会") <> 0 Then Call DeleteCopyOfSpecialFile(r, "弁護士会照会")
                    
                    ' 移動（復元）成功時、保管情報をクリアする
                    If brnMoveFile = True Then
                        Range("保管[保管日時]")(r).Value = ""
                        Range("保管[保管フォルダ名]")(r).Value = ""
                        Range("保管[保管ファイル名]")(r).Value = ""
                    End If
                    
                    ' 複製作成処理のクリア成功時、複製作成先をクリアする
                    If brnMakeShortcutsCopies = True Then
                        Range("保管[複製作成先]")(r).Value = ""
                    End If
                End If
            End If
            
            ' 書式を再設定する
            If brnMoveFile = True And brnMakeShortcutsCopies = True Then
                Call FormatLine(r)
            End If
    End Select
Skip:
    ' イベントの発生を再開する
    Evnt.Enable
    Cancel = True ' ダブルクリックによる編集モード移行をキャンセル
End Sub

'================================================================================
' イベント:  Worksheet_Change
' 機能:      ワークシートのセルが変更されたときの処理
' 概要:      特定のセルの変更に連動して、他のセルの値を自動入力・クリアします。
'            オブジェクト変数を使用し、高速化を図っています。
'================================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    ' ターゲットのテーブル行・列番号を計算する変数を宣言
    Dim r As Long, s As Long
    Dim rngTable As Range
    Set rngTable = Range("保管")
    
    r = Target.Row - rngTable.Row + 1          ' テーブル行番号
    s = Target.Column - rngTable.Column + 1    ' テーブル列番号

    ' テーブル行・列番号が入力範囲外の場合は終了する
    If r < 1 Or rngTable.Rows.Count < r Then Exit Sub
    If s < 1 Or rngTable.Columns.Count < s Then Exit Sub

    ' イベントの発生を停止する（無限ループ防止）
    Evnt.Disable

    ' 画面描画のみを再開する（UI反応のため）
    Application.ScreenUpdating = True

    Select Case rngTable(0, s).Value
        Case "事件呼名"
           ' 保管フォルダ名、保管ファイル名および複製作成先をクリアする
            Range("保管[保管フォルダ名]")(r).Value = ""
            Range("保管[保管ファイル名]")(r).Value = ""
            Range("保管[複製作成先]")(r).Value = ""
        
            ' 事件呼名が入力されたならば、次の入力列（内容）を選択する
            If Range("保管[事件呼名]")(r).Value <> "" Then Range("保管[内容]")(r).Select
        
        Case "内容"
            ' 内容が入力された場合のみ処理を行う
            If Range("保管[内容]")(r).Value <> "" Then
                ' 保管先の事件呼名を検索してフォルダ名を入力する
                With Sht2Hokansaki
                    Dim rngHokansakiCase As Range
                    Set rngHokansakiCase = .Range("保管先[事件呼名]")
                    
                    ' 保管先に同一の事件呼名があればフォルダ名および複製作成先を入力する
                    If WorksheetFunction.CountIf(rngHokansakiCase, Range("保管[事件呼名]")(r).Value) > 0 Then
                        Dim t As Long               ' 保管先シートのテーブル行番号
                        t = WorksheetFunction.Match(Range("保管[事件呼名]")(r).Value, rngHokansakiCase, 0)
                        Range("保管[保管フォルダ名]")(r).Value = .Range("保管先[保管フォルダ]")(t).Value
                        Range("保管[複製作成先]")(r).Value = .Range("保管先[担当区分]")(t).Value
                    End If
                End With

                ' 保管ファイル名を生成して入力する
                Call InputStoredFile(r)
            Else
                ' 内容がクリアされたらファイル名もクリア
                Range("保管[保管ファイル名]")(r).Value = ""
            End If
        
            ' 保管フォルダ名が入力済みの場合
            If Range("保管[保管フォルダ名]")(r).Value <> "" Then
                ' 次の入力列（保管日時）を選択する
                Range("保管[保管日時]")(r).Select
            ' 保管フォルダ名が未入力の場合
            Else
                ' 次の入力列（保管フォルダ名）を選択する
                Range("保管[保管フォルダ名]")(r).Select
            End If
            
        Case "保管フォルダ名"
            ' 保管ファイル名を再生成して入力する
            Call InputStoredFile(r)
            
            ' 次の入力列を選択する
            If Range("保管[保管フォルダ名]")(r).Value <> "" Then Range("保管[保管日時]")(r).Select
        
        Case "保管ファイル名"
            ' 保管ファイル名が入力できていれば、次の入力列を選択する
            If Range("保管[保管ファイル名]")(r).Value <> "" Then Range("保管[保管日時]")(r).Select

        Case "複製作成先"
            ' 複製作成先が入力できていれば、次の入力列を選択する
            If Range("保管[複製作成先]")(r).Value <> "" Then Range("保管[保管日時]")(r).Select
    End Select

    ' イベントの発生を再開する
    Evnt.Enable
End Sub

'================================================================================
' イベント:  Worksheet_SelectionChange
' 機能:      選択セルが変更されたときの処理
' 概要:      対象事件呼名を記録します。
'================================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' ターゲットのテーブル行・列番号を計算する
    Dim r As Long, s As Long
    Dim rngTable As Range
    Set rngTable = Range("保管")
    
    If rngTable.Rows.Count = 0 Then Exit Sub
    r = Target.Row - rngTable.Row + 1          ' テーブル行番号
    s = Target.Column - rngTable.Column + 1    ' テーブル列番号

    ' テーブル行・列番号が入力範囲外の場合は終了する
    If r < 1 Or rngTable.Rows.Count < r Then Exit Sub
    If s < 1 Or rngTable.Columns.Count < s Then Exit Sub

    ' 対象事件呼名をグローバル変数に記録
    strActCase = Range("保管[事件呼名]")(r).Value
End Sub

'================================================================================
' 機能:     指定された「読込フォルダ」内のファイル名を一覧として取得し、シートに
'           書き出します。
' 概要:     設定シートの読込フォルダを巡回し、「保管」テーブルに追加します。
'           フォルダ名の末尾に「\」がある場合はサブフォルダ内のファイルも取得し、
'           末尾に「\」がない場合は指定フォルダ直下のファイルのみを取得します。
'================================================================================
Sub ReadFileNames()
    ' ステータスバーに処理中である旨のメッセージを表示します。
    Application.StatusBar = "保管するファイルのファイル名を取得しています。"

    ' 入力するテーブル行の番号を格納する変数を宣言します。
    Dim r As Long
    ' 「保管」テーブルの1行目1列目が空白であるか判定します。
    If Range("保管")(1, 1).Value = "" Then
        ' 空白の場合は、入力開始行を1行目に設定します。
        r = 1
    ' 空白でない場合（すでにデータが存在する場合）の処理です。
    Else
        ' 既存のデータ行数に1を加算し、新規入力行を設定します。
        r = Range("保管").Rows.Count + 1
    ' 行番号取得の条件分岐を終了します。
    End If

    ' FileSystemObjectを格納するための変数を宣言します。
    Dim fso As Object
    ' FileSystemObjectのインスタンスを生成します。
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 読み込みフォルダの指定値を格納する配列変数を宣言します。
    Dim varFolders As Variant
    ' 設定シートの「読込フォルダ」セルの値を改行コード（vbLf）で分割し、配列に格納します。
    varFolders = Split(ShtOption.Range("読込フォルダ").Value, vbLf)

    ' 最終的に探索の対象となるフォルダを管理するコレクション変数を宣言します。
    Dim colTargetFolders As Collection
    ' 探索対象フォルダ用コレクションのインスタンスを生成します（安全な構文に変更）。
    Set colTargetFolders = New Collection

    ' サブフォルダ探索時の処理待ちキュー（待機列）となるコレクション変数を宣言します。
    Dim colQueue As Collection
    ' キュー用コレクションのインスタンスを生成します（安全な構文に変更）。
    Set colQueue = New Collection

    ' 配列のループ処理で使用するカウンタ変数を宣言します。
    Dim x As Long
    ' 処理中のフォルダ名を一時的に格納する変数を宣言します。
    Dim strFolder As String
    ' サブフォルダを探索するかどうかを判定するフラグ変数を宣言します。
    Dim blnSubFolder As Boolean

    ' 1. 指定フォルダの振り分け
    ' varFolders配列の最初の要素から最後の要素までループ処理を行います。
    For x = 0 To UBound(varFolders)
        ' 配列要素の前後の空白を取り除き、変数に格納します。
        strFolder = Trim(varFolders(x))
        ' フォルダ名が空白でない場合のみ処理を実行します。
        If strFolder <> "" Then
            ' フォルダ名の末尾が「\」であるかを判定し、結果（TrueまたはFalse）をフラグに代入します。
            blnSubFolder = (Right(strFolder, 1) = "\")
            
            ' FileSystemObjectでの処理に適した形にするため、末尾の「\」を取り除く処理です。
            ' フラグがTrue（末尾に「\」がある）の場合の処理です。
            If blnSubFolder Then
                ' 文字列全体の長さから1文字分（\）を差し引いた文字数を左側から取得して上書きします。
                strFolder = Left(strFolder, Len(strFolder) - 1)
            ' 末尾「\」除去の条件分岐を終了します。
            End If
            
            ' FileSystemObjectを使用して、指定されたフォルダが実在するか確認します。
            If fso.FolderExists(strFolder) Then
                ' フラグがTrue（サブフォルダも対象）の場合の処理です。
                If blnSubFolder Then
                    ' 探索キューにフォルダを追加します（のちほど再帰的に探索するため）。
                    colQueue.Add strFolder
                ' フラグがFalse（直下のファイルのみ対象）の場合の処理です。
                Else
                    ' 直接、探索対象フォルダのコレクションに追加します。
                    colTargetFolders.Add strFolder
                ' 探索対象追加先の条件分岐を終了します。
                End If
            ' フォルダが存在しない場合の処理です（追加修正）。
            Else
                ' フォルダが見つからない旨の警告メッセージを表示します。
                MsgBox strFolder & " が見つかりません。", vbExclamation
            ' フォルダ存在確認の条件分岐を終了します。
            End If
        ' 空白判定の条件分岐を終了します。
        End If
    ' 次の配列要素の処理へ移行します。
    Next x

    ' 2. サブフォルダの階層を再帰的に取得する
    ' キューから取り出した現在処理中のフォルダ名を格納する変数を宣言します。
    Dim currentFolder As String
    ' FileSystemObjectのFolderオブジェクトを格納する変数を宣言します。
    Dim oFolder As Object
    ' サブフォルダを列挙するためのFolderオブジェクト変数を宣言します。
    Dim oSubFolder As Object
    
    ' 探索キューの中にフォルダが1つ以上存在する限り、ループ処理を繰り返します。
    Do While colQueue.Count > 0
        ' キューの先頭（1番目）にあるフォルダ名を取得し、変数に格納します。
        currentFolder = colQueue.Item(1)
        ' 取得した先頭の要素をキューから削除します。
        colQueue.Remove 1
        
        ' 取り出したフォルダ自身を、最終的なファイル読み込み対象リストに追加します。
        colTargetFolders.Add currentFolder
        
        ' アクセス権限のないフォルダによる実行時エラーを回避するため、エラーを無視して処理を継続する設定にします。
        On Error Resume Next
        ' 現在のフォルダ名から、FileSystemObjectのFolderオブジェクトを取得します。
        Set oFolder = fso.GetFolder(currentFolder)
        ' Folderオブジェクトの取得に成功した（エラー番号が0である）か判定します。
        If Err.Number = 0 Then
            ' 取得したフォルダ内のすべてのサブフォルダに対してループ処理を行います。
            For Each oSubFolder In oFolder.SubFolders
                ' 発見したサブフォルダのパスを、次に探索すべき対象としてキューの末尾に追加します。
                colQueue.Add oSubFolder.Path
            ' 次のサブフォルダの処理へ移行します。
            Next oSubFolder
        ' エラー判定の条件分岐を終了します。
        End If
        ' 発生した可能性のあるエラーの情報をクリアします。
        Err.Clear
        ' エラー無視の設定を解除し、通常のエラー処理に戻します。
        On Error GoTo 0
    ' キューのループ処理の先頭に戻ります。
    Loop

    ' 3. 対象フォルダリストからファイルを読み込む
    ' コレクションから要素を取り出すための汎用変数を宣言します。
    Dim vTarget As Variant
    ' FileSystemObjectのFileオブジェクトを格納する変数を宣言します。
    Dim oFile As Object
    ' ファイル名を格納する文字列変数を宣言します。
    Dim strFile As String
    ' 読み込んだファイルの数を数えるカウンタ変数を宣言します。
    Dim lngReadCnt As Long
    ' ファイルの更新日時を格納する日付型変数を宣言します。
    Dim datToday As Date
    ' 各フォルダにおいて、エラーメッセージをすでに表示したかを判定するフラグ変数を宣言します。
    Dim blnErrorShown As Boolean

    ' 探索対象フォルダのコレクションから、フォルダを1つずつ取り出してループ処理を行います。
    For Each vTarget In colTargetFolders
        ' 取り出した要素を文字列型に変換し、フォルダ名変数に格納します。
        strFolder = CStr(vTarget)
        ' ステータスバーに、現在読み込み中のフォルダ名を表示します。
        Application.StatusBar = strFolder & "からファイル名を読み込んでいます。"
        ' 新しいフォルダの処理に入るため、エラー表示済みフラグをFalse（未表示）にリセットします。
        blnErrorShown = False
        
        ' フォルダ自体の取得エラー（アクセス権限等）を回避するため、エラーを無視する設定にします。
        On Error Resume Next
        ' 文字列のフォルダ名から、FileSystemObjectのFolderオブジェクトを取得します。
        Set oFolder = fso.GetFolder(strFolder)
        ' Folderオブジェクトの取得時にエラーが発生したか判定します。
        If Err.Number <> 0 Then
            ' エラーが発生した旨をメッセージボックスで表示します。
            MsgBox strFolder & "の読み込み時にエラーが発生しました。" & vbCr & _
                "次のフォルダを読み込みます。", vbExclamation
            ' エラーの情報をクリアします。
            Err.Clear
            ' 現在のフォルダ処理を中断し、次のフォルダの処理（NextFolderラベル）へジャンプします。
            GoTo NextFolder
        ' エラー判定の条件分岐を終了します。
        End If
        ' エラー無視の設定を解除します。
        On Error GoTo 0

        ' 取得したフォルダ内のすべてのファイルに対してループ処理を行います。
        For Each oFile In oFolder.Files
            ' 個別のファイル読み込み時にエラーが発生した場合、HandleFileErrorラベルへジャンプするよう設定します。
            On Error GoTo HandleFileError
            
            ' Fileオブジェクトからファイル名を取得し、変数に格納します。
            strFile = oFile.Name
            
            ' ファイル名の先頭1文字が「.（ドット）」であるか判定します。
            If Left(strFile, 1) = "." Then GoTo NextFile
            
            ' 現在入力しようとしている行が2行目以降（すでに1件以上データがある状態）か判定します。
            If r > 1 Then
                ' 読込フォルダ名および読込ファイル名が既存データと一致する行をカウントし、1以上か判定します。
                If WorksheetFunction.CountIfs(Range("保管[読込フォルダ名]"), strFolder, Range("保管[読込ファイル名]"), strFile) > 0 Then
                    ' 重複が存在する場合は、処理を行わず次のファイルの処理（NextFileラベル）へジャンプします。
                    GoTo NextFile
                ' 重複判定の条件分岐を終了します。
                End If
            ' 行番号判定の条件分岐を終了します。
            End If
            
            ' MOTTELのFAX機能で生成される特定のファイル名（先頭10文字がFAXMAILFAX）であるか判定し、該当する場合はスキップします。
            If Left(strFile, 10) = "FAXMAILFAX" Then GoTo NextFile
                                            
            ' 「保管」テーブルの「番号」列の該当行に、現在の日時を「yymmddhhnnss」形式の文字列として入力します。
            Range("保管[番号]")(r).Value = Format(Now, "yymmddhhnnss")
            
            ' 入力行が2行目以降であるか判定します。
            If r > 1 Then
                ' いま入力した番号と同じ番号がテーブル内に複数存在するか（同一秒内の処理で重複したか）判定するループ処理です。
                Do While WorksheetFunction.CountIf(Range("保管[番号]"), Range("保管[番号]")(r).Value) > 1
                    ' 重複している間、入力した番号を数値として扱い1を加算し、再度文字列としてセルに格納し直します。
                    Range("保管[番号]")(r).Value = CStr(CDbl(Range("保管[番号]")(r).Value) + 1)
                ' 重複回避のループ処理の先頭に戻り、重複がなくなるまで繰り返します。
                Loop
            ' 行番号判定の条件分岐を終了します。
            End If
            
            ' 「保管」テーブルの「読込フォルダ名」列の該当行に、フォルダ名を入力します。
            Range("保管[読込フォルダ名]")(r).Value = strFolder
            ' 「保管」テーブルの「読込ファイル名」列の該当行に、ファイル名を入力します。
            Range("保管[読込ファイル名]")(r).Value = strFile
            
            ' ファイル名の先頭2文字が"20"であり、かつ先頭から12文字が数値として扱えるか判定します。
            If Left(strFile, 2) = "20" And IsNumeric(Left(strFile, 12)) Then
                ' 判定を満たす場合、ファイル名の先頭12文字を日時データとして「日時」列に入力します。
                Range("保管[日時]")(r).Value = Left(strFile, 12)
            ' 判定を満たさない場合の処理です。
            Else
                ' FileSystemObjectのプロパティから、ファイルの最終更新日時を取得します。
                datToday = oFile.DateLastModified
                ' 取得した更新日時を「yyyymmddhhnn」形式の文字列に変換し、「日時」列に入力します。
                Range("保管[日時]")(r).Value = Format(datToday, "yyyymmddhhnn")
            ' 日時取得の条件分岐を終了します。
            End If
                        
            ' 別途定義されたFormatLineプロシージャを呼び出し、該当行の書式を設定します。
            Call FormatLine(r)
            
            ' 正常に読み込めたファイルのカウンタ変数に1を加算します。
            lngReadCnt = lngReadCnt + 1
            
            ' 次のデータを入力するため、テーブル行の番号に1を加算します。
            r = r + 1
NextFile:
            ' ファイルごとのエラーハンドリングを解除し、通常のエラー処理に戻します。
            On Error GoTo 0
        ' 次のファイルの処理へ移行します。
        Next oFile
        
        ' フォルダ内のすべてのファイルを正常に処理し終えた場合は、次のフォルダの処理へジャンプします。
        GoTo NextFolder

HandleFileError:
        ' フォルダ内で初めてエラーが発生したか（フラグがFalseか）判定します。
        If blnErrorShown = False Then
            ' 初回のみ、読み込めないファイルが存在した旨をメッセージボックスで表示します。
            MsgBox strFolder & "に読み込めないファイルがありました。", vbExclamation
            ' フラグをTrueに変更し、同フォルダ内でこれ以上メッセージが表示されないようにします。
            blnErrorShown = True
        ' エラー表示判定の条件分岐を終了します。
        End If
        ' エラー状態をクリアし、エラーが発生したファイルの次のファイル処理（NextFileラベル）から再開します。
        Resume NextFile

NextFolder:
    ' 次の対象フォルダの処理へ移行します。
    Next vTarget

    ' 別途定義されたSortAndFormatTableプロシージャを呼び出し、テーブル全体の並べ替えおよび書式設定を行います。
    Call SortAndFormatTable

    ' 処理が完了したため、ステータスバーのメッセージをクリア（非表示）にします。
    Application.StatusBar = ""
    ' 読み込んだファイル数が0であったか判定します。
    If lngReadCnt = 0 Then
        ' 0件であった旨の完了メッセージを表示します。
        MsgBox "読込ファイル名の読み込みを完了しました。" & vbCr & _
            "読み込むファイルがありませんでした。", vbInformation
    ' 1件以上のファイルを読み込んだ場合の処理です。
    Else
        ' 読み込んだファイル数を含む完了メッセージを表示します。
        MsgBox "読込ファイル名の読み込みを完了しました。" & vbCr & _
            lngReadCnt & "個のファイル名を読み込みました。", vbInformation
    ' 完了メッセージ表示の条件分岐を終了します。
    End If

    ' FileSystemObjectのインスタンスを破棄し、メモリを解放します。
    Set fso = Nothing
End Sub

'================================================================================
' プロシージャ: MoveFile
' 機能:         ファイルを「読込フォルダ」から「保管フォルダ」へ移動（リネーム）します。
' 概要:         FileSystemObject (FSO) を使用してファイルを移動します。
'               FSOを使用することで、異なるドライブ間での移動もサポートします。
'================================================================================
Private Sub MoveFile(r As Long)
    Dim strReadedFile As String, strStoredFile As String
    Dim strErrMsg As String
    Dim fso As Object

    ' 完了フラグをリセットする
    brnMoveFile = False

    Application.StatusBar = "ファイルのコピー元と移動先を確認しています"

    ' FSOオブジェクトを生成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' コピー元ファイルのパスを作成
    strReadedFile = Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value
    
    ' コピー元ファイルが存在しない場合のエラー処理
    If Not fso.FileExists(strReadedFile) Then
        strErrMsg = strReadedFile & "が見つかりません。"
        GoTo HdlPrelimErr
    End If

    ' 保管フォルダが存在しない場合のエラー処理
    If Not fso.FolderExists(Range("保管[保管フォルダ名]")(r).Value) Then
        strErrMsg = Range("保管[保管フォルダ名]")(r).Value & "が見つかりません。"
        GoTo HdlPrelimErr
    End If
    
    ' 移動先ファイルのパスを作成
    strStoredFile = Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value

    Application.StatusBar = "ファイルを保管フォルダに移動しています。"

    ' 同じ名前のファイルが移動先にない場合
    If Not fso.FileExists(strStoredFile) Then
        On Error GoTo HdlWriteErr
        ' FSOを使用してファイルを移動（ドライブ間移動も可能）
        fso.MoveFile Source:=strReadedFile, Destination:=strStoredFile
        On Error GoTo 0
    Else
        strErrMsg = strStoredFile & "と同じ名前のファイルが既にあります。"
        GoTo HdlPrelimErr
    End If

    Application.StatusBar = "保管先データを更新しています。"

    ' 保管先シートのデータを更新
    With Sht2Hokansaki
        Dim lo As ListObject
        Set lo = .ListObjects("保管先")
        
        Dim t As Long
        Dim isNew As Boolean
        isNew = True
        
        ' データが存在する場合のみ検索を行う（DataBodyRangeがNothingでないか確認）
        If Not lo.DataBodyRange Is Nothing Then
            ' 事件呼名列が存在するか確認し、存在すれば検索
            If WorksheetFunction.CountIf(lo.ListColumns("事件呼名").DataBodyRange, Range("保管[事件呼名]")(r).Value) > 0 Then
                t = WorksheetFunction.Match(Range("保管[事件呼名]")(r).Value, lo.ListColumns("事件呼名").DataBodyRange, 0)
                isNew = False
            End If
        End If
        
        ' 新規か更新かで処理を分ける
        If isNew Then
            ' 新規行を追加（これによりテーブルが空でも確実に内部に追加される）
            Dim newRow As ListRow
            Set newRow = lo.ListRows.Add
            
            ' 列インデックスを使用して書き込み
            newRow.Range(1, lo.ListColumns("事件呼名").Index).Value = Range("保管[事件呼名]")(r).Value
            newRow.Range(1, lo.ListColumns("保管フォルダ").Index).Value = Range("保管[保管フォルダ名]")(r).Value
            newRow.Range(1, lo.ListColumns("担当区分").Index).Value = Range("保管[複製作成先]")(r).Value
        Else
            ' 既存行を更新
            lo.ListColumns("事件呼名").DataBodyRange(t).Value = Range("保管[事件呼名]")(r).Value
            lo.ListColumns("保管フォルダ").DataBodyRange(t).Value = Range("保管[保管フォルダ名]")(r).Value
            lo.ListColumns("担当区分").DataBodyRange(t).Value = Range("保管[複製作成先]")(r).Value
        End If
    End With

    Application.StatusBar = ""
    brnMoveFile = True
    Set fso = Nothing
    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & "このファイルの移動を中止しました。", vbExclamation
    Application.StatusBar = ""
    Set fso = Nothing
    Exit Sub
    
HdlWriteErr:
    MsgBox "このファイルを移動できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイルが別のアプロケーションで開かれている。" & vbCr & _
        "・ファイル名に使うことのできない文字が含まれている。" & vbCr & _
        "詳細: " & Err.Description, vbExclamation
    Application.StatusBar = ""
    Set fso = Nothing
End Sub

'================================================================================
' プロシージャ: RstrFile
' 機能:         一度保管したファイルを、元の「読込フォルダ」に復元します。
' 概要:         FileSystemObject (FSO) を使用してファイルを元の場所に戻します。
'               FSOを使用することで、異なるドライブ間での移動もサポートします。
'================================================================================
Private Sub RstrFile(r As Long)
    Dim strReadedFile As String, strStoredFile As String
    Dim strErrMsg As String
    Dim fso As Object

    ' 完了フラグをリセットする
    brnMoveFile = False

    Application.StatusBar = "ファイルのコピー元と移動先を確認しています"

    ' FSOオブジェクトを生成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 現在の保管ファイル（移動元）のパスを作成
    strReadedFile = Range("保管[保管フォルダ名]")(r).Value & "\" & Range("保管[保管ファイル名]")(r).Value
    
    ' 保管ファイルが存在しない場合
    If Not fso.FileExists(strReadedFile) Then
        strErrMsg = strReadedFile & "が見つかりません。"
        GoTo HdlPrelimErr
    End If

    ' 元の読込フォルダが存在しない場合
    If Not fso.FolderExists(Range("保管[読込フォルダ名]")(r).Value) Then
        strErrMsg = Range("保管[読込フォルダ名]")(r).Value & "が見つかりません。"
        GoTo HdlPrelimErr
    End If

    ' 元のファイル名（移動先）のパスを作成
    strStoredFile = Range("保管[読込フォルダ名]")(r).Value & "\" & Range("保管[読込ファイル名]")(r).Value

    Application.StatusBar = "ファイルを読込フォルダに移動しています。"

    ' 同じ名前のファイルが移動先にない場合
    If Not fso.FileExists(strStoredFile) Then
        On Error GoTo HdlWriteErr
        ' FSOを使用してファイルを移動
        fso.MoveFile Source:=strReadedFile, Destination:=strStoredFile
        On Error GoTo 0
    Else
        strErrMsg = strStoredFile & "と同じ名前のファイルが既にあります。"
        GoTo HdlPrelimErr
    End If
            
    Application.StatusBar = ""
    brnMoveFile = True
    Set fso = Nothing
    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & "このファイルの元のフォルダへの移動を中止しました。", vbExclamation
    Application.StatusBar = ""
    Set fso = Nothing
    Exit Sub
    
HdlWriteErr:
    MsgBox "このファイルを元のフォルダに移動できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイルが別のアプロケーションで開かれている。" & vbCr & _
        "・ファイル名に使うことのできない文字が含まれている。" & vbCr & _
        "詳細: " & Err.Description, vbExclamation
    Application.StatusBar = ""
    Set fso = Nothing
End Sub

'================================================================================
' プロシージャ: MakeShortcuts
' 機能:         保管したファイルのショートカットを、指定されたフォルダに作成します。
' 概要:         WScript.Shellを使用してショートカット(.lnk)を作成します。
'               設定シートから出力先フォルダを取得してループ処理を行います。
'================================================================================
Private Sub MakeShortcuts(r As Long)
    ' 完了フラグをリセットする
    brnMakeShortcutsCopies = False

    Application.StatusBar = "担当弁護士に応じた未確認フォルダを確認しています。"

    ' SC作成先フォルダの行番号を取得する変数
    Dim v As Long
    ' 作成先区分の範囲をオブジェクト変数にセット
    Dim rngKbn As Range
    Set rngKbn = ShtOption.Range("SC作成先設定[複製作成先]")

    ' 区分が存在するか確認
    If WorksheetFunction.CountIf(rngKbn, Range("保管[複製作成先]")(r).Value) > 0 Then
        v = WorksheetFunction.Match(Range("保管[複製作成先]")(r).Value, rngKbn, 0)
    Else
        GoTo HdlPrelimErr ' 作成先区分が見つからない場合は終了
    End If

    ' SC作成先フォルダのリスト文字列を取得
    Dim strShortcutFolders As String, strShortcutFolder As String, strLinkFolder As String, strLinkFile As String
    strShortcutFolders = Replace(ShtOption.Range("SC作成先設定[フォルダ]")(v).Value, vbLf, vbCrLf)

    ' SC作成先フォルダごとにSCを作成するループ変数
    Dim x As Integer
    Dim objWshShell As Object
    Dim strShortcut As String
    Dim strErrMsg As String
    
    Set objWshShell = CreateObject("WScript.Shell") ' WshShellオブジェクトを作成
    x = 1
    Do Until SplitText(strShortcutFolders, x) = ""
        ' SC作成先フォルダを把握する
        strShortcutFolder = SplitText(strShortcutFolders, x)

        ' 未確認フォルダが存在しなければ中止する
        If Dir(strShortcutFolder, vbDirectory) = "" Then
            strErrMsg = strShortcutFolder & "がありません。"
            GoTo HdlPrelimErr
        End If
            
        ' リンク先のフォルダとファイルを取得する
        strLinkFolder = Range("保管[保管フォルダ名]")(r)
        strLinkFile = Range("保管[保管ファイル名]")(r)
        
        ' リンク先ファイルが存在しなければ中止する
        If Dir(strLinkFolder & "\" & strLinkFile) = "" Then
            strErrMsg = strLinkFolder & "\" & strLinkFile & "がありません。"
            GoTo HdlPrelimErr
        End If
        
        ' ショートカットファイルのフルパスを設定する
        strShortcut = strShortcutFolder & "\" & strLinkFile & ".lnk"
         
        ' 同名のショートカットがない場合
        If Dir(strShortcut) = "" Then
            ' ショートカットを作成する
            On Error GoTo HdlWriteErr
            With objWshShell.CreateShortcut(strShortcut)
                .targetPath = strLinkFolder & "\" & strLinkFile  ' ショートカットのリンク先を設定
                .WorkingDirectory = strLinkFolder                ' 作業フォルダ
                .Save                                            ' ショートカットの保存実行
            End With
            On Error GoTo 0
        ' 同名のショートカットがある場合
        Else
            strErrMsg = strShortcut & "と同じ名前のショートカットが既にあります。"
            GoTo HdlPrelimErr
        End If
        
        ' 次のSC作成先フォルダに移る
        x = x + 1
    Loop

    Set objWshShell = Nothing
    Application.StatusBar = ""
    brnMakeShortcutsCopies = True
    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & "このファイルのショートカットの作成を中止しました。", vbExclamation
    Set objWshShell = Nothing
    Application.StatusBar = ""
    Exit Sub

HdlWriteErr:
    MsgBox "このファイルのショートカットを作成できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイル名に使うことのできない文字が含まれている。", vbExclamation
    Set objWshShell = Nothing
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: DeleteShortcuts
' 機能:         作成されたショートカットを削除します。
' 概要:         指定されたフォルダ内のショートカットファイルをKillステートメントで削除します。
'================================================================================
Private Sub DeleteShortcuts(r As Long)
    ' 完了フラグをリセットする
    brnMakeShortcutsCopies = False

    Application.StatusBar = "担当弁護士の未確認ファイルを削除しています。"

    ' SC作成先フォルダを取得する
    Dim v As Long
    Dim rngKbn As Range
    Set rngKbn = ShtOption.Range("SC作成先設定[複製作成先]")

    If WorksheetFunction.CountIf(rngKbn, Range("保管[複製作成先]")(r).Value) > 0 Then
        v = WorksheetFunction.Match(Range("保管[複製作成先]")(r).Value, rngKbn, 0)
    Else
        GoTo HdlPrelimErr ' 作成先区分が見つからない場合は終了
    End If

    Dim strShortcutFolders As String, strShortcutFolder As String
    strShortcutFolders = Replace(ShtOption.Range("SC作成先設定[フォルダ]")(v).Value, vbLf, vbCrLf)

    ' SC作成先フォルダごとにSCを削除するループ
    Dim x As Integer
    Dim strShortcut As String
    Dim strErrMsg As String
    x = 1
    Do Until SplitText(strShortcutFolders, x) = ""
        strShortcutFolder = SplitText(strShortcutFolders, x)

        ' 未確認フォルダが存在しない場合
        If Dir(strShortcutFolder, vbDirectory) = "" Then
            strErrMsg = strShortcutFolder & "がありません。"
            GoTo HdlPrelimErr
        End If
            
        ' ショートカットのファイル名を取得する
        strShortcut = Range("保管[保管ファイル名]")(r) & ".lnk"
         
        ' 同名のショートカットがある場合
        If Dir(strShortcutFolder & "\" & strShortcut) <> "" Then
            ' ショートカットを削除する
            On Error GoTo HdlErr
            Kill strShortcutFolder & "\" & strShortcut
            On Error GoTo 0
        ' 同名のショートカットがない場合
        Else
            strErrMsg = strShortcut & "が見つかりません。"
            GoTo HdlPrelimErr
        End If
        
        ' 次のSC作成先フォルダに移る
        x = x + 1
    Loop

    Application.StatusBar = ""
    brnMakeShortcutsCopies = True
    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & "このファイルのショートカットの削除を中止しました。", vbExclamation
    Application.StatusBar = ""
    Exit Sub

HdlErr:
    MsgBox "このファイルのショートカットを削除できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイルが別のアプロケーションで開かれている。", vbExclamation
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: MakeCopies
' 機能:         保管したファイルのコピーを、指定されたフォルダに作成します。
' 概要:         FileCopyステートメントを使用してファイルのコピーを作成します。
'================================================================================
Private Sub MakeCopies(r As Long)
    ' 完了フラグをリセットする
    brnMakeShortcutsCopies = False

    Application.StatusBar = "担当弁護士の未確認フォルダにコピーを作成しています。"

    ' 複製作成先フォルダを取得する
    Dim v As Long
    Dim rngKbn As Range
    Set rngKbn = ShtOption.Range("コピー作成先設定[複製作成先]")

    If WorksheetFunction.CountIf(rngKbn, Range("保管[複製作成先]")(r).Value) > 0 Then
        v = WorksheetFunction.Match(Range("保管[複製作成先]")(r).Value, rngKbn, 0)
    Else
        GoTo HdlPrelimErr ' 作成先区分が見つからない場合は終了
    End If

    Dim strCopyFolders As String, strCopyFolder As String, strOriginalFolder As String, strOriginalFile As String
    strCopyFolders = Replace(ShtOption.Range("コピー作成先設定[フォルダ]")(v).Value, vbLf, vbCrLf)

    ' コピー作成先フォルダごとにコピーを作成するループ
    Dim x As Integer
    Dim strCopy As String
    Dim strErrMsg As String
    x = 1
    Do Until SplitText(strCopyFolders, x) = ""
        strCopyFolder = SplitText(strCopyFolders, x)

        ' コピー先フォルダが存在しない場合
        If Dir(strCopyFolder, vbDirectory) = "" Then
            strErrMsg = strCopyFolder & "がありません。"
            GoTo HdlPrelimErr
        End If
            
        ' コピー元のフォルダとファイルを取得する
        strOriginalFolder = Range("保管[保管フォルダ名]")(r).Value
        strOriginalFile = Range("保管[保管ファイル名]")(r).Value
        
        ' コピー元ファイルが存在しない場合
        If Dir(strOriginalFolder & "\" & strOriginalFile) = "" Then
            strErrMsg = strOriginalFolder & "\" & strOriginalFile & "がありません。"
            GoTo HdlPrelimErr
        End If
        
        ' ファイル名のフルパスを設定する
        strCopy = strCopyFolder & "\" & strOriginalFile
         
        ' 同名のファイルがない場合
        If Dir(strCopy) = "" Then
            ' コピーを作成する
            On Error GoTo HdlWriteErr
            FileCopy strOriginalFolder & "\" & strOriginalFile, strCopy
            On Error GoTo 0
        ' 同名のファイルがある場合
        Else
            strErrMsg = strCopy & "と同じ名前のファイルが既にあります。"
            GoTo HdlPrelimErr
        End If
        
        ' 次のコピー先フォルダに移る
        x = x + 1
    Loop

    brnMakeShortcutsCopies = True
    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & "このファイルのコピーの作成を中止しました。", vbExclamation
    Application.StatusBar = ""
    Exit Sub

HdlWriteErr:
    MsgBox "このファイルのコピーを作成できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイル名に使うことのできない文字が含まれている。", vbExclamation
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: DeleteCopies
' 機能:         作成されたファイルのコピーを削除します。
' 概要:         指定されたフォルダ内のコピーファイルをKillステートメントで削除します。
'================================================================================
Private Sub DeleteCopies(r As Long)
    ' 完了フラグをリセットする
    brnMakeShortcutsCopies = False

    Application.StatusBar = "担当弁護士の未確認フォルダからファイルを削除しています。"

    ' コピー作成先フォルダを取得する
    Dim v As Long
    Dim rngKbn As Range
    Set rngKbn = ShtOption.Range("コピー作成先設定[複製作成先]")

    If WorksheetFunction.CountIf(rngKbn, Range("保管[複製作成先]")(r).Value) > 0 Then
        v = WorksheetFunction.Match(Range("保管[複製作成先]")(r).Value, rngKbn, 0)
    Else
        GoTo HdlPrelimErr ' 作成先区分が見つからない場合は終了
    End If

    Dim strCopyFolders As String, strCopyFolder As String
    strCopyFolders = Replace(ShtOption.Range("コピー作成先設定[フォルダ]")(v).Value, vbLf, vbCrLf)

    ' コピー作成先フォルダごとにコピーを削除するループ
    Dim x As Integer
    Dim strCopy As String
    Dim strErrMsg As String
    x = 1
    Do Until SplitText(strCopyFolders, x) = ""
        strCopyFolder = SplitText(strCopyFolders, x)

        ' 未確認フォルダが存在しなければ中止する
        If Dir(strCopyFolder, vbDirectory) = "" Then
            strErrMsg = strCopyFolder & "がありません。"
            GoTo HdlPrelimErr
        End If
            
        ' コピー元のファイル名を取得する
        strCopy = Range("保管[保管ファイル名]")(r)
         
        ' 同名のファイルがある場合
        If Dir(strCopyFolder & "\" & strCopy) <> "" Then
            ' 削除する
            On Error GoTo HdlErr
            Kill strCopyFolder & "\" & strCopy
            On Error GoTo 0
        Else
            strErrMsg = strCopy & "が見つかりません。"
            GoTo HdlPrelimErr
        End If
        
        ' 次のコピー作成先フォルダに移行する
        x = x + 1
    Loop

    Application.StatusBar = ""
    brnMakeShortcutsCopies = True
    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & "このファイルのコピーの削除を中止しました。", vbExclamation
    Application.StatusBar = ""
    Exit Sub

HdlErr:
    MsgBox "このファイルのコピーを削除できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイルが別のアプロケーションで開かれている。", vbExclamation
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: MakeCopyOfSpecialFile
' 機能:         特定の文書種類のファイルのコピーを専用フォルダに作成します。
' 概要:         職務上請求や弁護士会照会などの特定文書を指定フォルダにコピーします。
'================================================================================
Private Sub MakeCopyOfSpecialFile(r As Long, strDocType As String)
    Application.StatusBar = strDocType & "保管フォルダにコピーを作成しています。"
        
    ' コピー先フォルダを把握します。
    Dim strCopyFolderName As String
    strCopyFolderName = strDocType & "保管フォルダ"

    Dim strCopyFolder As String
    strCopyFolder = ShtOption.Range(strCopyFolderName).Value

    ' コピー先フォルダが存在しない場合
    Dim strErrMsg  As String
    If Dir(strCopyFolder, vbDirectory) = "" Then
        strErrMsg = strCopyFolder & "がありません。"
        GoTo HdlPrelimErr
    End If
        
    ' コピー元のフォルダとファイルを取得する
    Dim strOriginalFolder As String, strOriginalFile As String
    strOriginalFolder = Range("保管[保管フォルダ名]")(r).Value
    strOriginalFile = Range("保管[保管ファイル名]")(r).Value

    ' コピー元ファイルが存在しない場合
    If Dir(strOriginalFolder & "\" & strOriginalFile) = "" Then
        strErrMsg = strOriginalFolder & "\" & strOriginalFile & "がありません。"
        GoTo HdlPrelimErr
    End If

    ' ファイル名を設定する
    Dim strCopy As String
    strCopy = strCopyFolder & "\" & strOriginalFile

    ' 同名のファイルがない場合
    If Dir(strCopy) = "" Then
        ' コピーを作成する
        On Error GoTo HdlWriteErr
        FileCopy strOriginalFolder & "\" & strOriginalFile, strCopy
        On Error GoTo 0
    ' 同名のファイルがある場合
    Else
        strErrMsg = strCopy & "と同じ名前のファイルが既にあります。"
        GoTo HdlPrelimErr
    End If

    Application.StatusBar = ""
    MsgBox strDocType & "保管フォルダにコピーを作成しました。"
    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & strDocType & "のコピーの作成を中止しました。", vbExclamation
    Application.StatusBar = ""
    Exit Sub

HdlWriteErr:
    MsgBox strDocType & "のコピーを作成できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイル名に使うことのできない文字が含まれている。", vbExclamation
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: DeleteCopyOfSpecialFile
' 機能:         専用フォルダに作成された特定の文書種類のファイルのコピーを削除します。
' 概要:         保管を取り消した際などに、専用フォルダ内のコピーを削除します。
'================================================================================
Private Sub DeleteCopyOfSpecialFile(r As Long, strDocType As String)
    Application.StatusBar = strDocType & "保管フォルダからファイルを削除しています。"

    ' コピー先フォルダを把握します。
    Dim strCopyFolderName As String
    strCopyFolderName = strDocType & "保管フォルダ"

    Dim strCopyFolder As String
    strCopyFolder = ShtOption.Range(strCopyFolderName).Value

    ' 保管フォルダが存在しなければ中止する
    Dim strErrMsg As String
    If Dir(strCopyFolder, vbDirectory) = "" Then
        strErrMsg = strCopyFolder & "がありません。"
        GoTo HdlPrelimErr
    End If
            
    ' コピー元のファイル名を取得する
    Dim strCopy As String
    strCopy = Range("保管[保管ファイル名]")(r)
         
    ' 同名のファイルがある場合
    If Dir(strCopyFolder & "\" & strCopy) <> "" Then
        ' 削除する
        On Error GoTo HdlErr
        Kill strCopyFolder & "\" & strCopy
        On Error GoTo 0
    Else
        strErrMsg = strCopy & "が見つかりません。"
        GoTo HdlPrelimErr
    End If
        
    Application.StatusBar = ""
    MsgBox strDocType & "保管フォルダのコピーを削除しました。"
    Exit Sub

HdlPrelimErr:
    MsgBox strErrMsg & vbCr & vbCr & strDocType & "のコピーの削除を中止しました。", vbExclamation
    Application.StatusBar = ""
    Exit Sub

HdlErr:
    MsgBox strDocType & "のコピーを削除できませんでした。" & vbCr & _
        "次のような原因が考えられます。" & vbCr & vbCr & _
        "・ファイルが別のアプロケーションで開かれている。", vbExclamation
    Application.StatusBar = ""
End Sub

'================================================================================
' プロシージャ: FormatLine
' 機能:         指定された行の書式設定（入力規則リスト、表示形式、フォント色）を行います。
' 概要:         オブジェクト変数を使用してセルにアクセスし、書式を整えます。
'================================================================================
Private Sub FormatLine(r As Long)
    ' リストを設定する
    Dim rngTarget As Range
    
    ' 複製作成先列の入力規則を設定
    Set rngTarget = Range("保管[複製作成先]")(r)
    With rngTarget.Validation
        .Delete
        ' SC作成先設定テーブルの作成先区分列を動的に参照するように設定
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(""SC作成先設定[複製作成先]"")"
    End With

    ' 保管日時列の表示形式を設定
    Range("保管[保管日時]")(r).NumberFormatLocal = "yyyy/m/d h:mm"

    ' 行全体の文字色を設定
    Dim rngRow As Range
    Set rngRow = Range("保管").Rows(r)
    With rngRow
        Select Case Range("保管[保管日時]")(r).Value
            Case ""
                .Font.ColorIndex = xlAutomatic
            Case Else
                .Font.ThemeColor = xlThemeColorAccent3
        End Select
    End With
End Sub

'================================================================================
' プロシージャ: SortAndFormatTable
' 機能:           「保管」テーブル全体の並べ替えと書式設定を再適用します。
' 概要:           テーブルオブジェクトを使用して並べ替えと書式設定を行います。
'                 【修正】Evntクラスでの一元管理のため、ScreenUpdating操作は削除しました。
'================================================================================
Public Sub SortAndFormatTable()
    Application.StatusBar = "「保管」テーブルの並び替えと書式設定を行っています。"
    
    ' ※画面更新制御は呼び出し元（Evntクラス）に委ねるため、ここでは操作しません。
    '   このメソッド単体で呼ぶ場合は、呼び出し元で Evnt.Disable / Enable を行ってください。

    Dim lo As ListObject
    Set lo = Me.ListObjects("保管")

    ' データが存在しない場合は処理を終了
    If lo.ListRows.Count = 0 Then GoTo Cleanup
    If lo.DataBodyRange Is Nothing Then GoTo Cleanup
    
    ' データが1行のみで、かつ空の場合は処理を終了（初期状態など）
    If lo.ListRows.Count = 1 And lo.DataBodyRange(1, 1).Value = "" Then GoTo Cleanup
    
    ' 並び替え
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add lo.ListColumns("保管日時").Range, xlSortOnValues, xlAscending
        .SortFields.Add lo.ListColumns("読込フォルダ名").Range, xlSortOnValues, xlAscending
        .SortFields.Add lo.ListColumns("読込ファイル名").Range, xlSortOnValues, xlAscending
        .Header = xlYes
        .Apply
    End With

    ' テーブル範囲だけを対象とした書式設定
    With lo.DataBodyRange
        ' セル内の文字列を折り返して全体を表示する
        .WrapText = True
        
        ' 文字色を一旦自動に戻す
        .Font.ColorIndex = xlAutomatic
    End With
        
    ' 各列の表示形式を設定
    On Error Resume Next
    lo.ListColumns("番号").DataBodyRange.NumberFormatLocal = "0"
    lo.ListColumns("日時").DataBodyRange.NumberFormatLocal = "0"
    lo.ListColumns("保管日時").DataBodyRange.NumberFormatLocal = "yyyy/m/d h:mm"
    On Error GoTo 0

    ' 文字色をまとめて設定する（保管日時が入力されている行のみ）
    Dim rngFilled As Range
    On Error Resume Next
    Set rngFilled = lo.ListColumns("保管日時").DataBodyRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    
    If Not rngFilled Is Nothing Then
        ' 【高速化ポイント】
        ' rngFilled.EntireRowだとシートの全列（16,384列）に対して処理が走るため、
        ' Intersectを使用してテーブル範囲内に限定する。
        Intersect(rngFilled.EntireRow, lo.DataBodyRange).Font.ThemeColor = xlThemeColorAccent3
    End If

    ' 【高速化ポイント】
    ' AutoFitは計算コストが高いため、全ての書式設定が完了した最後に一度だけ実行する。
    ' セルの高さを調節する
    lo.DataBodyRange.Rows.AutoFit

Cleanup:
    Application.StatusBar = ""
End Sub
                        
'================================================================================
' プロシージャ: InputStoredFolder
' 機能:         保管フォルダ名を自動で入力します。
' 概要:         事件呼名をキーにして保管先情報を検索し、フォルダ名を入力します。
'================================================================================
Private Sub InputStoredFolder(r As Long)
    ' 保管フォルダ名が入力済みの場合は終了する
    If Range("保管[保管フォルダ名]")(r).Value <> "" Then Exit Sub

    ' 保管先の事件呼名の有無に応じて処理を行う
    If WorksheetFunction.CountIf(Sht2Hokansaki.Range("保管先[事件呼名]"), Range("保管[事件呼名]")(r).Value) > 0 Then
        With Sht2Hokansaki
            Dim t As Long
            Dim strFolder As String
            
            ' 対象事件のテーブル行を取得する
            t = WorksheetFunction.Match(Range("保管[事件呼名]")(r).Value, .Range("保管先[事件呼名]"), 0)
            
            ' 対象事件のデータフォルダを取得する
            strFolder = .Range("保管先[保管フォルダ]")(t).Value
            
            ' 保管フォルダ名と担当区分を入力する
            Range("保管[保管フォルダ名]")(r).Value = .Range("保管先[保管フォルダ]")(t).Value
            Range("保管[複製作成先]")(r).Value = .Range("保管先[担当区分]")(t).Value
        End With
    Else
        ' 保管フォルダ名をダイアログで入力する
        With Application.FileDialog(msoFileDialogFolderPicker)
            .InitialFileName = ShtOption.Range("保管フォルダ")
            If .Show = True Then
                Range("保管[保管フォルダ名]")(r).Value = .SelectedItems(1)
            End If
        End With
    End If
End Sub
                    
'================================================================================
' プロシージャ: InputStoredFile
' 機能:         保管ファイル名をルールに基づいて自動生成し、入力します。
' 概要:         日時、事件呼名、内容などを組み合わせてファイル名を生成します。
'================================================================================
Private Sub InputStoredFile(r As Long)
    ' 保管ファイル名の各要素を取得する
    Dim strDate As String, strCase As String, strCont As String, strPrevName As String, strExt As String
    
    ' 拡張子の取得
    Dim strOriginalFileName As String
    strOriginalFileName = Range("保管[読込ファイル名]")(r).Value
    
    If InStr(strOriginalFileName, ".") > 0 Then
        strExt = Mid(strOriginalFileName, InStr(strOriginalFileName, "."))
    Else
        strExt = ""
    End If

    strDate = Range("保管[日時]")(r).Value & "00000"
    strCase = Range("保管[事件呼名]")(r).Value
    strCont = Range("保管[内容]")(r).Value
    
    ' 保管ファイル名を入力する
    If Right(strCase, 4) = "（文献）" Then
        Range("保管[保管ファイル名]")(r).Value = strCont & strExt
    Else
        Range("保管[保管ファイル名]")(r).Value = strDate & strCase & "　" & strCont & strExt
    End If
End Sub

'================================================================================
' プロシージャ: InputSCFor
' 機能:         複製作成先（担当区分）を自動で入力します。
' 概要:         事件呼名をキーにして保管先情報を検索し、担当区分を入力します。
'================================================================================
Private Sub InputSCFor(r As Long)
    ' SC作成先が入力済みの場合は終了する
    If Range("保管[複製作成先]")(r).Value <> "" Then Exit Sub

    ' SC作成先を入力する
    Dim t As Long
    With Sht2Hokansaki
        If WorksheetFunction.CountIf(.Range("保管先[事件呼名]"), Range("保管[事件呼名]")(r).Value) > 0 Then
            t = WorksheetFunction.Match(Range("保管[事件呼名]")(r).Value, .Range("保管先[事件呼名]"), 0)
            Range("保管[複製作成先]")(r).Value = .Range("保管先[担当区分]")(t).Value
        End If
    End With
End Sub

'================================================================================
' プロシージャ: DeleteInvalidData
' 機能:         保管未了データのうち、対応する読込ファイルが存在しない無効なデータを削除します。
' 概要:         テーブルを逆順にループし、実ファイルがない行を削除します。
'================================================================================
Private Sub DeleteInvalidData()
    ' 保管未了で読込ファイルが存在しないデータを削除
    Dim rTmp As Long
    For rTmp = Range("保管").Rows.Count To 1 Step -1
        If Range("保管[保管日時]")(rTmp).Value = "" Then
            If Dir(Range("保管[読込フォルダ名]")(rTmp).Value & "\" & _
                Range("保管[読込ファイル名]")(rTmp).Value) = "" Then
                Range("保管").Rows(rTmp).EntireRow.Delete
            End If
        End If
    Next
End Sub

'================================================================================
' プロシージャ: SetPhonetic
' 機能:         「事件呼名」列のセルに、ふりがな情報を設定します。
' 概要:         Excelの機能を利用してふりがなをセットします。
'================================================================================
Sub SetPhonetic()
    Dim rngCell As Range
    For Each rngCell In Range("保管[事件呼名]").Cells
        If rngCell.Value <> "" Then
            rngCell.Phonetic.Text = Application.GetPhonetic(rngCell)
        End If
    Next
End Sub

'================================================================================
' プロシージャ: SplitText
' 機能:         改行コード（vbCrLf）で区切られた文字列から、指定された番目の要素を抽出します。
' 概要:         文字列を指定区切り文字で配列化し、インデックスで要素を取得します。
'================================================================================
Private Function SplitText(strText As String, intSplit As Integer)
    On Error GoTo HdlErr
    ' セル内のテキスト（strText）を区切り文字で区切って、何番目か（intSplit）のデータを抜き出す
    SplitText = Split(strText, vbCrLf)(intSplit - 1)
    Exit Function

HdlErr:
    ' 該当する文字列が存在しない場合は空白にする
    SplitText = ""
End Function

