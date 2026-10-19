Option Explicit
'変数管理クラスを格納する変数を宣言する
Dim Variables As clsVariablesManager

'================================================================================
' リボンコールバック (Ribbon Callbacks)
'================================================================================

' 機能: リボンから「セットアップ」を実行するコールバック
Sub rxShowSetupMenu(control As IRibbonControl)
    Call ShowSetupMenu
End Sub


'================================================================================
' メイン処理プロシージャ (Main Procedures)
'================================================================================

'================================================================================
' 機能:         セットアップ用のユーザーフォームを表示します。
' 概要:         文書の書式設定を行うためのメインウィンドウ（Setupフォーム）を
'               モーダルで表示します。ユーザーがフォームを閉じるまで、
'               後続のコードは実行されません。
'================================================================================
Sub ShowSetupMenu()
    
    'フォームを表示する
    frmSetup.Show
    
End Sub

'================================================================================
' 機能:         ユーザーフォームで設定された内容に基づき、文書全体の書式設定を
'               一括で実行します。
' 概要:         このプロシージャは、書式設定プロセスの中心的な役割を担います。
'               スタイル、ページレイアウト、ヘッダー/フッターなど、複数の
'               サブプロシージャを適切な順序で呼び出します。
'               処理の開始時に画面更新を停止し、終了時に元に戻すことで、
'               パフォーマンスとユーザー体験を向上させます。
' 引数:         blnDummy (Boolean, Optional):
'               Wordの「マクロ」ダイアログボックスにこのマクロが表示されるのを
'               防ぐためのダミー引数です。
'================================================================================
Sub ApplyFormatSettings(Optional blnDummy As Boolean)
    'エラーのトラップを有効にする
    On Error GoTo HandleErr
    
    '変数オブジェクトを生成する
    Set Variables = New clsVariablesManager
        
    Dim blnRestart As Boolean
    blnRestart = False  ' デフォルトは再起動しない
    
    '実行を確認する
    If MsgBox("選択した設定で、現在のリスト、スタイルおよびページ設定を上書します。" & vbCrLf & vbCrLf & _
        "よろしいですか？", vbOKCancel, "VA公用文") = vbCancel Then
        blnRestart = True   'フォームを再起動する
        GoTo Cleanup        'クリーンアップ処理へ
    End If

    'マウスポインタをウエイト表示にする
    Application.System.Cursor = wdCursorWait
    
    '画面表示を停止する
    Application.ScreenUpdating = False

    'マークダウンに変換する
    'アウトラインを再設定するため、一旦マークダウンに変換します。
    Call basControl.ChangeToMarkdown

    'サンプルの追加が指示されている場合
    If Variables.strSample = "Add" Then
        '進捗状況を表示する
        Application.StatusBar = "サンプルを追加しています。"
        
        'サンプルを削除・更新・追加する
        Call InsertSample
    End If

    '進捗状況を表示する
    Application.StatusBar = "スタイルを設定しています。"
    
    'スタイルを設定する
    'マークダウンへの変換よりも後に行わないとリストの種類が変更されます。
    'アウトラインのよりも先に行わないと番号の振り直しが解除されます。
    Call ConfigStyle

    '進捗状況を表示する
    Application.StatusBar = "アウトラインを設定しています。"
    
    'アウトライン設定を初期化（強制適用）する
    '新規文書など、見出しがまだ存在しない状態でも、後からスタイルを選んだ際に
    '番号が表示されるよう、リストテンプレートとスタイルのリンクを確立する。
    Call InitOutlineSettings

    'アウトラインを設定する
    'マークダウンをアウトラインに変換する
    Call basControl.ChangeToOutline

    '進捗状況を表示する
    Application.StatusBar = "ページ設定を行っています。"
    'ページを設定する
    'アウトラインの設定より後でなければ、文字送りが揃いません。
    Call ConfigPage

    '進捗状況を表示する
    Application.StatusBar = "ページ番号を設定しています。"
    'ページ番号を設定する
    Call ConfigHeaderFooter
    
    '進捗状況を表示する
    Application.StatusBar = "スタイルギャラリーを設定しています。"
    If blnConfigGallery = True Then
        'スタイルギャラリーから不要なスタイルを削除する
        Call ConfigStyleGallery
        
        'リストギャラリーから不要なスタイルを削除する
        Call ConfigListGalleries
    End If
    
Cleanup:
    'マウスポインタのウエイト表示を解除する
    Application.System.Cursor = wdCursorNormal
    '画面表示を再開する
    Application.ScreenUpdating = True
    '進捗状況を消去する
    Application.StatusBar = ""
    '変数オブジェクトを解放する
    Set Variables = Nothing
    
    'フォームを停止する
    Unload frmSetup
    
    'フラグが立っている場合のみフォームを再起動
    If blnRestart Then
        frmSetup.Show
    End If
    
    Exit Sub

HandleErr:
    Resume Cleanup ' CleanUpセクションにジャンプして後処理を実行
End Sub

'================================================================================
' 機能:         アウトライン設定の初期化（強制適用）
' 概要:         文書の最後の段落の直前に一時的な段落を挿入してアウトライン定義を適用し、
'               見出しスタイルとリストテンプレートのリンクを確立します。
'               処理後、挿入した一時的な段落は削除されます。
'               これにより、既存の段落（特に1行目）の書式への影響を防ぎます。
'================================================================================
Private Sub InitOutlineSettings()
    'エラー処理
    On Error GoTo HandleErr

    'リストフォーマッターの準備
    Dim ListFormatter As clsListFormatter
    Set ListFormatter = New clsListFormatter

    '最後の段落の直前に、一時的な空の段落を挿入する
    'Paragraphs.Addは、新しい段落ではなく、それまで最後だった段落を返すため使用しない
    '（返り値を削除すると、利用者の最後の段落が削除される。Ver8D1までの不具合）
    '最後の段落の直前に挿入した段落は、自身の段落記号を持つため、
    '削除しても前後の段落の文字と書式に影響しない（文書が表で終わる場合も同じ）。
    Dim parTemp As Paragraph
    ActiveDocument.Paragraphs.Last.Range.InsertParagraphBefore
    Set parTemp = ActiveDocument.Paragraphs.Last.Previous

    '一時的な段落は最後の段落の書式を引き継ぐため、リストとスタイルを解除する
    '（最後の段落がリストに含まれていると、アウトラインの適用がそのリスト全体に及ぶため）
    parTemp.Range.ListFormat.RemoveNumbers
    parTemp.Style = "標準"

    'アウトライン定義を適用（ここでスタイルとリストのリンクが確立される）
    'この操作により、文書全体で「見出し 1」～「見出し 9」とアウトライン番号の定義が紐づきます。
    ListFormatter.ConfigOutline parTemp

Cleanup:
    '一時的に挿入した段落を削除する（エラーで中断した場合も削除する）
    '段落記号だけであることを確かめてから削除し、利用者の文字を消さないようにする
    On Error Resume Next
    If Not parTemp Is Nothing Then
        If parTemp.Range.Text = vbCr Then parTemp.Range.Delete
    End If
    On Error GoTo 0
    Set parTemp = Nothing
    Set ListFormatter = Nothing
    Exit Sub

HandleErr:
    'エラーが発生してもメイン処理は続行させる
    Resume Cleanup
End Sub


'================================================================================
' 機能:         文書のページ設定（用紙サイズ、余白、文字数、行数など）を
'               行います。
' 概要:         アクティブな文書のページ設定を、グローバル変数オブジェクト
'               （Variables）に格納された値に基づいて構成します。
'               行番号の表示/非表示や、各種余白の設定もここで行います。
'================================================================================
Private Sub ConfigPage()
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    'アクティブドキュメントのページ設定が対象
    With ActiveDocument.PageSetup
        .LayoutMode = wdLayoutModeGrid '文字数と行数の指定（文字数や行数より前に設定しないとエラーになる場合があります。）
        With .LineNumbering
            '行番号の表示オプションに応じて処理を行う
            .Active = (Variables.strLineNumber = "Show")
            .StartingNumber = 1             '開始番号
            .CountBy = 5                    '行番号の増分
            .RestartMode = wdRestartPage    'ページごとに振り直し
            .DistanceFromText = wdAutoPosition '文字列との間隔
        End With
        .Orientation = wdOrientPortrait
        .TopMargin = MillimetersToPoints(Variables.varMargines(1))      '余白（上）（変数で設定）
        .BottomMargin = MillimetersToPoints(Variables.varMargines(2))   '余白（下）（変数で設定）
        .LeftMargin = MillimetersToPoints(Variables.varMargines(3))     '余白（左）（変数で設定）
        .RightMargin = MillimetersToPoints(Variables.varMargines(4))    '余白（右）（修正値）（変数で設定）
        .Gutter = MillimetersToPoints(0)
        .HeaderDistance = MillimetersToPoints(20)                       '用紙の端からの距離（ヘッダー）
        .FooterDistance = MillimetersToPoints(15)                       '用紙の端からの距離（フッター）
        .PageWidth = MillimetersToPoints(Variables.sngPageWidth)        '用紙サイズの幅（変数で設定）
        .PageHeight = MillimetersToPoints(297)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        ' .SectionStart = wdSectionNewPage 'ページ区切りが「次のページから」に切り替わるのでコメントアウト
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        ' .TwoPagesOnOne = False 'Mac用WordではVBAが停止するのでコメントアウト
        ' .BookFoldPrinting = False 'Mac用WordではVBAが停止するのでコメントアウト
        ' .BookFoldRevPrinting = False 'Mac用WordではVBAが停止するのでコメントアウト
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .CharsLine = Variables.varCharsAndLines(1) '文字数（変数で設定）
        .LinesPage = Variables.varCharsAndLines(2) '行数（変数で設定）
    End With

    'マークダウンの記号に「.」が用いられている場合は「...」のオートコレクトを解除する
    '利用者のWord全体の設定を変える処理のため、エントリを削除した場合に限り通知する
    '（エントリが既に無い場合は、エラーを無視して何もしない）
    If strHead = "." Then
        Dim blnDeleted As Boolean ' オートコレクトのエントリを削除したかどうか
        On Error Resume Next
        AutoCorrect.Entries("...").Delete
        blnDeleted = (Err.Number = 0)
        Err.Clear
        On Error GoTo HandleErr
        'エントリを削除した場合は通知する
        If blnDeleted Then
            MsgBox "マークダウン記号に「.」を使用するため、オートコレクトの「...」→「…」の自動変換を無効にしました。" & vbCr & _
                "元に戻す場合は、Wordのオプション→文章校正→オートコレクトのオプションで「...」を再登録してください。", _
                vbOKOnly + vbInformation, "VA公用文"
        End If
    End If

    'タブ設定の既定値を1字に設定する
    ActiveDocument.DefaultTabStop = Variables.sngAdvWidth   'アドバンス幅(変数で設定)
    Exit Sub

HandleErr:
    MsgBox "ページの設定でエラーが発生したため、処理を停止しました。", vbOKOnly + vbExclamation, "VA公用文"
End Sub

'================================================================================
' 機能:         文書で使用される基本的なスタイル（標準、見出し、表題など）の
'               書式を定義します。
' 概要:         「標準」スタイルを基準に、「見出し 1」から「見出し 9」までの
'               見出しスタイル、および「リスト段落」「表題」などの特定の
'               スタイルのフォントや段落書式を一括で設定します。
'================================================================================
Private Sub ConfigStyle()
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    '文頭の段落書式を保持
    'カーソルのある段落にはスタイルが直接設定されるため、インデントが解除されてしまいます。
    Dim parFirst As Paragraph
    Set parFirst = ActiveDocument.Paragraphs(1)
    
    Dim sngLeftIndentBefore As Single, sngFirstLineIndentBefore As Single, sngRightIndentBefore As Single
    Dim wdpAlignBefore As WdParagraphAlignment
    
    With parFirst.Format
        sngLeftIndentBefore = .LeftIndent
        sngFirstLineIndentBefore = .FirstLineIndent
        sngRightIndentBefore = .RightIndent
        wdpAlignBefore = .alignment
    End With

    'すべての見出しレベルを対象に処理する
    Dim x As Long            '見出しレベル
    Dim styTarget As Style  'スタイルオブジェクト
    
    For x = 1 To 9
        'スタイル名を設定する
        '組み込み定数（標準スタイル：-1、見出し 1：-2、など）も利用可能ですが、使用していません。（可読性を重視するため）
        Set styTarget = ActiveDocument.Styles("見出し " & x)    '「見出し」の後に半角スペースが必要です
        
        'アクティブドキュメントのスタイルが対象
        With styTarget
            .AutomaticallyUpdate = False 'スタイルの再定義
            .NextParagraphStyle = "標準" '次の段落のスタイル
            .BaseStyle = "標準" '基準にするスタイル
            With .Font
                .NameFarEast = strFont                  'アジア言語のフォント名（定数で設定）
                .NameAscii = strFont                    '英数字に使用するフォント（定数で設定）
                .NameOther = strFont                    'その他の文字に使用するフォント（定数で設定）
                .Name = strFont                         'フォント名（定数で設定）
                .Bold = False                           '太字
                .Size = sngFont                         'フォントサイズ（定数で設定）
                .Italic = False                         'イタリック
                .Underline = wdUnderlineNone            'アンダーライン
                .Outline = False                        '中抜き
                .Emboss = False                         'エンボス
                .Shadow = False                         '影付き
                .Hidden = False                         '隠し文字
                .SmallCaps = False                      '小型英大文字
                .Color = wdColorAutomatic               '色
                .Engrave = False                        '浮き彫り
                .Superscript = False                    '上付き文字
                .Subscript = False                      '下付き文字
                .Spacing = 0                            '文字間隔（標準）
                .Scaling = 100                          'フォント倍率
                .Kerning = 0                            'カーニング
                .DisableCharacterSpaceGrid = False      '1 行あたりの文字数
                .EmphasisMark = wdEmphasisMarkNone      '傍点
                .Ligatures = wdLigaturesNone            '合字
                .NumberForm = wdNumberFormDefault       '数値の表示形式
                .StylisticSet = wdStylisticSetDefault   'スタイル セット
                .ContextualAlternates = 0               'コンテキスト代替
            End With
            With .ParagraphFormat
                .SpaceBefore = 0                        '段落前の間隔
                .SpaceAfter = 0                         '段落後の間隔
                .LineSpacingRule = wdLineSpaceSingle    '行間
                .alignment = lngAlign                   '配置（定数で設定）
                .OutlineLevel = x                       'アウトラインレベル
                .WidowControl = False                   '1行残して段落を区切る
                .KeepWithNext = False                   '次の段落と分離しない
                .KeepTogether = False                   '段落を分割しない
                .PageBreakBefore = False                '段落前で改ページ
                .NoLineNumber = False                   '行番号
                .Hyphenation = True                     'ハイフネーション
                .FarEastLineBreakControl = True         '禁則処理
                .WordWrap = True                        'ワードラップ
                .HangingPunctuation = True              '句読点のぶら下げ
                .HalfWidthPunctuationOnTopOfLine = False '行頭の記号を半角
                .AddSpaceBetweenFarEastAndAlpha = False  '日本語と英字の間隔を自動調整
                .AddSpaceBetweenFarEastAndDigit = False  '日本語と数字の間隔を自動調整
                .BaseLineAlignment = wdBaselineAlignAuto '文字の配置
                
                .LeftIndent = 0                          '左インデント(リセット)
                .RightIndent = 0                         '右インデント(リセット)
                .FirstLineIndent = 0                     '最初の行のインデント(リセット)
                .CharacterUnitLeftIndent = 0             '左インデント（文字数単位）(リセット)
                .CharacterUnitRightIndent = 0            '右インデント（文字数単位）(リセット)
                .CharacterUnitFirstLineIndent = 0        '最初の行のインデント（文字数単位）(リセット)
            End With
        End With
    Next x

    '標準スタイルを対象に処理する
    Set styTarget = ActiveDocument.Styles("標準")
    Call ConfigureStandardStyle(styTarget)
    
    '「リスト段落」スタイルを対象に処理する
    Set styTarget = ActiveDocument.Styles("リスト段落")
    Call ConfigureStandardStyle(styTarget)
    
    With styTarget
        .NextParagraphStyle = "リスト段落"
        ' スタイルにリンクされたリストテンプレートのインデント設定をリセット
        'リスト段落は、この処理を行わないとスタイルのインデント設定値が35mmになります。
        On Error Resume Next ' リストテンプレートが未設定の場合のエラーを回避
        Dim levlNumber As ListLevel
        Set levlNumber = .ListTemplate.ListLevels(.ListLevelNumber)
        If Err.Number = 0 Then
            ' 取得できた場合のみ、そのレベルのインデント関連設定を0にする
            With levlNumber
                .NumberFormat = ""        ' 番号/記号をなしに
                .NumberStyle = wdListNumberStyleNone
                .NumberPosition = 0       ' 番号の位置を 0
                .alignment = wdListLevelAlignLeft
                .TextPosition = 0         ' テキストの位置（インデント）を 0
                .TabPosition = 0          ' タブ位置を 0
            End With
        End If
        On Error GoTo HandleErr
    End With
    
    '「行間詰め」スタイルを対象に処理する
    'スタイルにリンクされたリストテンプレートのインデント設定をリセットしないと
    'スタイルのインデント設定値が35mmになります。
    'ただし、リセットすると最初の段落番号の設定がうまくいかなくなるので、
    'そのままにしています。
    Set styTarget = ActiveDocument.Styles("行間詰め")
    Call ConfigureStandardStyle(styTarget)
    With styTarget
         .NextParagraphStyle = "行間詰め" '次の段落のスタイル
    End With
    
    '「表題」スタイルを対象に処理する
    Set styTarget = ActiveDocument.Styles("表題")
    Call ConfigureStandardStyle(styTarget)
    
    With styTarget.ParagraphFormat
        .alignment = Variables.lngTitleAlign                                '配置（変数で設定）
        .LeftIndent = Variables.sngAdvWidth * Variables.lngTitleIndents     '左インデント（変数で設定）
        .RightIndent = Variables.sngAdvWidth * Variables.lngTitleIndents    '右インデント（左インデントと同一に設定）
    End With

    '「目次の見出し」スタイルを対象に処理する
    Set styTarget = ActiveDocument.Styles("目次の見出し")
    Call ConfigureStandardStyle(styTarget)

    '「行番号」スタイルを対象に処理する
    Set styTarget = ActiveDocument.Styles("行番号")
    With styTarget.Font
        .NameFarEast = strFont
        .Name = strFont
        .Size = sngLineNo
    End With

    'スタイルオブジェクトを解放する
    Set styTarget = Nothing

    '文字間隔の調整を設定する
    ActiveDocument.JustificationMode = lngJust

    '文頭の段落のインデントおよび配置を再適用する
    With parFirst.Format
        .LeftIndent = sngLeftIndentBefore
        .FirstLineIndent = sngFirstLineIndentBefore
        .RightIndent = sngRightIndentBefore
        .alignment = wdpAlignBefore
    End With

    Exit Sub

HandleErr:
    MsgBox "スタイルの設定でエラーが発生したため、処理を停止しました。", vbOKOnly + vbExclamation, "VA公用文"
End Sub

'================================================================================
' 機能:         指定されたスタイルに対し、共通のフォントおよび段落書式を設定します。
' 概要:         「標準」、「リスト段落」、「表題」など、多くのスタイルで共通する
'               基本設定（フォント名、サイズ、太字・斜体の解除、インデントの
'               リセット、行間設定など）を一括で適用します。
'               これにより、コードの重複を排除し、保守性を向上させます。
' 引数:         sty (Style):
'                   設定対象のスタイルオブジェクト。
'================================================================================
Private Sub ConfigureStandardStyle(ByVal sty As Style)
    With sty
        .AutomaticallyUpdate = False            'スタイルの再定義
        .NextParagraphStyle = "標準"            '次の段落のスタイル
        
        '「標準」スタイル自身にBaseStyleを設定するとエラーになるため除外
        If .NameLocal <> "標準" Then
            .BaseStyle = "標準"                 '基準にするスタイル
        End If
        
        With .Font
            .NameFarEast = strFont              'アジア言語のフォント名（定数で設定）
            .NameAscii = strFont                '英数字に使用するフォント（定数で設定）
            .NameOther = strFont                'その他の文字に使用するフォント（定数で設定）
            .Name = strFont                     'フォント名（定数で設定）
            .Bold = False                       '太字
            .Size = sngFont                     'フォントサイズ（定数で設定）
            .Italic = False                     'イタリック
            .Underline = wdUnderlineNone        'アンダーライン
            .Outline = False                    '中抜き
            .Emboss = False                     'エンボス
            .Shadow = False                     '影付き
            .Hidden = False                     '隠し文字
            .SmallCaps = False                  '小型英大文字
            .Color = wdColorAutomatic           '色
            .Engrave = False                    '浮き彫り
            .Superscript = False                '上付き文字
            .Subscript = False                  '下付き文字
            .Spacing = 0                        '文字間隔（標準）
            .Scaling = 100                      'フォント倍率
            .Kerning = 0                        'カーニング
            .DisableCharacterSpaceGrid = False  '1 行あたりの文字数
            .EmphasisMark = wdEmphasisMarkNone  '傍点
            .Ligatures = wdLigaturesNone        '合字
            .NumberForm = wdNumberFormDefault   '数値の表示形式
            .StylisticSet = wdStylisticSetDefault 'スタイル セット
            .ContextualAlternates = 0           'コンテキスト代替
        End With
        With .ParagraphFormat
            .CharacterUnitLeftIndent = 0        '左インデント（文字数単位）(一括して設定するのでリセットしています）
            .CharacterUnitRightIndent = 0       '右インデント（文字数単位）(一括して設定するのでリセットしています）
            .CharacterUnitFirstLineIndent = 0   '最初の行のインデント（文字数単位）(一括して設定するのでリセットしています）
            .LeftIndent = 0                     '左インデント(一括して設定するのでリセットしています）
            .RightIndent = 0                    '右インデント(一括して設定するのでリセットしています）
            .FirstLineIndent = 0                '最初の行のインデント(一括して設定するのでリセットしています）
            .SpaceBefore = 0                    '段落前の間隔
            .SpaceAfter = 0                     '段落後の間隔
            .LineSpacingRule = wdLineSpaceSingle '行間
            .alignment = lngAlign               '配置（定数で設定）
            .OutlineLevel = 10                  'アウトラインレベル(wdOutlineLevelBodyText)
            .WidowControl = False               '1行残して段落を区切る（どのレベルまで設定するか決められないので設定していません）
            .KeepWithNext = False               '次の段落と分離しない（どのレベルまで設定するか決められないので設定していません）
            .KeepTogether = False               '段落を分割しない（どのレベルまで設定するか決められないので設定していません）
            .PageBreakBefore = False            '段落前で改ページ（どのレベルまで設定するか決められないので設定していません）
            .NoLineNumber = False               '行番号
            .Hyphenation = True                 'ハイフネーション
            .FarEastLineBreakControl = True     '禁則処理
            .WordWrap = True                    'ワードラップ
            .HangingPunctuation = True          '句読点のぶら下げ
            .HalfWidthPunctuationOnTopOfLine = False '行頭の記号を半角
            .AddSpaceBetweenFarEastAndAlpha = False  '日本語と英字の間隔を自動調整
            .AddSpaceBetweenFarEastAndDigit = False  '日本語と数字の間隔を自動調整
            .BaseLineAlignment = wdBaselineAlignAuto '文字の配置
        End With
    End With
End Sub

Private Sub ConfigHeaderFooter()
    '--- 変数宣言 ---
    Dim sec As Section
    Dim rng As Range
    Dim hf As HeaderFooter
    Dim targetHFType As WdHeaderFooterIndex
    Dim targetAlign As WdParagraphAlignment
    Dim blnProcess As Boolean
    Dim isLinked As Boolean
    
    'エラーのトラップ
    On Error GoTo HandleErr

    'すべてのセクションを繰り返す
    For Each sec In ActiveDocument.Sections
        ' ヘッダー
        Set hf = sec.Headers(wdHeaderFooterPrimary)
        If sec.Index = 1 Or Not hf.LinkToPrevious Then
            If hf.Range.Text <> vbCr Then hf.Range.Delete ' 空の改行のみでない場合削除
        End If
        
        ' フッター
        Set hf = sec.Footers(wdHeaderFooterPrimary)
        If sec.Index = 1 Or Not hf.LinkToPrevious Then
            If hf.Range.Text <> vbCr Then hf.Range.Delete
        End If
    Next sec
    
    '--- 設定値の準備 ---
    blnProcess = True
    Select Case Variables.strPageNumber
        Case "TopRight"
            targetHFType = wdHeaderFooterPrimary
            targetAlign = wdAlignParagraphRight
            
        Case "BottomCenter", "BottomCenterDecorated"
            targetHFType = wdHeaderFooterPrimary ' フッターとして処理
            targetAlign = wdAlignParagraphCenter
            
        Case "BottomRight"
            targetHFType = wdHeaderFooterPrimary ' フッターとして処理
            targetAlign = wdAlignParagraphRight
            
        Case Else
            blnProcess = False
    End Select

    '--- ページ番号の挿入処理 ---
    If blnProcess Then
        'すべてのセクションを繰り返す
        For Each sec In ActiveDocument.Sections
            ' ヘッダーかフッターかを取得
            If Variables.strPageNumber = "TopRight" Then
                Set hf = sec.Headers(targetHFType)
            Else
                Set hf = sec.Footers(targetHFType)
            End If
            
            ' リンク設定を確認（2セクション目以降で、前と同じ設定なら処理不要）
            isLinked = False
            If sec.Index > 1 Then
                If hf.LinkToPrevious Then isLinked = True
            End If
            
            ' リンクされていない（＝設定の実体がある）場合のみ書き込む
            If Not isLinked Then
                Set rng = hf.Range
                
                ' 書式設定
                With rng
                    .ParagraphFormat.alignment = targetAlign
                    .Font.Name = strFont
                    .Font.Size = sngFont
                    ' 既存の改行などが残っている場合への安全策（基本は空のはず）
                    .Collapse Direction:=wdCollapseStart
                End With
                
                ' ページ番号設定（開始番号など）
                With hf.PageNumbers
                    .RestartNumberingAtSection = False
                    .StartingNumber = 1
                End With

                ' フィールドの挿入
                Select Case Variables.strPageNumber
                    Case "TopRight"
                        ' PAGE \* DBCHAR
                        rng.Fields.Add Range:=rng, Type:=wdFieldEmpty, Text:="PAGE  \* DBCHAR", PreserveFormatting:=True
                        
                        ' 最終ページ判定 (IF PAGE = NUMPAGES)
                        If blnPageNumberEnd = True Then
                            ' カーソル移動の代わりにRangeを末尾へ
                            Set rng = hf.Range
                            rng.Collapse Direction:=wdCollapseEnd
                            
                            ' RangeでIFフィールドを構築するのは複雑なため、
                            ' ここでは最小限のAdd Fieldメソッドチェーンで構築します
                            Dim fldIf As Field
                            Dim rngCode As Range
                            
                            ' IFフィールド枠を作成
                            Set fldIf = rng.Fields.Add(Range:=rng, Type:=wdFieldEmpty, Text:="IF ", PreserveFormatting:=True)
                            Set rngCode = fldIf.Code ' { IF } の中身
                            
                            ' "IF " の後ろへ移動
                            rngCode.End = rngCode.End - 1 ' 末尾スペース回避
                            rngCode.Collapse Direction:=wdCollapseEnd
                            
                            ' PAGE
                            rngCode.Fields.Add Range:=rngCode, Type:=wdFieldPage
                            
                            ' = NUMPAGES "」"
                            Set rngCode = fldIf.Code
                            rngCode.End = rngCode.End - 1
                            rngCode.Collapse Direction:=wdCollapseEnd
                            rngCode.InsertAfter " = "
                            
                            rngCode.Collapse Direction:=wdCollapseEnd
                            rngCode.Fields.Add Range:=rngCode, Type:=wdFieldNumPages
                            
                            Set rngCode = fldIf.Code
                            rngCode.End = rngCode.End - 1
                            rngCode.Collapse Direction:=wdCollapseEnd
                            rngCode.InsertAfter " ""」"" """""
                            
                            fldIf.Update
                        End If

                    Case "BottomCenter"
                        rng.Fields.Add Range:=rng, Type:=wdFieldPage

                    Case "BottomCenterDecorated"
                        ' "- 1 -" の形式
                        rng.InsertAfter "- "
                        rng.Collapse Direction:=wdCollapseEnd
                        rng.Fields.Add Range:=rng, Type:=wdFieldPage
                        ' 追加書き込みのためにRange再取得
                        Set rng = hf.Range
                        rng.Collapse Direction:=wdCollapseEnd
                        rng.InsertBefore " -"

                    Case "BottomRight"
                        ' "1/10" の形式
                        rng.Fields.Add Range:=rng, Type:=wdFieldPage
                        Set rng = hf.Range
                        rng.Collapse Direction:=wdCollapseEnd
                        rng.InsertBefore "/"
                        rng.Collapse Direction:=wdCollapseEnd
                        rng.Fields.Add Range:=rng, Type:=wdFieldNumPages
                End Select
            End If
        Next sec
    End If

Cleanup:
    '--- 終了処理 ---
    On Error Resume Next
    Set rng = Nothing
    Set sec = Nothing
    Set hf = Nothing
    Exit Sub

HandleErr:
    MsgBox "ヘッダー及びフッターの設定でエラーが発生したため、処理を停止しました。", vbOKOnly + vbExclamation, "VA公用文"
    Resume Cleanup
End Sub

'================================================================================
' 機能:         文書の末尾に、書式設定のサンプルとなるテキストブロックを
'               挿入または更新します。
' 概要:         文書内に特定の目印（"＜＜サンプル＞＞"）を持つサンプルテキストが
'               既に存在するかを検索します。存在する場合は、ユーザーに確認の上で
'               それを削除し、常に最新のサンプルテキストを文書の末尾に
'               挿入します。
'================================================================================
Private Sub InsertSample()
    'エラーのトラップを有効にする
    On Error GoTo HandleErr

    'サンプルの開始行を把握する
    Dim rngSearch As Range
    Set rngSearch = ActiveDocument.Content
    
    Dim blnSampleExist As Boolean
    
    '検索実行
    With rngSearch.Find
        .Text = "＜＜サンプル＞＞" & vbCr
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        If .Execute Then
            blnSampleExist = True
        End If
    End With
    
    'サンプルが存在する場合
    If blnSampleExist Then
        '実行を確認する
        If MsgBox("現在のサンプルを上書きします。" & vbCrLf & vbCrLf & _
            "よろしいですか？", vbOKCancel, "VA公用文") = vbCancel Then Exit Sub
            
        '既存のサンプルを削除する
        Dim rngDelete As Range
        Set rngDelete = ActiveDocument.Range(rngSearch.Start, ActiveDocument.Content.End)
        rngDelete.Delete
    End If
    
    'サンプルを追加する
    If Variables.strSample = "Add" Then
        'サンプルを挿入する空の段落を、文書の末尾に用意する
        'Paragraphs.Addは、新しい段落ではなく、それまで最後だった段落を返すため使用しない
        '（返り値の文字を書き換えると、利用者の最後の段落が上書きされる。Ver8D1までの不具合）
        Dim rngSample As Range
        Set rngSample = ActiveDocument.Paragraphs.Last.Range
        '最後の段落に文字がある場合は、その後ろに空の段落を追加する
        If rngSample.Text <> vbCr Then
            ActiveDocument.Content.InsertParagraphAfter
            Set rngSample = ActiveDocument.Paragraphs.Last.Range
        End If
        With rngSample
            'スタイルを標準に変更する
            .Style = "標準"
            '見出しの記号を置き換えてからサンプルを挿入する
            .InsertBefore Replace(Variables.strSampleContents, "#", strHead)
        End With
    End If
    
    Exit Sub

HandleErr:
    MsgBox "サンプルの追加でエラーが発生したため、処理を停止しました。", vbOKOnly + vbExclamation, "VA公用文"
End Sub

'================================================================================
' 機能: スタイル ギャラリーに表示されるスタイルを
'       公用文に必要なものだけに限定する。
'
'概要： アクティブなドキュメントの全スタイルを繰り返し処理します。
'       "標準", "行間詰め", "表題", "見出し 1"～"見出し 9" のスタイルを
'       ギャラリーに表示（QuickStyle = True）し、表示優先度（Priority）を設定します。
'       上記の許可リストに含まれない他のすべてのスタイルを
'       ギャラリーから非表示（QuickStyle = False）にします。
'================================================================================
Private Sub ConfigStyleGallery()
    ' スタイルが存在しない、またはQuickStyleプロパティを
    ' 持たない場合に発生するエラーを無視する
    On Error Resume Next
    
    Dim sty As Style
    ' 現在アクティブなドキュメントのすべてのスタイルを繰り返す
    For Each sty In ActiveDocument.Styles
        ' --- スタイルの種類をチェック ---
        ' 段落スタイル、またはリンクスタイル(段落と文字)のみを対象とする
        ' (表スタイルやリストスタイルはQuickStyleプロパティを持たないため)
        If sty.Type = wdStyleTypeParagraph Or sty.Type = wdStyleTypeLinked Then
            ' sty.NameLocal (ローカライズされた名前、つまり日本語名) で判断
            Select Case sty.NameLocal
                ' --- 許可リスト (表示するもの) ---
                Case "標準":      sty.QuickStyle = True: sty.Priority = 1 ' 常に1番目に表示
                Case "行間詰め":  sty.QuickStyle = True: sty.Priority = 2 ' 常に2番目に表示
                Case "見出し 1":  sty.QuickStyle = True: sty.Priority = 11 ' 見出しグループは11番目から
                Case "見出し 2":  sty.QuickStyle = True: sty.Priority = 12
                Case "見出し 3":  sty.QuickStyle = True: sty.Priority = 13
                Case "見出し 4":  sty.QuickStyle = True: sty.Priority = 14
                Case "見出し 5":  sty.QuickStyle = True: sty.Priority = 15
                Case "見出し 6":  sty.QuickStyle = True: sty.Priority = 16
                Case "見出し 7":  sty.QuickStyle = True: sty.Priority = 17
                Case "見出し 8":  sty.QuickStyle = True: sty.Priority = 18
                Case "見出し 9":  sty.QuickStyle = True: sty.Priority = 19
                Case "リスト段落":  sty.QuickStyle = True: sty.Priority = 20
                Case "表題":      sty.QuickStyle = True: sty.Priority = 21
                ' --- 拒否リスト (非表示にするもの) ---
                Case Else:        sty.QuickStyle = False ' 上記の許可リストに含まれないスタイルはすべて非表示
            End Select
        Else
            ' 段落・リンクスタイル以外 (文字, 表, リスト)
            ' これらもギャラリーから非表示にする
            sty.QuickStyle = False
        End If
    Next sty
    
    ' エラーハンドリングを通常に戻す
    On Error GoTo 0
End Sub

'================================================================================
' 機能:         リスト ギャラリー（箇条書き、段落番号、アウトライン）に
'               表示されるテンプレートを整理します。
' 概要:         「VA公用文」で使用するテンプレート（各ギャラリーの1番目）以外を
'               ギャラリーから非表示にする（既定の状態にリセットする）ことで、
'               ユーザーが必要なスタイルのみを選択できるようにします。
'               なお、リストは規定のテンプレートを非表示にすることはできません。
'================================================================================
Private Sub ConfigListGalleries()
    ' エラーが発生しても処理を続行する
    ' (存在しないインデックスをリセットしようとした場合などを考慮)
    On Error Resume Next
    Dim i As Long
    
    ' --- 箇条書きライブラリ (wdBulletGallery) ---
    ' テンプレート 2 から 7 までをリセット
    For i = 2 To 7
        ListGalleries(wdBulletGallery).Reset Index:=i
        ' --- 段落番号ライブラリ (wdNumberGallery) ---
        ' テンプレート 2 から 7 までをリセット
        ListGalleries(wdNumberGallery).Reset Index:=i
        ' --- アウトライン ライブラリ (wdOutlineNumberGallery) ---
        ' テンプレート 2 から 7 までをリセット
        ListGalleries(wdOutlineNumberGallery).Reset Index:=i
    Next i
    
    ' エラーハンドリングを通常に戻す
    On Error GoTo 0
End Sub
