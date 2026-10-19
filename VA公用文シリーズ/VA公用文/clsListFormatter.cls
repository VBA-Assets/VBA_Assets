Option Explicit
'================================================================================
' クラス名:     clsListFormatter
' 概要:         箇条書き、段落番号、アウトラインの書式定義と適用を行う
'================================================================================

'変数管理クラスを格納する変数を宣言する
Private Variables As clsVariablesManager

'================================================================================
' イベント: クラス初期化
'================================================================================
'クラス初期化時のプロシージャ
Private Sub Class_Initialize()
    '変数管理クラスのインスタンスを生成する
    Set Variables = New clsVariablesManager
'クラス初期化時のプロシージャを終了する
End Sub

'================================================================================
' イベント: クラス終了
'================================================================================
'クラス終了時のプロシージャ
Private Sub Class_Terminate()
    '変数管理クラスのインスタンスを解放する
    Set Variables = Nothing
'クラス終了時のプロシージャを終了する
End Sub

'==================================================================================================
' 概要:         指定された段落に対し、箇条書きまたは段落番号のリスト書式を定義・適用します。
'               前の段落がリストでない場合は新しいリストを開始し、連続している場合は番号を継続します。
' 引数:         strGallery (String) - 適用するリストの種類（wdBulletGallery または wdNumberGallery）。
' 引数:         parTarget (Paragraph) - 書式を適用する段落オブジェクト。
'==================================================================================================
'リストを設定するプロシージャ
Public Sub ConfigList(strGallery As String, parTarget As Paragraph)
    
    '#リストを設定する
    'リストギャラリーが対象
    With ListGalleries(strGallery).ListTemplates(1).ListLevels(1)
        '引数に応じて設定する
        Select Case strGallery
            '箇条書きの場合
            Case wdBulletGallery
                '記号を設定する
                .NumberFormat = Variables.strListBullFormat
                '記号の種類を設定する
                .NumberStyle = wdListNumberStyleBullet
            '段落番号の場合
            Case wdNumberGallery
                '記号を設定する
                .NumberFormat = Variables.strListNumbFormat
                '記号の種類を設定する
                .NumberStyle = Variables.lngListNumbStyle
        '条件分岐を終了する
        End Select
        
        '番号に続く空白の扱いを設定する
        .TrailingCharacter = wdTrailingNone
        '左インデントからの距離を設定する（一括して設定するのでリセットしています。）
        .NumberPosition = 0
        '番号の整列を設定する（スタイルでインデント調節ができないため左寄せにしています。）
        .alignment = wdListLevelAlignLeft
        'インデント位置を設定する(一括して設定するのでリセットしています。）
        .TextPosition = 0
        'リストを開始するレベルを設定する
        .ResetOnHigher = 0
        '開始番号を設定する
        .StartAt = 1
        'リンクするスタイルを設定する
        .LinkedStyle = "リスト段落"
        
        'フォントを対象
        With .Font
            '太字を解除する
            .Bold = False
            'イタリックを解除する
            .Italic = False
            '記号のフォントサイズを設定する（「リスト段落」のスタイルに合わせる）
            .Size = ActiveDocument.Styles("リスト段落").Font.Size
            '記号のフォント名を設定する（「リスト段落」のスタイルに合わせる）
            .Name = ActiveDocument.Styles("リスト段落").Font.Name
        'フォントの設定を終了する
        End With
    'リストギャラリーの設定を終了する
    End With

    '#リスト名を記録する
    'リストライブラリにVA公用文のテンプレートがあるかどうかを確認できるようにします。
    ListGalleries(strGallery).ListTemplates(1).Name = "VA公用文"
    
    '#番号を継続するかどうかを判断する
    '番号継続の有無を格納する変数を宣言する
    Dim blnContinue As Boolean
    '前の段落を格納する変数を宣言する
    Dim parPrevious As Paragraph
    
    '前の段落を取得する
    Set parPrevious = parTarget.Previous
    
    '前の段落が存在するかどうかを判定する
    If Not parPrevious Is Nothing Then
        '前の段落がリスト段落かどうかを判定する
        If parPrevious.Style = "リスト段落" Then
            '番号を継続する
            blnContinue = True
        'リスト段落以外の場合
        Else
            '番号を継続しない
            blnContinue = False
        '条件判定を終了する
        End If
    '1行目の場合
    Else
        '番号を継続しない
        blnContinue = False
    '条件判定を終了する
    End If

    '#リストを適用する
    'リストは、いずれかの段落に適用しないと有効になりません。
    parTarget.Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=ListGalleries(strGallery).ListTemplates(1), _
        ContinuePreviousList:=blnContinue, _
        ApplyTo:=wdListApplyToSelection, _
        DefaultListBehavior:=wdWord10ListBehavior
        
    '前の段落の変数を解放する
    Set parPrevious = Nothing
'リストを設定するプロシージャを終了する
End Sub

'==================================================================================================
' 概要:         見出し1から9のアウトライン書式を一括で定義し、指定された段落に適用して
'               新しいアウトラインリストを開始します。
' 引数:         parTarget (Paragraph) - 書式を適用する段落オブジェクト。
'==================================================================================================
'アウトラインを設定するプロシージャ
Public Sub ConfigOutline(parTarget As Paragraph)
    
    '番号と段落までの間隔は、番号の末尾に空白文字を入れることで行っています。
    '左揃えにしたほうが、インデントを解除した時の処理が簡単なのですが、文字間隔を揃えることを優先しています。
    'アウトライン番号を格納する変数を宣言する
    Dim x As Long

    '#アウトラインを設定する
    '見出しレベル1から9までのアウトラインを順次に設定します。
    For x = 1 To 9
        'リストギャラリーが対象
        With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(x)
            '番号の後に挿入する文字を設定する
            .TrailingCharacter = wdTrailingNone
            '番号書式（変数で設定）を設定する
            .NumberFormat = Variables.varNumberFormats(x)
            '番号のスタイル（変数で設定）を設定する
            .NumberStyle = Variables.varNumberStyles(x)
            '行の配置を設定する（左揃えでは全角数字と「（」を用いた際に文字間隔が揃いません。）
            .alignment = wdListLevelAlignLeft
            '対応付ける見出しスタイルを設定する
            .LinkedStyle = "見出し " & x
            'リストを開始するレベル（変数で設定）を設定する
            .ResetOnHigher = Variables.varResetOnHighers(x)
            '開始番号（変数で設定）を設定する
            .StartAt = Variables.varStartAts(x)
            'インデント位置を設定する(一括して設定するのでリセットしています）
            .TextPosition = 0
            
            '左インデントからの距離を設定する
            .NumberPosition = 0
            
            'フォントを対象
            With .Font
                '番号のフォント名を設定する（対応スタイルに合わせる）
                .Name = ActiveDocument.Styles("見出し " & x).Font.Name
                '太字を解除する
                .Bold = False
                'イタリックを解除する
                .Italic = False
                '取り消し線を未定義にする
                .StrikeThrough = wdUndefined
                '下付き文字を未定義にする
                .Subscript = wdUndefined
                '上付き文字を未定義にする
                .Superscript = wdUndefined
                '影付きを未定義にする
                .Shadow = wdUndefined
                '中抜きを未定義にする
                .Outline = wdUndefined
                'エンボスを未定義にする
                .Emboss = wdUndefined
                '浮き彫りを未定義にする
                .Engrave = wdUndefined
                'すべて大文字を未定義にする
                .AllCaps = wdUndefined
                '隠し文字を未定義にする
                .Hidden = wdUndefined
                '下線を解除する
                .Underline = wdUnderlineNone
                '色を未定義にする
                .Color = wdUndefined
                '番号のフォントサイズを設定する（対応スタイルに合わせる）
                .Size = ActiveDocument.Styles("見出し " & x).Font.Size
                'アニメーションを未定義にする
                .Animation = wdUndefined
                '二重取り消し線を未定義にする
                .DoubleStrikeThrough = wdUndefined
            'フォントの設定を終了する
            End With
        'リストギャラリーの設定を終了する
        End With
    '次のアウトラインレベルへ
    Next x

    '#リスト名を記録する
    'リストライブラリにVA公用文のテンプレートがあるかどうかを確認できるようにします。
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = "VA公用文"

    '#アウトラインを適用する
    'アウトラインは、いずれかの段落に適用しないと有効になりません。
    'ContinuePreviousList=Falseとすることで新しいリストを開始しています。
    parTarget.Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=False, _
        ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior
'アウトラインを設定するプロシージャを終了する
End Sub

