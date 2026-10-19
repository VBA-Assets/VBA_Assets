Option Explicit
'================================================================================
' クラスモジュール: clsVariablesManager
' 機能: アプリケーションの設定と関連変数を一元管理する
'================================================================================

'アウトラインおよびインデント
'番号書式を格納する変数を宣言する
Dim strNumberFormats As String
'番号スタイルを格納する変数を宣言する
Dim strNumberStyles As String
'リスト開始レベルを格納する変数を宣言する
Dim strResetOnHighers As String
'開始番号を格納する変数を宣言する
Dim strStartAts As String
'見出しの左インデントを格納する変数を宣言する
Dim strHeadingLeftIndents As String
'見出しの最初の行のインデントを格納する変数を宣言する
Dim strHeadingFirstLineIndents As String
'本文の左インデントを格納する変数を宣言する
Dim strTextLeftIndents As String
'本文の最初の行のインデントを格納する変数を宣言する
Dim strTextFirstLineIndents As String
'左インデント増加量を格納する変数を宣言する
Public lngLeftIndentsAdd As Long
'表題の配置を格納する変数を宣言する
Public lngTitleAlign As Long
'表題のインデントを格納する変数を宣言する
Public lngTitleIndents As Long

'リスト
'箇条書きの記号を格納する変数を宣言する
Public strListBullFormat As String
'段落番号の記号を格納する変数を宣言する
Public strListNumbFormat As String
'段落番号のスタイルを格納する変数を宣言する
Public lngListNumbStyle As Long
'見出しからの左インデント増加量を格納する変数を宣言する
Public lngListLeftIndent As Long
'最初の行のインデントを格納する変数を宣言する
Public lngListFirstLineIndent As Long

'ページ設定
'用紙サイズの幅（mm）を格納する変数を宣言する
Public sngPageWidth As Single
'ページ余白（mm）（上,下,左,右）を格納する変数を宣言する
Dim strMargines As String
'文字数及び行数（文字）（文字数,行数）を格納する変数を宣言する
Dim strCharsAndLines As String
'アドバンス幅を格納する変数を宣言する
Public sngAdvWidth As Single

'セットアップ
'文書の種類を格納する変数を宣言する
Public strDocumentType As String
'ページ番号の表示を格納する変数を宣言する
Public strPageNumber As String
'行番号の表示を格納する変数を宣言する
Public strLineNumber As String
'サンプルの追加を格納する変数を宣言する
Public strSample As String
'サンプル用文字列を格納する変数を宣言する
Public strSampleContents As String

'クラス変数
'番号書式設定値の配列を宣言する
Public varNumberFormats As Variant
'番号スタイル設定値の配列を宣言する
Public varNumberStyles As Variant
'見出しの左インデント設定値の配列を宣言する
Public varHeadingLeftIndents As Variant
'見出しの最初の行のインデント設定値の配列を宣言する
Public varHeadingFirstLineIndents As Variant
'本文の左インデント設定値の配列を宣言する
Public varTextLeftIndents As Variant
'本文の最初の行のインデント設定値の配列を宣言する
Public varTextFirstLineIndents As Variant
'ページ余白設定値の配列を宣言する
Public varMargines As Variant
'文字数および行数設定値の配列を宣言する
Public varCharsAndLines As Variant
'リスト開始レベルの配列を宣言する
Public varResetOnHighers As Variant
'開始番号の配列を宣言する
Public varStartAts As Variant

'================================================================================
' イベント: このクラスのインスタンスが作成されたときに自動実行される
'================================================================================
'クラスの初期化プロシージャ
Private Sub Class_Initialize()
    '変数の設定と配列の初期化を自動的に行う
    SetVariables
'クラスの初期化プロシージャを終了する
End Sub

'================================================================================
' 機能: 各種変数を設定する（内部処理用）
'================================================================================
'各種変数を設定するプロシージャ
Sub SetVariables()
    'イニシャライズによる自動実行は、起動しない場合があります。
    '設定のキー名配列を宣言する
    Dim varKeys() As Variant
    '設定の値配列を宣言する
    Dim varValues() As Variant
    'カウンタ変数を宣言する
    Dim i As Long

    '文書の種類をデフォルト値で初期化する
    strDocumentType = "Kouyou"
    'ページ番号の表示をデフォルト値で初期化する
    strPageNumber = "BottomCenter"
    '行番号の表示をデフォルト値で初期化する
    strLineNumber = "Hide"
    'サンプルの追加をデフォルト値で初期化する
    strSample = "DoNotAdd"

    '設定ファイル管理クラスのインスタンスを生成する
    Dim settingsMgr As New clsSettingsManager
    'ファイルから設定を読み込む
    settingsMgr.LoadSettingsFromProperty varKeys, varValues
    '設定ファイル管理クラスのインスタンスを解放する
    Set settingsMgr = Nothing

    'エラーが発生しても次の行へ進む
    On Error Resume Next
    '配列が空でないか確認する
    If UBound(varKeys) > 0 Then
        '配列の要素数分ループする
        For i = LBound(varKeys) To UBound(varKeys)
            'キー名によって処理を分岐する
            Select Case varKeys(i)
                '文書の種類の場合
                Case "DocumentType": strDocumentType = varValues(i)
                'ページ番号の場合
                Case "PageNumber": strPageNumber = varValues(i)
                '行番号の場合
                Case "LineNumber": strLineNumber = varValues(i)
                'サンプルの追加の場合
                Case "Sample": strSample = varValues(i)
            '条件分岐を終了する
            End Select
        '次の要素へ
        Next i
    '条件判定を終了する
    End If
    'エラーハンドリングを元に戻す
    On Error GoTo 0
    
    '文書の種類に応じて各種設定を読み込む
    LoadSettingsForDocumentType strDocumentType
    
    '配列を生成する
    SetArrays
    
    'アドバンス幅を把握する
    'ポイントに単位換算しています。
    sngAdvWidth = MillimetersToPoints((sngPageWidth - varMargines(3) - CalcMargine(varMargines, varCharsAndLines)) _
        / varCharsAndLines(1))
'変数を設定するプロシージャを終了する
End Sub

'================================================================================
' 機能: 文書の種類に応じて設定値を読み込む（内部ヘルパー）
'================================================================================
'文書の種類ごとに設定値を読み込むプロシージャ
Private Sub LoadSettingsForDocumentType(ByVal docType As String)
    '【注意】以下の定数は、別途標準モジュールで Public Const として定義されている必要があります。
    ' (例: Public Const strNumberFormats_Saiban As String = "...")

    ' --- アウトライン・ページ設定 ---
    '文書の種類によって処理を分岐する
    Select Case docType
        '裁判所提出書類の場合
        Case "Saiban"
            '番号書式を設定する
            strNumberFormats = strNumberFormats_Saiban
            '番号スタイルを設定する
            strNumberStyles = strNumberStyles_Saiban
            'リスト開始レベルを設定する
            strResetOnHighers = strResetOnHighers_Saiban
            '開始番号を設定する
            strStartAts = strStartAts_Saiban
            '見出しの左インデントを設定する
            strHeadingLeftIndents = strHeadingLeftIndents_Saiban
            '見出しの最初の行のインデントを設定する
            strHeadingFirstLineIndents = strHeadingFirstLineIndents_Saiban
            '本文の左インデントを設定する
            strTextLeftIndents = strTextLeftIndents_Saiban
            '本文の最初の行のインデントを設定する
            strTextFirstLineIndents = strTextFirstLineIndents_Saiban
            '左インデント増加量を設定する
            lngLeftIndentsAdd = lngLeftIndentsAdd_Saiban
            '用紙サイズの幅を設定する
            sngPageWidth = sngPageWidth_Saiban
            'ページ余白を設定する
            strMargines = strMargines_Saiban
            '文字数及び行数を設定する
            strCharsAndLines = strCharsAndLines_Saiban
            '表題の配置を設定する
            lngTitleAlign = lngTitleAlign_Saiban
            '表題のインデントを設定する
            lngTitleIndents = lngTitleIndents_Saiban
            '箇条書きの記号を設定する
            strListBullFormat = strListBullFormat_Saiban
            '段落番号の記号を設定する
            strListNumbFormat = strListNumbFormat_Saiban
            '段落番号のスタイルを設定する
            lngListNumbStyle = lngListNumbStyle_Saiban
            '見出しからの左インデント増加量を設定する
            lngListLeftIndent = lngListLeftIndent_Saiban
            '最初の行のインデントを設定する
            lngListFirstLineIndent = lngListFirstLineIndent_Saiban
            
        '内容証明の場合
        Case "Naiyou"
            '番号書式を設定する
            strNumberFormats = strNumberFormats_Naiyou
            '番号スタイルを設定する
            strNumberStyles = strNumberStyles_Naiyou
            'リスト開始レベルを設定する
            strResetOnHighers = strResetOnHighers_Naiyou
            '開始番号を設定する
            strStartAts = strStartAts_Naiyou
            '見出しの左インデントを設定する
            strHeadingLeftIndents = strHeadingLeftIndents_Naiyou
            '見出しの最初の行のインデントを設定する
            strHeadingFirstLineIndents = strHeadingFirstLineIndents_Naiyou
            '本文の左インデントを設定する
            strTextLeftIndents = strTextLeftIndents_Naiyou
            '本文の最初の行のインデントを設定する
            strTextFirstLineIndents = strTextFirstLineIndents_Naiyou
            '左インデント増加量を設定する
            lngLeftIndentsAdd = lngLeftIndentsAdd_Naiyou
            '用紙サイズの幅を設定する
            sngPageWidth = sngPageWidth_Naiyou
            'ページ余白を設定する
            strMargines = strMargines_Naiyou
            '文字数及び行数を設定する
            strCharsAndLines = strCharsAndLines_Naiyou
            '表題の配置を設定する
            lngTitleAlign = lngTitleAlign_Naiyou
            '表題のインデントを設定する
            lngTitleIndents = lngTitleIndents_Naiyou
            '箇条書きの記号を設定する
            strListBullFormat = strListBullFormat_Naiyou
            '段落番号の記号を設定する
            strListNumbFormat = strListNumbFormat_Naiyou
            '段落番号のスタイルを設定する
            lngListNumbStyle = lngListNumbStyle_Naiyou
            '見出しからの左インデント増加量を設定する
            lngListLeftIndent = lngListLeftIndent_Naiyou
            '最初の行のインデントを設定する
            lngListFirstLineIndent = lngListFirstLineIndent_Naiyou

        '公用文の場合
        Case "Kouyou"
            '番号書式を設定する
            strNumberFormats = strNumberFormats_Kouyou
            '番号スタイルを設定する
            strNumberStyles = strNumberStyles_Kouyou
            'リスト開始レベルを設定する
            strResetOnHighers = strResetOnHighers_Kouyou
            '開始番号を設定する
            strStartAts = strStartAts_Kouyou
            '見出しの左インデントを設定する
            strHeadingLeftIndents = strHeadingLeftIndents_Kouyou
            '見出しの最初の行のインデントを設定する
            strHeadingFirstLineIndents = strHeadingFirstLineIndents_Kouyou
            '本文の左インデントを設定する
            strTextLeftIndents = strTextLeftIndents_Kouyou
            '本文の最初の行のインデントを設定する
            strTextFirstLineIndents = strTextFirstLineIndents_Kouyou
            '左インデント増加量を設定する
            lngLeftIndentsAdd = lngLeftIndentsAdd_Kouyou
            '用紙サイズの幅を設定する
            sngPageWidth = sngPageWidth_Kouyou
            'ページ余白を設定する
            strMargines = strMargines_Kouyou
            '文字数及び行数を設定する
            strCharsAndLines = strCharsAndLines_Kouyou
            '表題の配置を設定する
            lngTitleAlign = lngTitleAlign_Kouyou
            '表題のインデントを設定する
            lngTitleIndents = lngTitleIndents_Kouyou
            '箇条書きの記号を設定する
            strListBullFormat = strListBullFormat_Kouyou
            '段落番号の記号を設定する
            strListNumbFormat = strListNumbFormat_Kouyou
            '段落番号のスタイルを設定する
            lngListNumbStyle = lngListNumbStyle_Kouyou
            '見出しからの左インデント増加量を設定する
            lngListLeftIndent = lngListLeftIndent_Kouyou
            '最初の行のインデントを設定する
            lngListFirstLineIndent = lngListFirstLineIndent_Kouyou

        '契約書1（法令型）の場合
        Case "Keiyaku1"
            '番号書式を設定する
            strNumberFormats = strNumberFormats_Keiyaku1
            '番号スタイルを設定する
            strNumberStyles = strNumberStyles_Keiyaku1
            'リスト開始レベルを設定する
            strResetOnHighers = strResetOnHighers_Keiyaku1
            '開始番号を設定する
            strStartAts = strStartAts_Keiyaku1
            '見出しの左インデントを設定する
            strHeadingLeftIndents = strHeadingLeftIndents_Keiyaku1
            '見出しの最初の行のインデントを設定する
            strHeadingFirstLineIndents = strHeadingFirstLineIndents_Keiyaku1
            '本文の左インデントを設定する
            strTextLeftIndents = strTextLeftIndents_Keiyaku1
            '本文の最初の行のインデントを設定する
            strTextFirstLineIndents = strTextFirstLineIndents_Keiyaku1
            '左インデント増加量を設定する
            lngLeftIndentsAdd = lngLeftIndentsAdd_Keiyaku1
            '用紙サイズの幅を設定する
            sngPageWidth = sngPageWidth_Keiyaku1
            'ページ余白を設定する
            strMargines = strMargines_Keiyaku1
            '文字数及び行数を設定する
            strCharsAndLines = strCharsAndLines_Keiyaku1
            '表題の配置を設定する
            lngTitleAlign = lngTitleAlign_Keiyaku1
            '表題のインデントを設定する
            lngTitleIndents = lngTitleIndents_Keiyaku1
            '箇条書きの記号を設定する
            strListBullFormat = strListBullFormat_Keiyaku1
            '段落番号の記号を設定する
            strListNumbFormat = strListNumbFormat_Keiyaku1
            '段落番号のスタイルを設定する
            lngListNumbStyle = lngListNumbStyle_Keiyaku1
            '見出しからの左インデント増加量を設定する
            lngListLeftIndent = lngListLeftIndent_Keiyaku1
            '最初の行のインデントを設定する
            lngListFirstLineIndent = lngListFirstLineIndent_Keiyaku1

        '契約書2（実務型）の場合
        Case "Keiyaku2"
            '番号書式を設定する
            strNumberFormats = strNumberFormats_Keiyaku2
            '番号スタイルを設定する
            strNumberStyles = strNumberStyles_Keiyaku2
            'リスト開始レベルを設定する
            strResetOnHighers = strResetOnHighers_Keiyaku2
            '開始番号を設定する
            strStartAts = strStartAts_Keiyaku2
            '見出しの左インデントを設定する
            strHeadingLeftIndents = strHeadingLeftIndents_Keiyaku2
            '見出しの最初の行のインデントを設定する
            strHeadingFirstLineIndents = strHeadingFirstLineIndents_Keiyaku2
            '本文の左インデントを設定する
            strTextLeftIndents = strTextLeftIndents_Keiyaku2
            '本文の最初の行のインデントを設定する
            strTextFirstLineIndents = strTextFirstLineIndents_Keiyaku2
            '左インデント増加量を設定する
            lngLeftIndentsAdd = lngLeftIndentsAdd_Keiyaku2
            '用紙サイズの幅を設定する
            sngPageWidth = sngPageWidth_keiyaku2
            'ページ余白を設定する
            strMargines = strMargines_keiyaku2
            '文字数及び行数を設定する
            strCharsAndLines = strCharsAndLines_keiyaku2
            '表題の配置を設定する
            lngTitleAlign = lngTitleAlign_Keiyaku2
            '表題のインデントを設定する
            lngTitleIndents = lngTitleIndents_Keiyaku2
            '箇条書きの記号を設定する
            strListBullFormat = strListBullFormat_Keiyaku2
            '段落番号の記号を設定する
            strListNumbFormat = strListNumbFormat_Keiyaku2
            '段落番号のスタイルを設定する
            lngListNumbStyle = lngListNumbStyle_Keiyaku2
            '見出しからの左インデント増加量を設定する
            lngListLeftIndent = lngListLeftIndent_Keiyaku2
            '最初の行のインデントを設定する
            lngListFirstLineIndent = lngListFirstLineIndent_Keiyaku2

        'ビジネス文書の場合
        Case "Business"
            '番号書式を設定する
            strNumberFormats = strNumberFormats_Business
            '番号スタイルを設定する
            strNumberStyles = strNumberStyles_Business
            'リスト開始レベルを設定する
            strResetOnHighers = strResetOnHighers_Business
            '開始番号を設定する
            strStartAts = strStartAts_Business
            '見出しの左インデントを設定する
            strHeadingLeftIndents = strHeadingLeftIndents_Business
            '見出しの最初の行のインデントを設定する
            strHeadingFirstLineIndents = strHeadingFirstLineIndents_Business
            '本文の左インデントを設定する
            strTextLeftIndents = strTextLeftIndents_Business
            '本文の最初の行のインデントを設定する
            strTextFirstLineIndents = strTextFirstLineIndents_Business
            '左インデント増加量を設定する
            lngLeftIndentsAdd = lngLeftIndentsAdd_Business
            '用紙サイズの幅を設定する
            sngPageWidth = sngPageWidth_Business
            'ページ余白を設定する
            strMargines = strMargines_Business
            '文字数及び行数を設定する
            strCharsAndLines = strCharsAndLines_Business
            '表題の配置を設定する
            lngTitleAlign = lngTitleAlign_Business
            '表題のインデントを設定する
            lngTitleIndents = lngTitleIndents_Business
            '箇条書きの記号を設定する
            strListBullFormat = strListBullFormat_Business
            '段落番号の記号を設定する
            strListNumbFormat = strListNumbFormat_Business
            '段落番号のスタイルを設定する
            lngListNumbStyle = lngListNumbStyle_Business
            '見出しからの左インデント増加量を設定する
            lngListLeftIndent = lngListLeftIndent_Business
            '最初の行のインデントを設定する
            lngListFirstLineIndent = lngListFirstLineIndent_Business
            
        '技術文書の場合
        Case "Technical"
            '番号書式を設定する
            strNumberFormats = strNumberFormats_Technical
            '番号スタイルを設定する
            strNumberStyles = strNumberStyles_Technical
            'リスト開始レベルを設定する
            strResetOnHighers = strResetOnHighers_Technical
            '開始番号を設定する
            strStartAts = strStartAts_Technical
            '見出しの左インデントを設定する
            strHeadingLeftIndents = strHeadingLeftIndents_Technical
            '見出しの最初の行のインデントを設定する
            strHeadingFirstLineIndents = strHeadingFirstLineIndents_Technical
            '本文の左インデントを設定する
            strTextLeftIndents = strTextLeftIndents_Technical
            '本文の最初の行のインデントを設定する
            strTextFirstLineIndents = strTextFirstLineIndents_Technical
            '左インデント増加量を設定する
            lngLeftIndentsAdd = lngLeftIndentsAdd_Technical
            '用紙サイズの幅を設定する
            sngPageWidth = sngPageWidth_Technical
            'ページ余白を設定する
            strMargines = strMargines_Technical
            '文字数及び行数を設定する
            strCharsAndLines = strCharsAndLines_Technical
            '表題の配置を設定する
            lngTitleAlign = lngTitleAlign_Technical
            '表題のインデントを設定する
            lngTitleIndents = lngTitleIndents_Technical
            '箇条書きの記号を設定する
            strListBullFormat = strListBullFormat_Technical
            '段落番号の記号を設定する
            strListNumbFormat = strListNumbFormat_Technical
            '段落番号のスタイルを設定する
            lngListNumbStyle = lngListNumbStyle_Technical
            '見出しからの左インデント増加量を設定する
            lngListLeftIndent = lngListLeftIndent_Technical
            '最初の行のインデントを設定する
            lngListFirstLineIndent = lngListFirstLineIndent_Technical
            
        'それ以外（デフォルト）の場合
        Case Else
            '番号書式を設定する
            strNumberFormats = strNumberFormats_Saiban
            '番号スタイルを設定する
            strNumberStyles = strNumberStyles_Saiban
            'リスト開始レベルを設定する
            strResetOnHighers = strResetOnHighers_Saiban
            '開始番号を設定する
            strStartAts = strStartAts_Saiban
            '見出しの左インデントを設定する
            strHeadingLeftIndents = strHeadingLeftIndents_Saiban
            '見出しの最初の行のインデントを設定する
            strHeadingFirstLineIndents = strHeadingFirstLineIndents_Saiban
            '本文の左インデントを設定する
            strTextLeftIndents = strTextLeftIndents_Saiban
            '本文の最初の行のインデントを設定する
            strTextFirstLineIndents = strTextFirstLineIndents_Saiban
            '左インデント増加量を設定する
            lngLeftIndentsAdd = lngLeftIndentsAdd_Saiban
            '用紙サイズの幅を設定する
            sngPageWidth = sngPageWidth_Saiban
            'ページ余白を設定する
            strMargines = strMargines_Saiban
            '文字数及び行数を設定する
            strCharsAndLines = strCharsAndLines_Saiban
            '表題の配置を設定する
            lngTitleAlign = lngTitleAlign_Saiban
            '表題のインデントを設定する
            lngTitleIndents = lngTitleIndents_Saiban
            '箇条書きの記号を設定する
            strListBullFormat = strListBullFormat_Saiban
            '段落番号の記号を設定する
            strListNumbFormat = strListNumbFormat_Saiban
            '段落番号のスタイルを設定する
            lngListNumbStyle = lngListNumbStyle_Saiban
            '見出しからの左インデント増加量を設定する
            lngListLeftIndent = lngListLeftIndent_Saiban
            '最初の行のインデントを設定する
            lngListFirstLineIndent = lngListFirstLineIndent_Saiban
    '条件分岐を終了する
    End Select

    ' --- サンプル文字列の設定 ---
    'サンプルの追加が指定されている場合
    If strSample = "Add" Then
        '文書の種類によって処理を分岐する
        Select Case docType
            '契約書1の場合
            Case "Keiyaku1"
                '契約書1のサンプル文字列を設定する
                strSampleContents = strSampleContents_Keiyaku1
            '契約書2の場合
            Case "Keiyaku2"
                '契約書2のサンプル文字列を設定する
                strSampleContents = strSampleContents_Keiyaku2
            'それ以外の場合
            Case Else
                '裁判所提出書類のサンプル文字列を設定する
                strSampleContents = strSampleContents_Saiban
        '条件分岐を終了する
        End Select
    'サンプルの追加が指定されていない場合
    Else
        'サンプル文字列を空にする
        strSampleContents = ""
    '条件判定を終了する
    End If
'文書の種類に応じた設定を読み込むプロシージャを終了する
End Sub

'================================================================================
' 機能: 文字列で定義された設定値を配列に変換する（内部処理用）
'================================================================================
'配列を設定するプロシージャ
Private Sub SetArrays()
    '番号書式を配列に変換する
    varNumberFormats = ConvertToArray(strNumberFormats)
    '番号スタイルを配列に変換する
    varNumberStyles = ConvertToArray(strNumberStyles)
    '見出しの左インデントを配列に変換する
    varHeadingLeftIndents = ConvertToArray(strHeadingLeftIndents)
    '見出しの最初の行のインデントを配列に変換する
    varHeadingFirstLineIndents = ConvertToArray(strHeadingFirstLineIndents)
    '本文の左インデントを配列に変換する
    varTextLeftIndents = ConvertToArray(strTextLeftIndents)
    '本文の最初の行のインデントを配列に変換する
    varTextFirstLineIndents = ConvertToArray(strTextFirstLineIndents)
    'ページ余白を配列に変換する
    varMargines = ConvertToArray(strMargines)
    '文字数および行数を配列に変換する
    varCharsAndLines = ConvertToArray(strCharsAndLines)
    'リスト開始レベルを配列に変換する
    varResetOnHighers = ConvertToArray(strResetOnHighers)
    '開始番号を配列に変換する
    varStartAts = ConvertToArray(strStartAts)

    '右余白を修正する
    varMargines(4) = Me.CalcMargine(varMargines, varCharsAndLines)
'配列を設定するプロシージャを終了する
End Sub

'================================================================================
' 機能: カンマ区切りの文字列を配列に変換する（ヘルパー関数）
'================================================================================
'文字列を配列に変換する関数
Private Function ConvertToArray(ByVal strArray As String) As Variant
    ' 先頭にカンマを追加することで、Split関数の結果の配列を
    ' 0番目を空要素にし、実質的に1始まりの配列として扱えるようにする
    ConvertToArray = Split("," & strArray, ",")
'文字列を配列に変換する関数を終了する
End Function

'================================================================================
' 機能: Wordの文字間隔の自動調整による誤差を補正し、設定すべき正確な右余白を計算する。
' 概要: Wordのページ設定で文字数と行数を指定し、段落の配置を「両端揃え」にすると、
'       Wordが内部的に行う「字送り（文字ごとの間隔）」の計算（0.05pt単位での
'       丸め処理）によって、意図した行幅と実際の行幅に微小な誤差が生じます。
'       この誤差が原因で、行末に改行がない行とある行とで文字間隔が不揃いになる
'       現象が発生することがあります。
'       この関数は、Wordの計算方法をシミュレートして誤差を算出し、その誤差分を
'       相殺するように右余白の値を補正します。これにより、レイアウトの崩れを防ぎます。
'
' 引数:
'   varMargines (Variant): ページ余白の設定値を格納した配列（1:上, 2:下, 3:左, 4:右）。単位はmm。
'   varCharsAndLines (Variant): 1行の文字数とページの行数を格納した配列（1:文字数, 2:行数）。
'
' 戻り値:
'   Single: 計算によって補正された、実際にページ設定に適用すべき右余白の値。単位はmm。
'================================================================================
'右余白を計算する関数
Function CalcMargine(varMargines As Variant, varCharsAndLines As Variant) As Single
'設定すべき右余白を計算します。
'標準スタイルを「両端揃え」にした場合、条件によっては行中に「改行がある行」と「改行がない行」
'との文字間隔が揃いません。
'ワードが「文字数」から割り出す「字送り」の値（0.05刻み）に誤差があるためです。
'「字送り」の値はVBAで操作できないので、「文字数」と「字送り」に応じて右余白を計算します。
'ページの設定には、計算で得られた右余白の値を用います。
'もっと簡単な計算方法もあるかも知れませんが、あまり算数は得意でないのでお許しください。
'セットアップの時しか計算しませんので、速度を追求する必要はないと思います。

    '設定しようとしている行幅を求める
    '左余白を格納する変数を宣言する
    Dim sngLeftMargine As Single
    '設定しようとしている左余白（mm）を取得する
    sngLeftMargine = varMargines(3)
    '右余白を格納する変数を宣言する
    Dim sngRightMargine As Single
    '設定しようとしている右余白（mm）を取得する
    sngRightMargine = varMargines(4)
    '行幅を格納する変数を宣言する
    Dim sngLineWidth As Single
    '設定しようとしている行幅（mm）を計算する
    sngLineWidth = sngPageWidth - sngLeftMargine - sngRightMargine
    
    '適正な字送り量を求める
    '文字数を格納する変数を宣言する
    Dim lngCharCount As Long
    '文字数（文字）を取得する
    lngCharCount = varCharsAndLines(1)
    '適正な字送り量（mm)を格納する変数を宣言する
    Dim sngAppCharFeed As Single
    '適正な字送り量（mm)を計算する
    sngAppCharFeed = sngLineWidth / lngCharCount
    '適正な字送り量（pt）を格納する変数を宣言する
    Dim sngAppCharFeedPt As Single
    '適正な字送り量（pt）に換算する（Word組み込みの換算関数を使用する）
    sngAppCharFeedPt = MillimetersToPoints(sngAppCharFeed)
    
    'ワードが算定する字送り量を求める
    '適正な字送り量を0.05刻みで四捨五入します。
    'エクセルのMROUD関数と同じことを行っています。
    'ワードが算定する字送り量（pt）を格納する変数を宣言する
    Dim sngCalCharFeedPt  As Single
    'ワードが算定する字送り量（pt）を計算する
    sngCalCharFeedPt = Round(sngAppCharFeedPt * 20, 0) / 20
    'ワードが算定する字送り量（mm）を格納する変数を宣言する
    Dim sngCalCharFeed  As Single
    'ワードが算定する字送り量（mm）に換算する（Word組み込みの換算関数を使用する）
    sngCalCharFeed = PointsToMillimeters(sngCalCharFeedPt)
    
    '適正な行幅を求める
    '適正な行幅(mm)を格納する変数を宣言する
    Dim sngAppLineWidth As Single
    '適正な行幅(mm)を計算する
    sngAppLineWidth = sngCalCharFeed * lngCharCount
    
    '設定すべき右余白を求める
    '設定すべき右余白（mm）を計算して返す
    CalcMargine = sngRightMargine - (sngAppLineWidth - sngLineWidth)

'右余白を計算する関数を終了する
End Function

