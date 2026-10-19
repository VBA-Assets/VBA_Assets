'================================================================================
' モジュール概要:
' Word文書の疑似アウトラインとマークダウンの相互変換マクロで使用される、
' グローバルな定数を定義します。
'
' 主な内容:
'   - アウトライン書式プリセット:
'     番号書式、番号スタイル、インデント設定などを「全半角混在」「すべて半角」
'     「全角主体」の3つのプリセットとして定義します。
'     これらの値を変更することで、出力形式をカスタマイズできます。
'
'   - リストおよびマークダウン記号:
'     箇条書きや段落番号の記号、変換処理で使用するマークダウン記号を定義します。
'
'   - マクロ動作設定:
'     設定ファイル名や、変換処理の挙動を制御する各種フラグを定義します。
'
' 注意:
' このモジュール内の定数を変更すると、マクロ全体の動作に影響します。
'================================================================================

Option Explicit

'アウトラインおよびインデント
'設定項目ごとに見出しレベル1から9までの設定値を「,」で区切って記載してください。
'番号書式
    '文字種および全・半角の別に関係なく、レベルに応じて%1～%9を入力してください。
    '番号の前後に入れる文字および空白をそのまま追加してください。
    '例：第１："第%1　"、第1："第%1 "、"１："%1　"、(1)："(%1) "
'番号スタイル
    '文字種および全・半角の別に応じて、組み込み定数を入力てください。
    '英字およびローマ数字には全角が使えません。
    '全角アラビア数字：14、半角アラビア数字：0、全角カタカナ：20、半角カタカナ：12
    '英字（小文字）：4、ローマ数字（小文字）：2、丸囲み数字：18
'見出しの左インデント
    '左インデントを文字数単位で入力してください。
    '例：0文字：0、1文字：1、1.5文字：1.5
'見出しの最初の行のインデント
    '最初の行のインデントを文字数単位で入力してください。
    '例：1文字の字下げ：1、1.5文字のぶら下げ：-1.5
'本文の左インデント
    '左インデントを文字数単位で入力してください。
    '例：0文字：0、1文字：1、1.5文字：1.5
'本文の最初の行のインデント
    '最初の行のインデントを文字数単位で入力してください。
    '例：1文字の字下げ：1、1.5文字のぶら下げ：-1.5
'左インデント増加量
    'アウトラインレベルに応じて自動的に増加させる左インデントの値を文字数単位で入力してください。
    '例：1文字ずつ増加：1、増加させない：0

'全半角混在　第１、１、(1)、ア、(ｱ)、a、(a)、i、(i)
Public Const strNumberFormats As String = "第%1,%2,(%3),%4,(%5),%6,(%7),%8,(%9)"      '番号書式
Public Const strNumberStyles As String = "14,14,0,20,12,4,4,2,2"                      '番号スタイル
Public Const strHeadingLeftIndents As String = "1,1,1,1,1,1,1,1,1"                    '見出しの左インデント
Public Const strHeadingFirstLineIndents As String = "-1,-1,-1,-1,-1,-1,-1,-1,-1"      '見出しの最初の行のインデント
Public Const strTextLeftIndents As String = "1,1,1,1,1,1,1,1,1"                       '本文の左インデント
Public Const strTextFirstLineIndents As String = "1,1,1,1,1,1,1,1,1"                  '本文の最初の行のインデント
Public Const lngLeftIndentsAdd As Long = 1                                            '左インデント増加量
    
'全半角混在（オリジナル）　第１、１、(1)、ア、(ｱ)、a、(a)、i、(i)
'Public Const strNumberFormats As String = "第%1,%2,(%3),%4,(%5),%6,(%7),%8,(%9)"      '*番号書式
'Public Const strNumberStyles As String = "14,14,0,20,12,4,4,2,2"                      '*番号スタイル
'Public Const strHeadingLeftIndents As String = "2,1,1,1,1,1,1,1,1"                    '*見出しの左インデント
'Public Const strHeadingFirstLineIndents As String = "-2,-1,-1,-1,-1,-1,-1,-1,-1"      '*見出しの最初の行のインデント
'Public Const strTextLeftIndents As String = "1,1,1,1,1,1,1,1,1"                       '*本文の左インデント
'Public Const strTextFirstLineIndents As String = "1,1,1,1,1,1,1,1,1"                  '*本文の最初の行のインデント
'Public Const lngLeftIndentsAdd As Long = 1                                            '*左インデント増加量

'すべて半角　第1、1、(1)、ｱ、(ｱ)、a、(a)、i、(i)
'Public Const strNumberFormats As String = "第%1,%2,(%3),%4,(%5),%6,(%7),%8,(%9)"        '番号書式
'Public Const strNumberStyles As String = "0,0,0,12,12,4,4,2,2"                          '番号スタイル
'Public Const strHeadingLeftIndents As String = "1,1,1,1,1,1,1,1,1"                      '見出しの左インデント
'Public Const strHeadingFirstLineIndents As String = "-1,-1,-1,-1,-1,-1,-1,-1,-1"        '見出しの最初の行のインデント
'Public Const strTextLeftIndents As String = "1,1,1,1,1,1,1,1,1"                         '本文の左インデント
'Public Const strTextFirstLineIndents As String = "1,1,1,1,1,1,1,1,1"                    '本文の最初の行のインデント
'Public Const lngLeftIndentsAdd As Long = 1                                              '左インデント増加量

'全角主体　第１、１、(１)、ア、(ア)、a、(a)、i、(i)
'Public Const strNumberFormats As String = "第%1,%2,(%3),%4,(%5),%6,(%7),%8,(%9)"        '番号書式
'Public Const strNumberStyles As String = "14,14,14,20,20,4,4,2,2"                       '番号スタイル
'Public Const strHeadingLeftIndents As String = "1,1,1,1,1,1,1,1,1"                      '見出しの左インデント
'Public Const strHeadingFirstLineIndents As String = "-1,-1,-1,-1,-1,-1,-1,-1,-1"        '見出しの最初の行のインデント
'Public Const strTextLeftIndents As String = "1,1,1,1,1,1,1,1,1"                         '本文の左インデント
'Public Const strTextFirstLineIndents As String = "1,1,1,1,1,1,1,1,1"                    '本文の最初の行のインデント
'Public Const lngLeftIndentsAdd As Long = 1                                              '左インデント増加量

'リスト
Public Const strListBullFormat As String = "・"     '箇条書きの記号（デフォルト："・"）
Public Const strListNumbFormat As String = "%1　"   '段落番号の記号（デフォルト："%1　"）
Public Const lngListLeftIndent As Long = 0          'リストの左インデント（デフォルト：0）
Public Const lngListFirstLineIndent As Long = -1    'リストの最初の行のインデント（デフォルト：-1）
    
'マークダウン記号
Public Const strHead As String = "#"               '*見出しレベル（デフォルト："#"）
Public Const strList As String = "-"               '箇条書き（デフォルト："-"）
Public Const strNumList As String = "1."           '段落番号（変更不可："1."）
Public Const blnMarkdownSpace As Boolean = True    '*マークダウン記号の後の空白（デフォルト：True）

'マクロ
Public Const blnReplaceSpaceWithIndent = True                           '行頭の空白をインデントに置き換える（デフォルト：True）
Public Const blnSkipUntil = True                                        '最初の見出しまでの処理を停止する（デフォルト:True）
Public Const blnStopAfter = True                                        '配置が右揃えの場合に処理を停止する（デフォルト：True）
Public Const blnApplyTableStyle = True                                  '表内の段落に「行間詰め」スタイルを適用する（デフォルト：True）
Public Const blnHeadingReset = True                                     'セクション区切りで見出し番号を振り直す（デフォルト：True）
'変換（疑似アウトライン・マークダウン・全角1文字）を Ctrl+Z 1回で元に戻せるようにする（デフォルト：False）
'元に戻す（Undo）の記録（Application.UndoRecord）の中で「すべて置換」を行うと、Wordが強制終了することが分かったため
'（Word 16.0.20430 で確認。VA公用文 Ver8D2 で発生）、Windows版・Mac版とも使わない。安定性を優先し、True にしないこと（Ver371）
Public Const blnUseUndoRecord As Boolean = False