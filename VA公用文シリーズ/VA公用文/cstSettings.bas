Option Explicit
'==================================================================================================
'   概要        : Microsoft Wordの文書整形マクロで使用する定数を定義します。
'                 このモジュールでは、文書の種類に応じた書式設定（アウトライン、インデント、
'                 ページ設定など）や、マクロ全体の動作を制御するための値を一元管理します。
'
'   使用方法    : 1. 各設定項目（例：「裁判所提出書類用」）には、複数の書式セットが
'                    コメントアウトされた状態で用意されています。
'                 2. 使用したい書式セットの行のコメントを解除し、他のセットはコメントアウト
'                    してください。
'                 3. 必要に応じて、各定数の値を直接編集することで、書式をカスタマイズする
'                    ことも可能です。
'
'   備考        : このモジュールで宣言されている定数（Public Const）は、プロジェクト内の
'                 他のすべてのモジュールから参照されることを想定しています。
'==================================================================================================

'すべてのモジュールで使用できるようにするため、Publicで宣言しています。
'クラスモジュール内のプロシージャにより、変数化および配列化して用いられます。

'準拠
'資料1：「裁判文書（裁判所提出書類）の標準的な書式」http://www5d.biglobe.ne.jp/Jusl/Bunsyo/Suujihyouki2.html
'資料2：「内容証明　ご利用の条件等」https://www.post.japanpost.jp/service/fuka_service/syomei/use.html
'資料3：「公用文の書き方」第2次改訂版　磯崎陽輔著　株式会社ぎょうせい
'資料4：「契約書フォーマット」経済産業省　https://www.meti.go.jp/information_2/publicoffer/keiyaku_format.html
'資料5：「契約書の達人」https://xn--wtsq13a09q.jp/how-to-write-operative-part/
'資料6：「公文書の書式と文例」文部科学省　https://upload.wikimedia.org/wikipedia/commons/f/fe/%E3%80%8E%E5%85%AC%E6%96%87%E6%9B%B8%E3%81%AE%E6%9B%B8%E5%BC%8F%E3%81%A8%E6%96%87%E4%BE%8B%EF%BC%88%E5%85%AD%E8%A8%82%EF%BC%89%E3%80%8F%EF%BC%88%E5%B9%B3%E6%88%9023%E5%B9%B4%EF%BC%93%E6%9C%88%E6%96%87%E9%83%A8%E7%A7%91%E5%AD%A6%E7%9C%81%EF%BC%89.pdf

    'アウトライン、リストおよびインデント
    '階層のあるものについては、見出しレベル1から9までの設定値を「,」で区切って記載してください。
    '番号書式
        '文字種および全・半角の別に関係なく、レベルに応じて%1～%9を指定してください。
        '番号の前後に入れる文字および空白をそのまま追加してください。
        '例：第１："第%1　"、第1："第%1 "、"１："%1　"、(1)："(%1) "
    '番号スタイル
        '文字種および全・半角の別に応じて、組み込み定数を指定してください。
        '英字およびローマ数字には全角が使えません。
        '全角アラビア数字：14、半角アラビア数字：0、全角カタカナ：20、半角カタカナ：12
        '英字（小文字）：4、ローマ数字（小文字）：2、丸囲み数字：18
    'リスト開始レベル
        '見出し番号をリセットするアウトラインレベルの値を指定してください。
        '例：直上のレベル：1（自分のレベルの値から1を減じたもの）、リセットしない：0
    '開始番号
        '見出し番号の開始番号を指定してください。
        '例：1から開始：1、2から開始：2
    '見出しの左インデント
        '左インデントを文字数単位で指定してください。
        '例：0文字：0、1文字：1、1.5文字：1.5
    '見出しの最初の行のインデント
        '最初の行のインデントを文字数単位で入力指定ください。
        '例：1文字の字下げ：1、1.5文字のぶら下げ：-1.5
    '本文の左インデント
        '左インデントを文字数単位で指定してください。
        '例：0文字：0、1文字：1、1.5文字：1.5
    '本文の最初の行のインデント
        '最初の行のインデントを文字数単位で指定してください。
        '例：1文字の字下げ：1、1.5文字のぶら下げ：-1.5
    '左インデント増加量
        'アウトラインレベルに応じて自動的に増加させる左インデントの値を文字数単位で指定してください。
        '例：1文字ずつ増加：1、増加させない：0
    '表題の配置
        '表題の配置を定数で指定してください。
        '例：中央揃え：wdAlignParagraphCenter、両端揃え：wdAlignParagraphJustify
    '表題のインデント
        '表題の左インデントを指定してください。（整数のみ）
    '箇条書きの記号
        '箇条書きの記号を指定してください。（1文字のみ）
    '段落番号の記号
        '段落番号の記号を指定してください。
        '番号の前後に入れる文字および空白をそのまま追加してください。
        '例：番号のみ："%1"、番号と半角空白："%1 "、番号と半角カッコ："%1)"
    '段落番号のスタイル
        '段落番号の記号を指定してください。
        '例：丸囲み数字：18、英小文字：4
    'リストの見出しからの左インデント増加量
        'リストの見出しからの左インデント増加量を指定してください。（整数のみ）
    'リストの最初の行のインデント
        'リストの最初の行のインデントを指定してください。（整数のみ）
    
        '◯裁判所提出書類用
        '次の定数の組み合わせの中から、1組をコメント解除してください。
        '「資料1」に準拠　第１　、１　、(1) 、ア　、(ｱ) 、a 　、(a) 、i　、(i)
        Public Const strNumberFormats_Saiban As String = "第%1　,%2　,(%3) ,%4　,(%5) ,%6　,(%7) ,%8　,(%9) "      '番号書式
        Public Const strNumberStyles_Saiban As String = "14,14,0,20,12,4,4,2,2"                                    '番号スタイル
        Public Const strResetOnHighers_Saiban As String = "0,1,2,3,4,5,6,7,8"                                      'リスト開始レベル
        Public Const strStartAts_Saiban As String = "1,1,1,1,1,1,1,1,1"                                            '開始番号
        Public Const strHeadingLeftIndents_Saiban As String = "1,1,1,1,1,1,1,1,1"                                  '見出しの左インデント
        Public Const strHeadingFirstLineIndents_Saiban As String = "-1,-1,-1,-1,-1,-.5,-1,-.5,-1"                  '見出しの最初の行のインデント
        Public Const strTextLeftIndents_Saiban As String = "1,1,1,1,1,1,1,1,1"                                     '本文の左インデント
        Public Const strTextFirstLineIndents_Saiban As String = "1,1,1,1,1,1,1,1,1"                                '本文の最初の行のインデント
        Public Const lngLeftIndentsAdd_Saiban As Long = 1                                                          '左インデント増加量
        Public Const lngTitleAlign_Saiban As Long = wdAlignParagraphCenter                                         '表題の配置（中央揃え）
        Public Const lngTitleIndents_Saiban As Long = 0                                                            '表題のインデント
        Public Const strListBullFormat_Saiban As String = "・"                                                     '箇条書きの記号
        Public Const strListNumbFormat_Saiban As String = "%1　"                                                   '段落番号の記号
        Public Const lngListNumbStyle_Saiban As Long = 18                                                          '段落番号のスタイル（丸囲み数字）
        Public Const lngListLeftIndent_Saiban As Long = 1                                                          'リストの見出しからの左インデント増加量
        Public Const lngListFirstLineIndent_Saiban As Long = -1                                                    'リストの最初の行のインデント
        
        'オリジナル　第１　、１　、(1)　、ア　、(ｱ)　、a　、(a)　、i　、(i)
'        Public Const strNumberFormats_Saiban As String = "第%1　,%2　,(%3)　,%4　,(%5)　,%6　,(%7)　,%8　,(%9)　" '*番号書式
'        Public Const strNumberStyles_Saiban As String = "14,14,0,20,12,4,4,2,2"                                   '*番号スタイル
'        Public Const strResetOnHighers_Saiban As String = "0,1,2,3,4,5,6,7,8"                                     '*リスト開始レベル
'        Public Const strStartAts_Saiban As String = "1,1,1,1,1,1,1,1,1"                                           '*開始番号
'        Public Const strHeadingLeftIndents_Saiban As String = "2,1,1,1,1,1,1,1,1"                                 '*見出しの左インデント
'        Public Const strHeadingFirstLineIndents_Saiban As String = "-2,-1,-1.5,-1,-1.5,-.5,-1.5,-.5,-1.5"         '*見出しの最初の行のインデント
'        Public Const strTextLeftIndents_Saiban As String = "2,1,1,1,1,1,1,1,1"                                    '*本文の左インデント
'        Public Const strTextFirstLineIndents_Saiban As String = "1,1,1,1,1,1,1,1,1"                               '*本文の最初の行のインデント
'        Public Const lngLeftIndentsAdd_Saiban As Long = 1                                                         '*左インデント増加量
'        Public Const lngTitleAlign_Saiban As Long = wdAlignParagraphCenter                                        '*表題の配置（中央揃え）
'        Public Const lngTitleIndents_Saiban As Long = 0                                                           '*表題のインデント
'        Public Const strListBullFormat_Saiban As String = "・"                                                    '*箇条書きの記号
'        Public Const strListNumbFormat_Saiban As String = "%1　"                                                  '*段落番号の記号
'        Public Const lngListNumbStyle_Saiban As Long = 18                                                         '*段落番号のスタイル
'        Public Const lngListLeftIndent_Saiban As Long = 1                                                         '*リストの見出しからの左インデント増加量
'        Public Const lngListFirstLineIndent_Saiban As Long = -1                                                   '*リストの最初の行のインデント
    
        '◯内容証明用
        '「資料1」及び「資料2」に準拠　第１　、１　、（１）、ア　、（ア）、a　、（a）、i　、（i）
        Public Const strNumberFormats_Naiyou As String = "第%1　,%2　,（%3）,%4　,（%5）,%6　,（%7）,%8　,（%9）"  '番号書式
        Public Const strNumberStyles_Naiyou As String = "14,14,14,20,20,4,4,2,2"                                   '番号スタイル
        Public Const strResetOnHighers_Naiyou As String = "0,1,2,3,4,5,6,7,8"                                      'リスト開始レベル
        Public Const strStartAts_Naiyou As String = "1,1,1,1,1,1,1,1,1"                                            '開始番号
        Public Const strHeadingLeftIndents_Naiyou As String = "1,1,1,1,1,1,1,1,1"                                  '見出しの左インデント
        Public Const strHeadingFirstLineIndents_Naiyou As String = "-1,-1,-1,-1,-1,-.5,-1.5,-.5,-1.5"              '見出しの最初の行のインデント
        Public Const strTextLeftIndents_Naiyou As String = "1,1,1,1,1,1,1,1,1"                                     '本文の左インデント
        Public Const strTextFirstLineIndents_Naiyou As String = "1,1,1,1,1,1,1,1,1"                                '本文の最初の行のインデント
        Public Const lngLeftIndentsAdd_Naiyou As Long = 1                                                          '左インデント増加量
        Public Const lngTitleAlign_Naiyou As Long = wdAlignParagraphCenter                                         '表題の配置（中央揃え）
        Public Const lngTitleIndents_Naiyou As Long = 0                                                            '表題のインデント
        Public Const strListBullFormat_Naiyou As String = "・"                                                     '箇条書きの記号
        Public Const strListNumbFormat_Naiyou As String = "%1　"                                                   '段落番号の記号
        Public Const lngListNumbStyle_Naiyou As Long = 18                                                          '段落番号のスタイル（丸囲み数字）
        Public Const lngListLeftIndent_Naiyou As Long = 1                                                          'リストの見出しからの左インデント増加量
        Public Const lngListFirstLineIndent_Naiyou As Long = -1                                                    'リストの最初の行のインデント
                
        'オリジナル及び「資料2」に準拠　第１　、１　、（１）、ア　、（ア）、a　、（a）、i　、（i）
'        Public Const strNumberFormats_Naiyou As String = "第%1　,%2　,（%3）,%4　,（%5）,%6　,（%7）,%8　,（%9）" '*番号書式
'        Public Const strNumberStyles_Naiyou As String = "14,14,14,20,20,4,4,2,2"                                  '*番号スタイル
'        Public Const strResetOnHighers_Naiyou As String = "0,1,2,3,4,5,6,7,8"                                     '*リスト開始レベル
'        Public Const strStartAts_Naiyou As String = "1,1,1,1,1,1,1,1,1"                                           '*開始番号
'        Public Const strHeadingLeftIndents_Naiyou As String = "2,1,1,1,1,1,1,1,1"                                 '*見出しの左インデント
'        Public Const strHeadingFirstLineIndents_Naiyou As String = "-2,-1,-1,-1,-1,-.5,-1.5,-.5,-1.5"             '*見出しの最初の行のインデント
'        Public Const strTextLeftIndents_Naiyou As String = "2,1,1,1,1,1,1,1,1"                                    '*本文の左インデント
'        Public Const strTextFirstLineIndents_Naiyou As String = "1,1,1,1,1,1,1,1,1"                               '*本文の最初の行のインデント
'        Public Const lngLeftIndentsAdd_Naiyou As Long = 1                                                         '*左インデント増加量
'        Public Const lngTitleAlign_Naiyou As Long = wdAlignParagraphCenter                                        '*表題の配置（中央揃え）
'        Public Const lngTitleIndents_Naiyou As Long = 0                                                           '*表題のインデント
'        Public Const strListBullFormat_Naiyou As String = "・"                                                    '*箇条書きの記号
'        Public Const strListNumbFormat_Naiyou As String = "%1　"                                                  '*段落番号の記号
'        Public Const lngListNumbStyle_Naiyou As Long = 18                                                         '*段落番号のスタイル
'        Public Const lngListLeftIndent_Naiyou As Long = 1                                                         '*リストの見出しからの左インデント増加量
'        Public Const lngListFirstLineIndent_Naiyou As Long = -1                                                   '*リストの最初の行のインデント
        
        '◯公用文用
        '次の定数の組み合わせの中から、1組をコメント解除してください。
        '「資料3」に準拠　第１　、１　、(1) 、ア　、(ｱ) 、a　、(a) 、i　、(i)
        Public Const strNumberFormats_Kouyou As String = "第%1　,%2　,(%3) ,%4　,(%5) ,%6　,(%7) ,%8　,(%9) "      '番号書式
        Public Const strNumberStyles_Kouyou As String = "14,14,0,20,12,4,4,2,2"                                    '番号スタイル
        Public Const strResetOnHighers_Kouyou As String = "0,1,2,3,4,5,6,7,8"                                      'リスト開始レベル
        Public Const strStartAts_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                            '開始番号
        Public Const strHeadingLeftIndents_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                  '見出しの左インデント
        Public Const strHeadingFirstLineIndents_Kouyou As String = "-1,-1,-1,-1,-1,-.5,-1,-.5,-1"                  '見出しの最初の行のインデント
        Public Const strTextLeftIndents_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                     '本文の左インデント
        Public Const strTextFirstLineIndents_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                '本文の最初の行のインデント
        Public Const lngLeftIndentsAdd_Kouyou As Long = 1                                                          '左インデント増加量
        Public Const lngTitleAlign_Kouyou As Long = wdAlignParagraphJustify                                        '表題の配置（両端揃え）
        Public Const lngTitleIndents_Kouyou As Long = 3                                                            '表題のインデント
        Public Const strListBullFormat_Kouyou As String = "・"                                                     '箇条書きの記号
        Public Const strListNumbFormat_Kouyou As String = "%1　"                                                   '段落番号の記号
        Public Const lngListNumbStyle_Kouyou As Long = 18                                                          '段落番号のスタイル（丸囲み数字）
        Public Const lngListLeftIndent_Kouyou As Long = 1                                                          'リストの見出しからの左インデント増加量
        Public Const lngListFirstLineIndent_Kouyou As Long = -1                                                    'リストの最初の行のインデント
    
        '「資料3」に準拠　第１　、１　、(1) 、ア　、(ア) 、a　、(a) 、i　、(i)
'        Public Const strNumberFormats_Kouyou As String = "第%1　,%2　,(%3) ,%4　,(%5) ,%6　,(%7) ,%8　,(%9) "     '番号書式
'        Public Const strNumberStyles_Kouyou As String = "14,14,0,20,20,4,4,2,2"                                   '番号スタイル
'        Public Const strResetOnHighers_Kouyou As String = "0,1,2,3,4,5,6,7,8"                                     'リスト開始レベル
'        Public Const strStartAts_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                           '開始番号
'        Public Const strHeadingLeftIndents_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                 '見出しの左インデント
'        Public Const strHeadingFirstLineIndents_Kouyou As String = "-1,-1,-1,-1,-1.5,-.5,-1,-.5,-1"               '見出しの最初の行のインデント
'        Public Const strTextLeftIndents_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                    '本文の左インデント
'        Public Const strTextFirstLineIndents_Kouyou As String = "1,1,1,1,1,1,1,1,1"                               '本文の最初の行のインデント
'        Public Const lngLeftIndentsAdd_Kouyou As Long = 1                                                         '左インデント増加量
'        Public Const lngTitleAlign_Kouyou As Long = wdAlignParagraphJustify                                       '表題の配置（両端揃え）
'        Public Const lngTitleIndents_Kouyou As Long = 3                                                           '表題のインデント
'        Public Const strListBullFormat_Kouyou As String = "・"                                                    '箇条書きの記号
'        Public Const strListNumbFormat_Kouyou As String = "%1　"                                                  '段落番号の記号
'        Public Const lngListNumbStyle_Kouyou As Long = 18                                                         '段落番号のスタイル
'        Public Const lngListLeftIndent_Kouyou As Long = 1                                                         'リストの見出しからの左インデント増加量
'        Public Const lngListFirstLineIndent_Kouyou As Long = -1                                                   'リストの最初の行のインデント
    
        '「資料3」に準拠　第１　、１　、(１) 、ア　、(ア) 、a　、(a) 、i　、(i)
'        Public Const strNumberFormats_Kouyou As String = "第%1　,%2　,(%3) ,%4　,(%5) ,%6　,(%7) ,%8　,(%9) "     '番号書式
'        Public Const strNumberStyles_Kouyou As String = "14,14,14,20,20,4,4,2,2"                                  '番号スタイル
'        Public Const strResetOnHighers_Kouyou As String = "0,1,2,3,4,5,6,7,8"                                     'リスト開始レベル
'        Public Const strStartAts_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                           '開始番号
'        Public Const strHeadingLeftIndents_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                 '見出しの左インデント
'        Public Const strHeadingFirstLineIndents_Kouyou As String = "-1,-1,-1,-1,-1.5,-.5,-1,-.5,-1"               '見出しの最初の行のインデント
'        Public Const strTextLeftIndents_Kouyou As String = "1,1,1,1,1,1,1,1,1"                                    '本文の左インデント
'        Public Const strTextFirstLineIndents_Kouyou As String = "1,1,1,1,1,1,1,1,1"                               '本文の最初の行のインデント
'        Public Const lngLeftIndentsAdd_Kouyou As Long = 1                                                         '左インデント増加量
'        Public Const lngTitleAlign_Kouyou As Long = wdAlignParagraphJustify                                       '表題の配置（両端揃え）
'        Public Const lngTitleIndents_Kouyou As Long = 3                                                           '表題のインデント
'        Public Const strListBullFormat_Kouyou As String = "・"                                                    '箇条書きの記号
'        Public Const strListNumbFormat_Kouyou As String = "%1　"                                                  '段落番号の記号
'        Public Const lngListNumbStyle_Kouyou As Long = 18                                                         '段落番号のスタイル
'        Public Const lngListLeftIndent_Kouyou As Long = 1                                                         'リストの見出しからの左インデント増加量
'        Public Const lngListFirstLineIndent_Kouyou As Long = -1                                                   'リストの最初の行のインデント
    
        '◯契約書1（法令型）（見出し→条番号タイプ）用
        '「資料4」に準拠　（）、第１条、２、（１）、、、、、
        Public Const strNumberFormats_Keiyaku1 As String = ",第%2条　,%3　,（%4）,%5　,,,,"                        '番号書式
        Public Const strNumberStyles_Keiyaku1 As String = "0,14,14,14,20,0,0,0,0"                                  '番号スタイル
        Public Const strResetOnHighers_Keiyaku1 As String = "0,0,2,3,4,5,6,7,8"                                    'リスト開始レベル
        Public Const strStartAts_Keiyaku1 As String = "1,1,2,1,1,1,1,1,1"                                          '開始番号
        Public Const strHeadingLeftIndents_Keiyaku1 As String = "1,1,1,3,4,0,0,0,0"                                '見出しの左インデント
        Public Const strHeadingFirstLineIndents_Keiyaku1 As String = "0,-1,-1,-2,-1,0,0,0,0"                       '見出しの最初の行のインデント
        Public Const strTextLeftIndents_Keiyaku1 As String = "1,1,1,1,1,1,1,1,1"                                   '本文の左インデント
        Public Const strTextFirstLineIndents_Keiyaku1 As String = "1,1,1,1,1,1,1,1,1"                              '本文の最初の行のインデント
        Public Const lngLeftIndentsAdd_Keiyaku1 As Long = 0                                                        '左インデント増加量
        Public Const lngTitleAlign_Keiyaku1 As Long = wdAlignParagraphJustify                                      '表題の配置（両端揃え）
        Public Const lngTitleIndents_Keiyaku1 As Long = 3                                                          '表題のインデント
        Public Const strListBullFormat_Keiyaku1 As String = "・"                                                   '箇条書きの記号
        Public Const strListNumbFormat_Keiyaku1 As String = "%1　"                                                 '段落番号の記号
        Public Const lngListNumbStyle_Keiyaku1 As Long = 18                                                        '段落番号のスタイル（丸囲み数字）
        Public Const lngListLeftIndent_Keiyaku1 As Long = 1                                                        'リストの見出しからの左インデント増加量
        Public Const lngListFirstLineIndent_Keiyaku1 As Long = -1                                                  'リストの最初の行のインデント

        '「資料4」に準拠　（）、第１条、２、（１）、、、、、（第1項と第2項のインデントが3文字）
'        Public Const strNumberFormats_Keiyaku1 As String = ",第%2条　,%3　,（%4）,%5　,,,,"                       '番号書式
'        Public Const strNumberStyles_Keiyaku1 As String = "0,14,14,14,20,0,0,0,0"                                 '番号スタイル
'        Public Const strResetOnHighers_Keiyaku1 As String = "0,0,2,3,4,5,6,7,8"                                   'リスト開始レベル
'        Public Const strStartAts_Keiyaku1 As String = "1,1,2,1,1,1,1,1,1"                                         '開始番号
'        Public Const strHeadingLeftIndents_Keiyaku1 As String = "0,3,3,4,5,0,0,0,0"                               '見出しの左インデント
'        Public Const strHeadingFirstLineIndents_Keiyaku1 As String = "0,-3,-1,-2,-1,0,0,0,0"                      '見出しの最初の行のインデント
'        Public Const strTextLeftIndents_Keiyaku1 As String = "1,1,1,1,1,1,1,1,1"                                  '本文の左インデント
'        Public Const strTextFirstLineIndents_Keiyaku1 As String = "1,1,1,1,1,1,1,1,1"                             '本文の最初の行のインデント
'        Public Const lngLeftIndentsAdd_Keiyaku1 As Long = 0                                                       '左インデント増加量
'        Public Const lngTitleAlign_Keiyaku1 As Long = wdAlignParagraphJustify                                     '表題の配置（両端揃え）
'        Public Const lngTitleIndents_Keiyaku1 As Long = 3                                                         '表題のインデント
'        Public Const strListBullFormat_Keiyaku1 As String = "・"                                                  '箇条書きの記号
'        Public Const strListNumbFormat_Keiyaku1 As String = "%1　"                                                '段落番号の記号
'        Public Const lngListNumbStyle_Keiyaku1 As Long = 18                                                       '段落番号のスタイル
'        Public Const lngListLeftIndent_Keiyaku1 As Long = 1                                                       'リストの見出しからの左インデント増加量
'        Public Const lngListFirstLineIndent_Keiyaku1 As Long = -1                                                 'リストの最初の行のインデント
        
        '◯契約書2（実務型）（条番号→見出しタイプ）用
        '「資料5」に準拠　第１条（）、１　、(1) 、ア、、、、、
        Public Const strNumberFormats_Keiyaku2 As String = "第%1条,%2　,(%3) ,%4　,,,,,"                           '番号書式
        Public Const strNumberStyles_Keiyaku2 As String = "14,14,0,20,0,0,0,0,0"                                   '番号スタイル
        Public Const strResetOnHighers_Keiyaku2 As String = "0,1,2,3,4,5,6,7,8"                                    'リスト開始レベル
        Public Const strStartAts_Keiyaku2 As String = "1,1,1,1,1,1,1,1,1"                                          '開始番号
        Public Const strHeadingLeftIndents_Keiyaku2 As String = "1,1,1,1,0,0,0,0,0"                                '見出しの左インデント
        Public Const strHeadingFirstLineIndents_Keiyaku2 As String = "-1,-1,-1,-1,0,0,0,0,0"                       '見出しの最初の行のインデント
        Public Const strTextLeftIndents_Keiyaku2 As String = "0,1,1,1,1,1,1,1,1"                                   '本文の左インデント
        Public Const strTextFirstLineIndents_Keiyaku2 As String = "1,1,1,1,1,1,1,1,1"                              '本文の最初の行のインデント
        Public Const lngLeftIndentsAdd_Keiyaku2 As Long = 1                                                        '左インデント増加量
        Public Const lngTitleAlign_Keiyaku2 As Long = wdAlignParagraphCenter                                       '表題の配置（中央揃え）
        Public Const lngTitleIndents_Keiyaku2 As Long = 0                                                          '表題のインデント
        Public Const strListBullFormat_Keiyaku2 As String = "・"                                                   '箇条書きの記号
        Public Const strListNumbFormat_Keiyaku2 As String = "%1　"                                                 '段落番号の記号
        Public Const lngListNumbStyle_Keiyaku2 As Long = 18                                                        '段落番号のスタイル（丸囲み数字）
        Public Const lngListLeftIndent_Keiyaku2 As Long = 1                                                        'リストの見出しからの左インデント増加量
        Public Const lngListFirstLineIndent_Keiyaku2 As Long = -1                                                  'リストの最初の行のインデント

        '◯ビジネス文書用（左インデント増加量1文字）
        '第1　、1　、(1) 、ア　、(ア)　、a　、(a) 、i　、(i)
        Public Const strNumberFormats_Business As String = "第%1 ,%2　,(%3) ,%4　,(%5) ,%6　,(%7) ,%8　,(%9) "     '番号書式
        Public Const strNumberStyles_Business As String = "0,0,0,20,20,4,4,2,2"                                    '番号スタイル
        Public Const strResetOnHighers_Business As String = "0,1,2,3,4,5,6,7,8"                                    'リスト開始レベル
        Public Const strStartAts_Business As String = "1,1,1,1,1,1,1,1,1"                                          '開始番号
        Public Const strHeadingLeftIndents_Business As String = "1,1,1,1,1,1,1,1,1"                                '見出しの左インデント
        Public Const strHeadingFirstLineIndents_Business As String = "-1,-.5,-1,-1,-1.5,-.5,-1,-.5,-1"             '見出しの最初の行のインデント
        Public Const strTextLeftIndents_Business As String = "1,1,1,1,1,1,1,1,1"                                   '本文の左インデント
        Public Const strTextFirstLineIndents_Business As String = "1,1,1,1,1,1,1,1,1"                              '本文の最初の行のインデント
        Public Const lngLeftIndentsAdd_Business As Long = 1                                                        '左インデント増加量
        Public Const lngTitleAlign_Business As Long = wdAlignParagraphCenter                                       '表題の配置（中央揃え）
        Public Const lngTitleIndents_Business As Long = 0                                                          '表題のインデント
        Public Const strListBullFormat_Business As String = "・"                                                   '箇条書きの記号
        Public Const strListNumbFormat_Business As String = "%1　"                                                 '段落番号の記号
        Public Const lngListNumbStyle_Business As Long = 18                                                        '段落番号のスタイル（丸囲み数字）
        Public Const lngListLeftIndent_Business As Long = 1                                                        'リストの見出しからの左インデント増加量
        Public Const lngListFirstLineIndent_Business As Long = -1                                                  'リストの最初の行のインデント
    
        '◯技術文書用（左インデント増加量0文字）
        '1.、1.1.、1.1.1.、1.1.1.1.、1.1.1.1.、1.1.1.1.1.、1.1.1.1.1.1.、1.1.1.1.1.1.1.、1.1.1.1.1.1.1.1.
        Public Const strNumberFormats_Technical As String = _
            "%1.,%1.%2.,%1.%2.%3.,%1.%2.%3.%4.,%1.%2.%3.%4.%5.,%1.%2.%3.%4.%5.%6.,%1.%2.%3.%4.%5.%6.%7.,%1.%2.%3.%4.%5.%6.%7.%8.,%1.%2.%3.%4.%5.%6.%7.%8.%9."      '番号書式
        Public Const strNumberStyles_Technical As String = "0,0,0,0,0,0,0,0,0"                                     '番号スタイル
        Public Const strResetOnHighers_Technical As String = "0,1,2,3,4,5,6,7,8"                                   'リスト開始レベル
        Public Const strStartAts_Technical As String = "1,1,1,1,1,1,1,1,1"                                         '開始番号
        Public Const strHeadingLeftIndents_Technical As String = "0,0,0,0,0,0,0,0,0"                               '見出しの左インデント
        Public Const strHeadingFirstLineIndents_Technical As String = "0,0,0,0,0,0,0,0,0"                          '見出しの最初の行のインデント
        Public Const strTextLeftIndents_Technical As String = "0,0,0,0,0,0,0,0,0"                                  '本文の左インデント
        Public Const strTextFirstLineIndents_Technical As String = "0,0,0,0,0,0,0,0,0"                             '本文の最初の行のインデント
        Public Const lngLeftIndentsAdd_Technical As Long = 0                                                       '左インデント増加量
        Public Const lngTitleAlign_Technical As Long = wdAlignParagraphLeft                                        '表題の配置（左揃え）
        Public Const lngTitleIndents_Technical As Long = 0                                                         '表題のインデント
        Public Const strListBullFormat_Technical As String = "-"                                                   '箇条書きの記号
        Public Const strListNumbFormat_Technical As String = "%1)"                                                 '段落番号の記号
        Public Const lngListNumbStyle_Technical As Long = 4                                                        '段落番号のスタイル（英小文字）
        Public Const lngListLeftIndent_Technical As Long = 0                                                       'リストの見出しからの左インデント増加量
        Public Const lngListFirstLineIndent_Technical As Long = 0                                                  'リストの最初の行のインデント
    
    'ページ設定
        '裁判所提出書類 用(「資料1」に準拠)
        Public Const sngPageWidth_Saiban As Single = 210              '用紙サイズの幅（mm）
        Public Const strMargines_Saiban As String = "35,27,30,17.5"   'ページ余白（mm）（上,下,左,右）'右余白は､修正されるのでわずかに増減します｡
        Public Const strCharsAndLines_Saiban As String = "37,26"      '文字数及び行数（文字）（文字数,行数）
                    
        '内容証明用
        '次の定数の組み合わせの中から、1組をコメント解除してください。
        '句読点のぶら下げに対応するため、文字数は制限値よりも1文字少なくしています。
        '「資料2」に準拠　横書き（その1）
        Public Const sngPageWidth_Naiyou As Single = 210              '用紙サイズの幅（mm）
        Public Const strMargines_Naiyou As String = "30,30,30,30"     'ページ余白（mm）（上,下,左,右）'右余白は､修正されるのでわずかに増減します｡
        Public Const strCharsAndLines_Naiyou As String = "19,26"      '文字数及び行数（文字）（文字数,行数）
    
        '「資料2」に準拠　横書き（その2）
'        Public Const sngPageWidth_Naiyou As Single = 210            '用紙サイズの幅（mm）
'        Public Const strMargines_Naiyou As String = "30,30,30,30"   'ページ余白（mm）（上,下,左,右）'右余白は､修正されるのでわずかに増減します｡
'        Public Const strCharsAndLines_Naiyou As String = "12,40"    '文字数及び行数（文字）（文字数,行数）
    
        '「資料2」に準拠　横書き（その3）
'        Public Const sngPageWidth_Naiyou As Single = 210            '用紙サイズの幅（mm）
'        Public Const strMargines_Naiyou As String = "30,30,30,30"   'ページ余白（mm）（上,下,左,右）'右余白は､修正されるのでわずかに増減します｡
'        Public Const strCharsAndLines_Naiyou As String = "25,20"    '文字数及び行数（文字）（文字数,行数）
        
        '公文書（「資料6」に準拠）
        Public Const sngPageWidth_Kouyou As Single = 210              '用紙サイズの幅（mm）
        Public Const strMargines_Kouyou As String = "25,20,25,20"     'ページ余白（mm）（上,下,左,右）'右余白は､修正されるのでわずかに増減します｡
        Public Const strCharsAndLines_Kouyou As String = "40,40"      '文字数及び行数（文字）（文字数,行数）

        '契約書1（法令型）用（公文書と共通）
        Public Const sngPageWidth_Keiyaku1 As Single = 210            '用紙サイズの幅（mm）
        Public Const strMargines_Keiyaku1 As String = "25,20,25,20"   'ページ余白（mm）（上,下,左,右）'右余白は､修正されるのでわずかに増減します｡
        Public Const strCharsAndLines_Keiyaku1 As String = "40,40"    '文字数及び行数（文字）（文字数,行数）
        
        '契約書2（実務型）用（ビジネス文書と共通）
        Public Const sngPageWidth_keiyaku2 As Single = 210            '用紙サイズの幅（mm）
        Public Const strMargines_keiyaku2 As String = "25,25,25,25"   'ページ余白（mm）（上,下,左,右）'右余白は､修正されるのでわずかに増減します｡
        Public Const strCharsAndLines_keiyaku2 As String = "38,38"    '文字数及び行数（文字）（文字数,行数）
        
        'ビジネス文書用
        Public Const sngPageWidth_Business As Single = 210           '用紙サイズの幅（mm）
        Public Const strMargines_Business As String = "25,25,25,25"  'ページ余白（mm）（上,下,左,右）'右余白は､修正されるのでわずかに増減します｡
        Public Const strCharsAndLines_Business As String = "38,38"   '文字数及び行数（文字）（文字数,行数）

        '技術文書用
        Public Const sngPageWidth_Technical As Single = 210          '用紙サイズの幅（mm）
        Public Const strMargines_Technical As String = "25,25,25,25" 'ページ余白（mm）（上,下,左,右）'右余白は､修正されるのでわずかに増減します｡
        Public Const strCharsAndLines_Technical As String = "38,38"  '文字数及び行数（文字）（文字数,行数）

'設定用定数（値が固定のもの）を宣言する
'すべてのモジュールで使用できるようにするため、Publicで宣言しています。
'各モジュール内のプロシージャから必要に応じて参照されます。
    'スタイル
    Public Const sngFont As Single = 12                         'フォントサイズ（ポイント単位で入力します。）
    Public Const strFont As String = "ＭＳ 明朝"                'フォント名（「ＭＳ」と「明朝」の間に半角スペースが必要です。）
'    Public Const strFont As String = "游明朝"                  'フォント名
    Public Const sngLineNo As Single = 8                        '行番号のフォントサイズ（ポイント単位で入力します。）
    Public Const sngTitleFont As Single = 12                    '表題のフォントサイズ（ポイント単位で入力します。）
    Public Const strTitleFont As String = "ＭＳ 明朝"           '表題のフォント名（「ＭＳ」と「明朝」の間に半角スペースが必要です。）
'    Public Const strTitleFont As String = "ＭＳ ゴシック"      '表題のフォント名（「ＭＳ」と「ゴシック」の間に半角スペースが必要です。）
'    Public Const strTitleFont As String = "游ゴシック"         '表題のフォント名
'    Public Const strTitleFont As String = "HGｺﾞｼｯｸE"           '表題のフォント名（すべて半角です。）
    Public Const lngJust As Long = wdJustificationModeExpand    '文字間隔の調整（「間隔を詰めない」に設定しています。）
    Public Const lngAlign As Long = wdAlignParagraphJustify     '配置（句読点のぶら下げを行うため、「両端揃え」に設定しています。）
    
    'ページ
    Public Const blnPageNumberEnd As Boolean = True             '右上のページ番号の最後に"」"を追加する（デフォルト：True）
    
    'マークダウン記号
    Public Const strHead As String = "#"                        '見出しレベル（デフォルト："#"）
                                                                '「.」を指定した場合、セットアップ時にオートコレクトの「...」→「…」のエントリが削除されます。
                                                                '元に戻すには、Wordのオプション→文章校正→オートコレクトのオプションで「...」を再登録してください。
    Public Const strList As String = "-"                        '箇条書き（デフォルト："-"）
    Public Const strNumList As String = "1."                    '段落番号（"1."変更不可）
    Public Const blnMarkdownSpace As Boolean = True             'マークダウン記号の後の空白（デフォルト：True）
    
    'マクロ
    Public Const blnReplaceSpaceWithIndent As Boolean = True    '行頭の空白をインデントに置き換える（デフォルト：True）
    Public Const blnSkipUntil As Boolean = True                 '最初の見出し、箇条書きまたは段落番号までの処理を停止する（デフォルト:True）
    Public Const blnStopAfter As Boolean = True                 '配置が右揃えで処理を停止する（デフォルト：True）
    Public Const blnHeadingReset As Boolean = True              'セクション区切りで見出し番号を振り直す（デフォルト：True）
    Public Const blnApplyTableStyle As Boolean = True           '表内の段落に「行間詰め」スタイルを適用する（デフォルト：True）
    Public Const blnSelectionBlock As Boolean = True            '選択範囲の「行間詰め」スタイルの適用を行う（デフォルト：True）
    Public Const blnConfigGallery = True                        'スタイルおよびリストのギャラリーから不要なスタイルを削除する（デフォルト：True）
    Public Const blnReconfigOutline As Boolean = True           'インデント設定時に最初のアウトラインやリストの設定を適用する（デフォルト：True）
    Public Const blnEnableExternalCleanup As Boolean = True     '外部ツール由来要素（手動改行・リンク・WikiLink・フロントマター・水平線・強調記号）の一括整形を有効にする（デフォルト：True）
    '変換（ニュートラル・マークダウン・アウトライン）を Ctrl+Z 1回で元に戻せるようにする（デフォルト：False）
    '元に戻す（Undo）の記録（Application.UndoRecord）の中で「すべて置換」を行うと、Wordが強制終了するため（Word 16.0.20430 で確認。
    '外部ツール由来要素の一括整形で必ず起きた）、Windows版・Mac版とも使わない。安定性を優先し、True にしないこと（Ver8D3）。
    Public Const blnUseUndoRecord As Boolean = False

    'サンプル文面
        '裁判所提出書類、内容証明、公文書用
        '説明タイプ
'        Public Const strSampleContents_Saiban As String = _
'            "＜＜サンプル＞＞" & vbCr & _
'            "↑↑↑最初の見出しまでは処理の対象外です。↑↑↑" & vbCr & _
'            "#見出し（「第１」ではなく、「１」から開始することもできます。）" & vbCr & _
'            "本文（見出しに合わせて自動的にインデントが設定されます。）" & vbCr & _
'            "##見出し（見出しレベルは、スタイルまたはアウトラインで設定できます。）" & vbCr & _
'            "本文" & vbCr & _
'            "-箇条書き（「・」を使ったリストになります。）" & vbCr & _
'            "-箇条書き" & vbCr & _
'            "##見出し" & vbCr & _
'            "本文" & vbCr & _
'            "1.段落番号（「①」を使った番号付きリストになります。）" & vbCr & _
'            "1.段落番号" & vbCr & _
'            "###見出し" & vbCr & _
'            "####見出し" & vbCr & _
'            "#####見出し" & vbCr & _
'            "######見出し" & vbCr & _
'            "#######見出し" & vbCr & _
'            "########見出し" & vbCr & _
'            "#########見出し" & vbCr & _
'            "↓↓↓右揃えの段落以降は処理の対象外です。↓↓↓" & vbCr & _
'            "付記"

        '確認タイプ
        '裁判所提出書類用等の確認タイプのサンプル文字列を設定する
        Public Const strSampleContents_Saiban As String = _
            "＜＜サンプル＞＞" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "#ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "##ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "-ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "-ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "##ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "1.ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "1.ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "###ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "####ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "#####ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "######ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "#######ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "########ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "#########ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ"

        '契約書1用
        '説明タイプ
'        Public Const strSampleContents_Keiyaku1 As String = _
'            "＜＜サンプル＞＞" & vbCr & _
'            "↑↑↑最初の見出しまではマクロが適用されません。↑↑↑" & vbCr & _
'            "#（見出し）" & vbCr & _
'            "##条番号および項" & vbCr & _
'            "###項（省略することもできます。）" & vbCr & _
'            "####号" & vbCr & _
'            "#####号の細分" & vbCr & _
'            "####号" & vbCr & _
'            "-箇条書き（「・」を使ったリストになります。）" & vbCr & _
'            "-箇条書き" & vbCr & _
'            "####号" & vbCr & _
'            "1.段落番号（「①」を使った番号付きリストになります。）" & vbCr & _
'            "1.段落番号" & vbCr & _
'            "↓↓↓右揃えの段落以降は処理の対象外です。↓↓↓" & vbCr & _
'            "付記"
        
        '確認タイプ
        '契約書1用の確認タイプのサンプル文字列を設定する
        Public Const strSampleContents_Keiyaku1 As String = _
            "＜＜サンプル＞＞" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "#（ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ）" & vbCr & _
            "##ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "###ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "####ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "#####ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "####ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "-ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "-ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "####ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "1.ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "1.ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ"

        '契約書2用
        '説明タイプ
'        Public Const strSampleContents_Keiyaku2 As String = _
'            "＜＜サンプル＞＞" & vbCr & _
'            "↑↑↑最初の見出しまではマクロが適用されません。↑↑↑" & vbCr & _
'            "#（見出し）" & vbCr & _
'            "本文" & vbCr & _
'            "##項" & vbCr & _
'            "###号" & vbCr & _
'            "####号の細分" & vbCr & _
'            "###号" & vbCr & _
'            "-箇条書き（「・」を使ったリストになります。）" & vbCr & _
'            "-箇条書き" & vbCr & _
'            "###号" & vbCr & _
'            "1.段落番号（「①」を使った番号付きリストになります。）" & vbCr & _
'            "1.段落番号" & vbCr & _
'            "↓↓↓右揃えの段落以降は処理の対象外です。↓↓↓" & vbCr & _
'            "付記"
    
        '確認タイプ
        '契約書2用の確認タイプのサンプル文字列を設定する
        Public Const strSampleContents_Keiyaku2 As String = _
            "＜＜サンプル＞＞" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "#（ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ）" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "##ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "###ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "####ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "###ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "-ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "-ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "###ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "1.ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "1.ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ" & vbCr & _
            "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ"


