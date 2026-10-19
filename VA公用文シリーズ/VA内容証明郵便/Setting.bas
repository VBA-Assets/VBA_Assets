'変数宣言を強制する
Option Explicit

'定数を宣言する
'すべてのモジュールで使用できるようにするため、Publicで宣言しています。

'準拠：「内容証明　ご利用の条件等」https://www.post.japanpost.jp/service/fuka_service/syomei/use.html

'次の定数の組み合わせの中から、1組をコメント解除してください。
'句読点のぶら下げに対応するため、文字数は制限値よりも1文字少なくしています。

'    横書き（その1）
    Public Const strMaxCharsAndLines = "20,26"             '文字数及び行数の制限値（文字）（文字数,行数）
'
'    横書き（その2）
'    Public Const strMaxCharsAndLines = "13,40"             '文字数及び行数の制限値（文字）（文字数,行数）
'
'    横書き（その3）
'    Public Const strMaxCharsAndLines = "26,20"             '文字数及び行数の制限値（文字）（文字数,行数）