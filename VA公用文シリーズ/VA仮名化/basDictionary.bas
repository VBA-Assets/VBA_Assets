Option Explicit
'==================================================================================================
'   概要        : 文書ごとの辞書（登録した語と記号の対応）を、文書の文書変数に読み書きします。
'                 また、記号の作成・判定と、登録できる語かどうかの判定を行います。
'
'   保存の形    : 辞書の各項目は、文書変数「<接頭辞>Entry<番号>」に「記号<Tab>語」の形で保存します。
'                 （ファイルの読み書きをしないのは、Mac版Wordのサンドボックスと文字コードの問題を避けるため）
'                 仮名化中かどうかは記録しません（文書変数の変更は Ctrl+Z で元に戻らないため、記録すると、
'                 仮名化を Ctrl+Z で取り消した後に、記録と本文が食い違うため）。
'
'   ファイルに残す期間 : 辞書は、本文に記号（〔人名Ａ〕の形のもの）がある間だけ、文書変数に残します。
'                 実名にすべて戻して本文に記号がなくなったら、文書変数から消し、その文書を閉じるまでは
'                 Word の中（docMemory・strMemory）にだけ覚えておきます（2026-10-17 班長決定）。
'                 覚えておく辞書は、文書の名前ではなく、開いている文書そのもの（Is で比べる）で見分けます
'                 （名前を付けて保存してフルパスが変わっても引き継ぎ、別のファイルの辞書が混ざらないようにするため）。
'                 消し忘れた辞書（実名）がファイルに隠れて残り、文書をほかの事件に使い回したときに
'                 漏れたり、別の事件の名前に戻ったりするのを防ぐため。ファイルに辞書が入っているのは、
'                 本文に記号が見えているときだけになる。
'                 各処理の初めに RestoreFromMemory、終わりに StoreToMemory を呼びます。
'==================================================================================================

'文書変数の名前（接頭辞の後に付ける部分）
Private Const strEntryName As String = "Entry" ' 辞書の項目の文書変数名

'ファイルから消した辞書を、その文書を閉じるまで覚えておく場所
'（docMemory(i) の文書の辞書が strMemory(i)。辞書は「記号<Tab>語」を改行でつないだ文字列）
Private docMemory() As Document ' 文書
Private strMemory() As String ' 辞書
Private lngMemoryCount As Long ' 覚えている文書の数

'英字の文字コード（AscW は &H8000 以上の文字で負の値を返すため、Long で扱う）
Private Const lngWideA As Long = &HFF21& ' 全角の「Ａ」
Private Const lngWideZ As Long = &HFF3A& ' 全角の「Ｚ」
Private Const lngNarrowA As Long = &H41& ' 半角の「A」
Private Const lngNarrowZ As Long = &H5A& ' 半角の「Z」
Private Const lngWideSmallA As Long = &HFF41& ' 全角の「ａ」
Private Const lngWideSmallZ As Long = &HFF5A& ' 全角の「ｚ」
Private Const lngNarrowSmallA As Long = &H61& ' 半角の「a」
Private Const lngNarrowSmallZ As Long = &H7A& ' 半角の「z」

'================================================================================
' 辞書の読み書き
'================================================================================

'================================================================================
' 機能:         文書の辞書の項目をすべて読み込みます。
' 引数:         doc - 対象の文書
'               strTokens - 記号を格納する配列（1始まり）
'               strWords - 語を格納する配列（1始まり）
'               strVarNames - 項目を保存している文書変数の名前を格納する配列（1始まり）
' 戻り値:       項目の数
'================================================================================
Public Function LoadEntries(ByVal doc As Document, ByRef strTokens() As String, _
                            ByRef strWords() As String, ByRef strVarNames() As String) As Long
    Dim v As Variable ' 文書変数
    Dim lngCount As Long ' 項目の数
    Dim lngTab As Long ' 記号と語の区切りの位置

    lngCount = 0
    'すべての文書変数を繰り返す
    For Each v In doc.Variables
        '辞書の項目の文書変数の場合
        If IsEntryVarName(v.Name) Then
            '記号と語の区切りを探す
            lngTab = InStr(v.Value, vbTab)
            '区切りがある場合のみ項目として扱う（壊れた値は無視する）
            If lngTab > 1 And lngTab < Len(v.Value) Then
                lngCount = lngCount + 1
                '配列を広げる
                ReDim Preserve strTokens(1 To lngCount)
                ReDim Preserve strWords(1 To lngCount)
                ReDim Preserve strVarNames(1 To lngCount)
                '記号・語・文書変数の名前を格納する
                strTokens(lngCount) = Left(v.Value, lngTab - 1)
                strWords(lngCount) = Mid(v.Value, lngTab + 1)
                strVarNames(lngCount) = v.Name
            End If
        End If
    Next v

    LoadEntries = lngCount
End Function

'================================================================================
' 機能:         指定した文字列が、登録済みの語または記号かどうかを調べます。
'               記号は、英字が半角・小文字になったもの（〔人名A〕〔人名a〕）も同じ記号とみなします。
' 引数:         doc - 対象の文書
'               strText - 調べる文字列
'               strToken, strWord, strVarName - 見つかった項目の記号・語・文書変数の名前（戻り値）
' 戻り値:       見つかった場合は True
'================================================================================
Public Function FindEntry(ByVal doc As Document, ByVal strText As String, ByRef strToken As String, _
                          ByRef strWord As String, ByRef strVarName As String) As Boolean
    Dim strTokens() As String ' 記号
    Dim strWords() As String ' 語
    Dim strVarNames() As String ' 文書変数の名前
    Dim lngCount As Long ' 項目の数
    Dim i As Long

    FindEntry = False
    '辞書を読み込む
    lngCount = LoadEntries(doc, strTokens, strWords, strVarNames)
    'すべての項目を繰り返す
    For i = 1 To lngCount
        '語または記号が一致する場合
        If strWords(i) = strText Or strTokens(i) = ToWideCapitalLetters(strText) Then
            strToken = strTokens(i)
            strWord = strWords(i)
            strVarName = strVarNames(i)
            FindEntry = True
            Exit Function
        End If
    Next i
End Function

'================================================================================
' 機能:         辞書に項目を追加します。
'               記号の英字は、種別ごとに、登録済みの記号の最大のものの次とします。
' 引数:         doc - 対象の文書
'               strType - 種別（例：人名）
'               strWord - 登録する語
' 戻り値:       追加した項目の記号（例：〔人名Ａ〕）
'================================================================================
Public Function AddEntry(ByVal doc As Document, ByVal strType As String, ByVal strWord As String) As String
    Dim v As Variable ' 文書変数
    Dim lngMaxId As Long ' 文書変数の名前の番号の最大値
    Dim lngMaxNumber As Long ' 同じ種別の記号の英字の番号の最大値
    Dim lngTab As Long ' 記号と語の区切りの位置
    Dim strType2 As String ' 登録済みの記号の種別
    Dim lngNumber As Long ' 登録済みの記号の英字の番号
    Dim strToken As String ' 追加する記号

    lngMaxId = 0
    lngMaxNumber = 0
    'すべての文書変数を繰り返す
    For Each v In doc.Variables
        '辞書の項目の文書変数の場合
        If IsEntryVarName(v.Name) Then
            '文書変数の名前の番号の最大値を求める
            If Val(Mid(v.Name, Len(strVarPrefix & strEntryName) + 1)) > lngMaxId Then
                lngMaxId = Val(Mid(v.Name, Len(strVarPrefix & strEntryName) + 1))
            End If
            '同じ種別の記号の英字の番号の最大値を求める
            lngTab = InStr(v.Value, vbTab)
            If lngTab > 1 Then
                If GetTokenParts(Left(v.Value, lngTab - 1), strType2, lngNumber) Then
                    If strType2 = strType And lngNumber > lngMaxNumber Then lngMaxNumber = lngNumber
                End If
            End If
        End If
    Next v

    '記号を作る（例：〔人名Ａ〕）
    strToken = strTokenOpen & strType & NumberToLetters(lngMaxNumber + 1) & strTokenClose
    '文書変数に保存する
    doc.Variables.Add strVarPrefix & strEntryName & CStr(lngMaxId + 1), strToken & vbTab & strWord

    AddEntry = strToken
End Function

'================================================================================
' 機能:         辞書から項目を削除します。
' 引数:         doc - 対象の文書
'               strVarName - 項目を保存している文書変数の名前
'================================================================================
Public Sub DeleteEntry(ByVal doc As Document, ByVal strVarName As String)
    Dim v As Variable ' 文書変数

    'すべての文書変数を繰り返す
    For Each v In doc.Variables
        '名前が一致する文書変数を削除する
        If v.Name = strVarName Then
            v.Delete
            Exit For
        End If
    Next v
End Sub

'================================================================================
' 機能:         ファイルから消して Word の中に覚えておいた辞書を、文書変数に書き戻します。
'               文書変数に辞書がある場合は、それを使うので何もしません。
'               各処理（登録・仮名化・実名に戻す）の初めに呼びます。
' 引数:         doc - 対象の文書
'================================================================================
Public Sub RestoreFromMemory(ByVal doc As Document)
    Dim strTokens() As String ' 記号
    Dim strWords() As String ' 語
    Dim strVarNames() As String ' 文書変数の名前
    Dim strValue As String ' 覚えておいた辞書
    Dim varLines As Variant ' 覚えておいた辞書の各項目
    Dim i As Long

    '文書変数に辞書がある場合は何もしない
    If LoadEntries(doc, strTokens, strWords, strVarNames) > 0 Then Exit Sub

    '覚えておいた辞書がない場合は何もしない
    strValue = GetMemory(doc)
    If strValue = "" Then Exit Sub

    '各項目を文書変数に書き戻す
    varLines = Split(strValue, vbLf)
    For i = LBound(varLines) To UBound(varLines)
        If varLines(i) <> "" Then doc.Variables.Add strVarPrefix & strEntryName & CStr(i + 1), varLines(i)
    Next i
End Sub

'================================================================================
' 機能:         辞書を Word の中に覚え、本文に記号が1つもない場合は、辞書を文書変数から消します。
'               各処理（登録・仮名化・実名に戻す）の終わりに呼びます。
'               本文に記号が残っている場合（仮名化中、または実名に戻せなかった記号がある場合）は、
'               実名に戻すのに必要なので、文書変数に残します。
' 引数:         doc - 対象の文書
' 戻り値:       辞書を文書変数から消した場合は True
'================================================================================
Public Function StoreToMemory(ByVal doc As Document) As Boolean
    Dim strTokens() As String ' 記号
    Dim strWords() As String ' 語
    Dim strVarNames() As String ' 文書変数の名前
    Dim lngCount As Long ' 項目の数
    Dim strValue As String ' 覚えておく辞書
    Dim i As Long

    StoreToMemory = False

    '辞書を読み込む
    lngCount = LoadEntries(doc, strTokens, strWords, strVarNames)
    '辞書がない場合（登録をすべて解除した場合など）は、覚えておいた辞書も消す
    If lngCount = 0 Then
        Call SetMemory(doc, "")
        Exit Function
    End If

    '辞書を Word の中に覚える（改行でつなぐ）
    strValue = ""
    For i = 1 To lngCount
        If i > 1 Then strValue = strValue & vbLf
        strValue = strValue & strTokens(i) & vbTab & strWords(i)
    Next i
    Call SetMemory(doc, strValue)

    '本文に記号が1つもない場合は、辞書を文書変数から消す
    If CountTokens(doc.Content.Text) = 0 Then
        For i = 1 To lngCount
            Call DeleteEntry(doc, strVarNames(i))
        Next i
        StoreToMemory = True
    End If
End Function

'================================================================================
' 機能:         Word の中に覚えておいた辞書を取り出します。
' 引数:         doc - 対象の文書
' 戻り値:       覚えておいた辞書（ない場合は空の文字列）
'================================================================================
Private Function GetMemory(ByVal doc As Document) As String
    Dim i As Long

    GetMemory = ""
    For i = 1 To lngMemoryCount
        '同じ文書かどうかは、文書そのもので比べる
        If docMemory(i) Is doc Then
            GetMemory = strMemory(i)
            Exit Function
        End If
    Next i
End Function

'================================================================================
' 機能:         辞書を Word の中に覚えます（同じ文書の古い内容は置き換え、空の場合は消します）。
'               あわせて、閉じた文書の分を消します（閉じた文書の辞書は使わないため）。
' 引数:         doc - 対象の文書
'               strValue - 覚える辞書（空の場合は、覚えておいた辞書を消す）
'================================================================================
Private Sub SetMemory(ByVal doc As Document, ByVal strValue As String)
    Dim docNew() As Document ' 残す文書
    Dim strNew() As String ' 残す辞書
    Dim lngNew As Long ' 残す数
    Dim blnOpen As Boolean ' 文書が開いているかどうか
    Dim strName As String ' 文書の名前（開いているかを確かめるために読む）
    Dim i As Long

    '対象の文書の分と、閉じた文書の分を除いて、残りを詰め直す
    lngNew = 0
    For i = 1 To lngMemoryCount
        If Not (docMemory(i) Is doc) Then
            '閉じた文書は、名前を読むとエラーになる
            On Error Resume Next
            strName = docMemory(i).Name
            blnOpen = (Err.Number = 0)
            Err.Clear
            On Error GoTo 0
            If blnOpen Then
                lngNew = lngNew + 1
                ReDim Preserve docNew(1 To lngNew)
                ReDim Preserve strNew(1 To lngNew)
                Set docNew(lngNew) = docMemory(i)
                strNew(lngNew) = strMemory(i)
            End If
        End If
    Next i

    '対象の文書の辞書を加える（空の場合は加えない＝消す）
    If strValue <> "" Then
        lngNew = lngNew + 1
        ReDim Preserve docNew(1 To lngNew)
        ReDim Preserve strNew(1 To lngNew)
        Set docNew(lngNew) = doc
        strNew(lngNew) = strValue
    End If

    '置き換える
    lngMemoryCount = lngNew
    If lngNew > 0 Then
        docMemory = docNew
        strMemory = strNew
    Else
        Erase docMemory
        Erase strMemory
    End If
End Sub

'================================================================================
' 記号
'================================================================================

'================================================================================
' 機能:         文字列が、種別の一覧にある種別かどうかを調べます。
' 引数:         strType - 調べる文字列
' 戻り値:       種別の一覧にある場合は True
'================================================================================
Public Function IsValidType(ByVal strType As String) As Boolean
    Dim varTypes As Variant ' 種別の一覧
    Dim i As Long

    IsValidType = False
    varTypes = Split(strTokenTypes, ",")
    For i = LBound(varTypes) To UBound(varTypes)
        If varTypes(i) = strType Then
            IsValidType = True
            Exit Function
        End If
    Next i
End Function

'================================================================================
' 機能:         文字列の中の記号（辞書にないものを含む）の数を数えます。
'               記号は、英字が半角・小文字になったもの（〔人名A〕〔人名a〕〔人名ａ〕）や、
'               全角と半角が混ざったもの（〔人名ＡA〕）も数えます（戻せなかった記号を知らせ漏らさないため）。
' 引数:         strText - 対象の文字列
' 戻り値:       記号の数
'================================================================================
Public Function CountTokens(ByVal strText As String) As Long
    Dim lngPos As Long ' 開き括弧の位置
    Dim lngEnd As Long ' 閉じ括弧の位置
    Dim strType As String ' 種別（GetTokenParts の戻り値。ここでは使わない）
    Dim lngNumber As Long ' 英字の番号（GetTokenParts の戻り値。ここでは使わない）
    Dim lngCount As Long ' 記号の数

    lngCount = 0
    '最初の開き括弧を探す
    lngPos = InStr(1, strText, strTokenOpen)
    Do While lngPos > 0
        'その後ろの閉じ括弧を探す
        lngEnd = InStr(lngPos + 1, strText, strTokenClose)
        '閉じ括弧がない場合は終了する
        If lngEnd = 0 Then Exit Do
        '開き括弧から閉じ括弧までが記号の場合は数える
        If GetTokenParts(Mid(strText, lngPos, lngEnd - lngPos + 1), strType, lngNumber) Then lngCount = lngCount + 1
        '次の開き括弧を探す（「〔〔人名Ａ〕」のように開き括弧が続く場合に備え、1文字後ろから探す）
        lngPos = InStr(lngPos + 1, strText, strTokenOpen)
    Loop

    CountTokens = lngCount
End Function

'================================================================================
' 機能:         記号を種別と英字の番号に分けます。
'               英字が半角・小文字になったもの（〔人名A〕〔人名a〕〔人名ａ〕）も記号とみなします。
' 引数:         strToken - 記号（例：〔人名Ａ〕）
'               strType - 種別（戻り値。例：人名）
'               lngNumber - 英字の番号（戻り値。Ａ=1、Ｚ=26、ＡＡ=27）
' 戻り値:       記号の形になっている場合は True
'================================================================================
Public Function GetTokenParts(ByVal strToken As String, ByRef strType As String, ByRef lngNumber As Long) As Boolean
    Dim strInner As String ' 括弧の内側
    Dim varTypes As Variant ' 種別の一覧
    Dim i As Long

    GetTokenParts = False
    '括弧で囲まれていない場合は記号ではない
    If Len(strToken) < Len(strTokenOpen) + Len(strTokenClose) + 2 Then Exit Function
    If Left(strToken, Len(strTokenOpen)) <> strTokenOpen Then Exit Function
    If Right(strToken, Len(strTokenClose)) <> strTokenClose Then Exit Function

    '括弧の内側を取り出す
    strInner = Mid(strToken, Len(strTokenOpen) + 1, Len(strToken) - Len(strTokenOpen) - Len(strTokenClose))

    'すべての種別を繰り返す
    varTypes = Split(strTokenTypes, ",")
    For i = LBound(varTypes) To UBound(varTypes)
        '種別で始まる場合
        If Left(strInner, Len(varTypes(i))) = varTypes(i) Then
            '残りが英字だけの場合は記号とみなす
            lngNumber = LettersToNumber(ToWideCapitalLetters(Mid(strInner, Len(varTypes(i)) + 1)))
            If lngNumber > 0 Then
                strType = varTypes(i)
                GetTokenParts = True
                Exit Function
            End If
        End If
    Next i
End Function

'================================================================================
' 機能:         記号の英字を半角にした文字列を返します（例：〔人名Ａ〕→〔人名A〕）。
'               AIが英字を半角にして出力した記号を戻すために使います。
' 引数:         strText - 対象の文字列
' 戻り値:       全角の英字を半角にした文字列
'================================================================================
Public Function ToNarrowLetters(ByVal strText As String) As String
    Dim i As Long
    Dim lngCode As Long ' 文字コード
    Dim strResult As String ' 結果

    strResult = ""
    '1文字ずつ繰り返す
    For i = 1 To Len(strText)
        lngCode = CharCode(Mid(strText, i, 1))
        '全角の英字の場合は半角にする
        If lngCode >= lngWideA And lngCode <= lngWideZ Then
            strResult = strResult & ChrW(lngCode - lngWideA + lngNarrowA)
        Else
            strResult = strResult & Mid(strText, i, 1)
        End If
    Next i

    ToNarrowLetters = strResult
End Function

'================================================================================
' 機能:         記号の英字を小文字にした文字列を返します（例：〔人名Ａ〕→〔人名ａ〕、〔人名A〕→〔人名a〕）。
'               全角・半角は変えません。AIが英字を小文字にして出力した記号を戻すために使います。
' 引数:         strText - 対象の文字列
' 戻り値:       英字の大文字を小文字にした文字列
'================================================================================
Public Function ToSmallLetters(ByVal strText As String) As String
    Dim i As Long
    Dim lngCode As Long ' 文字コード
    Dim strResult As String ' 結果

    strResult = ""
    '1文字ずつ繰り返す
    For i = 1 To Len(strText)
        lngCode = CharCode(Mid(strText, i, 1))
        '全角の英字の大文字の場合は、全角の小文字にする
        If lngCode >= lngWideA And lngCode <= lngWideZ Then
            strResult = strResult & ChrW(lngCode - lngWideA + lngWideSmallA)
        '半角の英字の大文字の場合は、半角の小文字にする
        ElseIf lngCode >= lngNarrowA And lngCode <= lngNarrowZ Then
            strResult = strResult & ChrW(lngCode - lngNarrowA + lngNarrowSmallA)
        Else
            strResult = strResult & Mid(strText, i, 1)
        End If
    Next i

    ToSmallLetters = strResult
End Function

'================================================================================
' 登録できる語の判定
'================================================================================

'================================================================================
' 機能:         選択範囲の文字列の前後から、段落記号・セルの終わりの記号・空白を取り除きます。
' 引数:         strText - 選択範囲の文字列
' 戻り値:       前後を取り除いた文字列
'================================================================================
Public Function TrimSelectionText(ByVal strText As String) As String
    Dim strChar As String ' 1文字

    '末尾の段落記号・セルの終わりの記号・改行・空白（半角・全角）を取り除く
    Do While Len(strText) > 0
        strChar = Right(strText, 1)
        If strChar = vbCr Or strChar = vbLf Or strChar = Chr(7) Or strChar = " " Or strChar = "　" Then
            strText = Left(strText, Len(strText) - 1)
        Else
            Exit Do
        End If
    Loop

    '先頭の空白（半角・全角）を取り除く
    Do While Len(strText) > 0
        strChar = Left(strText, 1)
        If strChar = " " Or strChar = "　" Then
            strText = Mid(strText, 2)
        Else
            Exit Do
        End If
    Loop

    TrimSelectionText = strText
End Function

'================================================================================
' 機能:         語が登録できるかどうかを調べます。
' 引数:         strWord - 登録する語（前後の空白などを取り除いたもの）
'               strReason - 登録できない理由（戻り値。メッセージに使う）
' 戻り値:       登録できる場合は True
'================================================================================
Public Function IsRegistrableWord(ByVal strWord As String, ByRef strReason As String) As Boolean
    Dim i As Long

    IsRegistrableWord = False

    '文字数が多すぎる場合（Wordの検索・置換の文字数の上限があるため）
    If Len(strWord) > lngMaxWordLength Then
        strReason = "選んだ文字列が長すぎます（" & lngMaxWordLength & "文字まで）。"
        Exit Function
    End If

    '段落記号・タブ・改行・図などの特殊な文字を含む場合
    For i = 1 To Len(strWord)
        If CharCode(Mid(strWord, i, 1)) < 32 Then
            strReason = "1つの段落の中で、文字だけを選んでください。" & vbCrLf & "（段落の区切り・タブ・改行・図などを含む範囲は仮名化できません）"
            Exit Function
        End If
    Next i

    '記号と重なる場合（記号の一部が置き換えられて、元に戻せなくなるため）
    If OverlapsToken(strWord) Then
        strReason = "選んだ文字列は、記号（" & strTokenOpen & "人名Ａ" & strTokenClose & " など）と重なるため、仮名化できません。"
        Exit Function
    End If

    IsRegistrableWord = True
End Function

'================================================================================
' 内部の処理
'================================================================================

'================================================================================
' 機能:         語が、記号の一部になりうるかどうかを調べます。
'               例：「〔」「〕」を含む語、「人名」「名Ａ」「Ａ」のような語
' 引数:         strWord - 調べる語
' 戻り値:       記号の一部になりうる場合は True
'================================================================================
Private Function OverlapsToken(ByVal strWord As String) As Boolean
    Dim strWide As String ' 英字を全角にした語
    Dim varTypes As Variant ' 種別の一覧
    Dim i As Long
    Dim k As Long

    OverlapsToken = True
    '英字を全角にする（半角の英字の記号も考慮するため）
    strWide = ToWideLetters(strWord)

    '括弧を含む場合
    If InStr(strWide, strTokenOpen) > 0 Or InStr(strWide, strTokenClose) > 0 Then Exit Function
    '英字だけの場合
    If LettersToNumber(strWide) > 0 Then Exit Function

    'すべての種別を繰り返す
    varTypes = Split(strTokenTypes, ",")
    For i = LBound(varTypes) To UBound(varTypes)
        '種別の一部の場合（例：「人名」「名」）
        If InStr(varTypes(i), strWide) > 0 Then Exit Function
        '種別の末尾の部分＋英字の場合（例：「人名Ａ」「名Ａ」）
        For k = 1 To Len(varTypes(i))
            If Len(strWide) > k And Left(strWide, k) = Right(varTypes(i), k) Then
                If LettersToNumber(Mid(strWide, k + 1)) > 0 Then Exit Function
            End If
        Next k
    Next i

    OverlapsToken = False
End Function

'================================================================================
' 機能:         文字列が辞書の項目の文書変数の名前かどうかを調べます。
'================================================================================
Private Function IsEntryVarName(ByVal strName As String) As Boolean
    IsEntryVarName = (Left(strName, Len(strVarPrefix & strEntryName)) = strVarPrefix & strEntryName)
End Function

'================================================================================
' 機能:         番号を全角の英字に変換します（1=Ａ、26=Ｚ、27=ＡＡ）。
'================================================================================
Private Function NumberToLetters(ByVal lngNumber As Long) As String
    Dim strResult As String ' 結果

    strResult = ""
    Do While lngNumber > 0
        lngNumber = lngNumber - 1
        strResult = ChrW(lngWideA + (lngNumber Mod 26)) & strResult
        lngNumber = lngNumber \ 26
    Loop

    NumberToLetters = strResult
End Function

'================================================================================
' 機能:         全角の英字を番号に変換します（Ａ=1、Ｚ=26、ＡＡ=27）。
' 戻り値:       番号（全角の英字以外を含む場合、または空の場合は 0）
'================================================================================
Private Function LettersToNumber(ByVal strLetters As String) As Long
    Dim i As Long
    Dim lngCode As Long ' 文字コード
    Dim lngResult As Long ' 結果

    LettersToNumber = 0
    If Len(strLetters) = 0 Or Len(strLetters) > 5 Then Exit Function

    lngResult = 0
    For i = 1 To Len(strLetters)
        lngCode = CharCode(Mid(strLetters, i, 1))
        '全角の英字以外を含む場合は 0 を返す
        If lngCode < lngWideA Or lngCode > lngWideZ Then Exit Function
        lngResult = lngResult * 26 + (lngCode - lngWideA + 1)
    Next i

    LettersToNumber = lngResult
End Function

'================================================================================
' 機能:         半角の英字（大文字）を全角にした文字列を返します（例：〔人名A〕→〔人名Ａ〕）。
'================================================================================
Private Function ToWideLetters(ByVal strText As String) As String
    Dim i As Long
    Dim lngCode As Long ' 文字コード
    Dim strResult As String ' 結果

    strResult = ""
    For i = 1 To Len(strText)
        lngCode = CharCode(Mid(strText, i, 1))
        '半角の英字（大文字）の場合は全角にする
        If lngCode >= lngNarrowA And lngCode <= lngNarrowZ Then
            strResult = strResult & ChrW(lngCode - lngNarrowA + lngWideA)
        Else
            strResult = strResult & Mid(strText, i, 1)
        End If
    Next i

    ToWideLetters = strResult
End Function

'================================================================================
' 機能:         英字（半角・全角の大文字と小文字）を、全角の大文字にした文字列を返します
'               （例：〔人名A〕〔人名a〕〔人名ａ〕→〔人名Ａ〕）。記号かどうかを判定するために使います。
'               （登録できる語の判定（OverlapsToken）には使いません。「Smith」のような小文字を含む語が、
'                 英字だけの語とみなされて登録できなくなるため）
'================================================================================
Private Function ToWideCapitalLetters(ByVal strText As String) As String
    Dim i As Long
    Dim lngCode As Long ' 文字コード
    Dim strResult As String ' 結果

    strResult = ""
    For i = 1 To Len(strText)
        lngCode = CharCode(Mid(strText, i, 1))
        '半角の英字（大文字）の場合は、全角の大文字にする
        If lngCode >= lngNarrowA And lngCode <= lngNarrowZ Then
            strResult = strResult & ChrW(lngCode - lngNarrowA + lngWideA)
        '半角の英字（小文字）の場合は、全角の大文字にする
        ElseIf lngCode >= lngNarrowSmallA And lngCode <= lngNarrowSmallZ Then
            strResult = strResult & ChrW(lngCode - lngNarrowSmallA + lngWideA)
        '全角の英字（小文字）の場合は、全角の大文字にする
        ElseIf lngCode >= lngWideSmallA And lngCode <= lngWideSmallZ Then
            strResult = strResult & ChrW(lngCode - lngWideSmallA + lngWideA)
        Else
            strResult = strResult & Mid(strText, i, 1)
        End If
    Next i

    ToWideCapitalLetters = strResult
End Function

'================================================================================
' 機能:         文字の文字コード（0～65535）を返します。
'               （AscW は &H8000 以上の文字で負の値を返すため、正の値に直す）
'================================================================================
Private Function CharCode(ByVal strChar As String) As Long
    CharCode = AscW(strChar)
    If CharCode < 0 Then CharCode = CharCode + 65536
End Function
