Option Explicit

'================================================================================
' クラス:    SearchLogic
' 機能:      汎用的な検索ロジックを提供します。
'            Settingsモジュールの定数設定に基づいて、検索の厳密さを制御します。
'================================================================================

'================================================================================
' メソッド:  ApplySearchAndHide
' 機能:      指定された範囲を検索し、ヒットしない行を非表示にします。
'================================================================================
Function ApplySearchAndHide(rngData As Range, strText As String, strMode As String) As Long
    Dim varData As Variant
    Dim r As Long, c As Long
    Dim blnFound As Boolean
    Dim strSearch As String
    Dim lngSearch As Long
    Dim strSearchSeireki As String
    Dim strSearchWareki As String
    Dim i As Long
    Dim varArray As Variant
    
    ' ヒット行数のカウンタ
    Dim hitCount As Long
    hitCount = 0
    
    ' データがない場合は終了
    If rngData Is Nothing Then
        ApplySearchAndHide = 0
        Exit Function
    End If
    
    ' 検索キーワードの処理
    strText = Replace(strText, "　", " ")
    strText = Trim(strText)
    varArray = Split(strText)
    
    ' データ全体を配列に読み込む
    If rngData.Cells.CountLarge > 10000000 Then
        ' 安全策: データ量が極端に多い場合の考慮
    End If
    
    varData = rngData.Value
    
    If Not IsArray(varData) Then
        ReDim varData(1 To 1, 1 To 1)
        varData(1, 1) = rngData.Value
    End If
    
    ' 非表示にする行範囲変数
    Dim rngToHide As Range
    Dim startHideRow As Long
    Dim endHideRow As Long
    Dim currentRowObj As Range
    
    startHideRow = 0
    
    ' 行ループ
    For r = 1 To UBound(varData, 1)
        
        ' モードによる初期値設定
        Dim isRowHit As Boolean
        If strMode = "AND" Then
            isRowHit = True
        Else
            isRowHit = False
        End If
        
        ' キーワードごとの判定
        For i = 0 To UBound(varArray)
            blnFound = False
            strSearch = varArray(i)
            If strSearch = "" Then GoTo NextKeyword
            
            ' --- 日付型変換の準備 (Settings参照) ---
            lngSearch = 0
            strSearchSeireki = ""
            strSearchWareki = ""
            
            If Settings.ENABLE_SMART_DATE_SEARCH Then
                If IsDate(strSearch) Then
                    lngSearch = Int(CDbl(CDate(strSearch))) ' シリアル値
                    strSearchSeireki = Format(DateValue(strSearch), "yyyy/m/d")
                    strSearchWareki = Format(DateValue(strSearch), "ge.m.d")
                End If
            End If
            
            ' --- 検索語の正規化 (Settings参照) ---
            Dim strSearchNormalized As String
            strSearchNormalized = NormalizeString(strSearch)
            
            ' 列ループ
            For c = 1 To UBound(varData, 2)
                Dim cellVal As Variant
                cellVal = varData(r, c)
                
                If Not IsError(cellVal) Then
                    If cellVal <> "" Then
                        Dim valStr As String
                        valStr = CStr(cellVal)
                        
                        ' 1. 文字列としての検索 (Settingsに従って正規化して比較)
                        If NormalizeString(valStr) Like "*" & strSearchNormalized & "*" Then
                            blnFound = True
                            Exit For
                        End If
                        
                        ' 2. 日付としての一致 (Settings参照)
                        If Settings.ENABLE_SMART_DATE_SEARCH Then
                            If lngSearch <> 0 Then
                                If IsDate(cellVal) Then
                                    If Int(CDbl(CDate(cellVal))) = lngSearch Then
                                        blnFound = True
                                        Exit For
                                    End If
                                End If
                            End If
                            
                            ' 西暦・和暦フォーマット文字列との部分一致救済
                            If strSearchSeireki <> "" And valStr Like "*" & strSearchSeireki & "*" Then blnFound = True: Exit For
                            If strSearchWareki <> "" And valStr Like "*" & strSearchWareki & "*" Then blnFound = True: Exit For
                        End If
                    End If
                End If
            Next c
            
            ' AND/ORによる最終判定
            If strMode = "AND" Then
                If blnFound = False Then
                    isRowHit = False
                    Exit For
                End If
            Else ' OR
                If blnFound = True Then
                    isRowHit = True
                    Exit For
                End If
            End If
            
NextKeyword:
        Next i
        
        ' --- 非表示範囲の結合ロジック ---
        If isRowHit Then
            hitCount = hitCount + 1
            If startHideRow > 0 Then
                endHideRow = r - 1
                Set currentRowObj = rngData.Rows(startHideRow & ":" & endHideRow)
                If rngToHide Is Nothing Then
                    Set rngToHide = currentRowObj
                Else
                    Set rngToHide = Union(rngToHide, currentRowObj)
                End If
                startHideRow = 0
            End If
        Else
            If startHideRow = 0 Then
                startHideRow = r
            End If
        End If
        
    Next r
    
    ' ループ終了後の処理
    If startHideRow > 0 Then
        endHideRow = UBound(varData, 1)
        Set currentRowObj = rngData.Rows(startHideRow & ":" & endHideRow)
        If rngToHide Is Nothing Then
            Set rngToHide = currentRowObj
        Else
            Set rngToHide = Union(rngToHide, currentRowObj)
        End If
    End If
    
    ' 一括非表示実行
    If hitCount > 0 Then
        If Not rngToHide Is Nothing Then
            rngToHide.EntireRow.Hidden = True
        End If
    End If
    
    ApplySearchAndHide = hitCount
End Function

'================================================================================
' 関数:      NormalizeString
' 機能:      Settingsの設定に従って文字列を正規化します。
' 概要:      比較のために文字列を「最も単純な形」に変換します。
'================================================================================
Private Function NormalizeString(ByVal strInput As String) As String
    Dim s As String
    s = strInput
    
    ' --- 1. 異字体を区別しない設定の場合 (Settings参照) ---
    If Settings.ENABLE_IGNORE_VARIANT_CHAR Then
        ' (A) IVS (Ideographic Variation Sequence) の除去
        If InStr(s, ChrW(&HDB40)) > 0 Then
            Dim i As Long
            Dim newStr As String
            Dim c1 As String, c2 As String
            newStr = ""
            i = 1
            Do While i <= Len(s)
                c1 = Mid(s, i, 1)
                If i < Len(s) Then
                    c2 = Mid(s, i + 1, 1)
                    If c1 = ChrW(&HDB40) And (c2 >= ChrW(&HDD00) And c2 <= ChrW(&HDDEF)) Then
                        i = i + 2 ' IVSなのでスキップ
                    Else
                        newStr = newStr & c1
                        i = i + 1
                    End If
                Else
                    newStr = newStr & c1
                    i = i + 1
                End If
            Loop
            s = newStr
        End If
        
        ' (B) 従来の異字体の置換
        ' ア行
        s = Replace(s, "亞", "亜")
        s = Replace(s, "醫", "医")
        s = Replace(s, "榮", "栄")
        s = Replace(s, "惠", "恵")
        ' カ行
        s = Replace(s, "會", "会")
        s = Replace(s, "學", "学")
        s = Replace(s, "國", "国")
        s = Replace(s, "劍", "剣")
        s = Replace(s, "廣", "広")
        s = Replace(s, "舊", "旧")
        ' 上が「土」の「吉」（U+20BB7）は、VBEで書けない文字のため、サロゲートペアで指定する
        s = Replace(s, ChrW(&HD842) & ChrW(&HDFB7), "吉")
        ' サ行
        s = Replace(s, "齋", "斉")
        s = Replace(s, "齊", "斉")
        s = Replace(s, "斎", "斉")
        s = Replace(s, "櫻", "桜")
        s = Replace(s, "澤", "沢")
        s = Replace(s, "嶋", "島")
        s = Replace(s, "嶌", "島")
        s = Replace(s, "壽", "寿")
        s = Replace(s, "眞", "真")
        s = Replace(s, "穣", "穣")
        s = Replace(s, "穰", "穣")
        ' タ行
        s = Replace(s, "髙", "高")
        s = Replace(s, "瀧", "滝")
        s = Replace(s, "塚", "塚")
        s = Replace(s, "鐵", "鉄")
        s = Replace(s, "德", "徳")
        s = Replace(s, "冨", "富")
        s = Replace(s, "豐", "豊")
        ' ハ行
        s = Replace(s, "濱", "浜")
        s = Replace(s, "濵", "浜")
        s = Replace(s, "邉", "辺")
        s = Replace(s, "邊", "辺")
        s = Replace(s, "檜", "桧")
        s = Replace(s, "淵", "渕")
        s = Replace(s, "渕", "淵")
        ' マ行
        s = Replace(s, "萬", "万")
        s = Replace(s, "峯", "峰")
        s = Replace(s, "楳", "梅")
        ' ヤ行
        s = Replace(s, "藪", "薮")
        ' ラ行
        s = Replace(s, "龍", "竜")
        s = Replace(s, "隆", "隆")
        s = Replace(s, "禮", "礼")
        ' その他
        s = Replace(s, "﨑", "崎")
        s = Replace(s, "區", "区")
    End If
    
    ' --- 2. カナの統一と半角化（正規化のベース処理） ---
    ' 以下のいずれかの日本語正規化オプションが有効な場合、
    ' 一旦「全角カタカナ」→「半角」に変換することで、
    ' 濁点分離や小書き文字の処理を確実に実行できるようにします。
    If Settings.ENABLE_IGNORE_KANA_TYPE Or _
       Settings.ENABLE_IGNORE_DAKUTEN Or _
       Settings.ENABLE_IGNORE_SMALL_KANA Or _
       Settings.ENABLE_IGNORE_CHOON Or _
       Settings.ENABLE_IGNORE_WIDTH_CASE Then
       
        ' (1) カナ種統一が有効、または濁点等の処理が必要な場合はカタカナに寄せる
        If Settings.ENABLE_IGNORE_KANA_TYPE Or _
           Settings.ENABLE_IGNORE_DAKUTEN Or _
           Settings.ENABLE_IGNORE_SMALL_KANA Or _
           Settings.ENABLE_IGNORE_CHOON Then
            s = StrConv(s, vbKatakana)
        End If
        
        ' (2) 全て半角に変換（濁点分離のため必須）
        s = StrConv(s, vbNarrow)
        
        ' (3) 大文字小文字の区別なし設定なら大文字化
        If Settings.ENABLE_IGNORE_WIDTH_CASE Then
            s = UCase(s)
        End If
    End If

    ' --- 3. 濁点・半濁点の除去 ---
    If Settings.ENABLE_IGNORE_DAKUTEN Then
        ' 半角カタカナの濁点(ﾞ)・半濁点(ﾟ)を除去
        s = Replace(s, Chr(222), "")
        s = Replace(s, Chr(223), "")
    End If
    
    ' --- 4. 小書き文字を区別しない ---
    If Settings.ENABLE_IGNORE_SMALL_KANA Then
        ' 半角カタカナの小書きを大文字に置換
        s = Replace(s, "ｧ", "ｱ")
        s = Replace(s, "ｨ", "ｲ")
        s = Replace(s, "ｩ", "ｳ")
        s = Replace(s, "ｪ", "ｴ")
        s = Replace(s, "ｫ", "ｵ")
        s = Replace(s, "ｯ", "ﾂ")
        s = Replace(s, "ｬ", "ﾔ")
        s = Replace(s, "ｭ", "ﾕ")
        s = Replace(s, "ｮ", "ﾖ")
        ' 念のためワ行
        s = Replace(s, "ヮ", "ﾜ")
    End If

    ' --- 5. 長音（ー）を区別しない ---
    If Settings.ENABLE_IGNORE_CHOON Then
        ' 半角化済みなので、半角ハイフンと半角カナの長音を除去
        s = Replace(s, "-", "")       ' 半角マイナス/ハイフン
        s = Replace(s, Chr(176), "")  ' 半角カナの長音 (ｰ)
        s = Replace(s, "ー", "")      ' 全角長音 (念のため)
    End If
    
    NormalizeString = s
End Function
