'//==================================================================================================
'// 機能: Word文書のスタイルを設定します。
'//
'// 概要: このモジュールには、見出しスタイルを順次に変更するためのプロシージャが含まれています。
'//==================================================================================================
Option Explicit

' 設定値を管理するためのクラスインスタンスをモジュールレベルで宣言
Dim Variables As VariablesManager

'//==================================================================================================
'// プロシージャ: ConfigStyle
'// 引数:        lngStep (Long) - スタイルを変更する方向 (1: 進む, -1: 戻る)
'// 機能:        文書全体の表題と見出しスタイルを段階的に変更します。
'//==================================================================================================
Sub ConfigStyle(lngStep As Long)
    ' ★修正点: エラー発生時はCleanupへ飛び、メッセージは出さない
    On Error GoTo Cleanup
    
    ' 画面描画を停止して高速化
    Application.ScreenUpdating = False
    
    '変数オブジェクトを生成する
    Set Variables = New VariablesManager
    
    'マウスポインタをウエイト表示にする
    Application.System.Cursor = wdCursorWait
    
    Dim doc As Document
    Set doc = ActiveDocument

    '--- 文書内の最初の見出しレベルを把握する ---
    Dim xFirst As Long          ' 最初の見出しレベル
    Dim blnTitle As Boolean     ' 表題の存在
    Dim para As Paragraph
    Dim targetPara As Paragraph ' 番号更新用の基準段落
    Dim strStyleName As String
    Dim foundHeader As Boolean
    
    xFirst = 999 ' 見つからなかった場合の初期値
    
    For Each para In doc.Paragraphs
        strStyleName = para.Style
        If strStyleName = "表題" Then
            blnTitle = True
        ' 見出しスタイルかどうかを判定
        ElseIf Left(strStyleName, 3) = "見出し" Then
            xFirst = para.OutlineLevel
            Set targetPara = para
            foundHeader = True
            Exit For
        End If
    Next para

    ' 見出し段落が見つからなかった場合、文書末尾を基準にする
    If Not foundHeader Then
        If doc.Paragraphs.count > 0 Then
            Set targetPara = doc.Paragraphs(doc.Paragraphs.count)
        End If
        ' 表題もなければ終了
        If Not blnTitle Then Exit Sub
    End If

    '--- 設定値のキャッシュ ---
    Dim bChangeFont As Boolean: bChangeFont = Variables.blnChangeFontType
    Dim bChangeSize As Boolean: bChangeSize = Variables.blnChangeFontSize
    Dim bLineUnit As Boolean:   bLineUnit = Variables.blnLineUnitBeforeOne
    Dim bKeepNext As Boolean:   bKeepNext = Variables.blnKeepParagraphWithNext
    Dim bKeepTog As Boolean:    bKeepTog = Variables.blnKeepParagraphTogether

    ' 行間隔を把握する
    Dim sngLineSpacing As Single
    With doc.PageSetup
        sngLineSpacing = (.PageHeight - .TopMargin - .BottomMargin) / .LinesPage
    End With
    
    '--- スタイル変更ループの開始・終了レベルを設定 ---
    Dim lngStart As Long, lngEnd As Long
    
    '--- 適用するスタイル設定値の準備 ---
    Dim strTargetFont As String, strCheckFont As String
    Dim sngTargetSize As Single, sngCheckSize As Single
    Dim blnDisableGrid As Boolean
    Dim sngTargetSpace As Single
    Dim blnTargetKeepNext As Boolean, blnCheckKeepNext As Boolean
    Dim blnTargetKeepTog As Boolean, blnCheckKeepTog As Boolean
    
    If lngStep = 1 Then
        '【進む】設定値
        lngStart = IIf(blnTitle, 0, 1)
        lngEnd = 9
        
        ' フォント: 標準 → 太字
        strTargetFont = strBoldFont
        strCheckFont = strFont
        
        ' サイズ: 標準 → 太字サイズ
        sngTargetSize = sngBoldFont
        sngCheckSize = sngFont
        blnDisableGrid = True
        
        ' 段落前: 1行空ける
        sngTargetSpace = sngLineSpacing
        
        ' 改ページ制御: Trueにする
        blnTargetKeepNext = True: blnCheckKeepNext = False
        blnTargetKeepTog = True: blnCheckKeepTog = False
    Else
        '【戻る】設定値
        lngStart = 9
        lngEnd = IIf(blnTitle, 0, 1)
        
        ' フォント: 太字 → 標準
        strTargetFont = strFont
        strCheckFont = strBoldFont
        
        ' サイズ: 太字サイズ → 標準
        sngTargetSize = sngFont
        sngCheckSize = sngBoldFont
        blnDisableGrid = False
        
        ' 段落前: 0に戻す
        sngTargetSpace = 0
        
        ' 改ページ制御: Falseに戻す
        blnTargetKeepNext = False: blnCheckKeepNext = True
        blnTargetKeepTog = False: blnCheckKeepTog = True
    End If


    '--- 表題およびすべての見出しスタイルをループ処理 ---
    Dim X As Long
    Dim blnStop As Boolean  '処理停止フラグ
    Dim currentStyle As Style
    Dim bSpaceCondition As Boolean
    
    For X = lngStart To lngEnd Step lngStep
        '見つかった最初の見出しレベルより上位のレベルはスキップ
        If X <> 0 And X < xFirst Then GoTo Continue
        
        '処理対象のスタイルオブジェクトを取得
        Set currentStyle = doc.Styles(IIf(X = 0, "表題", "見出し " & X))
            
        With currentStyle
            '--- フォントの変更 ---
            With .Font
                If .Name = strCheckFont Then
                    If bChangeFont = True Then
                        .NameFarEast = strTargetFont
                        .NameAscii = strTargetFont
                        .NameOther = strTargetFont
                        .Name = strTargetFont
                        If X = 0 Or X >= xFirst Then blnStop = True
                    End If
                End If
            End With
            
            '見出し番号のフォントも同期 (X=0以外)
            If X <> 0 Then
                With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(X).Font
                    If .Name = strCheckFont And bChangeFont = True Then
                        .Name = strTargetFont
                        If X >= xFirst Then blnStop = True
                    End If
                End With
            End If
        
            '--- フォントサイズの変更 ---
            With .Font
                If .Size = sngCheckSize Then
                    If bChangeSize = True Then
                        .Size = sngTargetSize
                        ' グリッド線の設定
                        currentStyle.ParagraphFormat.DisableLineHeightGrid = blnDisableGrid
                        If X = 0 Or X >= xFirst Then blnStop = True
                    End If
                End If
            End With
            
            '見出し番号のサイズも同期
            If X <> 0 Then
                With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(X).Font
                    If .Size = sngCheckSize And bChangeSize = True Then
                        .Size = sngTargetSize
                        If X >= xFirst Then blnStop = True
                    End If
                End With
            End If
             
            '--- 段落書式の設定 ---
            With .ParagraphFormat
                '--- 「段落前の間隔」を設定/解除 ---
                bSpaceCondition = False
                If lngStep = 1 And .SpaceBefore = 0 Then bSpaceCondition = True
                If lngStep = -1 And .SpaceBefore <> 0 Then bSpaceCondition = True
                
                If bSpaceCondition = True And bLineUnit = True And X <> 0 Then '表題は処理しない
                    .LineUnitBefore = 0
                    .SpaceBefore = sngTargetSpace
                    If X = 0 Or X >= xFirst Then blnStop = True
                End If
                
                '--- 「次の段落と分離しない」を設定/解除 ---
                If .KeepWithNext = blnCheckKeepNext Then
                    If bKeepNext = True And X <> 0 Then '表題は処理しない
                        .KeepWithNext = blnTargetKeepNext
                        If X = 0 Or X >= xFirst Then blnStop = True
                    End If
                End If
                
                '--- 「段落を分割しない」を設定/解除 ---
                If .KeepTogether = blnCheckKeepTog Then
                    If bKeepTog = True And X <> 0 Then '表題は処理しない
                        .KeepTogether = blnTargetKeepTog
                        If X = 0 Or X >= xFirst Then blnStop = True
                    End If
                End If
            End With
        End With
        
        'いずれかの書式が変更されたらループを抜ける
        If blnStop = True Then Exit For
Continue:
    Next X
        
    '見出し段落がない場合、かつ表題の処理も終わっていれば終了
    If targetPara Is Nothing Then Exit Sub

    '--- 見出し番号のフォントとサイズを、対応する見出しスタイルに同期させる ---
    Dim lstTmpl As ListTemplate
    Set lstTmpl = ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
    
    For X = 1 To 9
        With lstTmpl.ListLevels(X).Font
            .Name = doc.Styles("見出し " & X).Font.Name
            .Size = doc.Styles("見出し " & X).Font.Size
        End With
    Next X

    '--- アウトライン書式を再適用して番号を更新 ---
    With targetPara
        Dim dblFirstLineIndent As Double
        Dim dblLeftIndent As Double
        dblFirstLineIndent = .FirstLineIndent
        dblLeftIndent = .LeftIndent
        
        .Range.ListFormat.ApplyListTemplateWithLevel _
            ListTemplate:=lstTmpl, _
            ContinuePreviousList:=False, _
            ApplyTo:=wdListApplyToWholeList, _
            DefaultListBehavior:=wdWord10ListBehavior

        .FirstLineIndent = dblFirstLineIndent
        .LeftIndent = dblLeftIndent
    End With
    
Cleanup:
    '変数オブジェクトを解放する
    Set Variables = Nothing
    'マウスポインタと画面描画を戻す
    Application.System.Cursor = wdCursorNormal
    Application.ScreenUpdating = True
    
End Sub