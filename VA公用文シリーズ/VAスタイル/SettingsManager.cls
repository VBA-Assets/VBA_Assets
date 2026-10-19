'================================================================================
' クラスモジュール: SettingsManager
' 機能: 設定ファイルの読み書きを管理する
'================================================================================
Option Explicit

' このアドインが管理するプロパティであることを示すためのプレフィックス
Private Const c_Prefix As String = "VA_"

'==============================================================================
' 機能: キーと値の配列の内容を、アドインファイル（ThisDocument）の
'       カスタムドキュメントプロパティとして保存します。
'       保存後、アドインファイルを明示的に保存します。
'------------------------------------------------------------------------------
' 引数:
'   ByRef varKeys() As Variant: 設定のキー名（プロパティ名）を含む配列。
'   ByRef varValues() As Variant: 設定の値（プロパティ値）を含む配列。
' 注意:
'   varKeys と varValues の LBound, UBound は一致している必要があります。
'==============================================================================
Public Sub SaveSettingsToProperty(ByRef varKeys() As Variant, ByRef varValues() As Variant)
    Dim i As Long
    Dim doc As Document
    Dim keyName As String
    Dim prefixedKeyName As String
    Dim propType As Long
    Dim prop As DocumentProperty
    
    ' プロシージャを実行しているアドインファイル自身を参照
    Set doc = ThisDocument
    
    ' 配列の要素数が一致しない場合はエラー (LBoundもチェック)
    If LBound(varKeys) <> LBound(varValues) Or UBound(varKeys) <> UBound(varValues) Then
        MsgBox "キーと値の配列のインデックス範囲が一致しません。", vbCritical
        GoTo CleanExit
    End If
    
    On Error GoTo CleanExit
    
    For i = LBound(varKeys) To UBound(varKeys)
        keyName = varKeys(i)
        prefixedKeyName = c_Prefix & keyName ' プレフィックスを付加
        
        ' 値の型に応じてプロパティの型を決定
        Select Case VarType(varValues(i))
            Case vbInteger, vbLong
                propType = msoPropertyTypeNumber
            Case vbSingle, vbDouble, vbCurrency
                propType = msoPropertyTypeFloat
            Case vbDate
                propType = msoPropertyTypeDate
            Case vbBoolean
                propType = msoPropertyTypeBoolean
            Case Else ' vbString, vbEmpty, vbNull, etc.
                propType = msoPropertyTypeString
        End Select
        
        ' 既存のプロパティを検索
        On Error Resume Next
        Set prop = doc.CustomDocumentProperties(prefixedKeyName)
        
        If Err.Number <> 0 Then
            ' プロパティが存在しない場合：新規作成
            Err.Clear
        Else
            ' プロパティが存在する場合：更新
            ' 型の不一致エラーを避けるため、一度削除する
            prop.Delete
        End If
        
        ' エラーハンドリングを明示的にリセット
        On Error GoTo 0
        
        ' プロパティを（再）追加
        doc.CustomDocumentProperties.Add _
            Name:=prefixedKeyName, _
            LinkToContent:=False, _
            Type:=propType, _
            value:=varValues(i)
            
        On Error GoTo CleanExit ' エラー処理を元に戻す
    Next i
    
    ' Wordにファイルに変更があったことを明示的に通知する
    doc.Saved = False
    
    ' 変更を保存するために、アドインファイルを明示的に保存する
    doc.Save
    
CleanExit:
    ' オブジェクトを解放
    Set doc = Nothing
    Set prop = Nothing
End Sub

'==============================================================================
' 機能: アドインファイル（ThisDocument）のカスタムドキュメントプロパティから
'       このモジュールが管理する設定値（プレフィックス付き）のみを読み込み、
'       キーと値の配列に格納します。
'------------------------------------------------------------------------------
' 引数:
'   ByRef varKeys() As Variant: 読み込んだ設定のキー名（プレフィックス除去後）を格納する配列。
'   ByRef varValues() As Variant: 読み込んだ設定の値（プロパティ値）を格納する配列。
'==============================================================================
Public Sub LoadSettingsFromProperty(ByRef varKeys() As Variant, ByRef varValues() As Variant)
    Dim doc As Document
    Dim prop As DocumentProperty
    Dim tempKeys() As Variant
    Dim tempValues() As Variant
    Dim count As Long
    Dim i As Long
    Dim prefixLen As Long
    
    ' プロシージャを実行しているアドインファイル自身を参照
    Set doc = ThisDocument
    
    ' 配列を初期化
    ReDim varKeys(0)
    ReDim varValues(0)
    
    count = 0
    prefixLen = Len(c_Prefix)
    
    On Error GoTo CleanExit
    
    ' プレフィックスに一致するプロパティの数を先に数える
    For Each prop In doc.CustomDocumentProperties
        If Left(prop.Name, prefixLen) = c_Prefix Then
            count = count + 1
        End If
    Next prop
    
    ' データがない場合は空の配列のまま返す
    If count = 0 Then
        ReDim varKeys(0)
        ReDim varValues(0)
        GoTo CleanExit
    End If
    
    ' 配列のサイズを一度だけ確保 (1ベース配列)
    ReDim tempKeys(1 To count)
    ReDim tempValues(1 To count)
    
    i = 1
    ' 再度ループして、該当するキーと値を配列に格納
    For Each prop In doc.CustomDocumentProperties
        If Left(prop.Name, prefixLen) = c_Prefix Then
            ' プレフィックスを除去したキー名
            tempKeys(i) = Mid(prop.Name, prefixLen + 1)
            tempValues(i) = prop.value
            i = i + 1
            If i > count Then Exit For ' 念のため (count分取得したら抜ける)
        End If
    Next prop
    
    ' 結果を引数に代入
    varKeys = tempKeys
    varValues = tempValues
    
CleanExit:
    ' オブジェクトを解放
    Set doc = Nothing
    Set prop = Nothing
    Erase tempKeys
    Erase tempValues
End Sub