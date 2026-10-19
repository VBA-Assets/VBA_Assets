Option Explicit

' モジュール概要: CustomDocumentPropertiesを使用した設定情報の永続化およびモード管理
' Last Updated: 2026/02/01 11:15

' -----------------------------------------------------------------------------
' 設定値の取得 (Book単位)
' 指定されたブック（省略時はアクティブブック）のカスタムプロパティから値を取得する
' -----------------------------------------------------------------------------
Public Function GetBookSetting(ByVal propName As String, Optional ByVal TargetBook As Workbook = Nothing) As String
    ' 対象ブックが指定されていない場合は、アクティブブックを対象とする
    If TargetBook Is Nothing Then Set TargetBook = ActiveWorkbook
    
    ' プロパティが存在しない場合のエラーを無視して継続する
    On Error Resume Next
    
    ' 指定された名前のカスタムプロパティの値を取得して戻り値に設定
    GetBookSetting = TargetBook.CustomDocumentProperties(propName).value
    
    ' エラー処理を通常の動作に戻す
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------
' 設定値の保存 (Book単位)
' 指定されたブックのカスタムプロパティに値を保存する。存在しない場合は新規作成する。
' -----------------------------------------------------------------------------
Public Sub SaveBookSetting(ByVal propName As String, ByVal propValue As String, Optional ByVal TargetBook As Workbook = Nothing)
    ' 対象ブックが指定されていない場合は、アクティブブックを対象とする
    If TargetBook Is Nothing Then Set TargetBook = ActiveWorkbook
    
    ' カスタムプロパティコレクションの参照を取得
    Dim props As Object: Set props = TargetBook.CustomDocumentProperties
    
    ' プロパティが存在しない場合のエラーを捕捉するため、エラートラップを一時的に無効化
    On Error Resume Next
    
    ' 既存のプロパティへの値の書き込みを試みる
    props(propName).value = propValue
    
    ' エラーが発生した場合（プロパティが存在しない場合）
    If Err.Number <> 0 Then
        ' エラー情報をクリア
        Err.Clear
        ' 新しいプロパティとして追加する
        ' Type:=4 は msoPropertyTypeString（文字列型）を指す
        props.Add Name:=propName, LinkToContent:=False, Type:=4, value:=propValue
    End If
    
    ' エラー処理を元に戻す
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' 設定値の取得 (Sheet単位)
' シート名とサフィックスを組み合わせてキーを作成し、ブックの設定として取得する
' -----------------------------------------------------------------------------
Public Function GetSheetSetting(ByVal sh As Worksheet, ByVal suffix As String) As String
    ' シートオブジェクトが無効な場合は処理を抜ける
    If sh Is Nothing Then Exit Function
    
    ' シート名 + 接尾辞 をキーとして、ブックから設定値を取得する
    GetSheetSetting = GetBookSetting(sh.Name & suffix, sh.Parent)
End Function

' -----------------------------------------------------------------------------
' 設定値の保存 (Sheet単位)
' シート名とサフィックスを組み合わせてキーを作成し、ブックの設定として保存する
' -----------------------------------------------------------------------------
Public Sub SaveSheetSetting(ByVal sh As Worksheet, ByVal suffix As String, ByVal value As String)
    ' シートオブジェクトが無効な場合は処理を抜ける
    If sh Is Nothing Then Exit Sub
    
    ' シート名 + 接尾辞 をキーとして、ブックに設定値を保存する
    SaveBookSetting sh.Name & suffix, value, sh.Parent
End Sub

' -----------------------------------------------------------------------------
' 設定の削除 (Sheet単位および一括クリーンアップ)
' 指定されたシートに関する設定を削除し、不要になった共有設定もクリーンアップする
' -----------------------------------------------------------------------------
Public Sub DeleteSheetSetting(ByVal sh As Worksheet, ByVal suffix As String)
    ' シートオブジェクトが無効な場合は処理を抜ける
    If sh Is Nothing Then Exit Sub
    
    ' エラーが発生しても処理を継続する（存在しないプロパティ削除時のエラー無視など）
    On Error Resume Next
    
    ' 指定されたシート固有のプロパティ紐付け情報: MasterSheetNameを削除
    sh.Parent.CustomDocumentProperties(sh.Name & suffix).Delete
    
    ' ★追加修正: 編集モード用のスタイル情報（OriginalStyle）も同時に削除する
    sh.Parent.CustomDocumentProperties(sh.Name & cstSettings.PROP_SUFFIX_STYLE).Delete
    
    ' 最終同期日時（LastSyncTime）もクリアする
    ' ※設定変更によって同期状態がリセットされるべきため
    sh.Parent.CustomDocumentProperties(cstSettings.PROP_LAST_SYNC).Delete
    
    ' 他に管理対象となっているシートがブック内に残っているかチェックする
    Dim otherSh As Worksheet, hasOther As Boolean: hasOther = False
    For Each otherSh In sh.Parent.Worksheets
        ' 自分以外のシートについて確認
        If otherSh.Name <> sh.Name Then
            ' 管理用の設定値（マスタシート名）が存在するか確認
            If GetSheetSetting(otherSh, cstSettings.PROP_SUFFIX_SHEET) <> "" Then
                ' 他に管理対象が見つかった場合はフラグを立ててループを抜ける
                hasOther = True: Exit For
            End If
        End If
    Next otherSh
    
    ' もし他に管理対象シートが一つもなければ、ブック全体の共有設定（マスタパス）も削除してクリーンにする
    If Not hasOther Then
        sh.Parent.CustomDocumentProperties(cstSettings.PROP_MASTER_PATH).Delete
    End If
    
    ' エラー処理を元に戻す
    On Error GoTo 0
End Sub