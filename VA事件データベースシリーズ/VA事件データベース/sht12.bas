Option Explicit

'================================================================================
' モジュール名: sht12 (事件メモ)
' 概要:         事件メモシートの操作を制御するクラスモジュールです。
'               選択中の事件データをメモテンプレートに反映し、別ブックとして
'               保存する機能などを提供します。
' 更新日時:     2026/02/01 05:37
'================================================================================

'================================================================================
' イベント:     Worksheet_Activate
' 機能:         事件メモの作成開始
' 概要:         シートアクティブ時に、選択中の事件番号に基づいてメモを作成するか
'               ユーザーに確認し、データをセットします。
'================================================================================
Private Sub Worksheet_Activate()
    Dim r As Long
    Evnt.Disable
    
    '事件番号が選択されていなければ中止
    If lngActCase = 0 Then
        MsgBox "事件番号が選択されていないため、事件メモは作成できません。", vbInformation, "事件データベース"
        GoTo Cleanup
    End If

    '選択された事件番号のデータを事件シートから探す
    On Error Resume Next
    r = WorksheetFunction.Match(lngActCase, sht11.Range("事件[事件番号]"), 0)
    On Error GoTo 0
    
    If r = 0 Then GoTo Cleanup

    '作成確認
    If MsgBox(sht11.Range("事件[事件呼名]")(r) & "の事件メモを作成しますか？", vbYesNo, "事件データベース") = vbNo Then
        GoTo Cleanup
    End If
    
    '事件番号をセットしてセルを選択
    Evnt.Enable
    Range("事件メモ事件番号").Value = lngActCase
    Range("事件メモ事件番号").Select
    Evnt.Disable

    '保存処理を呼び出し
    Call SheetSave

Cleanup:
    Evnt.Enable
End Sub

'================================================================================
' プロシージャ: SheetSave
' 機能:         事件メモのファイル保存
' 概要:         作成された事件メモを新規ブックにコピーし、値貼り付けを行って
'               当該事件のデータフォルダに保存します。
'================================================================================
Private Sub SheetSave()
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim savePath As String
    Dim fileName As String

    '保存確認
    If MsgBox("事件メモを別ファイルに保存しますか？" & vbNewLine & _
              "作成されたファイルは、当該事件のデータフォルダに保存されます。", vbYesNo, "事件データベース") = vbNo Then
        Exit Sub
    End If
    
    '現在のシートを新規ブックへコピー
    Me.Copy
    Set wbNew = ActiveWorkbook
    Set wsNew = wbNew.Worksheets(1)
    
    '数式を値に変換（データの固定化）
    With wsNew.UsedRange
        .Value = .Value
    End With
    
    '保存先パスの構築
    savePath = Me.Range("事件メモデータフォルダ").Value
    fileName = "事件メモ" & Format(Date, "yyyymmdd") & ".xlsx"
    
    '保存処理
    On Error GoTo ErrHandler
    wbNew.SaveAs fileName:=savePath & Application.PathSeparator & fileName, _
                   FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    '保存して閉じる
    wbNew.Close SaveChanges:=False
    
    '元のシートに戻る
    Me.Activate
    MsgBox "保存が完了しました。", vbInformation, "事件データベース"
    Exit Sub

ErrHandler:
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    Me.Activate
    MsgBox "ファイルの保存に失敗しました。" & vbCr & "パスを確認してください: " & savePath, vbExclamation, "事件データベース"
End Sub

