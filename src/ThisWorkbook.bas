Attribute VB_Name = "ThisWorkbook"
Option Explicit

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim ws As Worksheet
    Dim triggerRange As Range
    Dim prevEvents As Boolean

    prevEvents = Application.EnableEvents
    On Error GoTo EH

    ' ワークシート以外の変更イベントは対象外
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    Set ws = Sh

    ' modDeepL の設定に従って対象シートを判定
    If Not IsDeepLTargetSheet(ws) Then Exit Sub

    ' 設定されたトリガーセル以外の変更は対象外
    Set triggerRange = GetDeepLTriggerRange(ws)
    If triggerRange Is Nothing Then Exit Sub
    If Intersect(Target, triggerRange) Is Nothing Then Exit Sub

    Application.EnableEvents = False

    ' 氏名元セルが数式でも最新値を取得できるよう再計算
    RecalculateDeepLSource ws

    ' 氏名をDeepLへ送り、設定された出力セルへ書き込む
    UpdateKana_FromDeepL ws

SafeExit:
    Application.EnableEvents = prevEvents
    Exit Sub

EH:
    Debug.Print "Workbook_SheetChange error: " & Err.Number & " - " & Err.Description
    Resume SafeExit
End Sub
