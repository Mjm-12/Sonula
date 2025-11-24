Attribute VB_Name = "Module4"
Option Explicit

Public Sub DuplicateLastSheetFive()
    Dim src As Object ' Worksheet または Chart を許容
    Dim i As Long
    Dim wasScreenUpdating As Boolean, wasEnableEvents As Boolean, wasDisplayAlerts As Boolean
    
    If ThisWorkbook.Sheets.Count = 0 Then
        MsgBox "シートが存在しません。", vbExclamation
        Exit Sub
    End If
    
    Set src = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count) ' 最後尾のシートを取得
    
    ' 動作の安定化
    wasScreenUpdating = Application.ScreenUpdating
    wasEnableEvents = Application.EnableEvents
    wasDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    On Error GoTo ERR_HANDLER
    
    ' 5枚複製して最後尾へ
    For i = 1 To 5
        src.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ' ここでExcelが自動的にシート名を付けます（例：「Sheet1 (2)」など）
        ' 独自の命名をしたい場合は、以下の例を参考に:
        ' ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = src.Name & "_copy" & i
    Next i
    
    GoTo CLEANUP

ERR_HANDLER:
    MsgBox "複製中にエラーが発生しました: " & Err.Description, vbCritical

CLEANUP:
    Application.DisplayAlerts = wasDisplayAlerts
    Application.EnableEvents = wasEnableEvents
    Application.ScreenUpdating = wasScreenUpdating
End Sub

