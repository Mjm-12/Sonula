Attribute VB_Name = "Module3"
Option Explicit

Public Sub A1_1x()
    Dim curSheet As Object
    Set curSheet = ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo CleanExit

    ' --- ワークシート：ズーム=100%、A1をアクティブにしてスクロール ---
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            On Error Resume Next
            ActiveWindow.Zoom = 100
            On Error GoTo 0
            Application.Goto Reference:=ws.Range("A1"), Scroll:=True
        End If
    Next ws

    ' --- グラフシート：ズーム=100%（A1は存在しないため未設定） ---
    Dim cs As chart
    For Each cs In ThisWorkbook.Charts
        If cs.Visible = xlSheetVisible Then
            cs.Activate
            On Error Resume Next
            ActiveWindow.Zoom = 100
            On Error GoTo 0
        End If
    Next cs

CleanExit:
    ' 元のシートに戻る
    If Not curSheet Is Nothing Then
        On Error Resume Next
        curSheet.Activate
        On Error GoTo 0
    End If

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

