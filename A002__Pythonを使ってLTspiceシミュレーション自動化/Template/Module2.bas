Attribute VB_Name = "Module2"
Option Explicit

Public Sub ExportAllChartsAsPng()
    Dim basePath As String
    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        MsgBox "このブックを一度保存してください（保存先フォルダが取得できません）。", vbExclamation
        Exit Sub
    End If
    
    ' フォルダ名：images_MM_dd_hh-mm-ss（:は使えないため-に置換）
    Dim folderName As String
    folderName = "images_" & Format(Now, "MM_dd_hh-mm-ss")
    Dim outDir As String
    outDir = basePath & Application.PathSeparator & folderName
    
    On Error Resume Next
    MkDir outDir
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "出力フォルダの作成に失敗しました：" & vbCrLf & outDir, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' ?? ここで全シートの表示倍率を150%に揃える ??
    Call SetAllSheetsZoom(150)

    Dim savedCount As Long
    savedCount = 0
    
    ' ワークシート上の埋め込みグラフ（ChartObject）
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.ChartObjects.Count > 0 Then
            Dim chObj As ChartObject
            Set chObj = ws.ChartObjects(1)
            If Not chObj Is Nothing Then
                Dim f1 As String
                f1 = outDir & Application.PathSeparator & "image_" & SanitizeForFileName(ws.Name) & ".png"
                On Error Resume Next
                chObj.chart.Export Filename:=f1, FilterName:="PNG"
                If Err.Number = 0 Then
                    savedCount = savedCount + 1
                Else
                    MsgBox "保存に失敗: " & f1, vbExclamation
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next ws
    
    ' グラフシート（Chartsコレクション）
    Dim cs As chart
    For Each cs In ThisWorkbook.Charts
        Dim f2 As String
        f2 = outDir & Application.PathSeparator & "image_" & SanitizeForFileName(cs.Name) & ".png"
        On Error Resume Next
        cs.Export Filename:=f2, FilterName:="PNG"
        If Err.Number = 0 Then
            savedCount = savedCount + 1
        Else
            MsgBox "保存に失敗: " & f2, vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    Next cs

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "完了しました。" & vbCrLf & _
           "保存先: " & outDir & vbCrLf & _
           "保存枚数: " & savedCount, vbInformation
End Sub

Private Sub SetAllSheetsZoom(ByVal pct As Long)
    Dim curSheet As Object
    Set curSheet = ActiveSheet

    ' ワークシート
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ActiveWindow.Zoom = pct
        End If
    Next ws
    
    ' グラフシート
    Dim cs As chart
    For Each cs In ThisWorkbook.Charts
        If cs.Visible = xlSheetVisible Then
            cs.Activate
            ActiveWindow.Zoom = pct
        End If
    Next cs

    ' 元のシートに戻す
    If Not curSheet Is Nothing Then
        On Error Resume Next
        curSheet.Activate
        On Error GoTo 0
    End If
End Sub

Private Function SanitizeForFileName(ByVal s As String) As String
    Dim badChars As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, CStr(badChars(i)), "_")
    Next i
    s = Trim$(s)
    Do While Len(s) > 0 And Right$(s, 1) = "."
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "Sheet"
    SanitizeForFileName = s
End Function


