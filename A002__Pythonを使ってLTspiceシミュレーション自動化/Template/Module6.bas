Attribute VB_Name = "Module6"
Option Explicit

'========================
' 一括処理マクロ
' ImportCsvsIntoTemplateCopies → ExportAllChartsAsPng → A1_1x
'========================
Public Sub BatchProcess_Import_Export_Reset()
    Dim startTime As Double
    Dim errorOccurred As Boolean
    Dim errorMsg As String

    startTime = Timer
    errorOccurred = False
    errorMsg = ""

    On Error GoTo ErrorHandler

    ' ステップ1: CSVファイルをTemplateシートのコピーにインポート
    MsgBox "ステップ1/3: CSVファイルをインポートします...", vbInformation, "一括処理"
    Call ImportCsvsIntoTemplateCopies

    ' ステップ2: 全グラフをPNG画像として出力
    MsgBox "ステップ2/3: グラフをPNGとして出力します...", vbInformation, "一括処理"
    Call ExportAllChartsAsPng

    ' ステップ3: 全シートのズームを100%にしてA1セルにスクロール
    MsgBox "ステップ3/3: 全シートをリセットします...", vbInformation, "一括処理"
    Call A1_1x

    ' 完了メッセージ
    Dim elapsed As Double
    elapsed = Timer - startTime
    MsgBox "すべての処理が完了しました。" & vbCrLf & _
           "処理時間: " & Format(elapsed, "0.0") & "秒", vbInformation, "一括処理完了"

    Exit Sub

ErrorHandler:
    errorMsg = "エラーが発生しました: " & Err.Description & vbCrLf & _
               "エラー番号: " & Err.Number
    MsgBox errorMsg, vbCritical, "一括処理エラー"
End Sub
