Attribute VB_Name = "Module1"
Option Explicit

'========================
' グラフタイトルなし版
'========================
Sub ChangeFontSizeAndBold_NoTitle()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As chart
    Dim fontSize As Integer
    Dim series As series

    ' フォントサイズを指定
    fontSize = 16

    ' 現在のシートを対象にする
    Set ws = ActiveSheet

    ' シート内の全てのグラフオブジェクトに対して処理
    For Each chartObj In ws.ChartObjects
        Set chart = chartObj.chart

        ' 枠線(グラフオブジェクト)の枠線をなしに設定
        chartObj.ShapeRange.Line.Visible = msoFalse

        ' プロットエリアの枠線を灰色、太さ1.25に設定
        chart.PlotArea.Format.Line.ForeColor.RGB = RGB(120, 120, 120)
        chart.PlotArea.Format.Line.Weight = 1.25

        ' 縦軸の線を灰色、太さ1.25に設定
        chart.Axes(xlValue).Format.Line.ForeColor.RGB = RGB(120, 120, 120)
        chart.Axes(xlValue).Format.Line.Weight = 1.25

        ' 横軸の線を灰色、太さ1.25に設定
        chart.Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(120, 120, 120)
        chart.Axes(xlCategory).Format.Line.Weight = 1.25

        ' 縦目盛線(主)をグレー、太さ0.5に設定
        chart.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(160, 160, 160)
        chart.Axes(xlValue).MajorGridlines.Format.Line.Weight = 0.5

        ' 縦目盛線(補助)を薄いグレー、太さ0.5に設定
        chart.Axes(xlValue).MinorGridlines.Format.Line.ForeColor.RGB = RGB(225, 225, 225)
        chart.Axes(xlValue).MinorGridlines.Format.Line.Weight = 0.5

        ' 横目盛線(主)をグレー、太さ0.5に設定
        chart.Axes(xlCategory).MajorGridlines.Format.Line.ForeColor.RGB = RGB(160, 160, 160)
        chart.Axes(xlCategory).MajorGridlines.Format.Line.Weight = 0.5

        ' 横目盛線(補助)を薄いグレー、太さ0.5に設定
        chart.Axes(xlCategory).MinorGridlines.Format.Line.ForeColor.RGB = RGB(225, 225, 225)
        chart.Axes(xlCategory).MinorGridlines.Format.Line.Weight = 0.5

        ' X軸ラベルのフォント設定
        chart.Axes(xlCategory).TickLabels.Font.Size = fontSize
        chart.Axes(xlCategory).TickLabels.Font.Bold = False
        chart.Axes(xlCategory).TickLabels.Font.Color = RGB(0, 0, 0)

        ' Y軸ラベルのフォント設定
        chart.Axes(xlValue).TickLabels.Font.Size = fontSize
        chart.Axes(xlValue).TickLabels.Font.Bold = False
        chart.Axes(xlValue).TickLabels.Font.Color = RGB(0, 0, 0)

        ' X軸タイトルの設定
        If Not chart.Axes(xlCategory, xlPrimary).HasTitle Then
            chart.Axes(xlCategory, xlPrimary).HasTitle = True
        End If
        chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Frequency (Hz)"
        chart.Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = fontSize + 2
        chart.Axes(xlCategory, xlPrimary).AxisTitle.Font.Bold = False
        chart.Axes(xlCategory, xlPrimary).AxisTitle.Font.Color = RGB(0, 0, 0)

        '=== X軸(横軸)対数設定・範囲・目盛 ===
        Dim xAx As Axis
        On Error Resume Next
        Set xAx = chart.Axes(xlCategory, xlPrimary)
        On Error GoTo 0
        If Not xAx Is Nothing Then
            Call SetupLogXAxis(xAx, 10#, 20000#)
        End If

        ' Y軸タイトルの設定
        If Not chart.Axes(xlValue, xlPrimary).HasTitle Then
            chart.Axes(xlValue, xlPrimary).HasTitle = True
        End If
        chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Gain (dB)"
        chart.Axes(xlValue, xlPrimary).AxisTitle.Font.Size = fontSize + 2
        chart.Axes(xlValue, xlPrimary).AxisTitle.Font.Bold = False
        chart.Axes(xlValue, xlPrimary).AxisTitle.Font.Color = RGB(0, 0, 0)

        '=== Y軸(縦軸)のスケーリング設定 ===
        ' 主目盛線：10 dB刻み、補助目盛線：5 dB刻み
        ' 最小値：-40 dB、最大値：プロットの最大値を10の位で切り上げ
        Dim yMax As Double, yMaxCeil As Double
        yMax = GetChartMaxY(chart)
        yMaxCeil = WorksheetFunction.Ceiling(yMax, 10)
        With chart.Axes(xlValue)
            .MinimumScaleIsAuto = False
            .MaximumScaleIsAuto = False
            .MinimumScale = -40
            .MaximumScale = yMaxCeil
            .MajorUnit = 10
            .MinorUnit = 5
            .HasMinorGridlines = True
            .MajorTickMark = xlOutside
            .MinorTickMark = xlOutside
        End With

        ' 凡例の処理
        If chart.HasLegend Then
            With chart.Legend
                With .Font
                    .Size = fontSize - 2
                    .Bold = False
                    .Color = RGB(0, 0, 0)
                End With
                With .Format.Fill
                    .ForeColor.RGB = RGB(255, 255, 255)
                    .Transparency = 0.1
                End With
            End With
        End If

        '--- 単色＆透明度をシリーズ番号に応じて設定(1～11で不透明度UP) ---
        Const BASE_R As Long = 10
        Const BASE_G As Long = 62
        Const BASE_B As Long = 85
        Const MAX_SERIES As Long = 11
        Const TRANS_START As Double = 0.65
        Const TRANS_END   As Double = 0.05

        Dim s As series
        Dim idx As Long
        Dim frac As Double
        Dim t As Double

        idx = 1
        For Each s In chart.SeriesCollection
            With s.Format.Line
                .Visible = msoTrue
                .Weight = 1.75
                .ForeColor.RGB = RGB(BASE_R, BASE_G, BASE_B)
                frac = WorksheetFunction.Min(idx - 1, MAX_SERIES - 1) / (MAX_SERIES - 1)
                t = TRANS_START + (TRANS_END - TRANS_START) * frac
                If t < 0 Then t = 0
                If t > 1 Then t = 1
                .Transparency = t
            End With
            idx = idx + 1
        Next s

        ' データラベルのフォント設定
        For Each series In chart.SeriesCollection
            If series.HasDataLabels Then
                series.DataLabels.Font.Size = fontSize
                series.DataLabels.Font.Bold = False
                series.DataLabels.Font.Color = RGB(0, 0, 0)
            End If
        Next series

        '=== グラフタイトルは使用しない ===
        ' 上部タイトルを無効化
        chart.HasTitle = False
        ' 以前の下部タイトル(BottomTitle)があれば削除
        Dim shp As Shape
        For Each shp In chart.Shapes
            If shp.Name = "BottomTitle" Then
                shp.Delete
                Exit For
            End If
        Next shp
        ' ラベルが詰まりすぎないように、X軸ラベルの更に下に最低パディングを確保
        ' 第2引数は確保したい余白パディング(ポイント)。フォントサイズ係数的 + 余白を足す。
        Call EnsureBottomPadding(chart, fontSize * 0.9 + 8, 0)
    Next chartObj
End Sub

Private Sub SetupLogXAxis(xAx As Axis, minVal As Double, maxVal As Double)
    ' X軸を対数(底10)に設定。値軸の場合のみ MajorUnit=1 を適用。
    ' カテゴリ軸では MajorUnit は設定不可のため、エラーは無視し補助目盛のみ判定する。
    On Error Resume Next
    With xAx
        .ScaleType = xlLogarithmic
        .LogBase = 10
        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = minVal
        .MaximumScale = maxVal

        ' 値軸(xlValue)なら MajorUnit が設定可。カテゴリ軸ではエラーになるため無視。
        .MajorUnitIsAuto = False
        Err.Clear
        .MajorUnit = 1     ' 1デケード刻み(10^n)
        If Err.Number <> 0 Then
            ' カテゴリ軸など：MajorUnit が非対応 → 何もしない
            Err.Clear
        End If

        ' 補助目盛は Excel の対数仕様(2～9倍)を利用
        .MinorUnitIsAuto = True
        .HasMinorGridlines = True
        .MinorTickMark = xlOutside
        .MajorTickMark = xlOutside

        ' ラベルの書式は参照セルに追従
        .TickLabels.NumberFormatLinked = True
    End With
    On Error GoTo 0
End Sub

Private Function GetChartMaxY(ch As chart) As Double
    Dim s As series
    Dim v As Variant
    Dim curMax As Double
    Dim tmp As Double
    curMax = -1E+308
    On Error Resume Next
    For Each s In ch.SeriesCollection
        v = s.Values
        If IsArray(v) Then
            tmp = Application.WorksheetFunction.Max(v)
        Else
            tmp = CDbl(v)
        End If
        If Err.Number = 0 Then
            If tmp > curMax Then curMax = tmp
        Else
            Err.Clear
        End If
    Next s
    On Error GoTo 0
    If curMax = -1E+308 Then
        curMax = 0  ' フォールバック
    End If
    GetChartMaxY = curMax
End Function

Private Sub EnsureBottomPadding(ch As chart, paddingPts As Single, Optional minPlotHeight As Single = 60)
    On Error Resume Next

    Dim axisBottom As Single
    Dim currentGap As Single
    Dim diff As Single

    ' X軸ラベル最下の位置と、ChartArea底とのギャップを取得
    axisBottom = ch.Axes(xlCategory).Top + ch.Axes(xlCategory).Height
    currentGap = ch.ChartArea.Height - axisBottom

    ' 目標との差分(+：余白が不足、-：余白が過剰)
    diff = paddingPts - currentGap

    ' 微妙な差分は調整不要
    If Abs(diff) <= 0.5 Then GoTo CleanExit

    If diff > 0 Then
        ' --- 余白が足りない：プロットエリアを上に押し上げる(高さを縮める優先) ---
        Dim deltaUp As Single
        deltaUp = diff

        If ch.PlotArea.Height > deltaUp + minPlotHeight Then
            ch.PlotArea.Height = ch.PlotArea.Height - deltaUp
        Else
            ' 最低高を守りつつ分をTopを上に上げて確保
            Dim needMoveUp As Single
            needMoveUp = deltaUp - (ch.PlotArea.Height - minPlotHeight)
            ch.PlotArea.Height = minPlotHeight
            ch.PlotArea.Top = Application.Max(0, ch.PlotArea.Top - needMoveUp)
        End If

    Else
        ' --- 余白が余っている：プロットエリアを下側に拡張^上に移動 ---
        Dim needReduce As Single
        needReduce = -diff   ' 減らしたいギャップ量

        ' 1) まず高さを伸ばしてギャップを埋める
        Dim freeBottom As Single
        Dim grow As Single
        freeBottom = ch.ChartArea.Height - (ch.PlotArea.Top + ch.PlotArea.Height)
        grow = Application.Min(needReduce, Application.Max(0, freeBottom - 1))
        If grow > 0 Then
            ch.PlotArea.Height = ch.PlotArea.Height + grow
            needReduce = needReduce - grow
        End If

        ' 2) まだ余るなら、プロットエリアを下に移動
        If needReduce > 0 Then
            freeBottom = ch.ChartArea.Height - (ch.PlotArea.Top + ch.PlotArea.Height)
            Dim moveDown As Single
            moveDown = Application.Min(needReduce, Application.Max(0, freeBottom - 1))
            If moveDown > 0 Then
                ch.PlotArea.Top = ch.PlotArea.Top + moveDown
            End If
        End If
    End If

CleanExit:
    On Error GoTo 0
End Sub

