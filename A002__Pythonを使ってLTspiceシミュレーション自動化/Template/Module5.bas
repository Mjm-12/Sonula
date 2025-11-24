Attribute VB_Name = "Module5"
Option Explicit

Public Sub ImportCsvsIntoTemplateCopies()
    Dim basePath As String, csvName As String
    Dim tpl As Worksheet, outWS As Worksheet
    Dim wasScreenUpdating As Boolean, wasEnableEvents As Boolean, wasDisplayAlerts As Boolean, wasCalc As XlCalculation
    
    Set tpl = Nothing
    On Error Resume Next
    Set tpl = ThisWorkbook.Worksheets("Template")
    On Error GoTo 0
    If tpl Is Nothing Then
        MsgBox "Template シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        MsgBox "このブックを一度保存してください（保存先フォルダが取得できません）。", vbExclamation
        Exit Sub
    End If
    
    ' 高速化
    wasScreenUpdating = Application.ScreenUpdating
    wasEnableEvents = Application.EnableEvents
    wasDisplayAlerts = Application.DisplayAlerts
    wasCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CLEAN_FAIL
    
    csvName = Dir(basePath & Application.PathSeparator & "*.csv")
    If Len(csvName) = 0 Then
        MsgBox "フォルダ内にCSVが見つかりませんでした。" & vbCrLf & basePath, vbInformation
        GoTo CLEAN_EXIT
    End If
    
    Do While Len(csvName) > 0
        ' --- Template を複製して処理用シートを作成 ---
        tpl.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set outWS = ActiveSheet
        On Error Resume Next
        outWS.Name = Left$(SanitizeSheetName(RemoveCsvExt(csvName)), 31) ' 31文字制限
        On Error GoTo 0
        
        ' --- CSVを開いて配列に読み込み ---
        Dim csvFullPath As String
        csvFullPath = basePath & Application.PathSeparator & csvName
        
        Dim wbCsv As Workbook, wsCsv As Worksheet
        Set wbCsv = Nothing: Set wsCsv = Nothing
        
        ' UTF-8(BOM 無/有) に広く対応するため OpenText を使用
        Workbooks.OpenText Filename:=csvFullPath, _
                           Origin:=65001, DataType:=xlDelimited, Comma:=True, _
                           Local:=False
        Set wbCsv = ActiveWorkbook
        Set wsCsv = wbCsv.Worksheets(1)
        
        Dim data As Variant
        data = wsCsv.UsedRange.Value ' 1-based 2次元配列
        
        ' --- データを索引化： index→(frequency→mag) ---
        Dim dictIdx As Object
        Set dictIdx = CreateObject("Scripting.Dictionary") ' key: index (String), item: Dictionary(freqKey → mag)
        
        Dim i As Long, nRow As Long
        Dim freq As Double, mag As Double
        Dim idx As Variant ' indexは数値のはずだがキー化時は文字列で統一
        Dim freqKey As String
        
        If IsArray(data) Then
            nRow = UBound(data, 1)
            
            ' ヘッダ行を想定（1行目）→ 2行目から読み込む
            For i = 2 To nRow
                If Not IsError(data(i, 1)) And Not IsError(data(i, 2)) And Not IsError(data(i, 4)) Then
                    If Len(data(i, 1)) > 0 And Len(data(i, 2)) > 0 And Len(data(i, 4)) > 0 Then
                        freq = CDbl(data(i, 1))          ' A列 frequency_Hz
                        mag = CDbl(data(i, 2))           ' B列 mag
                        idx = CStr(CLng(data(i, 4)))     ' D列 step_index（文字列キー化）
                        
                        ' indexごとの周波数→mag辞書を用意
                        If Not dictIdx.Exists(idx) Then
                            Set dictIdx(idx) = CreateObject("Scripting.Dictionary")
                        End If
                        freqKey = GetFreqKey(freq)       ' 浮動小数の一致誤差対策
                        dictIdx(idx)(freqKey) = mag
                    End If
                End If
            Next i
        End If
        
        ' --- 複製シートへ貼り付け ---
        ' 1) 行見出し（B2↓）： step_index = 0 の frequency を順番に並べる
        Dim listFreq As Variant, r As Long
        listFreq = BuildSortedFreqList(dictIdx, "0") ' index=0 の周波数リスト（並び保持）
        If IsEmpty(listFreq) Then
            ' index 0 がない場合は、最小のindexを使って列挙（保険）
            Dim fallbackIdx As String
            fallbackIdx = FirstKey(dictIdx)
            listFreq = BuildSortedFreqList(dictIdx, fallbackIdx)
        End If
        
        If Not IsEmpty(listFreq) Then
            ' 貼り付け
            outWS.Range("B2").Resize(UBound(listFreq) - LBound(listFreq) + 1, 1).Value = _
                ToVerticalRange(listFreq)
        End If
        
        ' 2) 列見出し（C1→）の index を読み取り、frequency×indexで mag を埋める
        Dim lastCol As Long, c As Long
        lastCol = outWS.Cells(1, outWS.Columns.Count).End(xlToLeft).Column
        If lastCol < 3 Then lastCol = 2 ' C列未満なら処理スキップ
        
        For c = 3 To lastCol
            Dim headerVal As Variant
            headerVal = outWS.Cells(1, c).Value
            
            If IsNumeric(headerVal) Then
                Dim idxKey As String
                idxKey = CStr(CLng(headerVal))
                
                If dictIdx.Exists(idxKey) Then
                    ' 各行の周波数に対する mag を充填
                    For r = LBound(listFreq) To UBound(listFreq)
                        freqKey = GetFreqKey(CDbl(listFreq(r)))
                        If dictIdx(idxKey).Exists(freqKey) Then
                            outWS.Cells(r - LBound(listFreq) + 2, c).Value = dictIdx(idxKey)(freqKey)
                        Else
                            ' 該当frequencyが無ければ空白（必要ならNA等）
                            outWS.Cells(r - LBound(listFreq) + 2, c).Value = vbNullString
                        End If
                    Next r
                Else
                    ' ヘッダーにある index にCSV側のデータが無い
                    ' → 空欄のまま（必要なら0を入れる等に変更可）
                End If
            End If
        Next c
        
        ' CSVを閉じる（保存しない）
        wbCsv.Close SaveChanges:=False
        
        ' 次のCSVへ
        csvName = Dir()
    Loop
    
    MsgBox "完了しました。", vbInformation
    GoTo CLEAN_EXIT

CLEAN_FAIL:
    MsgBox "エラー: " & Err.Description, vbCritical

CLEAN_EXIT:
    ' 復帰
    Application.Calculation = wasCalc
    Application.DisplayAlerts = wasDisplayAlerts
    Application.EnableEvents = wasEnableEvents
    Application.ScreenUpdating = wasScreenUpdating
End Sub

' ===== ヘルパー =====

Private Function GetFreqKey(ByVal f As Double) As String
    ' 周波数をキー化（丸め誤差対策：有効桁を揃えて文字列化）
    GetFreqKey = Format$(f, "0.##############") ' 必要に応じて桁数調整
End Function

Private Function BuildSortedFreqList(ByVal dictIdx As Object, ByVal idxKey As String) As Variant
    ' 指定 index の frequency キー一覧（キー順のまま＝CSVの順序）を Double 配列で返す
    ' Scripting.Dictionary は挿入順を保持する（現行Excel/VBAの実装前提）
    If Not dictIdx.Exists(idxKey) Then Exit Function
    Dim k As Variant, i As Long
    ReDim arr(1 To dictIdx(idxKey).Count) As Double
    i = 1
    For Each k In dictIdx(idxKey).Keys
        arr(i) = CDbl(k)
        i = i + 1
    Next k
    BuildSortedFreqList = arr
End Function

Private Function ToVerticalRange(ByVal arr As Variant) As Variant
    ' 1次元配列 → 縦ベクトル(2次元)にして返す
    Dim i As Long, lo As Long, hi As Long
    lo = LBound(arr): hi = UBound(arr)
    Dim m As Variant
    ReDim m(1 To hi - lo + 1, 1 To 1)
    For i = lo To hi
        m(i - lo + 1, 1) = arr(i)
    Next i
    ToVerticalRange = m
End Function

Private Function RemoveCsvExt(ByVal s As String) As String
    If LCase$(Right$(s, 4)) = ".csv" Then
        RemoveCsvExt = Left$(s, Len(s) - 4)
    Else
        RemoveCsvExt = s
    End If
End Function

Private Function SanitizeSheetName(ByVal s As String) As String
    Dim badChars As Variant, i As Long
    badChars = Array("/", "\", "[", "]", ":", "*", "?", """")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, CStr(badChars(i)), "_")
    Next i
    ' 先頭・末尾のシングルクォート回避
    If Left$(s, 1) = "'" Then s = "_" & Mid$(s, 2)
    If Right$(s, 1) = "'" Then s = Left$(s, Len(s) - 1) & "_"
    If Len(Trim$(s)) = 0 Then s = "Sheet"
    SanitizeSheetName = s
End Function

Private Function FirstKey(ByVal dict As Object) As String
    Dim k As Variant
    For Each k In dict.Keys
        FirstKey = CStr(k)
        Exit Function
    Next k
    FirstKey = vbNullString
End Function


