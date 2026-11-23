Attribute VB_Name = "BikeSharingAnalysis"
Option Explicit

'
' Prerequisites:
' - Import the original data/hour.csv onto a worksheet named Data
' - Convert/import the data to an Excel Table named tblRentals
' - Import this .bas module into the VBA editor (Alt+F11)
' Run: RunReport
' Result: five CSVs + three chart PNGs in <project-root>\results\excel\
'
Public Sub RunReport()
    Dim previousScreenUpdating As Boolean
    previousScreenUpdating = Application.ScreenUpdating

    On Error GoTo Failed
    Application.ScreenUpdating = False

    CheckSourceData
    BuildResultSheets
    Application.CalculateFull
    BuildDashboard
    ExportResults

    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Analysis complete: five CSVs exported and Excel dashboard created. Optional PNGs are in results\excel if supported.", vbInformation
    Exit Sub

Failed:
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Analysis failed: " & Err.Description, vbExclamation
End Sub

' Optional: use this only AFTER the first run if you want to reload hour.csv.
' The original CSV must remain at the path configured by Excel's data import.
Public Sub RefreshAndRun()
    On Error GoTo Failed
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    RunReport
    Exit Sub
Failed:
    MsgBox "Refresh failed: " & Err.Description, vbExclamation
End Sub

' Validate that Excel loaded the same complete CSV as the Python / R runs.
Private Sub CheckSourceData()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Data").ListObjects("tblRentals")

    Dim required As Variant, field As Variant
    required = Array("hr", "yr", "mnth", "workingday", "weathersit", "cnt", "casual", "registered")
    For Each field In required
        Dim col As ListColumn
        Set col = tbl.ListColumns(CStr(field))   ' Raises an error when a column is absent
    Next field

    If tbl.ListRows.Count <> 17379 Then
        Err.Raise vbObjectError + 101, , "Expected 17,379 rows; found " & tbl.ListRows.Count
    End If

    Dim total As Double
    total = Application.WorksheetFunction.Sum(tbl.ListColumns("cnt").DataBodyRange)
    If total <> 3292679 Then
        Err.Raise vbObjectError + 102, , "Expected 3,292,679 rentals; found " & total
    End If
End Sub

' Create each analytical table with normal Excel formulas.
' Structured references such as tblRentals[cnt] automatically cover the full data table.
Private Sub BuildResultSheets()
    Dim ws As Worksheet, r As Long, h As Long, w As Long
    Dim yearNum As Long, monthNum As Long, yearCode As Long

    ' 1. Overall metrics: two rows, matching summary.csv
    Set ws = PrepareSheet("Summary", "metric", "value")
    ws.Range("A2").Value = "total_rentals"
    ws.Range("B2").Formula = "=SUM(tblRentals[cnt])"
    ws.Range("A3").Value = "avg_hourly_demand"
    ws.Range("B3").Formula = "=AVERAGE(tblRentals[cnt])"
    ws.Range("B2").NumberFormat = "0"
    ws.Range("B3").NumberFormat = "0.0000"

    ' 2. Mean of cnt for each hour of the day: 0..23
    Set ws = PrepareSheet("Hourly", "hr", "avg_rentals")
    For h = 0 To 23
        r = h + 2
        ws.Cells(r, 1).Value = h
        ws.Cells(r, 2).Formula = "=AVERAGEIF(tblRentals[hr],A" & r & ",tblRentals[cnt])"
    Next h
    ws.Range("B2:B25").NumberFormat = "0.0000"

    ' 3. Mean of cnt for workingday 0 (non-working) and 1 (working)
    Set ws = PrepareSheet("Workingday", "workingday", "avg_rentals")
    For w = 0 To 1
        r = w + 2
        ws.Cells(r, 1).Value = w
        ws.Cells(r, 2).Formula = "=AVERAGEIF(tblRentals[workingday],A" & r & ",tblRentals[cnt])"
    Next w
    ws.Range("B2:B3").NumberFormat = "0.0000"

    ' 4. Mean of cnt for each weather code PRESENT in the original data.
    Set ws = PrepareSheet("Weather", "weathersit", "avg_rentals")
    r = 2
    For w = 1 To 4
        If Application.WorksheetFunction.CountIf( _
             ThisWorkbook.Worksheets("Data").ListObjects("tblRentals").ListColumns("weathersit").DataBodyRange, w) > 0 Then
            ws.Cells(r, 1).Value = w
            ws.Cells(r, 2).Formula = "=AVERAGEIF(tblRentals[weathersit],A" & r & ",tblRentals[cnt])"
            r = r + 1
        End If
    Next w
    ws.Columns("B").NumberFormat = "0.0000"

    ' 5. Monthly totals: 2011-01 to 2012-12.
    ' In hour.csv, yr=0 represents 2011 and yr=1 represents 2012.
    ' We use existing yr/mnth columns rather than adding a redundant date field.
    Set ws = PrepareSheet("Monthly", "year_month", "total_rentals")
    r = 2
    For yearNum = 2011 To 2012
        yearCode = yearNum - 2011
        For monthNum = 1 To 12
            ws.Cells(r, 1).NumberFormat = "@"  ' Keep YYYY-MM as TEXT, not a date
            ws.Cells(r, 1).Value = CStr(yearNum) & "-" & Right$("0" & CStr(monthNum), 2)
            ws.Cells(r, 2).Formula = "=SUMIFS(tblRentals[cnt],tblRentals[yr]," & _
                yearCode & ",tblRentals[mnth]," & monthNum & ")"
            r = r + 1
        Next monthNum
    Next yearNum
    ws.Columns("B").NumberFormat = "0"
End Sub

' Find or create an output worksheet; replace its previous results.
Private Function PrepareSheet(ByVal sheetName As String, _
    ByVal firstHeader As String, ByVal secondHeader As String) As Worksheet

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    ws.Cells(1, 1).Value = firstHeader
    ws.Cells(1, 2).Value = secondHeader
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Interior.Color = RGB(221, 235, 247)
    ws.Columns("A").ColumnWidth = 23
    ws.Columns("B").ColumnWidth = 20
    Set PrepareSheet = ws
End Function

' Create an Excel dashboard with three charts that read the result worksheets.
Private Sub BuildDashboard()
    Dim ws As Worksheet, obj As ChartObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Dashboard"
    End If

    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
    ws.Cells.Clear
    ws.Range("A1").Value = "BIKE SHARING | EXCEL + VBA"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Total rentals"
    ws.Range("B3").Formula = "=Summary!B2"
    ws.Range("A4").Value = "Avg. rentals per observed hour"
    ws.Range("B4").Formula = "=Summary!B3"
    ws.Range("B3").NumberFormat = "#,##0"
    ws.Range("B4").NumberFormat = "0.00"
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 20

    Set obj = AddSimpleChart(ws, "Hourly", 25, xlLineMarkers, _
        "Average rentals by hour", 15, 125, 480, 270)
    Set obj = AddSimpleChart(ws, "Workingday", 3, xlColumnClustered, _
        "Mean rentals: non-working (0) vs working (1)", 515, 125, 430, 270)
    Set obj = AddSimpleChart(ws, "Monthly", 25, xlLineMarkers, _
        "Monthly total rentals", 15, 415, 930, 290)
End Sub

Private Function AddSimpleChart(ByVal destination As Worksheet, ByVal sourceName As String, _
    ByVal lastRow As Long, ByVal chartKind As XlChartType, ByVal chartTitle As String, _
    ByVal leftPos As Double, ByVal topPos As Double, _
    ByVal chartWidth As Double, ByVal chartHeight As Double) As ChartObject

    Dim src As Worksheet
    Set src = ThisWorkbook.Worksheets(sourceName)

    Dim ch As ChartObject
    Set ch = destination.ChartObjects.Add(leftPos, topPos, chartWidth, chartHeight)
    With ch.Chart
        .ChartType = chartKind
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        .HasLegend = False
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = src.Range("B1").Value
        .SeriesCollection(1).XValues = src.Range("A2:A" & lastRow)
        .SeriesCollection(1).Values = src.Range("B2:B" & lastRow)
    End With
    Set AddSimpleChart = ch
End Function

' Save five result tables as true comma-separated CSV files, plus chart PNGs.
Public Sub ExportResults()
    Dim outFolder As String
    outFolder = GetOutputFolder()

    ExportTwoColumnSheet ThisWorkbook.Worksheets("Summary"), outFolder & "\summary.csv"
    ExportTwoColumnSheet ThisWorkbook.Worksheets("Hourly"), outFolder & "\hourly.csv"
    ExportTwoColumnSheet ThisWorkbook.Worksheets("Workingday"), outFolder & "\workingday.csv"
    ExportTwoColumnSheet ThisWorkbook.Worksheets("Weather"), outFolder & "\weather.csv"
    ExportTwoColumnSheet ThisWorkbook.Worksheets("Monthly"), outFolder & "\monthly.csv"

    ' Excel's chart export usually works on Windows; PNG export is optional.
    ' The CSV exports are the required cross-technology outputs.
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").ChartObjects(1).Chart.Export outFolder & "\chart_hourly.png", "PNG"
    ThisWorkbook.Worksheets("Dashboard").ChartObjects(2).Chart.Export outFolder & "\chart_workingday.png", "PNG"
    ThisWorkbook.Worksheets("Dashboard").ChartObjects(3).Chart.Export outFolder & "\chart_monthly.png", "PNG"
    On Error GoTo 0
End Sub

' Create <project-root>\results\excel\ even if the results folder does not exist.
Private Function GetOutputFolder() As String
    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 103, , "Save the workbook in the project's excel folder first."
    End If

    Dim fso As Object, rootFolder As String, resultsFolder As String, excelFolder As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    rootFolder = fso.GetParentFolderName(ThisWorkbook.Path)
    resultsFolder = fso.BuildPath(rootFolder, "results")
    excelFolder = fso.BuildPath(resultsFolder, "excel")

    If Not fso.FolderExists(resultsFolder) Then fso.CreateFolder resultsFolder
    If Not fso.FolderExists(excelFolder) Then fso.CreateFolder excelFolder
    GetOutputFolder = excelFolder
End Function

' Export the full precision of the calculated Excel values, not the rounded display.
Private Sub ExportTwoColumnSheet(ByVal ws As Worksheet, ByVal filename As String)
    Dim fileNum As Integer, lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    fileNum = FreeFile
    Open filename For Output As #fileNum
    On Error GoTo CloseAndFail
    For r = 1 To lastRow
        Print #fileNum, CsvField(ws.Cells(r, 1).Value2) & "," & CsvField(ws.Cells(r, 2).Value2)
    Next r
    Close #fileNum
    Exit Sub

CloseAndFail:
    Close #fileNum
    Err.Raise Err.Number, , "Export failed for " & filename & ": " & Err.Description
End Sub

' Str$ uses a period for decimals regardless of Windows locale.
' Text fields are quoted only when required by the CSV standard.
Private Function CsvField(ByVal value As Variant) As String
    If IsNumeric(value) And VarType(value) <> vbString Then
        CsvField = Trim$(Str$(CDbl(value)))
        Exit Function
    End If

    Dim t As String
    t = CStr(value)
    If InStr(t, ",") > 0 Or InStr(t, """") > 0 Or _
       InStr(t, vbCr) > 0 Or InStr(t, vbLf) > 0 Then
        t = """" & Replace(t, """", """""") & """"
    End If
    CsvField = t
End Function
