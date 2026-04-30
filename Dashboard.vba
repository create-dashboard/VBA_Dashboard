'-------------------------------------------------------------
'  MAIN ENTRY â€” for first-time build
'-------------------------------------------------------------
Sub BuildDashboard()
    Call DashboardMain
    MsgBox "Dashboard created successfully!", vbInformation, "Done"
End Sub


'-------------------------------------------------------------
'  CORE LOGIC
'  FIX 2: Never deletes the Dashboard sheet.
'         Finds or creates it, then clears content/charts
'         and rebuilds everything in-place.
'-------------------------------------------------------------
Private Sub DashboardMain()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Validate Pivot sheet
    Dim wsPvt As Worksheet
    Set wsPvt = Nothing
    On Error Resume Next
    Set wsPvt = ThisWorkbook.Sheets("Pivot")
    On Error GoTo 0
    If wsPvt Is Nothing Then
        MsgBox "Sheet 'Pivot' not found!", vbCritical
        GoTo CleanUp
    End If

    ' FIX 2: Get existing Dashboard sheet OR create it once
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0

    If ws Is Nothing Then
        ' First run â€” create the sheet
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "Dashboard"
        ws.Tab.Color = RGB(94, 59, 183)
    End If

    ' FIX 2: Clear existing content and charts in-place (preserves sheet)
    ws.Cells.Clear
    ws.Cells.Interior.ColorIndex = xlNone

    ' Remove existing shapes safely (backwards loop)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        ws.Shapes(i).Delete
        On Error GoTo 0
    Next i

    ' Rebuild all sections
    Call SetupDashboardLayout(ws)
    Call BuildHeader(ws)
    Call BuildKPICards(ws, wsPvt)
    Call CreateAllCharts(ws, wsPvt)
    Call AddMacroButton(ws)
    Call AddSlicers(ws, wsPvt)

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


'-------------------------------------------------------------
'  Layout: column widths + row heights
'-------------------------------------------------------------
Sub SetupDashboardLayout(ws As Worksheet)
    Dim c As Integer
    For c = 1 To 17
        ws.Columns(c).ColumnWidth = 9.5
    Next c

    ws.Rows("1:2").RowHeight = 20    ' Header
    ws.Rows("3:3").RowHeight = 5     ' Spacer
    ws.Rows("4:6").RowHeight = 16    ' KPI rows
    ws.Rows("7:7").RowHeight = 5     ' Spacer

    Dim r As Integer
    For r = 8 To 31
        ws.Rows(r).RowHeight = 15
    Next r

    ws.Range("A1:Q31").Interior.Color = RGB(248, 247, 255)
End Sub


'-------------------------------------------------------------
'  Header banner
'-------------------------------------------------------------
Sub BuildHeader(ws As Worksheet)
    With ws.Range("A1:Q2")
        .UnMerge
        .Interior.Color = RGB(94, 59, 183)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Merge
    End With
    
    With ws.Range("A1")
        .Value = "SALES PERFORMANCE DASHBOARD"
        .Font.Name = "Calibri"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub


'-------------------------------------------------------------
'  KPI Cards
'  FIX 1: All borders (top, sides, bottom) now use the SAME
'         style: xlContinuous + xlThin + purpleDeep
'         The shadow row keeps its lighter fill but matches
'         the border weight of the header row â€” no heavy
'         xlMedium bottom that looked mismatched.
'-------------------------------------------------------------
Sub BuildKPICards(ws As Worksheet, wsPvt As Worksheet)
    Dim purple     As Long: purple = RGB(94, 59, 183)
    Dim purpleDeep As Long: purpleDeep = RGB(60, 30, 130)
    Dim bgWhite    As Long: bgWhite = RGB(255, 255, 255)
    Dim shadow     As Long: shadow = RGB(210, 205, 230)     ' lighter shadow
    Dim textDark   As Long: textDark = RGB(30, 30, 30)

    Dim kpiLabels(5) As String
    Dim kpiFmlas(5)  As String
    Dim kpiFmts(5)   As String
    Dim kpiCols(5)   As Integer

    kpiLabels(0) = "Total Sales"
    kpiLabels(1) = "Total Profit"
    kpiLabels(2) = "Total Orders"
    kpiLabels(3) = "Total Qty Sold"
    kpiLabels(4) = "Avg Discount"
    kpiLabels(5) = "Profit Margin"

    'kpiFmlas(0) = "=Pivot!B6"
    'kpiFmlas(1) = "=Pivot!F46"
    'kpiFmlas(2) = "=COUNTA(Data!A2:A10000)"
    'kpiFmlas(3) = "=SUM(Data!L2:L10000)"
    'kpiFmlas(4) = "=Pivot!W20"
    'kpiFmlas(5) = "=Pivot!F46/Pivot!B6"

    kpiFmlas(0) = "=GETPIVOTDATA(""Total Sales"",Pivot!$D$2)"
    kpiFmlas(1) = "=GETPIVOTDATA(""Total Profit"",Pivot!$D$2)"
    kpiFmlas(2) = "=COUNTA(Pivot!Y:Y)"
    kpiFmlas(3) = "=GETPIVOTDATA(""Total Quantity"",Pivot!$AB$2)"
    kpiFmlas(4) = "=GETPIVOTDATA(""Avg Discount"",Pivot!$U$2)"
    kpiFmlas(5) = "=D5/B5"

    kpiFmts(0) = "[$$-en-US]#,##0"
    kpiFmts(1) = "[$$-en-US]#,##0"
    
'    kpiFmts(0) = """$""#,##0"
'    kpiFmts(1) = """$""#,##0"
    kpiFmts(2) = "#,##0"
    kpiFmts(3) = "#,##0"
    kpiFmts(4) = "0.0%"
    kpiFmts(5) = "0.0%"

    kpiCols(0) = 2: kpiCols(1) = 4: kpiCols(2) = 6
    kpiCols(3) = 8: kpiCols(4) = 10: kpiCols(5) = 12

    Dim k As Integer
    For k = 0 To 5
        Dim c As Integer: c = kpiCols(k)

        ' â”€â”€ ROW 4: Purple header (label) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        With ws.Range(ws.Cells(4, c), ws.Cells(4, c + 1))
            .Merge
            .Value = kpiLabels(k)
            .Interior.Color = purple
            .Font.Bold = True
            .Font.Size = 9
            .Font.Name = "Calibri"
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            ' FIX 1: All borders uniform - xlThin + purpleDeep
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).Color = purpleDeep
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).Color = purpleDeep
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeRight).Color = purpleDeep
        End With

        ' â”€â”€ ROW 5: White body (live formula value) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        With ws.Range(ws.Cells(5, c), ws.Cells(5, c + 1))
            .Merge
            .Formula = kpiFmlas(k)
            .NumberFormat = kpiFmts(k)
            .Interior.Color = bgWhite
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = "Calibri"
            .Font.Color = textDark
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            ' FIX 1: Side borders only - same xlThin + purpleDeep
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).Color = purpleDeep
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeRight).Color = purpleDeep
        End With

        ' â”€â”€ ROW 6: Shadow strip (subtle 3D base) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        ' FIX 1: Bottom border now matches top â€” xlThin + purpleDeep
        '        (was xlMedium which looked heavier than the top)
        With ws.Range(ws.Cells(6, c), ws.Cells(6, c + 1))
            .Merge
            .Interior.Color = shadow
            ' All four borders uniform â€” xlThin + purpleDeep
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).Color = purpleDeep
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin         ' FIX 1: was xlMedium
            .Borders(xlEdgeBottom).Color = purpleDeep      ' FIX 1: same as top
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).Color = purpleDeep
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin          ' FIX 1: was xlMedium
            .Borders(xlEdgeRight).Color = purpleDeep
        End With
    Next k
End Sub


'-------------------------------------------------------------
'  CREATE ALL CHARTS
'-------------------------------------------------------------
Sub CreateAllCharts(ws As Worksheet, wsPvt As Worksheet)
    Dim cPurple As Long: cPurple = RGB(94, 59, 183)
    Dim cPink   As Long: cPink = RGB(197, 90, 240)
    Dim cOrange As Long: cOrange = RGB(255, 140, 0)
    Dim cBlue   As Long: cBlue = RGB(68, 114, 196)
    Dim cWhite  As Long: cWhite = RGB(255, 255, 255)

    Dim L1  As Double: L1 = ws.Cells(8, 1).Left
    Dim L2  As Double: L2 = ws.Cells(8, 6).Left
    Dim L3  As Double: L3 = ws.Cells(8, 12).Left
    Dim T1  As Double: T1 = ws.Cells(8, 1).Top
    Dim T2  As Double: T2 = ws.Cells(19, 1).Top
    Dim CW1 As Double: CW1 = ws.Cells(8, 6).Left - ws.Cells(8, 1).Left - 3
    Dim CW2 As Double: CW2 = ws.Cells(8, 12).Left - ws.Cells(8, 6).Left - 3
    Dim CW3 As Double: CW3 = ws.Cells(8, 17).Left - ws.Cells(8, 12).Left - 3
    Dim CH  As Double: CH = ws.Cells(19, 1).Top - ws.Cells(8, 1).Top - 3

    Dim cht As ChartObject
    Dim sc  As Integer
    Dim p   As Integer

    ' CHART 1 â€” Sales by Segment (Doughnut)
    Set cht = ws.ChartObjects.Add(Left:=L1, Top:=T1, Width:=CW1, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("A3:B5")
        .ChartType = xlDoughnut
        .HasTitle = True
        .ChartTitle.Text = "Sales by Segment"
        Call ApplyTitleStyle(.ChartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Points(1).Interior.Color = cOrange
        .SeriesCollection(1).Points(2).Interior.Color = cPurple
        .SeriesCollection(1).Points(3).Interior.Color = cPink
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowPercentage = True
        .SeriesCollection(1).DataLabels.ShowValue = False
        .SeriesCollection(1).DataLabels.Font.Size = 9
        .SeriesCollection(1).DataLabels.Font.Bold = True
        .SeriesCollection(1).DataLabels.Font.Color = cWhite
        On Error GoTo 0
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .Legend.Font.Size = 8
        .ShowAllFieldButtons = False
    End With

    ' CHART 2 â€” Sales & Profit Trend (Line)
    Set cht = ws.ChartObjects.Add(Left:=L2, Top:=T1, Width:=CW2, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("D3:F46")
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Sales & Profit Trend"
        Call ApplyTitleStyle(.ChartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        With .SeriesCollection(1)
            .Name = "Total Sales"
            .Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent2
            .Format.Line.ForeColor.Brightness = 0
            .Format.Line.Weight = 2
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
            .Format.MarkerFill.ForeColor.ObjectThemeColor = msoThemeColorAccent2
            .Format.MarkerLine.ForeColor.ObjectThemeColor = msoThemeColorAccent2
        End With
        With .SeriesCollection(2)
            .Name = "Total Profit"
            .Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent6
            .Format.Line.ForeColor.Brightness = 0
            .Format.Line.Weight = 2
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
            .Format.MarkerFill.ForeColor.ObjectThemeColor = msoThemeColorAccent6
            .Format.MarkerLine.ForeColor.ObjectThemeColor = msoThemeColorAccent6
        End With
        On Error GoTo 0
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 7
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    ' CHART 3 â€” Monthly Sales Performance (Column)
    Set cht = ws.ChartObjects.Add(Left:=L3, Top:=T1, Width:=CW3, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("H3:L15")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Monthly Sales Performance"
        Call ApplyTitleStyle(.ChartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        Dim yearCols(3) As Long
        yearCols(0) = cPurple
        yearCols(1) = cPink
        yearCols(2) = cOrange
        yearCols(3) = cBlue
        For sc = 1 To .SeriesCollection.Count
            If sc <= 4 Then .SeriesCollection(sc).Interior.Color = yearCols(sc - 1)
        Next sc
        On Error GoTo 0
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 8
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    ' CHART 4 â€” Profit vs Discount (Column)
    Set cht = ws.ChartObjects.Add(Left:=L1, Top:=T2, Width:=CW1, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("U3:V19")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Profit vs Discount"
        Call ApplyTitleStyle(.ChartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Interior.Color = cPink
        .SeriesCollection(1).Name = "Total Profit"
        On Error GoTo 0
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 6
        .Axes(xlCategory).TickLabelSpacing = 1
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    ' CHART 5 â€” Sales by Category (Bar)
    Set cht = ws.ChartObjects.Add(Left:=L2, Top:=T2, Width:=CW2, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("O3:P5")
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Sales by Category"
        Call ApplyTitleStyle(.ChartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Points(1).Interior.Color = cPink
        .SeriesCollection(1).Points(2).Interior.Color = cPurple
        .SeriesCollection(1).Points(3).Interior.Color = cOrange
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(1).DataLabels.NumberFormat = "#,##0"
        .SeriesCollection(1).DataLabels.Font.Size = 8
        On Error GoTo 0
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .Legend.Font.Size = 8
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 9
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    ' CHART 6 â€” Sales by Region (Column)
    Set cht = ws.ChartObjects.Add(Left:=L3, Top:=T2, Width:=CW3, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("R3:S6")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Sales by Region"
        Call ApplyTitleStyle(.ChartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        Dim regCols(3) As Long
        regCols(0) = cPurple
        regCols(1) = cPink
        regCols(2) = cOrange
        regCols(3) = cBlue
        For p = 1 To .SeriesCollection(1).Points.Count
            If p <= 4 Then .SeriesCollection(1).Points(p).Interior.Color = regCols(p - 1)
        Next p
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(1).DataLabels.NumberFormat = "$#,##0"
        .SeriesCollection(1).DataLabels.Font.Size = 8
        On Error GoTo 0
        .HasLegend = False
        With .Axes(xlValue)
            .MinimumScaleIsAuto = True
            '.MinimumScale = 2050000
            .MaximumScaleIsAuto = True
            .TickLabels.NumberFormat = "#,##0"
            .TickLabels.Font.Size = 8
        End With
        .Axes(xlCategory).TickLabels.Font.Size = 9
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With
End Sub


'-------------------------------------------------------------
'  MACRO BUTTON â€” top-right of header, stays in same sheet
'  .OnAction -> RefreshDashboard (rebuilds in-place)
'-------------------------------------------------------------
Sub AddMacroButton(ws As Worksheet)
    ' Button is already deleted in DashboardMain shape-clear loop
    Dim btnLeft   As Double: btnLeft = ws.Cells(1, 15).Left + 4
    Dim btnTop    As Double: btnTop = ws.Cells(1, 1).Top + 4
    Dim btnWidth  As Double: btnWidth = ws.Cells(1, 17).Left + ws.Columns(17).Width - ws.Cells(1, 15).Left - 8
    Dim btnHeight As Double: btnHeight = ws.Cells(3, 1).Top - ws.Cells(1, 1).Top - 8

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                  btnLeft, btnTop, btnWidth, btnHeight)
    With shp
        .Name = "btnRefreshDashboard"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Solid
        .Line.ForeColor.RGB = RGB(94, 59, 183)
        .Line.Weight = 1.5
        .Line.Visible = msoTrue
        .TextFrame2.TextRange.Text = "Refresh Dashboard"
        With .TextFrame2.TextRange.Font
            .Bold = msoTrue
            .Size = 9
            .Fill.ForeColor.RGB = RGB(94, 59, 183)
        End With
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = "RefreshDashboard"
    End With
End Sub


'-------------------------------------------------------------
'  HELPER: Title styling
'-------------------------------------------------------------
Sub ApplyTitleStyle(ct As ChartTitle, titleColor As Long)
    ct.Font.Bold = True
    ct.Font.Color = titleColor
    ct.Font.Size = 11
    ct.Font.Name = "Calibri"
End Sub
'-------------------------------------------------------------
'  HELPER: Chart area styling
'-------------------------------------------------------------
Sub ApplyChartAreaStyle(cht As Chart)
    cht.ChartArea.Interior.Color = RGB(245, 242, 255)
    With cht.ChartArea.Border
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(94, 59, 183)
    End With
    cht.PlotArea.Interior.ColorIndex = xlNone
    On Error Resume Next
    cht.PlotArea.Border.LineStyle = xlNone
    On Error GoTo 0
End Sub
'-------------------------------------------------------------
'  HELPER: Remove all gridlines
'-------------------------------------------------------------
Sub RemoveGridlines(cht As Chart)
    On Error Resume Next
    cht.Axes(xlValue).HasMajorGridlines = False
    cht.Axes(xlValue).HasMinorGridlines = False
    cht.Axes(xlCategory).HasMajorGridlines = False
    cht.Axes(xlCategory).HasMinorGridlines = False
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
'  SLICERS (Dynamic Visualization)
'-------------------------------------------------------------
Sub AddSlicers(ws As Worksheet, wsPvt As Worksheet)
    Dim sc As SlicerCache
    Dim pt As PivotTable
    
    ' 1. Delete all existing slicer caches to prevent duplication
    For Each sc In ThisWorkbook.SlicerCaches
        sc.Delete
    Next sc
    
    ' 2. Create SlicerCaches based on PT_Trend (which has Year and Region fields in Data sheet)
    Dim scYear As SlicerCache
    Set scYear = ThisWorkbook.SlicerCaches.Add(wsPvt.PivotTables("PT_Trend"), "Year", "SlicerCache_Year")
    
    Dim scRegion As SlicerCache
    Set scRegion = ThisWorkbook.SlicerCaches.Add(wsPvt.PivotTables("PT_Trend"), "Region", "SlicerCache_Region")
    
    ' 3. Connect SlicerCaches to ALL other PivotTables
    For Each pt In wsPvt.PivotTables
        If pt.Name <> "PT_Trend" Then
            On Error Resume Next
            scYear.PivotTables.AddPivotTable pt
            scRegion.PivotTables.AddPivotTable pt
            On Error GoTo 0
        End If
    Next pt
    
    ' 4. Create 'Year' Slicer on Dashboard (Cols 14-15, Rows 4-6)
    Dim slYear As Slicer
    Set slYear = scYear.Slicers.Add(ws, , "Slicer_Year", "Year", _
                                    Top:=ws.Cells(4, 1).Top, _
                                    Left:=ws.Cells(4, 14).Left, _
                                    Width:=ws.Cells(4, 16).Left - ws.Cells(4, 14).Left - 2, _
                                    Height:=ws.Cells(7, 1).Top - ws.Cells(4, 1).Top)
    
    ' 5. Create 'Region' Slicer on Dashboard (Cols 16-17, Rows 4-6)
    Dim slRegion As Slicer
    Set slRegion = scRegion.Slicers.Add(ws, , "Slicer_Region", "Region", _
                                        Top:=ws.Cells(4, 1).Top, _
                                        Left:=ws.Cells(4, 16).Left, _
                                        Width:=ws.Cells(4, 18).Left - ws.Cells(4, 16).Left - 2, _
                                        Height:=ws.Cells(7, 1).Top - ws.Cells(4, 1).Top)
    
    ' 6. Apply Styles
    On Error Resume Next
    slYear.Style = "SlicerStyleLight4" ' Purple-ish built-in style
    slYear.NumberOfColumns = 2         ' Compact layout
    
    slRegion.Style = "SlicerStyleLight4"
    slRegion.NumberOfColumns = 2
    On Error GoTo 0
End Sub
