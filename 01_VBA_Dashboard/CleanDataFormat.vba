VBA Code: '=============================================================
'  DATA CLEANING & FORMATTING MACRO - PURPLE THEME
'  Sheet: "Data"
'  Columns: Order ID | Order Date | Year | Month | Day |
'           Segment | Category | Sub-Category | Region |
'           Sales | Profit | Quantity | Discount
'
'  BUTTON: Assign this macro -> CleanAndFormatData
'  STYLE:  Purple theme (TableStyleMedium13)
'          Header  = Dark purple  RGB(94, 59, 183)
'          Odd rows  = Light purple RGB(235, 228, 255)
'          Even rows = White        RGB(255, 255, 255)
'=============================================================

Option Explicit

'-------------------------------------------------------------
'  MAIN - Assign this to your macro button
'-------------------------------------------------------------
Sub CleanAndFormatData()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long

    Set ws = ThisWorkbook.Sheets("Data")

    ' Find actual data boundaries
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Run each cleaning & formatting step
    Call StepRemoveBlankRowsCols(ws, lastRow, lastCol)
    Call StepTrimWhitespace(ws, lastRow, lastCol)
    Call StepStandardiseText(ws, lastRow)
    Call StepFormatNumbers(ws, lastRow)
    Call StepCentreAlignment(ws, lastRow, lastCol)
    Call StepApplyPurpleTheme(ws, lastRow, lastCol)  ' Purple theme step
    Call StepFormatHeader(ws, lastCol)
    Call StepSetColumnWidths(ws)
    Call StepFreezeHeader(ws)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Purple Theme Applied Successfully!" & vbNewLine & _
           lastRow - 1 & " records processed.", vbInformation, "Done"
End Sub


'-------------------------------------------------------------
'  STEP 1 - Remove entirely blank rows and columns
'-------------------------------------------------------------
Sub StepRemoveBlankRowsCols(ws As Worksheet, lastRow As Long, lastCol As Long)
    Dim r As Long
    Dim c As Long

    ' Delete blank rows (bottom-up to avoid row shift issues)
    For r = lastRow To 2 Step -1
        If WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r

    ' Delete blank columns (right-to-left)
    For c = lastCol To 1 Step -1
        If WorksheetFunction.CountA(ws.Columns(c)) = 0 Then
            ws.Columns(c).Delete
        End If
    Next c
End Sub


'-------------------------------------------------------------
'  STEP 2 - Trim leading/trailing spaces from all text cells
'-------------------------------------------------------------
Sub StepTrimWhitespace(ws As Worksheet, lastRow As Long, lastCol As Long)
    Dim r    As Long
    Dim c    As Long
    Dim cell As Range

    For r = 1 To lastRow
        For c = 1 To lastCol
            Set cell = ws.Cells(r, c)
            If cell.Value <> "" And VarType(cell.Value) = vbString Then
                cell.Value = Trim(cell.Value)
            End If
        Next c
    Next r
End Sub


'-------------------------------------------------------------
'  STEP 3 - Standardise text columns (Proper Case)
'  Applies to: Segment, Category, Sub-Category, Region
'-------------------------------------------------------------
Sub StepStandardiseText(ws As Worksheet, lastRow As Long)
    Dim r      As Long
    Dim col    As Variant
    Dim colNum As Integer
    Dim headers As Variant

    headers = Array("Segment", "Category", "Sub-Category", "Region")

    For Each col In headers
        colNum = FindColumn(ws, CStr(col))
        If colNum > 0 Then
            For r = 2 To lastRow
                If ws.Cells(r, colNum).Value <> "" Then
                    ws.Cells(r, colNum).Value = _
                        WorksheetFunction.Proper(ws.Cells(r, colNum).Value)
                End If
            Next r
        End If
    Next col
End Sub


'-------------------------------------------------------------
'  STEP 4 - Format number columns with correct formats
'  Sales, Profit  -> #,##0.00
'  Quantity       -> #,##0
'  Discount       -> 0%
'  Year, Day      -> 0 (plain integer)
'  Order Date     -> yyyy-mm-dd (convert serials first)
'-------------------------------------------------------------
Sub StepFormatNumbers(ws As Worksheet, lastRow As Long)
    Dim colSales  As Integer: colSales = FindColumn(ws, "Sales")
    Dim colProfit As Integer: colProfit = FindColumn(ws, "Profit")
    Dim colQty    As Integer: colQty = FindColumn(ws, "Quantity")
    Dim colDisc   As Integer: colDisc = FindColumn(ws, "Discount")
    Dim colYear   As Integer: colYear = FindColumn(ws, "Year")
    Dim colDay    As Integer: colDay = FindColumn(ws, "Day")
    Dim colDate   As Integer: colDate = FindColumn(ws, "Order Date")

    If colSales > 0 Then ws.Range(ws.Cells(2, colSales), ws.Cells(lastRow, colSales)).NumberFormat = "#,##0.00"
    If colProfit > 0 Then ws.Range(ws.Cells(2, colProfit), ws.Cells(lastRow, colProfit)).NumberFormat = "#,##0.00"
    If colQty > 0 Then ws.Range(ws.Cells(2, colQty), ws.Cells(lastRow, colQty)).NumberFormat = "#,##0"
    If colDisc > 0 Then ws.Range(ws.Cells(2, colDisc), ws.Cells(lastRow, colDisc)).NumberFormat = "0%"
    If colYear > 0 Then ws.Range(ws.Cells(2, colYear), ws.Cells(lastRow, colYear)).NumberFormat = "0"
    If colDay > 0 Then ws.Range(ws.Cells(2, colDay), ws.Cells(lastRow, colDay)).NumberFormat = "0"

    ' Order Date: convert Excel serial numbers to readable dates
    If colDate > 0 Then
        Dim dateCell As Range
        Dim r        As Long
        For r = 2 To lastRow
            Set dateCell = ws.Cells(r, colDate)
            If IsNumeric(dateCell.Value) And dateCell.Value > 1 Then
                dateCell.Value = CDate(dateCell.Value)
            End If
        Next r
        ws.Range(ws.Cells(2, colDate), ws.Cells(lastRow, colDate)).NumberFormat = "yyyy-mm-dd"
    End If
End Sub


'-------------------------------------------------------------
'  STEP 5 - Centre-align ALL cells (header + data)
'-------------------------------------------------------------
Sub StepCentreAlignment(ws As Worksheet, lastRow As Long, lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Uniform row height
    ws.Rows("1:" & lastRow).RowHeight = 18
End Sub


'-------------------------------------------------------------
'  STEP 6 - Apply Purple Theme (TableStyleMedium13)
'           Matches your ApplyPurpleTheme macro exactly,
'           integrated into the full cleaning workflow
'-------------------------------------------------------------
Sub StepApplyPurpleTheme(ws As Worksheet, lastRow As Long, lastCol As Long)
    Dim tbl As ListObject

    ' Detect data range automatically (same as your original macro)
    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion

    ' Unlist any existing tables (removes table format, KEEPS data intact)
    On Error Resume Next
    For Each tbl In ws.ListObjects
        tbl.Unlist   ' FIX: Unlist preserves data; Delete would erase it
    Next tbl
    On Error GoTo 0

    ' Create table (same as your original macro)
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "tblSalesData"

    ' Apply Purple Table Style (your chosen style)
    tbl.TableStyle = "TableStyleMedium13"

    ' Show filter dropdowns
    tbl.ShowAutoFilter = True

    ' ---- Override row banding with exact purple shades ----
    ' Odd rows  -> light purple tint
    ' Even rows -> white
    Dim r       As Long
    Dim oddBg   As Long: oddBg = RGB(235, 228, 255)    ' Light purple
    Dim evenBg  As Long: evenBg = RGB(255, 255, 255)   ' White

    For r = 2 To lastRow
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior
            If (r Mod 2) = 1 Then   ' Odd row  -> light purple
                .Color = oddBg
            Else                     ' Even row -> white
                .Color = evenBg
            End If
        End With
    Next r

    ' Autofit columns (same as your original macro)
    rng.Columns.AutoFit
End Sub


'-------------------------------------------------------------
'  STEP 7 - Format header row: purple theme styling
'           Dark purple background, white bold text, centred
'-------------------------------------------------------------
Sub StepFormatHeader(ws As Worksheet, lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)            ' White text
        .Font.Name = "Calibri"
        .Interior.Color = RGB(94, 59, 183)          ' Dark purple header
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With
End Sub


'-------------------------------------------------------------
'  STEP 8 - Set column widths tailored to each column
'-------------------------------------------------------------
Sub StepSetColumnWidths(ws As Worksheet)
    Dim colMap As Variant
    Dim i      As Integer
    Dim colNum As Integer

    colMap = Array( _
        "Order ID", 13, _
        "Order Date", 14, _
        "Year", 8, _
        "Month", 9, _
        "Day", 7, _
        "Segment", 14, _
        "Category", 16, _
        "Sub-Category", 16, _
        "Region", 10, _
        "Sales", 12, _
        "Profit", 12, _
        "Quantity", 10, _
        "Discount", 10)

    For i = 0 To UBound(colMap) - 1 Step 2
        colNum = FindColumn(ws, CStr(colMap(i)))
        If colNum > 0 Then
            ws.Columns(colNum).ColumnWidth = CDbl(colMap(i + 1))
        End If
    Next i
End Sub


'-------------------------------------------------------------
'  STEP 9 - Freeze the header row
'-------------------------------------------------------------
Sub StepFreezeHeader(ws As Worksheet)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub


'-------------------------------------------------------------
'  HELPER - Find column number by header name (case-insensitive)
'-------------------------------------------------------------
Function FindColumn(ws As Worksheet, headerName As String) As Integer
    Dim c       As Integer
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If LCase(Trim(ws.Cells(1, c).Value)) = LCase(Trim(headerName)) Then
            FindColumn = c
            Exit Function
        End If
    Next c

    FindColumn = 0  ' Not found
End Function


'=============================================================
'  STANDALONE: Run this alone to apply only the purple theme
'  (Your original macro - preserved as-is)
'=============================================================
Sub ApplyPurpleTheme()
    Dim ws  As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Set ws = ActiveSheet

    ' Detect data range automatically
    Set rng = ws.Range("A1").CurrentRegion

    ' Unlist existing table if already exists (keeps data, removes table format)
    On Error Resume Next
    For Each tbl In ws.ListObjects
        tbl.Unlist   ' FIX: Unlist preserves data; Delete would erase it
    Next tbl
    On Error GoTo 0

    ' Create table
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)

    ' Apply Purple Table Style
    tbl.TableStyle = "TableStyleMedium13"   ' Purple theme

    ' Optional: Center align
    rng.HorizontalAlignment = xlCenter

    ' Autofit columns
    rng.Columns.AutoFit

    MsgBox "Data Clean&Format Successfully!"
End Sub
