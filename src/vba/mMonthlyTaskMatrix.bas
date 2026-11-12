Attribute VB_Name = "mMonthlyTaskMatrix"
Option Explicit

' ==============================================================================
' Module: mMonthlyTaskMatrix
' Description: Unified Single-Sheet Monthly Task Completion Matrix Orchestrator
'              Powered natively by cDataFrame (TryParseExcelDate, BuildCalendarMatrix),
'              cReportBuilder (AddSectionHeader, AddQuickJumpBar, AddActionButton),
'              and cTableStyler (ApplyWeeklyCalendarGrid, SpotlightColumn).
' ==============================================================================

Private Const MATRIX_SHEET_NAME As String = "TaskCompletionMatrix"
Private Const START_YEAR As Long = 2026
Private Const MAX_DAYS_COLS As Long = 31
Private Const TOTAL_MATRIX_COLS As Long = 35 ' Assignee + Dept + D01..D31 + Month Total + Daily Avg

Public gMatrixDirty As Boolean
Private mIsRefreshing As Boolean
Private mLastRefreshDate As Date

' Called by Workbook_SheetChange on RawTasks / Employees (0ms dirty flag)
Public Sub MarkMatrixDirty()
    gMatrixDirty = True
End Sub

' Called by Workbook_Open and Workbook_SheetActivate on TaskCompletionMatrix
Public Sub AutoRefreshIfNeeded(Optional ByVal ForceRefresh As Boolean = False)
    If mIsRefreshing Then Exit Sub
    If Not Application.EnableEvents Then Exit Sub

    Dim sheetMissing As Boolean
    sheetMissing = (FindSheet(ThisWorkbook, MATRIX_SHEET_NAME) Is Nothing)

    If ForceRefresh Or gMatrixDirty Or sheetMissing Or (mLastRefreshDate <> Date) Then
        GenerateCombinedMonthlyMatrix True, False
    End If
End Sub

Public Sub Btn_RefreshMatrix()
    GenerateCombinedMonthlyMatrix False, False
End Sub

Public Sub Btn_GenerateAll()
    GenerateCombinedMonthlyMatrix False, False
End Sub

Public Sub Btn_JumpToCurrentMonth()
    AutoRefreshIfNeeded False
    Dim ws As Worksheet
    Set ws = FindSheet(ThisWorkbook, MATRIX_SHEET_NAME)
    If ws Is Nothing Then
        GenerateCombinedMonthlyMatrix True, True
        Exit Sub
    End If

    ws.Activate
    Dim curM As Long
    curM = Month(Date)
    If curM < 1 Or curM > 12 Then curM = 1

    On Error Resume Next
    Application.GoTo ThisWorkbook.Names("JumpMonth_" & Format$(curM, "00")).RefersToRange, True
    On Error GoTo 0
End Sub

Public Sub Btn_ClearSheets()
    ClearGeneratedSheets False
End Sub

Public Sub GenerateAllMonthlyReports(Optional ByVal Silent As Boolean = False)
    GenerateCombinedMonthlyMatrix Silent, False
End Sub

Public Sub GenerateCombinedMonthlyMatrix(Optional ByVal Silent As Boolean = False, _
                                         Optional ByVal JumpToCurrent As Boolean = False)
    If mIsRefreshing Then Exit Sub
    mIsRefreshing = True

    Dim t0 As Double
    t0 = Timer

    Dim prevScreen As Boolean, prevCalc As Long, prevEvents As Boolean, prevAlerts As Boolean
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.Calculation = -4135 ' xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim refDate As Date
    refDate = Date

    ' 1. Load Employees & Completed Tasks once into cDataFrame
    Dim dfEmployees As New cDataFrame
    dfEmployees.LoadFromRange wb.Sheets("Employees").ListObjects("tbEmployees").Range

    Dim dfAllTasks As New cDataFrame
    dfAllTasks.LoadFromRange wb.Sheets("RawTasks").ListObjects("tbTasks").Range

    Dim dfCompleted As cDataFrame
    Set dfCompleted = dfAllTasks.Filter("Status", "=", "Completed")

    ' 2. Compute top-level YTD / Today / MTD KPI metrics via cDataFrame.TryParseExcelDate
    Dim colDate As cColumn
    Set colDate = dfCompleted.Column("TaskDate")
    Dim ytdCompleted As Long, todayCompleted As Long, mtdCompleted As Long
    Dim curY As Long, curM As Long, r As Long, dt As Date
    curY = Year(refDate)
    curM = Month(refDate)

    For r = 1 To dfCompleted.RowCount
        If dfCompleted.TryParseExcelDate(colDate.Item(r), dt) Then
            If Year(dt) = START_YEAR Then
                If dt <= refDate Then ytdCompleted = ytdCompleted + 1
                If dt = refDate Then todayCompleted = todayCompleted + 1
                If Year(dt) = curY And Month(dt) = curM Then mtdCompleted = mtdCompleted + 1
            End If
        End If
    Next r

    DeleteLegacySplitSheets wb

    ' 3. Prepare single unified output sheet
    Dim wsMatrix As Worksheet
    Set wsMatrix = FindSheet(wb, MATRIX_SHEET_NAME)
    If wsMatrix Is Nothing Then
        Set wsMatrix = wb.Sheets.Add(Before:=wb.Sheets(1))
        wsMatrix.Name = MATRIX_SHEET_NAME
    Else
        wsMatrix.Cells.Clear
        wsMatrix.Hyperlinks.Delete
        If wsMatrix.Index <> 1 Then wsMatrix.Move Before:=wb.Sheets(1)
    End If
    wsMatrix.Tab.Color = RGB(16, 185, 129)

    ' 4. Compose Executive Banner + Daily KPI Row via cReportBuilder
    Dim rb As New cReportBuilder
    rb.Attach wsMatrix, "B2"
    rb.BeginUpdate

    rb.AddBanner "2026 Master Task Completion Matrix - All 12 Months (Daily Detail)", _
                 "Unified Single-Sheet View | Reference Date: " & Format$(refDate, "yyyy-mm-dd") & _
                 " | Last Updated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss"), _
                 "Emerald", TOTAL_MATRIX_COLS
    rb.AddSpace 1

    rb.AddKpiRow Array("Completed (YTD)", "Today (" & Format$(refDate, "dd mmm") & ")", _
                       "Current Month (" & Format$(refDate, "mmm") & ")", "Active Members"), _
                 Array(ytdCompleted, todayCompleted, mtdCompleted, dfEmployees.RowCount), _
                 Empty, 4, Array("#,##0", "#,##0", "#,##0", "0")
    rb.AddSpace 1

    Dim jumpBarRow As Long
    jumpBarRow = rb.CurrentRow
    rb.AddSpace 2

    ' 5. Stack all 12 monthly calendar matrices using cDataFrame + cReportBuilder + cTableStyler
    Dim pillLabels(0 To 11) As String
    Dim pillTargets(0 To 11) As String
    Dim m As Long, secRow As Long

    For m = 1 To 12
        secRow = RenderMonthBlock(rb, dfCompleted, dfEmployees, START_YEAR, m, refDate)
        pillLabels(m - 1) = MonthName(m, True) & " '" & Right$(CStr(START_YEAR), 2)
        pillTargets(m - 1) = "'" & MATRIX_SHEET_NAME & "'!B" & CStr(secRow)
        If m < 12 Then rb.AddSpace 1
    Next m

    ' 6. Populate same-sheet Quick-Jump navigation pills & Refresh button via cReportBuilder
    rb.AddQuickJumpBar "Quick Jump to Month:", pillLabels, pillTargets, curM, 2, jumpBarRow
    rb.AddActionButton "btnRefreshMatrix", "Refresh Today's Matrix", "Btn_RefreshMatrix", "AD", jumpBarRow, 175#, 25#

    rb.AddSpace 1
    rb.AddProvenanceFooter "cReportBuilder Unified Matrix Engine", _
                           "Auto-Refreshes Daily on Open & RawTasks Edits (Ref: " & Format$(refDate, "yyyy-mm-dd") & ")", _
                           "", TOTAL_MATRIX_COLS

    ' 7. Consistent column widths across all 12 stacked blocks
    wsMatrix.Columns("B").ColumnWidth = 22
    wsMatrix.Columns("C").ColumnWidth = 16
    Dim dayColIdx As Long
    For dayColIdx = 4 To 34
        wsMatrix.Columns(dayColIdx).ColumnWidth = 5.5
    Next dayColIdx
    wsMatrix.Columns("AI").ColumnWidth = 12
    wsMatrix.Columns("AJ").ColumnWidth = 10

    rb.EndUpdate

    gMatrixDirty = False
    mLastRefreshDate = refDate

    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    mIsRefreshing = False

    Dim elapsed As Double
    elapsed = Timer - t0

    On Error Resume Next
    wsMatrix.Activate
    If JumpToCurrent Then
        Application.GoTo wb.Names("JumpMonth_" & Format$(curM, "00")).RefersToRange, True
    Else
        wsMatrix.Range("B2").Select
    End If
    On Error GoTo 0

    If Not Silent Then
        Dim isVis As Boolean
        On Error Resume Next
        isVis = Application.Visible
        On Error GoTo 0
        If isVis Then
            MsgBox "Updated unified 12-month Task Completion Matrix as of " & _
                   Format$(refDate, "yyyy-mm-dd") & " in " & Format$(elapsed, "0.00") & "s!", _
                   vbInformation, "Daily Matrix Updated"
        End If
    End If
    Exit Sub

ErrorHandler:
    mIsRefreshing = False
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    Err.Raise Err.Number, "mMonthlyTaskMatrix.GenerateCombinedMonthlyMatrix", Err.Description
End Sub

Public Sub ClearGeneratedSheets(Optional ByVal Silent As Boolean = False)
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    DeleteLegacySplitSheets ThisWorkbook
    Dim wsMatrix As Worksheet
    Set wsMatrix = FindSheet(ThisWorkbook, MATRIX_SHEET_NAME)
    If Not wsMatrix Is Nothing Then wsMatrix.Delete

    Application.DisplayAlerts = prevAlerts

    On Error Resume Next
    ThisWorkbook.Sheets("ControlPanel").Activate
    On Error GoTo 0

    If Not Silent Then
        MsgBox "Generated matrix sheet cleared successfully.", vbInformation, "Sheet Cleared"
    End If
End Sub

Private Function RenderMonthBlock(ByVal rb As cReportBuilder, _
                                  ByVal dfCompleted As cDataFrame, _
                                  ByVal dfEmployees As cDataFrame, _
                                  ByVal YearVal As Long, _
                                  ByVal MonthVal As Long, _
                                  ByVal refDate As Date) As Long
    Dim dfMatrix As cDataFrame
    Set dfMatrix = dfCompleted.BuildCalendarMatrix("TaskDate", "Assignee", YearVal, MonthVal, _
                                                   Array("Department"), dfEmployees, "Name", MAX_DAYS_COLS)

    Dim colEmp As cColumn, colTot As cColumn
    Set colEmp = dfMatrix.Column("Assignee")
    Set colTot = dfMatrix.Column("Month Total")

    Dim totalMonthCompleted As Long, topPerson As String, topCount As Long
    Dim r As Long, rowTot As Long
    topPerson = "None"
    For r = 1 To dfMatrix.RowCount
        rowTot = CLng(colTot.Item(r))
        totalMonthCompleted = totalMonthCompleted + rowTot
        If rowTot > topCount Then
            topCount = rowTot
            topPerson = CStr(colEmp.Item(r))
        End If
    Next r

    Dim daysInMonth As Long
    daysInMonth = Day(DateSerial(YearVal, MonthVal + 1, 0))
    Dim teamDailyAvg As Double
    teamDailyAvg = Round(CDbl(totalMonthCompleted) / CDbl(daysInMonth), 1)
    Dim topPerformerText As String
    topPerformerText = IIf(topCount > 0, topPerson & " (" & CStr(topCount) & ")", "None")

    Dim isCurrentMonth As Boolean, statusBadge As String, sectionColor As Long
    isCurrentMonth = (YearVal = Year(refDate) And MonthVal = Month(refDate))
    If YearVal < Year(refDate) Or (YearVal = Year(refDate) And MonthVal < Month(refDate)) Then
        statusBadge = "[CLOSED - ACTUAL]"
        sectionColor = RGB(30, 41, 59)
    ElseIf isCurrentMonth Then
        statusBadge = "[CURRENT MONTH - LIVE TODAY: " & Format$(refDate, "dd mmm yyyy") & "]"
        sectionColor = RGB(6, 95, 70)
    Else
        statusBadge = "[UPCOMING / SCHEDULED]"
        sectionColor = RGB(71, 85, 105)
    End If

    Dim secTitle As String
    secTitle = "  " & UCase$(MonthName(MonthVal)) & " " & CStr(YearVal) & _
               "   |   Completed: " & Format$(totalMonthCompleted, "#,##0") & _
               "   |   Daily Avg: " & Format$(teamDailyAvg, "0.0") & _
               "   |   Top: " & topPerformerText & _
               "   |   " & statusBadge

    Dim secRow As Long
    secRow = rb.CurrentRow
    rb.AddSectionHeader secTitle, "Slate", TOTAL_MATRIX_COLS, "JumpMonth_" & Format$(MonthVal, "00"), sectionColor
    rb.AddSpace 1

    dfMatrix.Sort Array("Department", "Assignee"), Array(True, True)
    rb.AddTable dfMatrix, "", "Slate", True

    rb.LastTableStyler.ApplyWeeklyCalendarGrid "D01", daysInMonth, MAX_DAYS_COLS, "Month Total", "Daily Avg"
    If isCurrentMonth Then
        rb.LastTableStyler.SpotlightColumn "D" & Format$(Day(refDate), "00")
    End If

    RenderMonthBlock = secRow
End Function

Private Function FindSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    Set FindSheet = ws
End Function

Private Sub DeleteLegacySplitSheets(ByVal wb As Workbook)
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Dim i As Long
    For i = wb.Sheets.Count To 1 Step -1
        Dim ws As Worksheet
        Set ws = wb.Sheets(i)
        If ws.Name Like "####-##" Or ws.Name = "TOC" Then
            ws.Delete
        End If
    Next i

    Application.DisplayAlerts = prevAlerts
End Sub
