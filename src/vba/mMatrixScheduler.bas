Attribute VB_Name = "mMatrixScheduler"
Option Explicit

' ==============================================================================
' Module: mMatrixScheduler
' Description: High-performance 2D Matrix Schedule Generator and Cell Mapper
'              powered by the cDataFrame in-memory columnar engine.
' Features:
'   1. Generates 2D Matrix (Rows: Grouped Teams & Persons, Cols: Day 1..N)
'   2. Maps Event/Shift transactions into Matrix cell intersections via O(1) Loc
'   3. Multi-conflict handling (CONCAT, OVERWRITE, SUM)
'   4. Beautiful presentation styling with team banners & status color tags
' ==============================================================================

Public Sub Btn_GenerateMatrix()
    Dim wsEmp As Worksheet, wsMatrix As Worksheet
    Dim tbEmp As ListObject
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Set wsEmp = ThisWorkbook.Sheets("Employees")
    Set wsMatrix = ThisWorkbook.Sheets("ScheduleMatrix")
    Set tbEmp = wsEmp.ListObjects("tbEmployees")
    
    GenerateMatrixFromTable tbEmp, wsMatrix.Range("B5"), 30, "Team", "Name", True
    
    Application.ScreenUpdating = True
    MsgBox "Matrix Schedule generated successfully for 30 days!", vbInformation, "Matrix Generator"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error generating matrix: " & Err.Description, vbCritical, "Generator Error"
End Sub

Public Sub Btn_MapEvents()
    Dim wsEvents As Worksheet, wsMatrix As Worksheet
    Dim tbEvents As ListObject
    Dim matchCount As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Set wsEvents = ThisWorkbook.Sheets("Events")
    Set wsMatrix = ThisWorkbook.Sheets("ScheduleMatrix")
    Set tbEvents = wsEvents.ListObjects("tbEvents")
    
    matchCount = MapEventsToMatrix(tbEvents, wsMatrix.Range("B5"), "Person", "Day", "Status", "CONCAT")
    
    ' Apply color formatting to status cells
    ApplyStatusFormatting wsMatrix.Range("B5")
    
    Application.ScreenUpdating = True
    MsgBox "Successfully mapped " & matchCount & " event entries into Matrix!", vbInformation, "Event Mapper"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error mapping events: " & Err.Description, vbCritical, "Mapper Error"
End Sub

Public Sub Btn_ClearMatrix()
    Dim wsMatrix As Worksheet
    Dim startCell As Range
    Dim lastRow As Long, lastCol As Long
    
    Set wsMatrix = ThisWorkbook.Sheets("ScheduleMatrix")
    Set startCell = wsMatrix.Range("B5")
    
    lastRow = wsMatrix.Cells(wsMatrix.Rows.Count, startCell.Column).End(xlUp).Row
    lastCol = wsMatrix.Cells(startCell.Row, wsMatrix.Columns.Count).End(xlToLeft).Column
    
    If lastRow >= startCell.Row And lastCol >= startCell.Column Then
        With wsMatrix.Range(startCell, wsMatrix.Cells(lastRow, lastCol))
            .ClearContents
            .ClearFormats
        End With
    End If
End Sub

Public Sub Btn_RunAll()
    Btn_GenerateMatrix
    Btn_MapEvents
End Sub

' ==============================================================================
' Core Engine Procedures
' ==============================================================================

Public Sub GenerateMatrixFromTable(ByVal SourceTable As ListObject, _
                                  ByVal StartCell As Range, _
                                  ByVal NumDays As Long, _
                                  Optional ByVal GroupCol As String = "Team", _
                                  Optional ByVal NameCol As String = "Name", _
                                  Optional ByVal SpacerRows As Boolean = True)
    Dim dfSource As New cDataFrame
    Dim dfMatrix As New cDataFrame
    Dim r As Long, d As Long
    Dim currentGroup As String, lastGroup As String, empName As String
    Dim targetWs As Worksheet
    Dim clearEndRow As Long, clearEndCol As Long
    
    Set targetWs = StartCell.Worksheet
    
    ' 1. Clear existing matrix range
    clearEndRow = targetWs.Cells(targetWs.Rows.Count, StartCell.Column).End(xlUp).Row
    clearEndCol = targetWs.Cells(StartCell.Row, targetWs.Columns.Count).End(xlToLeft).Column
    If clearEndRow < StartCell.Row Then clearEndRow = StartCell.Row + 200
    If clearEndCol < StartCell.Column + NumDays Then clearEndCol = StartCell.Column + NumDays
    
    targetWs.Range(StartCell, targetWs.Cells(clearEndRow, clearEndCol)).ClearContents
    targetWs.Range(StartCell, targetWs.Cells(clearEndRow, clearEndCol)).ClearFormats
    
    ' 2. Ingest Source Table & Sort in RAM
    dfSource.LoadFromListObject SourceTable
    dfSource.Sort Array(GroupCol, NameCol), Array(True, True)
    
    ' 3. Construct Matrix Columns: Name + Day 1..NumDays
    dfMatrix.AddColumn "Name"
    For d = 1 To NumDays
        dfMatrix.AddColumn CStr(d)
    Next d
    
    ' 4. Build Group Hierarchy with Section Banners
    lastGroup = ""
    For r = 1 To dfSource.RowCount
        currentGroup = Trim$(CStr(dfSource(r, GroupCol)))
        empName = Trim$(CStr(dfSource(r, NameCol)))
        
        If currentGroup <> lastGroup Then
            If SpacerRows And Len(lastGroup) > 0 Then
                dfMatrix.AddRow ""  ' Empty spacer row between teams
            End If
            dfMatrix.AddRow "[ TEAM: " & UCase$(currentGroup) & " ]"
            lastGroup = currentGroup
        End If
        
        dfMatrix.AddRow empName
    Next r
    
    ' 5. Dump Matrix to Target Sheet in a Single COM Call
    dfMatrix.ToRange StartCell
    
    ' 6. Apply Header & Banner Styling
    FormatMatrixLayout StartCell, dfMatrix.RowCount, dfMatrix.ColumnCount
End Sub

Public Function MapEventsToMatrix(ByVal EventsTable As ListObject, _
                                 ByVal StartCell As Range, _
                                 Optional ByVal PersonCol As String = "Person", _
                                 Optional ByVal DateCol As String = "Day", _
                                 Optional ByVal StatusCol As String = "Status", _
                                 Optional ByVal ConflictMode As String = "CONCAT") As Long
    Dim dfMatrix As New cDataFrame
    Dim dfEvents As New cDataFrame
    Dim targetWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim matrixRange As Range
    Dim r As Long, matchCount As Long
    Dim person As String, dateKey As String
    Dim statusVal As Variant, existingVal As Variant
    Dim rawDate As Variant
    
    Set targetWs = StartCell.Worksheet
    
    ' 1. Determine dynamic boundary of generated matrix
    lastRow = targetWs.Cells(targetWs.Rows.Count, StartCell.Column).End(xlUp).Row
    lastCol = targetWs.Cells(StartCell.Row, targetWs.Columns.Count).End(xlToLeft).Column
    
    If lastRow < StartCell.Row Or lastCol <= StartCell.Column Then
        Err.Raise vbObjectError + 513, "mMatrixScheduler", "Matrix not found. Please generate matrix first."
    End If
    
    Set matrixRange = targetWs.Range(StartCell, targetWs.Cells(lastRow, lastCol))
    dfMatrix.LoadFromRange matrixRange
    
    ' 2. Build 2-Way Hash Index on 'Name' column (skips blank spacer rows automatically)
    dfMatrix.SetIndex "Name", SkipEmptyKeys:=True
    
    ' 3. Ingest Events Table
    dfEvents.LoadFromListObject EventsTable
    
    ' 4. In-Memory Matrix Lookup and Mutation
    matchCount = 0
    For r = 1 To dfEvents.RowCount
        person = Trim$(CStr(dfEvents(r, PersonCol)))
        rawDate = dfEvents(r, DateCol)
        
        ' Normalize dateKey: handles Day number, Date serial, or Date string
        If IsDate(rawDate) Then
            dateKey = CStr(Day(CDate(rawDate)))
        Else
            dateKey = Trim$(CStr(rawDate))
        End If
        
        statusVal = dfEvents(r, StatusCol)
        
        If dfMatrix.HasIndexKey(person) And dfMatrix.HasColumn(dateKey) Then
            existingVal = dfMatrix.Loc(person, dateKey)
            
            Select Case UCase$(ConflictMode)
                Case "CONCAT"
                    If IsEmpty(existingVal) Or Len(Trim$(CStr(existingVal))) = 0 Then
                        dfMatrix.Loc(person, dateKey) = statusVal
                    ElseIf InStr(1, CStr(existingVal), CStr(statusVal), vbTextCompare) = 0 Then
                        dfMatrix.Loc(person, dateKey) = existingVal & " / " & statusVal
                    End If
                Case "SUM"
                    If IsNumeric(existingVal) And IsNumeric(statusVal) Then
                        dfMatrix.Loc(person, dateKey) = CDbl(existingVal) + CDbl(statusVal)
                    Else
                        dfMatrix.Loc(person, dateKey) = statusVal
                    End If
                Case Else ' OVERWRITE
                    dfMatrix.Loc(person, dateKey) = statusVal
            End Select
            
            matchCount = matchCount + 1
        End If
    Next r
    
    ' 5. Write enriched matrix back to worksheet in Single COM Call
    dfMatrix.ToRange StartCell
    
    MapEventsToMatrix = matchCount
End Function

' ==============================================================================
' Styling & Presentation Helpers
' ==============================================================================

Private Sub FormatMatrixLayout(ByVal StartCell As Range, ByVal RowCount As Long, ByVal ColCount As Long)
    Dim ws As Worksheet
    Dim headerRange As Range
    Dim fullRange As Range
    Dim r As Long, rowNum As Long
    Dim cellVal As String
    
    Set ws = StartCell.Worksheet
    Set headerRange = ws.Range(StartCell, ws.Cells(StartCell.Row, StartCell.Column + ColCount - 1))
    Set fullRange = ws.Range(StartCell, ws.Cells(StartCell.Row + RowCount - 1, StartCell.Column + ColCount - 1))
    
    ' Header Style (Deep Navy Blue)
    With headerRange
        .Interior.Color = RGB(31, 58, 82)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 26
    End With
    
    ' Grid Lines
    fullRange.Borders.LineStyle = xlContinuous
    fullRange.Borders.Color = RGB(218, 224, 233)
    fullRange.Borders.Weight = xlThin
    
    ' Highlight Team Section Banners
    For r = 2 To RowCount
        rowNum = StartCell.Row + r - 1
        cellVal = Trim$(CStr(ws.Cells(rowNum, StartCell.Column).Value))
        
        If Left$(cellVal, 8) = "[ TEAM: " Then
            With ws.Range(ws.Cells(rowNum, StartCell.Column), ws.Cells(rowNum, StartCell.Column + ColCount - 1))
                .Interior.Color = RGB(225, 235, 245)
                .Font.Color = RGB(20, 45, 75)
                .Font.Bold = True
            End With
        ElseIf Len(cellVal) = 0 Then
            ' Empty spacer row
            ws.Range(ws.Cells(rowNum, StartCell.Column), ws.Cells(rowNum, StartCell.Column + ColCount - 1)).Interior.Color = RGB(250, 250, 250)
            ws.Rows(rowNum).RowHeight = 8
        Else
            ws.Cells(rowNum, StartCell.Column).Font.Bold = False
            ws.Cells(rowNum, StartCell.Column).HorizontalAlignment = xlLeft
        End If
    Next r
    
    ' Center align day columns
    ws.Range(ws.Cells(StartCell.Row + 1, StartCell.Column + 1), _
             ws.Cells(StartCell.Row + RowCount - 1, StartCell.Column + ColCount - 1)).HorizontalAlignment = xlCenter
    
    ' Column Widths
    ws.Columns(StartCell.Column).ColumnWidth = 26
    ws.Range(ws.Columns(StartCell.Column + 1), ws.Columns(StartCell.Column + ColCount - 1)).ColumnWidth = 5.5
End Sub

Private Sub ApplyStatusFormatting(ByVal StartCell As Range)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim cell As Range
    Dim valStr As String
    
    Set ws = StartCell.Worksheet
    lastRow = ws.Cells(ws.Rows.Count, StartCell.Column).End(xlUp).Row
    lastCol = ws.Cells(StartCell.Row, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow <= StartCell.Row Or lastCol <= StartCell.Column Then Exit Sub
    
    Set dataRange = ws.Range(ws.Cells(StartCell.Row + 1, StartCell.Column + 1), ws.Cells(lastRow, lastCol))
    
    For Each cell In dataRange
        valStr = UCase$(Trim$(CStr(cell.Value)))
        If Len(valStr) > 0 Then
            If InStr(1, valStr, "LEAVE") > 0 Then
                cell.Interior.Color = RGB(255, 224, 224) ' Light Coral / Red
                cell.Font.Color = RGB(153, 27, 27)
                cell.Font.Bold = True
            ElseIf InStr(1, valStr, "WFH") > 0 Then
                cell.Interior.Color = RGB(224, 242, 254) ' Light Sky Blue
                cell.Font.Color = RGB(3, 105, 161)
                cell.Font.Bold = True
            ElseIf InStr(1, valStr, "OT") > 0 Then
                cell.Interior.Color = RGB(243, 232, 255) ' Light Purple
                cell.Font.Color = RGB(107, 33, 168)
                cell.Font.Bold = True
            ElseIf InStr(1, valStr, "OFF") > 0 Then
                cell.Interior.Color = RGB(241, 245, 249) ' Slate Gray
                cell.Font.Color = RGB(71, 85, 105)
            Else
                cell.Interior.Color = RGB(220, 252, 231) ' Light Green
                cell.Font.Color = RGB(22, 101, 52)
            End If
        End If
    Next cell
End Sub
