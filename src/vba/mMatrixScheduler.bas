Attribute VB_Name = "mMatrixScheduler"
Option Explicit

' ==============================================================================
' Module: mMatrixScheduler
' Description: 30-Day Employee Scheduling Matrix & Event Mapping Engine.
'              Demonstrates high-performance In-Memory Matrix transformation,
'              O(1) Loc 2-way lookups, and presentation styling using cDataFrame
'              and cTableStyler.
' Author: Korn (AI Backend Developer, Technodevmoo)
' Architecture: Kit (Tech Lead)
' ==============================================================================

' ==============================================================================
' Interactive Button Handlers (Assigned to Worksheet Form Controls)
' ==============================================================================

Public Sub Btn_GenerateMatrix()
    Dim wsEmp As Worksheet
    Dim wsMatrix As Worksheet
    Dim tbEmp As ListObject
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Set wsEmp = ThisWorkbook.Sheets("Employees")
    Set wsMatrix = ThisWorkbook.Sheets("ScheduleMatrix")
    Set tbEmp = wsEmp.ListObjects("tbEmployees")
    
    ' Generate 30-day matrix starting at cell B5 with SpacerRows between teams
    GenerateMatrixFromTable tbEmp, wsMatrix.Range("B5"), 30, "Team", "Name", True
    
    Application.ScreenUpdating = True
    MsgBox "Schedule Matrix successfully generated in RAM and dumped to sheet!", vbInformation, "Matrix Generator"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error generating matrix: " & Err.Description, vbCritical, "Generator Error"
End Sub

Public Sub Btn_MapEvents()
    Dim wsEvents As Worksheet
    Dim wsMatrix As Worksheet
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

' Headless initializers (Zero-argument wrappers for automated build scripts)
Public Sub GenerateDefaultMatrix()
    Dim wsEmp As Worksheet, wsMatrix As Worksheet
    Dim tbEmp As ListObject
    Set wsEmp = ThisWorkbook.Sheets("Employees")
    Set wsMatrix = ThisWorkbook.Sheets("ScheduleMatrix")
    Set tbEmp = wsEmp.ListObjects("tbEmployees")
    GenerateMatrixFromTable tbEmp, wsMatrix.Range("B5"), 30, "Team", "Name", True
End Sub

Public Sub MapDefaultEvents()
    Dim wsEvents As Worksheet, wsMatrix As Worksheet
    Dim tbEvents As ListObject
    Set wsEvents = ThisWorkbook.Sheets("Events")
    Set wsMatrix = ThisWorkbook.Sheets("ScheduleMatrix")
    Set tbEvents = wsEvents.ListObjects("tbEvents")
    MapEventsToMatrix tbEvents, wsMatrix.Range("B5"), "Person", "Day", "Status", "CONCAT"
    ApplyStatusFormatting wsMatrix.Range("B5")
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
    
    ' 1. Clear target area
    clearEndRow = targetWs.Cells(targetWs.Rows.Count, StartCell.Column).End(xlUp).Row
    clearEndCol = targetWs.Cells(StartCell.Row, targetWs.Columns.Count).End(xlToLeft).Column
    If clearEndRow < StartCell.Row + 50 Then clearEndRow = StartCell.Row + 50
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
    
    ' 6. Apply Header & Layout Styling via cTableStyler
    FormatMatrixLayout StartCell, dfMatrix
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
                Case "OVERWRITE"
                    dfMatrix.Loc(person, dateKey) = statusVal
                Case "CONCAT"
                    If Len(Trim$(CStr(existingVal))) > 0 Then
                        dfMatrix.Loc(person, dateKey) = CStr(existingVal) & "/" & CStr(statusVal)
                    Else
                        dfMatrix.Loc(person, dateKey) = statusVal
                    End If
                Case "KEEP_FIRST"
                    If Len(Trim$(CStr(existingVal))) = 0 Then
                        dfMatrix.Loc(person, dateKey) = statusVal
                    End If
                Case Else
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
' Styling & Presentation Helpers (Powered by cTableStyler)
' ==============================================================================

Public Sub FormatMatrixLayout(ByVal StartCell As Range, ByVal dfMatrix As cDataFrame)
    Dim styler As New cTableStyler
    styler.Attach(StartCell, dfMatrix, HasHeaders:=True) _
          .ApplyTheme("NAVY") _
          .SetFont("Segoe UI", DataSize:=9.5, HeaderSize:=10#) _
          .SetHeaderRowHeight(28#) _
          .SetDataRowHeight(20#) _
          .SetSectionRowHeight("Name", "[ TEAM: * ]", Height:=24#, Bold:=True, _
                               BgColor:=RGB(225, 235, 245), FontColor:=RGB(20, 45, 75)) _
          .SetColumnWidth("Name", 24#) _
          .SetUniformColumnWidth("[0-9]*", Padding:=1.2#, MinWidth:=5.5) _
          .AlignColumn("[0-9]*", -4108) _
          .ApplyBorders()
End Sub

Public Sub ApplyStatusFormatting(ByVal StartCell As Range)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim matrixRange As Range
    Dim dfMatrix As New cDataFrame
    
    Set ws = StartCell.Worksheet
    lastRow = ws.Cells(ws.Rows.Count, StartCell.Column).End(xlUp).Row
    lastCol = ws.Cells(StartCell.Row, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow <= StartCell.Row Or lastCol <= StartCell.Column Then Exit Sub
    
    Set matrixRange = ws.Range(StartCell, ws.Cells(lastRow, lastCol))
    dfMatrix.LoadFromRange matrixRange
    
    Dim styler As New cTableStyler
    styler.Attach(StartCell, dfMatrix, HasHeaders:=True) _
          .HighlightValues("[0-9]*", "LEAVE", RGB(255, 224, 224), RGB(153, 27, 27), Bold:=True) _
          .HighlightValues("[0-9]*", "WFH", RGB(224, 242, 254), RGB(3, 105, 161), Bold:=True) _
          .HighlightValues("[0-9]*", "OT", RGB(243, 232, 255), RGB(107, 33, 168), Bold:=True) _
          .HighlightValues("[0-9]*", "OFF", RGB(241, 245, 249), RGB(71, 85, 105))
End Sub
