Attribute VB_Name = "mDemoWorkflow"
Option Explicit

' ==============================================================================
' Module: mDemoWorkflow
' Description: Demonstrates real-world end-to-end workflows using cDataFrame:
'              Step 1: Pulls tbEmployees, groups by Team, builds Day 1..30 matrix in Sheet A.
'              Step 2: Pulls Sheet B records, maps into Sheet A via Loc(Person, Day) in RAM.
'              Step 3: Multi-Branch Sales ETL, Deduplication, Missing Data Cleaning,
'                      Left Hash Join, Vectorized Column Pricing, GroupBy & Executive Pivot.
' ==============================================================================

Public Sub RunStep1()
    Step1_GenerateMatrixFromTable False
End Sub

Public Sub RunStep2()
    Step2_MapSheetBToSheetA False
End Sub

Public Sub RunStep3()
    Step3_RunSalesAnalyticsPipeline False
End Sub

Public Sub RunStep3Silent()
    Step3_RunSalesAnalyticsPipeline True
End Sub

Public Sub TestRunAll()
    Step1_GenerateMatrixFromTable True
    Step2_MapSheetBToSheetA True
    Step3_RunSalesAnalyticsPipeline True
End Sub

Public Sub Step1_GenerateMatrixFromTable(Optional ByVal Silent As Boolean = False)
    Dim dfSource As New cDataFrame
    Dim dfMatrix As New cDataFrame
    Dim r As Long, d As Long
    Dim currentTeam As String, lastTeam As String, empName As String
    Dim wsA As Worksheet
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' 1. Load employee master list from Excel Table tbEmployees
    dfSource.LoadFromRange ThisWorkbook.Sheets("Employees").ListObjects("tbEmployees").Range
    
    ' 2. Sort by Team, then by Name in RAM
    dfSource.Sort Array("Team", "Name"), Array(True, True)
    
    ' 3. Create Matrix columns: Name + Day 1..30
    dfMatrix.AddColumn "Name"
    For d = 1 To 30
        dfMatrix.AddColumn CStr(d)
    Next d
    
    ' 4. Group rows by Team (inserting section header and empty spacer rows)
    lastTeam = ""
    For r = 1 To dfSource.RowCount
        currentTeam = Trim$(CStr(dfSource(r, "Team")))
        empName = Trim$(CStr(dfSource(r, "Name")))
        
        If currentTeam <> lastTeam Then
            If Len(lastTeam) > 0 Then dfMatrix.AddRow ""
            dfMatrix.AddRow "[ TEAM: " & currentTeam & " ]"
            lastTeam = currentTeam
        End If
        
        dfMatrix.AddRow empName
    Next r
    
    ' 5. Dump Matrix to Sheet A starting at A5
    Set wsA = ThisWorkbook.Sheets("A")
    wsA.Range("A5:AF100").ClearContents
    wsA.Range("A5:AF100").ClearFormats
    
    dfMatrix.ToRange wsA.Range("A5")
    
    ' Format Header Row
    With wsA.Range("A5:AF5")
        .Interior.Color = RGB(41, 128, 185)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    wsA.Columns("A:AF").AutoFit
    wsA.Columns("A").ColumnWidth = 28
    
    Application.ScreenUpdating = True
    Dim isVisible1 As Boolean
    On Error Resume Next
    isVisible1 = Application.Visible
    On Error GoTo 0
    If (Not Silent) And isVisible1 Then
        MsgBox "Step 1 Completed: Generated 30-Day Matrix grouped by team!", vbInformation, "Step 1 Success"
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Silent Or Not Application.Visible Then
        Err.Raise Err.Number, "mDemoWorkflow.Step1", Err.Description
    Else
        MsgBox "Error: " & Err.Description, vbCritical, "Step 1 Error"
    End If
End Sub

Public Sub Step2_MapSheetBToSheetA(Optional ByVal Silent As Boolean = False)
    Dim dfA As New cDataFrame
    Dim dfB As New cDataFrame
    Dim wsA As Worksheet, wsB As Worksheet
    Dim r As Long, matchCount As Long
    Dim person As String, dayKey As String, statusVal As Variant
    Dim matrixRange As Range
    Dim lastRow As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Set wsA = ThisWorkbook.Sheets("A")
    Set wsB = ThisWorkbook.Sheets("B")
    
    ' Check if Matrix table exists on Sheet A
    If Len(Trim$(wsA.Range("A5").Value)) = 0 Then
        If Not Silent Then
            MsgBox "Please click 'Step 1' to create the Matrix table first.", vbExclamation
        End If
        Exit Sub
    End If
    
    ' 1. Load full Matrix range from Sheet A (spanning past empty spacer rows)
    lastRow = wsA.Cells(wsA.Rows.Count, "A").End(xlUp).Row
    If lastRow < 5 Then lastRow = 5
    Set matrixRange = wsA.Range(wsA.Cells(5, 1), wsA.Cells(lastRow, 31))
    
    dfA.LoadFromRange matrixRange
    
    ' 2. Set 'Name' as 2-Way Matrix Index (skips blank spacer rows automatically)
    dfA.SetIndex "Name", True
    
    ' 3. Load transaction records from Sheet B
    dfB.LoadFromRange wsB.Range("A1").CurrentRegion
    
    ' 4. Iterate Sheet B and map into Sheet A in RAM via O(1) Matrix Loc
    matchCount = 0
    For r = 1 To dfB.RowCount
        person = Trim$(CStr(dfB(r, "Person")))
        dayKey = Trim$(CStr(dfB(r, "Day")))
        statusVal = dfB(r, "Status")
        
        If dfA.HasIndexKey(person) And dfA.HasColumn(dayKey) Then
            dfA.Loc(person, dayKey) = statusVal
            matchCount = matchCount + 1
        End If
    Next r
    
    ' 5. Dump updated matrix back to Sheet A in a single COM write
    dfA.ToRange wsA.Range("A5")
    
    Application.ScreenUpdating = True
    Dim isVisible2 As Boolean
    On Error Resume Next
    isVisible2 = Application.Visible
    On Error GoTo 0
    If (Not Silent) And isVisible2 Then
        MsgBox "Step 2 Completed: Successfully mapped " & matchCount & " entries from Sheet B!", vbInformation, "Step 2 Success"
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Silent Or Not Application.Visible Then
        Err.Raise Err.Number, "mDemoWorkflow.Step2", Err.Description
    Else
        MsgBox "Error: " & Err.Description, vbCritical, "Step 2 Error"
    End If
End Sub

' ==============================================================================
' Step 3: End-to-End Multi-Branch Sales ETL, Vector Pricing & Executive Pivot
' ==============================================================================

Public Sub Step3_RunSalesAnalyticsPipeline(Optional ByVal Silent As Boolean = False)
    On Error GoTo ErrorHandler

    Dim dfEnriched As cDataFrame
    Dim dfSummary As cDataFrame
    Dim dfPivot As cDataFrame
    Dim csvPath As String

    csvPath = Environ$("TEMP") & "\vba_sales_executive_summary.csv"
    RunSalesPipelineInMemory dfEnriched, dfSummary, dfPivot, csvPath

    ' Render output tables to sheet "Analytics_Demo" if running inside Excel
    Dim xlApp As Object
    On Error Resume Next
    Set xlApp = Application
    On Error GoTo ErrorHandler

    If Not xlApp Is Nothing Then
        If Not xlApp.ActiveWorkbook Is Nothing Then
            xlApp.ScreenUpdating = False

            Dim ws As Object
            On Error Resume Next
            Set ws = xlApp.ActiveWorkbook.Worksheets("Analytics_Demo")
            On Error GoTo ErrorHandler

            If ws Is Nothing Then
                Set ws = xlApp.ActiveWorkbook.Worksheets.Add
                ws.Name = "Analytics_Demo"
            Else
                ws.Cells.Clear
                Do While ws.ListObjects.Count > 0
                    ws.ListObjects(1).Delete
                Loop
                On Error Resume Next
                ws.DrawingObjects.Delete
                On Error GoTo ErrorHandler
            End If

            ' Dashboard Header Banner
            With ws.Range("A1:M1")
                .Merge
                .Value = "EXECUTIVE SALES ANALYTICS (Powered by cDataFrame & cTableStyler)"
                .Font.Name = "Segoe UI"
                .Font.Size = 13
                .Font.Bold = True
                .Interior.Color = RGB(31, 78, 120)
                .Font.Color = RGB(255, 255, 255)
                .HorizontalAlignment = -4108
                .VerticalAlignment = -4108
            End With
            ws.Rows(1).RowHeight = 28

            ws.Range("A2").Value = "End-to-End Pipeline: Diagonal Concat, DropDuplicates, DropNA/FillNA, Hash Merge, Vector Math, GroupBy & Pivot."
            ws.Range("A2").Font.Italic = True
            ws.Range("A2").Font.Size = 9.5
            ws.Range("A2").Font.Color = RGB(110, 110, 110)

            ' Interactive Run Button
            Dim btn As Object
            On Error Resume Next
            Set btn = ws.Buttons.Add(ws.Range("K2").Left, ws.Range("K2").Top, 160, 24)
            If Not btn Is Nothing Then
                btn.Caption = "Re-Run Analytics Pipeline"
                btn.OnAction = "Step3_RunSalesAnalyticsPipeline"
            End If
            On Error GoTo ErrorHandler

            ' 1. Write Enriched Orders as structured Excel Table (ListObject) at A4
            ws.Range("A4").Value = "1. Clean & Enriched Orders (Vectorized Pricing & Catalog Merge)"
            ws.Range("A4").Font.Bold = True
            ws.Range("A4").Font.Size = 11
            dfEnriched.ToListObject ws.Range("A5"), "tblEnrichedOrders"

            ' 2. Write Category x Channel GroupBy Summary at A14
            ws.Range("A14").Value = "2. Category & Sales Channel Performance (O(N) GroupBy)"
            ws.Range("A14").Font.Bold = True
            ws.Range("A14").Font.Size = 11
            dfSummary.ToRange ws.Range("A15"), True

            Dim stylerSummary As New cTableStyler
            stylerSummary.Attach ws.Range("A15"), dfSummary, True
            stylerSummary.ApplyTheme "NAVY"
            stylerSummary.FormatColumn "NetRevenue_SUM", "$#,##0.00"
            stylerSummary.FormatColumn "NetRevenue_AVG", "$#,##0.00"
            stylerSummary.FormatColumn "OrderID_COUNT", "#,##0"
            stylerSummary.ApplyBorders
            stylerSummary.ApplyZebraStripes
            stylerSummary.AutoFitColumns

            ' 3. Write Category x Quarter Executive Pivot Table at H14
            ws.Range("H14").Value = "3. Quarterly Revenue Matrix (Executive Pivot Table)"
            ws.Range("H14").Font.Bold = True
            ws.Range("H14").Font.Size = 11
            dfPivot.ToRange ws.Range("H15"), True

            Dim stylerPivot As New cTableStyler
            stylerPivot.Attach ws.Range("H15"), dfPivot, True
            stylerPivot.ApplyTheme "EMERALD"
            stylerPivot.FormatColumn "Q1", "$#,##0.00"
            stylerPivot.FormatColumn "Q2", "$#,##0.00"
            stylerPivot.ApplyBorders
            stylerPivot.ApplyZebraStripes
            stylerPivot.AutoFitColumns

            ws.Columns("A:M").AutoFit
            xlApp.ScreenUpdating = True
        End If
    End If

    Dim isVisible3 As Boolean
    On Error Resume Next
    isVisible3 = Application.Visible
    On Error GoTo 0
    If (Not Silent) And isVisible3 Then
        MsgBox "Step 3 Sales Analytics Pipeline Complete!" & vbCrLf & _
               "Clean Orders: " & dfEnriched.RowCount & vbCrLf & _
               "Total Net Revenue: " & Format$(dfEnriched.Column("NetRevenue").Sum(), "#,##0.00") & vbCrLf & _
               "Hardware Q1 Revenue (via O(1) Pivot.Loc): " & Format$(dfPivot.Loc("Hardware", "Q1"), "#,##0.00"), _
               vbInformation, "Step 3 Analytics Pipeline"
    End If
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    If Silent Or Not Application.Visible Then
        Err.Raise Err.Number, "mDemoWorkflow.Step3", Err.Description
    Else
        MsgBox "Error: " & Err.Description, vbCritical, "Step 3 Error"
    End If
End Sub

' Pure in-memory pipeline executing Concat, DropDuplicates, DropNA, FillNA,
' LEFT Hash Merge, Vectorized Column Arithmetic (Multiply/Subtract/Unique),
' Multi-Column GroupBy + Sort, Cross-Tabulation Pivot + O(1) Loc, and ToCSV.
Public Sub RunSalesPipelineInMemory(ByRef OutEnriched As cDataFrame, _
                                    ByRef OutSummary As cDataFrame, _
                                    ByRef OutPivot As cDataFrame, _
                                    Optional ByVal CsvExportPath As String = "")
    ' --------------------------------------------------------------------------
    ' 1. Multi-Branch Raw Data Ingestion & Diagonal Concat (UNION ALL)
    ' --------------------------------------------------------------------------
    Dim branchBkk As New cDataFrame
    branchBkk.AddColumn "OrderID"
    branchBkk.AddColumn "Quarter"
    branchBkk.AddColumn "SKU"
    branchBkk.AddColumn "Qty"
    branchBkk.AddColumn "DiscountRate"

    branchBkk.AddRow "ORD-101", "Q1", "SKU-LAP", 2#, 0.05  ' Superseded by re-transmission below
    branchBkk.AddRow "ORD-101", "Q1", "SKU-LAP", 4#, 0.1   ' Latest revision of ORD-101
    branchBkk.AddRow "ORD-102", "Q1", "SKU-MON", 5#, Empty ' Missing DiscountRate -> FillNA(0)
    branchBkk.AddRow "ORD-103", "Q2", "", 2#, 0#           ' Corrupt blank SKU -> DropNA
    branchBkk.AddRow "ORD-104", "Q2", "SKU-KEY", 10#, 0.05

    Dim branchCnx As New cDataFrame
    branchCnx.AddColumn "OrderID"
    branchCnx.AddColumn "Quarter"
    branchCnx.AddColumn "SKU"
    branchCnx.AddColumn "Qty"
    branchCnx.AddColumn "DiscountRate"
    branchCnx.AddColumn "SalesChannel"                     ' Extra column not in branchBkk!

    branchCnx.AddRow "ORD-201", "Q1", "SKU-MON", 3#, 0.1, "Online"
    branchCnx.AddRow "ORD-202", "Q2", "SKU-LAP", 2#, 0#, "Corporate"
    branchCnx.AddRow "ORD-203", "Q2", "SKU-LIC", 8#, Empty, "Online"
    branchCnx.AddRow "ORD-204", "Q2", "SKU-KEY", Empty, 0#, "Walk-In" ' Corrupt missing Qty -> DropNA

    ' Combine both branches with schema superset expansion (AllowMissingColumns:=True)
    Dim dfRaw As cDataFrame
    Set dfRaw = branchBkk.Concat(branchCnx, True)

    ' --------------------------------------------------------------------------
    ' 2. Data Cleaning: Deduplicate Revisions, Drop Corrupt Rows, Impute Defaults
    ' --------------------------------------------------------------------------
    Dim dfDedup As cDataFrame
    Set dfDedup = dfRaw.DropDuplicates("OrderID", "LAST")

    Dim dfClean As cDataFrame
    Set dfClean = dfDedup.DropNA(Array("SKU", "Qty"), "ANY", True)

    dfClean.FillNA 0#, "DiscountRate", False
    dfClean.FillNA "Walk-In", "SalesChannel", True

    ' --------------------------------------------------------------------------
    ' 3. Product Catalog Hash Join (LEFT Merge) & Vectorized Pricing Arithmetic
    ' --------------------------------------------------------------------------
    Dim dfCatalog As New cDataFrame
    dfCatalog.AddColumn "SKU"
    dfCatalog.AddColumn "ProductName"
    dfCatalog.AddColumn "Category"
    dfCatalog.AddColumn "UnitPrice"

    dfCatalog.AddRow "SKU-LAP", "ProBook 14", "Hardware", 1000#
    dfCatalog.AddRow "SKU-MON", "UltraWide 27", "Hardware", 400#
    dfCatalog.AddRow "SKU-KEY", "Mech Keyboard", "Accessories", 100#
    dfCatalog.AddRow "SKU-LIC", "Cloud Suite", "Software", 250#

    Set OutEnriched = dfClean.Merge(dfCatalog, "SKU", "LEFT")

    ' Vectorized column calculations (O(N) 1D array math without VBA row loops)
    Dim colGross As cColumn
    Dim colDisc As cColumn
    Dim colNet As cColumn

    Set colGross = OutEnriched.Column("Qty").Multiply(OutEnriched.Column("UnitPrice"), "GrossAmount")
    Set colDisc = colGross.Multiply(OutEnriched.Column("DiscountRate"), "DiscountAmt")
    Set colNet = colGross.Subtract(colDisc, "NetRevenue")

    OutEnriched.AddColumn "GrossAmount", , colGross.ToArray()
    OutEnriched.AddColumn "DiscountAmt", , colDisc.ToArray()
    OutEnriched.AddColumn "NetRevenue", , colNet.ToArray()

    ' --------------------------------------------------------------------------
    ' 4. Executive GroupBy Summary (Sorted) & Cross-Tabulation Pivot Table
    ' --------------------------------------------------------------------------
    Set OutSummary = OutEnriched.GroupBy( _
        Array("Category", "SalesChannel"), _
        Array("NetRevenue", "NetRevenue", "OrderID"), _
        Array("SUM", "AVG", "COUNT"))

    OutSummary.Sort "NetRevenue_SUM", False

    Set OutPivot = OutEnriched.Pivot("Category", "Quarter", "NetRevenue", "SUM", 0#)

    ' --------------------------------------------------------------------------
    ' 5. Optional RFC 4180 CSV Export
    ' --------------------------------------------------------------------------
    If Len(Trim$(CsvExportPath)) > 0 Then
        OutSummary.ToCSV CsvExportPath, ",", True
    End If
End Sub
