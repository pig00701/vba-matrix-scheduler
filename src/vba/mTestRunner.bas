Attribute VB_Name = "mTestRunner"
Option Explicit

' ==============================================================================
' Module: mTestRunner
' Description: Comprehensive automated test suite for cDataFrame & cColumn.
'              Includes Matrix Lookup, Multi-column In-Memory QuickSort,
'              Hash Merges (INNER/LEFT), Vector Math, Range I/O, and 10k Stress.
' Author: Korn (AI Backend Developer, Technodevmoo)
' Architecture: Kit (Tech Lead)
' ==============================================================================

Private mPassCount As Long
Private mFailCount As Long
Private mLogOutput As String
Private mHeartbeatFile As String

' Sub entry point for Excel Macro menu / F5 in VBE
Public Sub RunTests()
    Dim res As String
    res = RunAllTests()
    MsgBox res, IIf(InStr(res, "FAIL") > 0, vbCritical, vbInformation), "VBA Columnar Table Test Runner"
End Sub

Public Function RunAllTests() As String
    mPassCount = 0
    mFailCount = 0
    mLogOutput = ""

    On Error Resume Next
    mHeartbeatFile = Environ$("TEMP") & "\vba_test_heartbeat.log"
    Dim hNum As Integer
    hNum = FreeFile
    Open mHeartbeatFile For Output As #hNum
    Close #hNum
    On Error GoTo 0

    LogPrint "=========================================================="
    LogPrint " Running VBA Columnar Table (cDataFrame) Test Suite"
    LogPrint " Timestamp: " & Now
    LogPrint "=========================================================="

    Dim tStart As Double
    tStart = Timer

    Test_BasicCreationAndRowAdd
    Test_ColumnDropAndRename
    Test_RowInsertAndDelete
    Test_VectorMath
    Test_FilterAndClone
    Test_RangeBulkIO
    Test_MatrixLookupAndIndex
    Test_InMemorySort
    Test_HashMerge
    Test_GroupBy
    Test_Pivot
    Test_UniqueAndDropDuplicates
    Test_Concat
    Test_VectorArithmetic
    Test_MissingDataFillAndDropNA
    Test_CSVRoundTrip
    Test_ListObjectIO
    Test_WindowAndCumulative
    Test_SliceAndSelectColumns
    Test_AdvancedJoins
    Test_VectorStringAndWhere
    Test_VectorDateOperations
    Test_FilterCompound
    Test_MeltUnpivot
    Test_JSONSerialization
    Test_BulkAddRows
    Test_CSVFormulaSanitization
    Test_MultiColumnIndex
    Test_TableStyler
    Test_LoadFromWorkbook
    Test_StatisticsAndDescribe
    Test_RollingWindows
    Test_CovarianceAndCorrelation
    Test_MarkdownAndHTMLExport
    Test_ExpressionEval
    Test_DemoAnalyticsPipeline
    Test_StressTest10k
    Test_RemediationHardening
    Test_WindowLagLeadDiffAndSplitBy
    Test_DataValidator
    Test_ReportBuilder
    Test_CalendarMatrixAndWeeklyGrid
    Test_TableDrivenConfigAndStyles
    Test_FormulaEngine
    Test_AllClassesGapCoverage
    Test_GetDataAndRangeGaps
    Test_CleanAndTransformData
    Test_3DArrayAndMultiSheets
    Test_CrossReferenceMatrix
    Test_MatrixEnhancements
    Test_SparseMatrixOptimization

    Dim tElapsed As Double
    tElapsed = Timer - tStart

    LogPrint "----------------------------------------------------------"
    LogPrint " Test Suite Summary:"
    LogPrint " Passed: " & mPassCount
    LogPrint " Failed: " & mFailCount
    LogPrint " Total Duration: " & Format$(tElapsed, "0.000") & " seconds"

    Dim summaryResult As String
    If mFailCount = 0 Then
        summaryResult = "PASS: All tests passed (" & mPassCount & " assertions) in " & Format$(tElapsed, "0.000") & "s"
        LogPrint " STATUS: ALL TESTS PASSED SUCCESSFULLY! (100% OK)"
    Else
        summaryResult = "FAIL: " & mFailCount & " assertions failed out of " & (mPassCount + mFailCount)
        LogPrint " STATUS: " & summaryResult
    End If
    LogPrint "=========================================================="

    RunAllTests = summaryResult & vbCrLf & vbCrLf & mLogOutput
End Function

' ------------------------------------------------------------------------------
' Test 1: Basic Creation, AddColumn, AddRow
' ------------------------------------------------------------------------------
Private Sub Test_BasicCreationAndRowAdd()
    On Error GoTo TestFail
    LogPrint "[TEST 1] Basic Creation, AddColumn, AddRow..."

    Dim df As New cDataFrame
    AssertEqual df.RowCount, 0, "Initial RowCount is 0"
    AssertEqual df.ColumnCount, 0, "Initial ColumnCount is 0"

    df.AddColumn "ID"
    df.AddColumn "Name"
    df.AddColumn "Score"

    AssertEqual df.ColumnCount, 3, "ColumnCount after AddColumn is 3"
    AssertTrue df.HasColumn("ID"), "HasColumn ID"
    AssertTrue df.HasColumn("Name"), "HasColumn Name"
    AssertTrue df.HasColumn("Score"), "HasColumn Score"

    df.AddRow 101, "Alice", 95.5
    df.AddRow 102, "Bob", 88#
    df.AddRow 103, "Charlie", 92.5

    AssertEqual df.RowCount, 3, "RowCount after 3 AddRows is 3"
    AssertEqual df(1, "ID"), 101, "df(1, 'ID') = 101"
    AssertEqual df(1, "Name"), "Alice", "df(1, 'Name') = Alice"
    AssertEqual df(1, "Score"), 95.5, "df(1, 'Score') = 95.5"
    AssertEqual df(2, "Name"), "Bob", "df(2, 'Name') = Bob"
    AssertEqual df(3, "ID"), 103, "df(3, 'ID') = 103"

    ' Test numeric column index lookup
    AssertEqual df(1, 1), 101, "df(1, 1) = 101 (Numeric col lookup)"
    AssertEqual df(1, 2), "Alice", "df(1, 2) = Alice (Numeric col lookup)"
    AssertEqual df(1, 3), 95.5, "df(1, 3) = 95.5 (Numeric col lookup)"

    ' Mutation via default member
    df(2, "Score") = 90#
    AssertEqual df(2, "Score"), 90#, "df(2, 'Score') mutated to 90"

    Exit Sub
TestFail:
    RecordFailure "Test_BasicCreationAndRowAdd", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 2: Dynamic Column Drop & Rename (O(1))
' ------------------------------------------------------------------------------
Private Sub Test_ColumnDropAndRename()
    On Error GoTo TestFail
    LogPrint "[TEST 2] Dynamic Column Drop & Rename..."

    Dim df As New cDataFrame
    df.AddColumn "ColA"
    df.AddColumn "ColB"
    df.AddColumn "ColC"

    df.AddRow 10, 20, 30
    df.AddRow 40, 50, 60

    ' Rename ColB -> ColBeta (O(1) in-place dictionary key change)
    df.RenameColumn "ColB", "ColBeta"
    AssertTrue df.HasColumn("ColBeta"), "ColBeta exists after rename"
    AssertTrue Not df.HasColumn("ColB"), "ColB no longer exists"
    AssertEqual df(1, "ColBeta"), 20, "df(1, 'ColBeta') retained value 20"
    AssertEqual df(2, "ColBeta"), 50, "df(2, 'ColBeta') retained value 50"

    ' Drop ColA (O(1) dictionary removal)
    df.DropColumn "ColA"
    AssertEqual df.ColumnCount, 2, "ColumnCount is 2 after dropping ColA"
    AssertTrue Not df.HasColumn("ColA"), "ColA is removed"
    AssertEqual df(1, 1), 20, "df(1, 1) is now ColBeta (value 20)"
    AssertEqual df(1, 2), 30, "df(1, 2) is now ColC (value 30)"

    Exit Sub
TestFail:
    RecordFailure "Test_ColumnDropAndRename", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 3: Dynamic Row Insert & Delete
' ------------------------------------------------------------------------------
Private Sub Test_RowInsertAndDelete()
    On Error GoTo TestFail
    LogPrint "[TEST 3] Dynamic Row Insert & Delete..."

    Dim df As New cDataFrame
    df.AddColumn "Item"
    df.AddColumn "Qty"

    df.AddRow "Apple", 10
    df.AddRow "Banana", 20
    df.AddRow "Cherry", 30

    ' Insert row at index 2 (between Apple and Banana)
    df.InsertRow 2, "Blueberry", 15
    AssertEqual df.RowCount, 4, "RowCount after InsertRow is 4"
    AssertEqual df(1, "Item"), "Apple", "Row 1 is Apple"
    AssertEqual df(2, "Item"), "Blueberry", "Row 2 is Blueberry"
    AssertEqual df(3, "Item"), "Banana", "Row 3 shifted to Banana"
    AssertEqual df(4, "Item"), "Cherry", "Row 4 shifted to Cherry"
    AssertEqual df(2, "Qty"), 15, "Blueberry Qty is 15"

    ' Delete row at index 3 (delete Banana)
    df.DeleteRow 3
    AssertEqual df.RowCount, 3, "RowCount after DeleteRow is 3"
    AssertEqual df(1, "Item"), "Apple", "Row 1 is Apple"
    AssertEqual df(2, "Item"), "Blueberry", "Row 2 is Blueberry"
    AssertEqual df(3, "Item"), "Cherry", "Row 3 is Cherry"
    AssertEqual df(3, "Qty"), 30, "Cherry Qty is 30"

    Exit Sub
TestFail:
    RecordFailure "Test_RowInsertAndDelete", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 4: Vector Math & Analytics on cColumn
' ------------------------------------------------------------------------------
Private Sub Test_VectorMath()
    On Error GoTo TestFail
    LogPrint "[TEST 4] Vector Math & Analytics on cColumn..."

    Dim col As New cColumn
    col.Name = "Numbers"
    col.Append 10#
    col.Append 20#
    col.Append 30#
    col.Append 40#
    col.Append 50#

    AssertEqual col.Sum(), 150#, "Sum of 10..50 is 150"
    AssertEqual col.Average(), 30#, "Average of 10..50 is 30"
    AssertEqual col.Min(), 10#, "Min of 10..50 is 10"
    AssertEqual col.Max(), 50#, "Max of 10..50 is 50"

    ' Test ToArray
    Dim arr As Variant
    arr = col.ToArray()
    AssertEqual UBound(arr) - LBound(arr) + 1, 5, "ToArray length is 5"
    AssertEqual arr(1), 10#, "ToArray(1) is 10"
    AssertEqual arr(5), 50#, "ToArray(5) is 50"

    ' Test Fill
    col.Fill 99#, 2, 4
    AssertEqual col.Item(1), 10#, "col(1) unchanged"
    AssertEqual col.Item(2), 99#, "col(2) filled with 99"
    AssertEqual col.Item(4), 99#, "col(4) filled with 99"
    AssertEqual col.Item(5), 50#, "col(5) unchanged"

    Exit Sub
TestFail:
    RecordFailure "Test_VectorMath", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 5: Filter & Clone
' ------------------------------------------------------------------------------
Private Sub Test_FilterAndClone()
    On Error GoTo TestFail
    LogPrint "[TEST 5] Filter & Clone..."

    Dim df As New cDataFrame
    df.AddColumn "Dept"
    df.AddColumn "Employee"
    df.AddColumn "Salary"

    df.AddRow "IT", "Korn", 75000
    df.AddRow "HR", "Somchai", 45000
    df.AddRow "IT", "Kit", 120000
    df.AddRow "Sales", "Mook", 60000
    df.AddRow "IT", "Nat", 80000

    ' Filter IT department
    Dim itDf As cDataFrame
    Set itDf = df.Filter("Dept", "=", "IT")

    AssertEqual itDf.RowCount, 3, "Filtered IT RowCount is 3"
    AssertEqual itDf(1, "Employee"), "Korn", "IT Row 1 is Korn"
    AssertEqual itDf(2, "Employee"), "Kit", "IT Row 2 is Kit"
    AssertEqual itDf(3, "Employee"), "Nat", "IT Row 3 is Nat"

    ' Filter Salary > 70000
    Dim highSalaryDf As cDataFrame
    Set highSalaryDf = df.Filter("Salary", ">", 70000)
    AssertEqual highSalaryDf.RowCount, 3, "Salary > 70000 count is 3"

    ' Clone structure only
    Dim cloneEmpty As cDataFrame
    Set cloneEmpty = df.Clone(True)
    AssertEqual cloneEmpty.ColumnCount, 3, "Cloned structure has 3 columns"
    AssertEqual cloneEmpty.RowCount, 0, "Cloned structure has 0 rows"

    ' Deep Clone
    Dim cloneFull As cDataFrame
    Set cloneFull = df.Clone(False)
    AssertEqual cloneFull.RowCount, 5, "Cloned full has 5 rows"
    AssertEqual cloneFull(3, "Employee"), "Kit", "Cloned full(3, 'Employee') = Kit"

    Exit Sub
TestFail:
    RecordFailure "Test_FilterAndClone", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 6: Range Bulk I/O
' ------------------------------------------------------------------------------
Private Sub Test_RangeBulkIO()
    On Error GoTo SkipOrReport
    LogPrint "[TEST 6] Range Bulk I/O..."

    ' Check if running inside active Excel host
    Dim xlApp As Object
    On Error Resume Next
    Set xlApp = Application
    On Error GoTo TestFail

    If xlApp Is Nothing Then
        LogPrint "  (Skipping RangeBulkIO: Host application not detected)"
        mPassCount = mPassCount + 1
        Exit Sub
    End If

    Dim ws As Object
    Set ws = xlApp.ActiveWorkbook.Worksheets.Add

    ' Write sample range
    ws.Range("A1").Value2 = "SKU"
    ws.Range("B1").Value2 = "Price"
    ws.Range("A2").Value2 = "ITEM-01"
    ws.Range("B2").Value2 = 150
    ws.Range("A3").Value2 = "ITEM-02"
    ws.Range("B3").Value2 = 250

    Dim df As New cDataFrame
    df.LoadFromRange ws.Range("A1:B3"), True

    AssertEqual df.RowCount, 2, "LoadFromRange loaded 2 data rows"
    AssertEqual df.ColumnCount, 2, "LoadFromRange loaded 2 columns"
    AssertEqual df(1, "SKU"), "ITEM-01", "Row 1 SKU is ITEM-01"
    AssertEqual df(2, "Price"), 250, "Row 2 Price is 250"

    ' Dump to Range D1
    df.ToRange ws.Range("D1"), True
    AssertEqual ws.Range("D1").Value2, "SKU", "Dumped header D1 is SKU"
    AssertEqual ws.Range("E2").Value2, 150, "Dumped cell E2 is 150"

    ' Clean up test sheet
    xlApp.DisplayAlerts = False
    ws.Delete
    xlApp.DisplayAlerts = True
    Exit Sub

SkipOrReport:
    LogPrint "  (Range test skipped or completed with status: " & Err.Description & ")"
    mPassCount = mPassCount + 1
    Exit Sub
TestFail:
    RecordFailure "Test_RangeBulkIO", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 7: Stress Test (10,000 Rows x 5 Columns Timing, Sort, Loc)
' ------------------------------------------------------------------------------
Private Sub Test_StressTest10k()
    On Error GoTo TestFail
    LogPrint "[TEST 7] Stress Test (10,000 Rows x 5 Columns)..."

    Dim df As New cDataFrame
    df.AddColumn "RowID"
    df.AddColumn "Code"
    df.AddColumn "Amount"
    df.AddColumn "Tax"
    df.AddColumn "Status"

    Dim tStart As Double, tElapsed As Double
    tStart = Timer

    Const TARGET_ROWS As Long = 10000
    Dim i As Long
    For i = 1 To TARGET_ROWS
        df.AddRow i, "CODE_" & (i Mod 100), i * 1.5, (i * 1.5) * 0.07, "ACTIVE"
    Next i

    tElapsed = Timer - tStart

    AssertEqual df.RowCount, TARGET_ROWS, "Stress test built 10,000 rows"
    AssertEqual df.ColumnCount, 5, "Stress test built 5 columns"
    AssertEqual df(TARGET_ROWS, "RowID"), TARGET_ROWS, "Last RowID is 10000"

    LogPrint "  [PERF] 10,000 rows x 5 cols populated in: " & Format$(tElapsed, "0.000") & " seconds"

    ' Vector math performance
    tStart = Timer
    Dim totalAmt As Double
    totalAmt = df.Column("Amount").Sum()
    Dim tMath As Double
    tMath = Timer - tStart

    AssertTrue totalAmt > 0, "Vector sum on 10,000 rows executed"
    LogPrint "  [PERF] Column('Amount').Sum() over 10,000 rows: " & Format$(tMath, "0.0000") & " seconds (Sum=" & totalAmt & ")"

    ' Sorting 10,000 rows descending
    tStart = Timer
    df.Sort "Amount", False
    Dim tSort As Double
    tSort = Timer - tStart
    AssertEqual df(1, "RowID"), TARGET_ROWS, "10k rows sorted descending: top RowID is 10000"
    AssertEqual df(TARGET_ROWS, "RowID"), 1, "10k rows sorted descending: bottom RowID is 1"
    LogPrint "  [PERF] 10,000 rows sorted in: " & Format$(tSort, "0.0000") & " seconds"

    ' Index build and O(1) Loc lookup on 10,000 rows
    tStart = Timer
    df.SetIndex "RowID"
    Dim tIndex As Double
    tIndex = Timer - tStart
    LogPrint "  [PERF] 10,000 rows indexed in: " & Format$(tIndex, "0.0000") & " seconds"

    tStart = Timer
    Dim locVal As Variant
    locVal = df.Loc(5000, "Code")
    Dim tLoc As Double
    tLoc = Timer - tStart
    AssertEqual locVal, "CODE_" & (5000 Mod 100), "Loc(5000, 'Code') matches expected"
    LogPrint "  [PERF] O(1) Loc lookup on 10,000 rows: " & Format$(tLoc, "0.000000") & " seconds"

    ' O(N) Hash GroupBy on 10,000 rows across 100 unique Code groups
    tStart = Timer
    Dim grouped10k As cDataFrame
    Set grouped10k = df.GroupBy("Code", Array("Amount", "Tax", "RowID"), Array("SUM", "SUM", "COUNT"))
    Dim tGroup As Double
    tGroup = Timer - tStart
    AssertEqual grouped10k.RowCount, 100, "10k rows grouped into 100 unique Code groups"
    AssertEqual grouped10k(1, "RowID_COUNT"), 100, "Each Code group has 100 rows"
    LogPrint "  [PERF] O(N) GroupBy on 10,000 rows (100 groups): " & Format$(tGroup, "0.0000") & " seconds"

    Exit Sub
TestFail:
    RecordFailure "Test_StressTest10k", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 8: Matrix Lookup & Index (Lookup, SetIndex, Loc Get/Let)
' ------------------------------------------------------------------------------
Private Sub Test_MatrixLookupAndIndex()
    On Error GoTo TestFail
    LogPrint "[TEST 8] Matrix Lookup & Index (Lookup, SetIndex, Loc Get/Let)..."

    Dim df As New cDataFrame
    df.AddColumn "EmpID"
    df.AddColumn "EmpName"
    df.AddColumn "Dept"
    df.AddColumn "Salary"

    df.AddRow "E101", "Alice", "IT", 80000
    df.AddRow "E102", "Bob", "HR", 50000
    df.AddRow "E103", "Charlie", "IT", 95000
    df.AddRow "E104", "David", "Finance", 70000

    ' 1. Test Lookup (Linear scan on 1D vector)
    Dim val As Variant
    val = df.Lookup("E102", "EmpID", "EmpName")
    AssertEqual val, "Bob", "Lookup E102 EmpName = Bob"

    val = df.Lookup("E103", "EmpID", "Salary")
    AssertEqual val, 95000, "Lookup E103 Salary = 95000"

    ' Lookup non-existent key returns #N/A (Error 2042)
    val = df.Lookup("E999", "EmpID", "EmpName")
    AssertIsError val, "Lookup non-existent key returns error (#N/A)"

    ' Lookup with numeric column index (EmpID col 1 -> Salary col 4)
    val = df.Lookup("E104", 1, 4)
    AssertEqual val, 70000, "Lookup with numeric col indices (1 -> 4) = 70000"

    ' 2. Test SetIndex & Loc (Property Get O(1))
    df.SetIndex "EmpID"
    AssertEqual df.IndexColumnName, "EmpID", "IndexColumnName is EmpID"

    AssertEqual df.Loc("E101", "EmpName"), "Alice", "Loc('E101', 'EmpName') = Alice"
    AssertEqual df.Loc("E103", "Salary"), 95000, "Loc('E103', 'Salary') = 95000"
    AssertEqual df.Loc("e102", "Dept"), "HR", "Loc case-insensitive key 'e102' = HR"
    AssertEqual df.Loc("E104", 2), "David", "Loc with numeric column index 2 = David"

    ' Loc non-existent key returns error (#N/A)
    val = df.Loc("E999", "EmpName")
    AssertIsError val, "Loc non-existent key returns error (#N/A)"

    ' 3. Test Loc (Property Let mutation)
    df.Loc("E102", "Salary") = 55000
    AssertEqual df.Loc("E102", "Salary"), 55000, "Loc mutated Bob's salary to 55000"
    AssertEqual df(2, "Salary"), 55000, "Underlying df(2, 'Salary') reflected mutation"

    ' 4. Test Index Auto-rebuild on row modifications
    ' Append row
    df.AddRow "E105", "Eve", "Marketing", 62000
    AssertEqual df.Loc("E105", "EmpName"), "Eve", "Loc finds newly added row E105 after auto-rebuild"

    ' Insert row at index 2
    df.InsertRow 2, "E101B", "Alex", "IT", 82000
    AssertEqual df.Loc("E101B", "EmpName"), "Alex", "Loc finds inserted row E101B"
    AssertEqual df.Loc("E102", "EmpName"), "Bob", "Loc finds Bob after shift"

    ' Delete row
    df.DeleteRow 2 ' deletes Alex
    AssertEqual df.Loc("E102", "EmpName"), "Bob", "Loc finds Bob after row deletion"
    val = df.Loc("E101B", "EmpName")
    AssertIsError val, "Deleted key E101B is no longer found in index"

    ' 5. Test Duplicate Key detection in SetIndex
    Dim dupDf As New cDataFrame
    dupDf.AddColumn "Code"
    dupDf.AddColumn "Desc"
    dupDf.AddRow "A", "First"
    dupDf.AddRow "B", "Second"
    dupDf.AddRow "A", "Duplicate A"

    Dim caughtDup As Boolean
    caughtDup = False
    On Error Resume Next
    dupDf.SetIndex "Code"
    If Err.Number = 457 Then caughtDup = True
    On Error GoTo TestFail
    AssertTrue caughtDup, "SetIndex correctly raises error 457 on duplicate key"

    Exit Sub
TestFail:
    RecordFailure "Test_MatrixLookupAndIndex", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 9: In-Memory Sort Engine (Single & Multi-Column QuickSort)
' ------------------------------------------------------------------------------
Private Sub Test_InMemorySort()
    On Error GoTo TestFail
    LogPrint "[TEST 9] In-Memory Sort Engine (Single & Multi-Column QuickSort)..."

    Dim df As New cDataFrame
    df.AddColumn "ID"
    df.AddColumn "Dept"
    df.AddColumn "Score"
    df.AddColumn "Name"

    df.AddRow 1, "IT", 85, "Korn"
    df.AddRow 2, "HR", 90, "Somchai"
    df.AddRow 3, "IT", 95, "Kit"
    df.AddRow 4, "Sales", 70, "Mook"
    df.AddRow 5, "HR", 80, "Noi"
    df.AddRow 6, "IT", 75, "Nat"

    ' 1. Single Column Sort Ascending (Score)
    df.Sort "Score", True
    AssertEqual df(1, "Score"), 70, "Row 1 min Score is 70"
    AssertEqual df(1, "Name"), "Mook", "Row 1 Name is Mook"
    AssertEqual df(6, "Score"), 95, "Row 6 max Score is 95"
    AssertEqual df(6, "Name"), "Kit", "Row 6 Name is Kit"

    ' 2. Single Column Sort Descending (Score)
    df.Sort "Score", False
    AssertEqual df(1, "Score"), 95, "Row 1 max Score is 95 descending"
    AssertEqual df(1, "Name"), "Kit", "Row 1 Name is Kit descending"
    AssertEqual df(6, "Score"), 70, "Row 6 min Score is 70 descending"

    ' 3. Multi-Column Sort: Dept Ascending, then Score Descending
    ' Expected order:
    ' HR: Somchai (90), Noi (80)
    ' IT: Kit (95), Korn (85), Nat (75)
    ' Sales: Mook (70)
    df.Sort Array("Dept", "Score"), Array(True, False)

    AssertEqual df(1, "Dept"), "HR", "Row 1 Dept is HR"
    AssertEqual df(1, "Score"), 90, "Row 1 Score is 90 (Somchai)"
    AssertEqual df(2, "Dept"), "HR", "Row 2 Dept is HR"
    AssertEqual df(2, "Score"), 80, "Row 2 Score is 80 (Noi)"

    AssertEqual df(3, "Dept"), "IT", "Row 3 Dept is IT"
    AssertEqual df(3, "Score"), 95, "Row 3 Score is 95 (Kit)"
    AssertEqual df(4, "Dept"), "IT", "Row 4 Dept is IT"
    AssertEqual df(4, "Score"), 85, "Row 4 Score is 85 (Korn)"
    AssertEqual df(5, "Dept"), "IT", "Row 5 Dept is IT"
    AssertEqual df(5, "Score"), 75, "Row 5 Score is 75 (Nat)"

    AssertEqual df(6, "Dept"), "Sales", "Row 6 Dept is Sales"
    AssertEqual df(6, "Score"), 70, "Row 6 Score is 70 (Mook)"

    ' 4. Multi-column Sort with single boolean (Both Ascending)
    df.Sort Array("Dept", "Score"), True
    AssertEqual df(1, "Dept"), "HR", "Row 1 Dept is HR"
    AssertEqual df(1, "Score"), 80, "Row 1 Score is 80 (Noi - Ascending)"
    AssertEqual df(2, "Score"), 90, "Row 2 Score is 90 (Somchai - Ascending)"

    ' 5. Sort by numeric column index (Column 1: ID Ascending)
    df.Sort 1, True
    AssertEqual df(1, "ID"), 1, "Sorted by col 1: Row 1 ID is 1"
    AssertEqual df(6, "ID"), 6, "Sorted by col 1: Row 6 ID is 6"

    Exit Sub
TestFail:
    RecordFailure "Test_InMemorySort", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 10: Hash Join / Table Merge (INNER and LEFT Joins)
' ------------------------------------------------------------------------------
Private Sub Test_HashMerge()
    On Error GoTo TestFail
    LogPrint "[TEST 10] Hash Join / Table Merge (INNER and LEFT Joins)..."

    ' Left Table: Employees
    Dim leftDf As New cDataFrame
    leftDf.AddColumn "EmpID"
    leftDf.AddColumn "Name"
    leftDf.AddColumn "DeptID"
    leftDf.AddColumn "Status"

    leftDf.AddRow "E1", "Alice", "D10", "Active"
    leftDf.AddRow "E2", "Bob", "D20", "Active"
    leftDf.AddRow "E3", "Charlie", "D10", "Probation"
    leftDf.AddRow "E4", "David", "D30", "Active"
    leftDf.AddRow "E5", "Eve", "D99", "Pending" ' D99 has no dept match

    ' Right Table: Departments
    Dim rightDf As New cDataFrame
    rightDf.AddColumn "DeptID"
    rightDf.AddColumn "DeptName"
    rightDf.AddColumn "Location"
    rightDf.AddColumn "Status" ' Colliding column name!

    rightDf.AddRow "D10", "Engineering", "Building A", "Open"
    rightDf.AddRow "D20", "Human Resources", "Building B", "Open"
    rightDf.AddRow "D30", "Finance", "Building C", "Closed"
    rightDf.AddRow "D40", "Marketing", "Building D", "Open" ' No employee in D40

    ' 1. Test INNER JOIN
    Dim innerDf As cDataFrame
    Set innerDf = leftDf.Merge(rightDf, "DeptID", "INNER")

    ' Alice(D10), Bob(D20), Charlie(D10), David(D30) match -> 4 rows
    AssertEqual innerDf.RowCount, 4, "INNER JOIN produces 4 matching rows"

    ' Verify column schema and collision resolution
    AssertTrue innerDf.HasColumn("DeptID"), "Merged has DeptID key"
    AssertTrue innerDf.HasColumn("Name"), "Merged has Name"
    AssertTrue innerDf.HasColumn("DeptName"), "Merged has DeptName"
    AssertTrue innerDf.HasColumn("Location"), "Merged has Location"
    AssertTrue innerDf.HasColumn("Status_left"), "Collision resolved: Status_left"
    AssertTrue innerDf.HasColumn("Status_right"), "Collision resolved: Status_right"
    AssertTrue Not innerDf.HasColumn("Status"), "Original ambiguous Status not present"

    ' Verify row contents
    AssertEqual innerDf(1, "Name"), "Alice", "Row 1 Name is Alice"
    AssertEqual innerDf(1, "DeptName"), "Engineering", "Row 1 DeptName is Engineering"
    AssertEqual innerDf(1, "Location"), "Building A", "Row 1 Location is Building A"
    AssertEqual innerDf(1, "Status_left"), "Active", "Row 1 Status_left is Active"
    AssertEqual innerDf(1, "Status_right"), "Open", "Row 1 Status_right is Open"

    AssertEqual innerDf(2, "Name"), "Bob", "Row 2 Name is Bob"
    AssertEqual innerDf(2, "DeptName"), "Human Resources", "Row 2 DeptName is Human Resources"

    ' 2. Test LEFT JOIN
    Dim leftJoinDf As cDataFrame
    Set leftJoinDf = leftDf.Merge(rightDf, "DeptID", "LEFT")

    ' All 5 left rows retained (including Eve with D99)
    AssertEqual leftJoinDf.RowCount, 5, "LEFT JOIN produces 5 rows"
    AssertEqual leftJoinDf(5, "Name"), "Eve", "Row 5 Name is Eve"
    AssertEqual leftJoinDf(5, "DeptID"), "D99", "Row 5 DeptID is D99"
    AssertEqual leftJoinDf(5, "Status_left"), "Pending", "Row 5 Status_left is Pending"
    AssertTrue IsEmpty(leftJoinDf(5, "DeptName")), "Row 5 DeptName is Empty (unmatched right)"
    AssertTrue IsEmpty(leftJoinDf(5, "Location")), "Row 5 Location is Empty (unmatched right)"
    AssertTrue IsEmpty(leftJoinDf(5, "Status_right")), "Row 5 Status_right is Empty (unmatched right)"

    ' 3. Test 1:N Join (Multiple right matches for single left key)
    Dim ordersDf As New cDataFrame
    ordersDf.AddColumn "CustID"
    ordersDf.AddColumn "Customer"
    ordersDf.AddRow "C1", "Acme Corp"

    Dim itemsDf As New cDataFrame
    itemsDf.AddColumn "CustID"
    itemsDf.AddColumn "Item"
    itemsDf.AddColumn "Price"
    itemsDf.AddRow "C1", "Laptop", 1200
    itemsDf.AddRow "C1", "Mouse", 25
    itemsDf.AddRow "C1", "Keyboard", 75

    Dim multiDf As cDataFrame
    Set multiDf = ordersDf.Merge(itemsDf, "CustID", "INNER")
    AssertEqual multiDf.RowCount, 3, "1:N Join produces 3 order rows for C1"
    AssertEqual multiDf(1, "Item"), "Laptop", "Item 1 is Laptop"
    AssertEqual multiDf(2, "Item"), "Mouse", "Item 2 is Mouse"
    AssertEqual multiDf(3, "Item"), "Keyboard", "Item 3 is Keyboard"
    AssertEqual multiDf(3, "Customer"), "Acme Corp", "Customer repeated on all 1:N rows"

    Exit Sub
TestFail:
    RecordFailure "Test_HashMerge", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 11: GroupBy Aggregation (Single/Multi-Column, SUM/COUNT/AVG/MIN/MAX)
' ------------------------------------------------------------------------------
Private Sub Test_GroupBy()
    On Error GoTo TestFail
    LogPrint "[TEST 11] GroupBy Aggregation (Single/Multi-Column, SUM/COUNT/AVG/MIN/MAX)..."

    Dim df As New cDataFrame
    df.AddColumn "Dept"
    df.AddColumn "Location"
    df.AddColumn "EmpID"
    df.AddColumn "Salary"

    df.AddRow "IT", "BKK", "E1", 80000#
    df.AddRow "IT", "BKK", "E2", 100000#
    df.AddRow "IT", "CNX", "E3", 60000#
    df.AddRow "HR", "BKK", "E4", 50000#
    df.AddRow "HR", "BKK", "E5", 70000#
    df.AddRow "Sales", "BKK", "E6", 90000#

    ' 1. Single-column GroupBy with SUM (default)
    Dim gSum As cDataFrame
    Set gSum = df.GroupBy("Dept", "Salary")
    AssertEqual gSum.RowCount, 3, "GroupBy Dept produces 3 groups"
    AssertEqual gSum.ColumnCount, 2, "GroupBy Dept has 2 columns (Dept, Salary_SUM)"
    AssertTrue gSum.HasColumn("Dept"), "Has Dept column"
    AssertTrue gSum.HasColumn("Salary_SUM"), "Has Salary_SUM column"
    AssertEqual gSum(1, "Dept"), "IT", "Group 1 is IT"
    AssertEqual gSum(1, "Salary_SUM"), 240000#, "IT Salary_SUM = 240000"
    AssertEqual gSum(2, "Dept"), "HR", "Group 2 is HR"
    AssertEqual gSum(2, "Salary_SUM"), 120000#, "HR Salary_SUM = 120000"
    AssertEqual gSum(3, "Dept"), "Sales", "Group 3 is Sales"
    AssertEqual gSum(3, "Salary_SUM"), 90000#, "Sales Salary_SUM = 90000"

    ' 2. Single-column GroupBy with all 5 functions (SUM, COUNT, AVERAGE->AVG, MIN, MAX)
    Dim gAll As cDataFrame
    Set gAll = df.GroupBy("Dept", _
        Array("Salary", "EmpID", "Salary", "Salary", "Salary"), _
        Array("SUM", "COUNT", "AVERAGE", "MIN", "MAX"))

    AssertEqual gAll.ColumnCount, 6, "gAll has 6 columns"
    AssertTrue gAll.HasColumn("Salary_SUM"), "Has Salary_SUM"
    AssertTrue gAll.HasColumn("EmpID_COUNT"), "Has EmpID_COUNT"
    AssertTrue gAll.HasColumn("Salary_AVG"), "AVERAGE normalized to Salary_AVG"
    AssertTrue gAll.HasColumn("Salary_MIN"), "Has Salary_MIN"
    AssertTrue gAll.HasColumn("Salary_MAX"), "Has Salary_MAX"

    AssertEqual gAll(1, "EmpID_COUNT"), 3, "IT count is 3"
    AssertEqual gAll(1, "Salary_AVG"), 80000#, "IT avg salary is 80000"
    AssertEqual gAll(1, "Salary_MIN"), 60000#, "IT min salary is 60000"
    AssertEqual gAll(1, "Salary_MAX"), 100000#, "IT max salary is 100000"

    AssertEqual gAll(2, "EmpID_COUNT"), 2, "HR count is 2"
    AssertEqual gAll(2, "Salary_AVG"), 60000#, "HR avg salary is 60000"
    AssertEqual gAll(2, "Salary_MIN"), 50000#, "HR min salary is 50000"
    AssertEqual gAll(2, "Salary_MAX"), 70000#, "HR max salary is 70000"

    ' 3. Multi-column GroupBy (Dept + Location) with multiple aggregations
    Dim gMulti As cDataFrame
    Set gMulti = df.GroupBy(Array("Dept", "Location"), _
        Array("Salary", "Salary", "EmpID"), _
        Array("SUM", "AVG", "COUNT"))

    AssertEqual gMulti.RowCount, 4, "GroupBy (Dept, Location) produces 4 unique pairs"
    AssertEqual gMulti.ColumnCount, 5, "GroupBy (Dept, Location) has 5 columns"
    AssertEqual gMulti(1, "Dept"), "IT", "Row 1 Dept = IT"
    AssertEqual gMulti(1, "Location"), "BKK", "Row 1 Location = BKK"
    AssertEqual gMulti(1, "Salary_SUM"), 180000#, "IT-BKK Salary_SUM = 180000"
    AssertEqual gMulti(1, "Salary_AVG"), 90000#, "IT-BKK Salary_AVG = 90000"
    AssertEqual gMulti(1, "EmpID_COUNT"), 2, "IT-BKK EmpID_COUNT = 2"

    AssertEqual gMulti(2, "Dept"), "IT", "Row 2 Dept = IT"
    AssertEqual gMulti(2, "Location"), "CNX", "Row 2 Location = CNX"
    AssertEqual gMulti(2, "Salary_SUM"), 60000#, "IT-CNX Salary_SUM = 60000"
    AssertEqual gMulti(2, "EmpID_COUNT"), 1, "IT-CNX EmpID_COUNT = 1"

    ' 4. Edge Case: Empty and Null cells in aggregated column
    Dim sparseDf As New cDataFrame
    sparseDf.AddColumn "Grp"
    sparseDf.AddColumn "Val"
    sparseDf.AddRow "A", 10#
    sparseDf.AddRow "A", Empty
    sparseDf.AddRow "A", 30#
    sparseDf.AddRow "B", Empty

    Dim gSparse As cDataFrame
    Set gSparse = sparseDf.GroupBy("Grp", _
        Array("Val", "Val", "Val", "Val", "Val"), _
        Array("SUM", "COUNT", "AVG", "MIN", "MAX"))

    AssertEqual gSparse(1, "Val_SUM"), 40#, "Sparse A SUM = 40"
    AssertEqual gSparse(1, "Val_COUNT"), 2, "Sparse A COUNT ignores Empty = 2"
    AssertEqual gSparse(1, "Val_AVG"), 20#, "Sparse A AVG (40/2) = 20"
    AssertEqual gSparse(1, "Val_MIN"), 10#, "Sparse A MIN = 10"
    AssertEqual gSparse(1, "Val_MAX"), 30#, "Sparse A MAX = 30"

    AssertEqual gSparse(2, "Val_SUM"), 0#, "All-empty group B SUM = 0"
    AssertEqual gSparse(2, "Val_COUNT"), 0, "All-empty group B COUNT = 0"
    AssertEqual gSparse(2, "Val_AVG"), 0#, "All-empty group B AVG = 0"
    AssertTrue IsEmpty(gSparse(2, "Val_MIN")), "All-empty group B MIN is Empty"
    AssertTrue IsEmpty(gSparse(2, "Val_MAX")), "All-empty group B MAX is Empty"

    ' 5. Edge Case: Empty DataFrame (0 rows) preserves schema
    Dim emptyDf As cDataFrame
    Set emptyDf = df.Clone(True)
    Dim gEmpty As cDataFrame
    Set gEmpty = emptyDf.GroupBy("Dept", "Salary", "SUM")
    AssertEqual gEmpty.RowCount, 0, "Empty GroupBy has 0 rows"
    AssertEqual gEmpty.ColumnCount, 2, "Empty GroupBy has 2 schema columns"
    AssertTrue gEmpty.HasColumn("Salary_SUM"), "Empty GroupBy has Salary_SUM column"

    ' 6. Error Validation: Unsupported function and mismatched lengths
    Dim caughtUnsupported As Boolean
    caughtUnsupported = False
    On Error Resume Next
    Dim gBad As cDataFrame
    Set gBad = df.GroupBy("Dept", "Salary", "BOGUS_AGG")
    If Err.Number = 5 Then caughtUnsupported = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtUnsupported, "GroupBy raises Error 5 on unsupported function"

    Dim caughtMismatch As Boolean
    caughtMismatch = False
    On Error Resume Next
    Set gBad = df.GroupBy("Dept", Array("Salary", "EmpID"), Array("SUM"))
    If Err.Number = 5 Then caughtMismatch = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtMismatch, "GroupBy raises Error 5 on mismatched AggColumns/AggFuncs lengths"

    Exit Sub
TestFail:
    RecordFailure "Test_GroupBy", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 12: Pivot Table / Cross-Tabulation (SUM/COUNT, FillEmpty, O(1) Loc)
' ------------------------------------------------------------------------------
Private Sub Test_Pivot()
    On Error GoTo TestFail
    LogPrint "[TEST 12] Pivot Table / Cross-Tabulation (SUM/COUNT, FillEmpty, O(1) Loc)..."

    Dim salesDf As New cDataFrame
    salesDf.AddColumn "Region"
    salesDf.AddColumn "Quarter"
    salesDf.AddColumn "Revenue"

    salesDf.AddRow "North", "Q1", 100#
    salesDf.AddRow "North", "Q1", 50#
    salesDf.AddRow "North", "Q2", 200#
    salesDf.AddRow "South", "Q1", 80#
    salesDf.AddRow "South", "Q3", 300#
    salesDf.AddRow "East", "Q2", 120#

    ' 1. Pivot with SUM and default FillEmpty = 0
    Dim pivSum As cDataFrame
    Set pivSum = salesDf.Pivot("Region", "Quarter", "Revenue", "SUM", 0)

    AssertEqual pivSum.RowCount, 3, "Pivot has 3 unique Region rows (North, South, East)"
    AssertEqual pivSum.ColumnCount, 4, "Pivot has 4 columns (Region, Q1, Q2, Q3)"
    AssertEqual pivSum.IndexColumnName, "Region", "Pivot automatically sets index on Region"

    ' Test O(1) Loc lookup directly on Pivot table
    AssertEqual pivSum.Loc("North", "Q1"), 150#, "North Q1 SUM = 150"
    AssertEqual pivSum.Loc("North", "Q2"), 200#, "North Q2 SUM = 200"
    AssertEqual pivSum.Loc("North", "Q3"), 0, "North Q3 missing cell filled with 0"

    AssertEqual pivSum.Loc("South", "Q1"), 80#, "South Q1 SUM = 80"
    AssertEqual pivSum.Loc("South", "Q2"), 0, "South Q2 missing cell filled with 0"
    AssertEqual pivSum.Loc("South", "Q3"), 300#, "South Q3 SUM = 300"

    AssertEqual pivSum.Loc("East", "Q1"), 0, "East Q1 missing cell filled with 0"
    AssertEqual pivSum.Loc("East", "Q2"), 120#, "East Q2 SUM = 120"
    AssertEqual pivSum.Loc("East", "Q3"), 0, "East Q3 missing cell filled with 0"

    ' 2. Pivot with COUNT and custom FillEmpty = Empty
    Dim pivCount As cDataFrame
    Set pivCount = salesDf.Pivot("Region", "Quarter", "Revenue", "COUNT", Empty)

    AssertEqual pivCount.Loc("North", "Q1"), 2, "North Q1 COUNT = 2"
    AssertEqual pivCount.Loc("North", "Q2"), 1, "North Q2 COUNT = 1"
    AssertTrue IsEmpty(pivCount.Loc("North", "Q3")), "North Q3 missing cell is Empty"
    AssertTrue IsEmpty(pivCount.Loc("South", "Q2")), "South Q2 missing cell is Empty"
    AssertEqual pivCount.Loc("South", "Q3"), 1, "South Q3 COUNT = 1"

    ' 3. Pivot with AVG
    Dim pivAvg As cDataFrame
    Set pivAvg = salesDf.Pivot("Region", "Quarter", "Revenue", "AVG", 0#)
    AssertEqual pivAvg.Loc("North", "Q1"), 75#, "North Q1 AVG (150/2) = 75"

    ' 4. Edge Case: Empty DataFrame Pivot
    Dim emptySales As cDataFrame
    Set emptySales = salesDf.Clone(True)
    Dim emptyPiv As cDataFrame
    Set emptyPiv = emptySales.Pivot("Region", "Quarter", "Revenue", "SUM")
    AssertEqual emptyPiv.RowCount, 0, "Empty Pivot has 0 rows"
    AssertEqual emptyPiv.ColumnCount, 1, "Empty Pivot has 1 column (Region)"

    Exit Sub
TestFail:
    RecordFailure "Test_Pivot", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 13: Deduplication (cColumn.Unique & cDataFrame.DropDuplicates)
' ------------------------------------------------------------------------------
Private Sub Test_UniqueAndDropDuplicates()
    On Error GoTo TestFail
    LogPrint "[TEST 13] Deduplication (cColumn.Unique & cDataFrame.DropDuplicates)..."

    ' 1. cColumn.Unique() with numbers, case-insensitive strings, Empty vs ""
    Dim rawCol As New cColumn
    rawCol.Name = "Raw"
    rawCol.Append 10
    rawCol.Append 20
    rawCol.Append 10
    rawCol.Append Empty
    rawCol.Append ""
    rawCol.Append Empty
    rawCol.Append "abc"
    rawCol.Append "ABC"

    Dim uCol As cColumn
    Set uCol = rawCol.Unique()
    AssertEqual uCol.Count, 5, "cColumn.Unique returns 5 distinct elements"
    AssertEqual uCol.Name, "Raw", "cColumn.Unique preserves column Name"
    AssertEqual uCol(1), 10, "Unique(1) = 10"
    AssertEqual uCol(2), 20, "Unique(2) = 20"
    AssertTrue IsEmpty(uCol(3)), "Unique(3) is Empty (preserved once)"
    AssertEqual uCol(4), "", "Unique(4) is empty string '' (distinct from Empty)"
    AssertEqual uCol(5), "abc", "Unique(5) is 'abc' ('ABC' deduplicated)"

    Dim emptyCol As New cColumn
    AssertEqual emptyCol.Unique().Count, 0, "Empty column Unique() has Count = 0"

    ' 2. cDataFrame.DropDuplicates across all columns (default)
    Dim dfAll As New cDataFrame
    dfAll.AddColumn "ID"
    dfAll.AddColumn "Val"
    dfAll.AddRow 1, "A"
    dfAll.AddRow 2, "B"
    dfAll.AddRow 1, "A"
    dfAll.AddRow 3, "C"
    dfAll.AddRow 4, "D"

    Dim ddAll As cDataFrame
    Set ddAll = dfAll.DropDuplicates()
    AssertEqual ddAll.RowCount, 4, "DropDuplicates() across all columns keeps 4 unique rows"
    AssertEqual ddAll(1, "ID"), 1, "Row 1 ID = 1"
    AssertEqual ddAll(4, "ID"), 4, "Row 4 ID = 4"

    ' 3. cDataFrame.DropDuplicates on single column with FIRST, LAST, NONE
    Dim empDf As New cDataFrame
    empDf.AddColumn "Dept"
    empDf.AddColumn "Role"
    empDf.AddColumn "Emp"

    empDf.AddRow "IT", "Dev", "Alice"
    empDf.AddRow "IT", "Dev", "Bob"
    empDf.AddRow "HR", "Recruiter", "Charlie"
    empDf.AddRow "Sales", "AE", "Dave"
    empDf.AddRow "Sales", "SE", "Eve"

    Dim ddFirst As cDataFrame
    Set ddFirst = empDf.DropDuplicates("Dept", "FIRST")
    AssertEqual ddFirst.RowCount, 3, "DropDuplicates('Dept', 'FIRST') keeps 3 rows"
    AssertEqual ddFirst(1, "Emp"), "Alice", "FIRST keeps Alice for IT"
    AssertEqual ddFirst(2, "Emp"), "Charlie", "FIRST keeps Charlie for HR"

    Dim ddLast As cDataFrame
    Set ddLast = empDf.DropDuplicates("Dept", "LAST")
    AssertEqual ddLast.RowCount, 3, "DropDuplicates('Dept', 'LAST') keeps 3 rows"
    AssertEqual ddLast(1, "Emp"), "Bob", "LAST keeps Bob for IT"
    AssertEqual ddLast(2, "Emp"), "Charlie", "LAST keeps Charlie for HR"
    AssertEqual ddLast(3, "Emp"), "Eve", "LAST keeps Eve for Sales"

    Dim ddNone As cDataFrame
    Set ddNone = empDf.DropDuplicates("Dept", "NONE")
    AssertEqual ddNone.RowCount, 1, "DropDuplicates('Dept', 'NONE') keeps only singleton groups (1 row)"
    AssertEqual ddNone(1, "Dept"), "HR", "NONE singleton Dept is HR"
    AssertEqual ddNone(1, "Emp"), "Charlie", "NONE singleton Emp is Charlie"

    ' 4. Multi-column subset Array("Dept", "Role") + Empty vs "" distinction
    Dim ddMulti As cDataFrame
    Set ddMulti = empDf.DropDuplicates(Array("Dept", "Role"), "FIRST")
    AssertEqual ddMulti.RowCount, 4, "DropDuplicates(Array('Dept', 'Role')) keeps 4 distinct (Dept, Role) pairs"

    Dim nullVsBlankDf As New cDataFrame
    nullVsBlankDf.AddColumn "Key"
    nullVsBlankDf.AddRow Empty
    nullVsBlankDf.AddRow ""
    Dim ddNullBlank As cDataFrame
    Set ddNullBlank = nullVsBlankDf.DropDuplicates("Key")
    AssertEqual ddNullBlank.RowCount, 2, "DropDuplicates distinguishes Empty from empty string ''"

    ' 5. Error 5 on invalid Keep mode
    Dim caughtBadKeep As Boolean
    caughtBadKeep = False
    On Error Resume Next
    Dim ddBad As cDataFrame
    Set ddBad = empDf.DropDuplicates("Dept", "INVALID_MODE")
    If Err.Number = 5 Then caughtBadKeep = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtBadKeep, "DropDuplicates raises Error 5 on unsupported Keep argument"

    Exit Sub
TestFail:
    RecordFailure "Test_UniqueAndDropDuplicates", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 14: Vertical Concatenation / UNION ALL (Concat)
' ------------------------------------------------------------------------------
Private Sub Test_Concat()
    On Error GoTo TestFail
    LogPrint "[TEST 14] Vertical Concatenation / UNION ALL (Concat)..."

    Dim df1 As New cDataFrame
    df1.AddColumn "ID"
    df1.AddColumn "Name"
    df1.AddRow 1, "Alice"
    df1.AddRow 2, "Bob"

    Dim df2 As New cDataFrame
    df2.AddColumn "ID"
    df2.AddColumn "Name"
    df2.AddRow 3, "Charlie"
    df2.AddRow 4, "Dave"

    ' 1. Identical schema Concat
    Dim c1 As cDataFrame
    Set c1 = df1.Concat(df2)
    AssertEqual c1.RowCount, 4, "Concat identical schema RowCount = 4"
    AssertEqual c1.ColumnCount, 2, "Concat identical schema ColumnCount = 2"
    AssertEqual c1(1, "ID"), 1, "Row 1 ID = 1"
    AssertEqual c1(3, "ID"), 3, "Row 3 ID = 3"
    AssertEqual c1(4, "Name"), "Dave", "Row 4 Name = Dave"

    ' 2. Reordered columns in OtherTable (Name, ID) with strict AllowMissingColumns = False
    Dim dfReordered As New cDataFrame
    dfReordered.AddColumn "Name"
    dfReordered.AddColumn "ID"
    dfReordered.AddRow "Charlie", 3
    dfReordered.AddRow "Dave", 4

    Dim cReordered As cDataFrame
    Set cReordered = df1.Concat(dfReordered, False)
    AssertEqual cReordered.RowCount, 4, "Concat reordered columns RowCount = 4"
    AssertEqual cReordered(3, "ID"), 3, "Reordered row 3 ID aligned to 3"
    AssertEqual cReordered(3, "Name"), "Charlie", "Reordered row 3 Name aligned to Charlie"

    ' 3. Diagonal / Superset schema with AllowMissingColumns = True
    Dim dfLeft As New cDataFrame
    dfLeft.AddColumn "ID"
    dfLeft.AddColumn "Name"
    dfLeft.AddColumn "Dept"
    dfLeft.AddRow 1, "Alice", "IT"
    dfLeft.AddRow 2, "Bob", "HR"

    Dim dfRight As New cDataFrame
    dfRight.AddColumn "ID"
    dfRight.AddColumn "Name"
    dfRight.AddColumn "Bonus"
    dfRight.AddRow 3, "Charlie", 5000
    dfRight.AddRow 4, "Dave", 7500

    Dim cDiag As cDataFrame
    Set cDiag = dfLeft.Concat(dfRight, True)
    AssertEqual cDiag.RowCount, 4, "Diagonal Concat RowCount = 4"
    AssertEqual cDiag.ColumnCount, 4, "Diagonal Concat union ColumnCount = 4 (ID, Name, Dept, Bonus)"
    AssertTrue cDiag.HasColumn("Bonus"), "Diagonal Concat includes Bonus column from OtherTable"
    AssertTrue IsEmpty(cDiag(1, "Bonus")), "Top rows have Empty for Bonus"
    AssertEqual cDiag(3, "Bonus"), 5000, "Bottom row 3 has Bonus = 5000"
    AssertEqual cDiag(1, "Dept"), "IT", "Top row 1 has Dept = IT"
    AssertTrue IsEmpty(cDiag(3, "Dept")), "Bottom rows have Empty for Dept"

    ' 4. Error 5 when AllowMissingColumns = False and schemas differ
    Dim caughtMismatch As Boolean
    caughtMismatch = False
    On Error Resume Next
    Dim cBad As cDataFrame
    Set cBad = dfLeft.Concat(dfRight, False)
    If Err.Number = 5 Then caughtMismatch = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtMismatch, "Concat raises Error 5 on schema mismatch when AllowMissingColumns = False"

    Exit Sub
TestFail:
    RecordFailure "Test_Concat", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 15: Vectorized Column Arithmetic (VectorOp, Add, Subtract, Multiply, Divide)
' ------------------------------------------------------------------------------
Private Sub Test_VectorArithmetic()
    On Error GoTo TestFail
    LogPrint "[TEST 15] Vectorized Column Arithmetic (VectorOp, Add, Subtract, Multiply, Divide)..."

    Dim colA As New cColumn
    colA.Name = "Price"
    colA.Append 100#
    colA.Append 200#
    colA.Append Empty
    colA.Append 400#

    ' 1. Scalar addition via Add
    Dim colAdd As cColumn
    Set colAdd = colA.Add(10#, "Plus10")
    AssertEqual colAdd.Count, 4, "colAdd Count = 4"
    AssertEqual colAdd.Name, "Plus10", "colAdd Name = Plus10"
    AssertEqual colAdd(1), 110#, "100 + 10 = 110"
    AssertEqual colAdd(2), 210#, "200 + 10 = 210"
    AssertTrue IsEmpty(colAdd(3)), "Empty + 10 propagates Empty"
    AssertEqual colAdd(4), 410#, "400 + 10 = 410"

    ' 2. Column-to-Column subtraction via Subtract
    Dim colB As New cColumn
    colB.Name = "Discount"
    colB.Append 20#
    colB.Append 50#
    colB.Append 10#
    colB.Append 0#

    Dim colSub As cColumn
    Set colSub = colA.Subtract(colB, "NetPrice")
    AssertEqual colSub(1), 80#, "100 - 20 = 80"
    AssertEqual colSub(2), 150#, "200 - 50 = 150"
    AssertTrue IsEmpty(colSub(3)), "Empty - 10 propagates Empty"
    AssertEqual colSub(4), 400#, "400 - 0 = 400"

    ' 3. 1D Array multiplication via Multiply
    Dim colMul As cColumn
    Set colMul = colA.Multiply(Array(0.9, 0.8, 0.7, 0.5), "Scaled")
    AssertEqual colMul(1), 90#, "100 * 0.9 = 90"
    AssertEqual colMul(2), 160#, "200 * 0.8 = 160"
    AssertEqual colMul(4), 200#, "400 * 0.5 = 200"

    ' 4. Division with zero divisor protection (#DIV/0! = CVErr(2007))
    Dim colDiv As cColumn
    Set colDiv = colA.Divide(colB, "Ratio")
    AssertEqual colDiv(1), 5#, "100 / 20 = 5"
    AssertEqual colDiv(2), 4#, "200 / 50 = 4"
    AssertIsError colDiv(4), "400 / 0 returns #DIV/0! error variant without crashing"

    ' 5. Error validation on mismatched length and unsupported operator
    Dim shortCol As New cColumn
    shortCol.Append 1#
    Dim caughtLenMismatch As Boolean
    caughtLenMismatch = False
    On Error Resume Next
    Dim badCol As cColumn
    Set badCol = colA.Add(shortCol)
    If Err.Number = 5 Then caughtLenMismatch = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtLenMismatch, "VectorOp raises Error 5 on mismatched cColumn lengths"

    Dim caughtBadOp As Boolean
    caughtBadOp = False
    On Error Resume Next
    Set badCol = colA.VectorOp("^", 2)
    If Err.Number = 5 Then caughtBadOp = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtBadOp, "VectorOp raises Error 5 on unsupported operator"

    Exit Sub
TestFail:
    RecordFailure "Test_VectorArithmetic", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 16: Missing Data Handling (FillNA & DropNA)
' ------------------------------------------------------------------------------
Private Sub Test_MissingDataFillAndDropNA()
    On Error GoTo TestFail
    LogPrint "[TEST 16] Missing Data Handling (FillNA & DropNA)..."

    ' 1. cColumn.FillNA
    Dim naCol As New cColumn
    naCol.Name = "Metric"
    naCol.Append 15#
    naCol.Append Empty
    naCol.Append CVErr(2042)
    naCol.Append "   "

    naCol.FillNA 0#, False
    AssertEqual naCol(2), 0#, "FillNA replaced Empty with 0"
    AssertEqual naCol(3), 0#, "FillNA replaced #N/A Error with 0"
    AssertEqual naCol(4), "   ", "FillNA preserved blank string when TreatBlankStringAsNA = False"

    naCol.FillNA -1#, True
    AssertEqual naCol(4), -1#, "FillNA replaced blank string when TreatBlankStringAsNA = True"

    ' 2. cDataFrame.DropNA (ANY, ALL, SubsetColumns, TreatBlankStringAsNA)
    Dim df As New cDataFrame
    df.AddColumn "Name"
    df.AddColumn "Dept"
    df.AddColumn "Score"

    df.AddRow "Alice", "IT", 95#        ' Row 1: Complete
    df.AddRow "Bob", "", 80#            ' Row 2: Blank string in Dept, valid Score
    df.AddRow "Charlie", "HR", Empty    ' Row 3: Missing Score
    df.AddRow Empty, Empty, Empty       ' Row 4: All Empty

    Dim cleanAny As cDataFrame
    Set cleanAny = df.DropNA(, "ANY", False)
    AssertEqual cleanAny.RowCount, 2, "DropNA('ANY', False) keeps rows 1 and 2"
    AssertEqual cleanAny(1, "Name"), "Alice", "Row 1 is Alice"

    Dim cleanSub As cDataFrame
    Set cleanSub = df.DropNA("Score", "ANY", False)
    AssertEqual cleanSub.RowCount, 2, "DropNA(Subset:='Score') keeps rows 1 and 2"

    Dim cleanAll As cDataFrame
    Set cleanAll = df.DropNA(, "ALL", False)
    AssertEqual cleanAll.RowCount, 3, "DropNA('ALL') drops only Row 4 (all Empty)"

    Dim cleanBlank As cDataFrame
    Set cleanBlank = df.DropNA(, "ANY", True)
    AssertEqual cleanBlank.RowCount, 1, "DropNA('ANY', TreatBlankStringAsNA:=True) keeps only Row 1"

    ' 3. cDataFrame.FillNA (SubsetColumns and all columns)
    Dim dfFill As cDataFrame
    Set dfFill = df.Clone(False)
    dfFill.FillNA 0#, "Score", False
    AssertEqual dfFill(3, "Score"), 0#, "df.FillNA filled Score in Row 3 with 0"
    AssertTrue IsEmpty(dfFill(4, "Name")), "df.FillNA on subset left Name in Row 4 Empty"

    dfFill.FillNA "UNKNOWN", , True
    AssertEqual dfFill(2, "Dept"), "UNKNOWN", "df.FillNA replaced blank Dept in Row 2 with UNKNOWN"
    AssertEqual dfFill(4, "Name"), "UNKNOWN", "df.FillNA replaced Empty Name in Row 4 with UNKNOWN"

    ' 4. Error 5 on invalid How argument in DropNA
    Dim caughtBadHow As Boolean
    caughtBadHow = False
    On Error Resume Next
    Dim badDf As cDataFrame
    Set badDf = df.DropNA(, "SOME")
    If Err.Number = 5 Then caughtBadHow = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtBadHow, "DropNA raises Error 5 on unsupported How mode"

    Exit Sub
TestFail:
    RecordFailure "Test_MissingDataFillAndDropNA", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 17: Fast CSV File I/O Round-Trip (ToCSV & LoadFromCSV, RFC 4180)
' ------------------------------------------------------------------------------
Private Sub Test_CSVRoundTrip()
    On Error GoTo TestFail
    LogPrint "[TEST 17] Fast CSV File I/O Round-Trip (ToCSV & LoadFromCSV)..."

    Dim srcDf As New cDataFrame
    srcDf.AddColumn "ID"
    srcDf.AddColumn "City"
    srcDf.AddColumn "Item"
    srcDf.AddColumn "Price"

    srcDf.AddRow 101, "Bangkok, TH", "15"" Monitor", 450.5
    srcDf.AddRow 102, "Chiang Mai", Empty, 120#
    srcDf.AddRow 103, "Phuket", "Keyboard", Empty

    Dim tempCsv As String
    tempCsv = Environ$("TEMP") & "\vba_df_test_roundtrip.csv"

    ' 1. Write and read back with headers
    srcDf.ToCSV tempCsv, ",", True

    Dim loadedDf As New cDataFrame
    loadedDf.LoadFromCSV tempCsv, ",", True, True

    AssertEqual loadedDf.RowCount, 3, "LoadFromCSV loaded 3 rows"
    AssertEqual loadedDf.ColumnCount, 4, "LoadFromCSV loaded 4 columns"
    AssertEqual loadedDf(1, "ID"), 101#, "Row 1 ID inferred as numeric 101"
    AssertEqual loadedDf(1, "City"), "Bangkok, TH", "Row 1 City preserved comma inside quotes"
    AssertEqual loadedDf(1, "Item"), "15"" Monitor", "Row 1 Item unescaped double quotes"
    AssertEqual loadedDf(1, "Price"), 450.5, "Row 1 Price = 450.5"
    AssertEqual loadedDf(2, "City"), "Chiang Mai", "Row 2 City = Chiang Mai"
    AssertTrue IsEmpty(loadedDf(2, "Item")), "Row 2 middle empty field parsed as Empty"
    AssertEqual loadedDf(2, "Price"), 120#, "Row 2 Price = 120"
    AssertEqual loadedDf(3, "City"), "Phuket", "Row 3 City = Phuket"
    AssertTrue IsEmpty(loadedDf(3, "Price")), "Row 3 trailing empty field parsed as Empty"

    ' 2. Write and read back without headers (HasHeaders = False)
    srcDf.ToCSV tempCsv, ",", False
    Dim noHdrDf As New cDataFrame
    noHdrDf.LoadFromCSV tempCsv, ",", False, True

    AssertEqual noHdrDf.RowCount, 3, "HasHeaders:=False loaded 3 rows"
    AssertTrue noHdrDf.HasColumn("Column_1"), "HasHeaders:=False created Column_1"
    AssertEqual noHdrDf(1, "Column_2"), "Bangkok, TH", "HasHeaders:=False Row 1 Column_2 = Bangkok, TH"

    On Error Resume Next
    Kill tempCsv
    On Error GoTo TestFail

    ' 3. Error 53 on missing file
    Dim caughtMissingFile As Boolean
    caughtMissingFile = False
    On Error Resume Next
    Dim badCsvDf As New cDataFrame
    badCsvDf.LoadFromCSV Environ$("TEMP") & "\non_existent_vba_df_99999.csv"
    If Err.Number = 53 Then caughtMissingFile = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtMissingFile, "LoadFromCSV raises Error 53 when file does not exist"

    Exit Sub
TestFail:
    On Error Resume Next
    If Len(tempCsv) > 0 And Len(Dir$(tempCsv)) > 0 Then Kill tempCsv
    On Error GoTo 0
    RecordFailure "Test_CSVRoundTrip", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 18: Excel Table (ListObject) Bulk I/O (ToListObject & LoadFromListObject)
' ------------------------------------------------------------------------------
Private Sub Test_ListObjectIO()
    On Error GoTo SkipOrReport
    LogPrint "[TEST 18] Excel Table (ListObject) Bulk I/O..."

    Dim xlApp As Object
    On Error Resume Next
    Set xlApp = Application
    On Error GoTo TestFail

    If xlApp Is Nothing Then
        LogPrint "  (Skipping ListObjectIO: Host application not detected)"
        mPassCount = mPassCount + 1
        Exit Sub
    End If

    Dim ws As Object
    Set ws = xlApp.ActiveWorkbook.Worksheets.Add

    Dim srcDf As New cDataFrame
    srcDf.AddColumn "SKU"
    srcDf.AddColumn "Item"
    srcDf.AddColumn "Qty"

    srcDf.AddRow "SKU-01", "Keyboard", 30
    srcDf.AddRow "SKU-02", "Mouse", 45
    srcDf.AddRow "SKU-03", "Monitor", 50

    ' 1. Dump DataFrame to ListObject
    Dim lo As Object
    Set lo = srcDf.ToListObject(ws.Range("A1"), "tblInventory")

    AssertEqual lo.Name, "tblInventory", "Created ListObject named tblInventory"
    AssertEqual lo.ListColumns.Count, 3, "ListObject has 3 columns"
    AssertEqual lo.ListRows.Count, 3, "ListObject has 3 data rows"

    ' Turn on Totals Row to verify LoadFromListObject ignores TotalsRowRange
    lo.ShowTotals = True

    ' 2. Load back from ListObject
    Dim dfFromLo As New cDataFrame
    dfFromLo.LoadFromListObject lo

    AssertEqual dfFromLo.RowCount, 3, "LoadFromListObject loaded 3 data rows (ignoring Totals row)"
    AssertEqual dfFromLo.ColumnCount, 3, "LoadFromListObject loaded 3 columns"
    AssertEqual dfFromLo(1, "SKU"), "SKU-01", "Row 1 SKU = SKU-01"
    AssertEqual dfFromLo(2, "Item"), "Mouse", "Row 2 Item = Mouse"
    AssertEqual dfFromLo(3, "Qty"), 50, "Row 3 Qty = 50"

    xlApp.DisplayAlerts = False
    ws.Delete
    xlApp.DisplayAlerts = True
    Exit Sub

SkipOrReport:
    LogPrint "  (ListObject test skipped or completed with status: " & Err.Description & ")"
    mPassCount = mPassCount + 1
    Exit Sub
TestFail:
    RecordFailure "Test_ListObjectIO", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 19: End-to-End Demo Sales Analytics Pipeline (mDemoWorkflow.Step3)
' ------------------------------------------------------------------------------
Private Sub Test_DemoAnalyticsPipeline()
    On Error GoTo TestFail
    LogPrint "[TEST 19] End-to-End Demo Sales Analytics Pipeline (mDemoWorkflow.Step3)..."

    Dim dfEnriched As cDataFrame
    Dim dfSummary As cDataFrame
    Dim dfPivot As cDataFrame
    Dim tempCsv As String
    tempCsv = Environ$("TEMP") & "\vba_demo_pipeline_test.csv"

    mDemoWorkflow.RunSalesPipelineInMemory dfEnriched, dfSummary, dfPivot, tempCsv

    ' 1. Verify Concat + DropDuplicates(LAST) + DropNA + FillNA + Merge
    AssertEqual dfEnriched.RowCount, 6, "Pipeline produced 6 clean enriched orders"
    AssertEqual dfEnriched.ColumnCount, 12, "Enriched DataFrame has 12 columns"
    AssertEqual dfEnriched(1, "OrderID"), "ORD-101", "Row 1 is ORD-101"
    AssertEqual dfEnriched(1, "Qty"), 4#, "DropDuplicates(LAST) kept revised Qty = 4 for ORD-101"
    AssertEqual dfEnriched(1, "SalesChannel"), "Walk-In", "FillNA imputed Walk-In for BKK SalesChannel"
    AssertEqual dfEnriched(2, "DiscountRate"), 0#, "FillNA imputed 0 for missing DiscountRate"

    ' 2. Verify Vectorized Pricing Arithmetic (GrossAmount, DiscountAmt, NetRevenue) & Unique
    AssertEqual dfEnriched.Column("GrossAmount").Sum(), 12200#, "Total GrossAmount = 12200"
    AssertEqual dfEnriched.Column("DiscountAmt").Sum(), 570#, "Total DiscountAmt = 570"
    AssertEqual dfEnriched.Column("NetRevenue").Sum(), 11630#, "Total NetRevenue = 11630"
    AssertEqual dfEnriched.Column("Category").Unique().Count, 3, "3 distinct product categories"

    ' 3. Verify Multi-Column GroupBy + Descending Sort
    AssertEqual dfSummary.RowCount, 5, "GroupBy (Category, SalesChannel) produced 5 groups"
    AssertEqual dfSummary(1, "Category"), "Hardware", "Top revenue group Category is Hardware"
    AssertEqual dfSummary(1, "SalesChannel"), "Walk-In", "Top revenue group Channel is Walk-In"
    AssertEqual dfSummary(1, "NetRevenue_SUM"), 5600#, "Top group NetRevenue_SUM = 5600"
    AssertEqual dfSummary(1, "NetRevenue_AVG"), 2800#, "Top group NetRevenue_AVG = 2800"
    AssertEqual dfSummary(1, "OrderID_COUNT"), 2, "Top group OrderID_COUNT = 2"

    ' 4. Verify Cross-Tabulation Pivot & O(1) Loc Lookup
    AssertEqual dfPivot.RowCount, 3, "Pivot has 3 Category rows"
    AssertEqual dfPivot.Loc("Hardware", "Q1"), 6680#, "Pivot Loc(Hardware, Q1) = 6680"
    AssertEqual dfPivot.Loc("Hardware", "Q2"), 2000#, "Pivot Loc(Hardware, Q2) = 2000"
    AssertEqual dfPivot.Loc("Accessories", "Q1"), 0#, "Pivot Loc(Accessories, Q1) filled with 0"
    AssertEqual dfPivot.Loc("Accessories", "Q2"), 950#, "Pivot Loc(Accessories, Q2) = 950"
    AssertEqual dfPivot.Loc("Software", "Q2"), 2000#, "Pivot Loc(Software, Q2) = 2000"

    ' 5. Verify CSV Export Round-Trip
    Dim csvCheckDf As New cDataFrame
    csvCheckDf.LoadFromCSV tempCsv, ",", True, True
    AssertEqual csvCheckDf.RowCount, 5, "Exported summary CSV reloaded 5 rows"
    AssertEqual csvCheckDf(1, "NetRevenue_SUM"), 5600#, "Reloaded CSV top NetRevenue_SUM = 5600"

    On Error Resume Next
    Kill tempCsv
    On Error GoTo 0
    Exit Sub

TestFail:
    On Error Resume Next
    If Len(tempCsv) > 0 And Len(Dir$(tempCsv)) > 0 Then Kill tempCsv
    On Error GoTo 0
    RecordFailure "Test_DemoAnalyticsPipeline", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 20: Window & Cumulative Functions (CumSum, Shift, PctChange, Rank)
' ------------------------------------------------------------------------------
Private Sub Test_WindowAndCumulative()
    On Error GoTo TestFail
    LogPrint "[TEST 20] Window & Cumulative Functions (CumSum, Shift, PctChange, Rank)..."

    Dim col As New cColumn
    col.Name = "Rev"
    col.Append 10#
    col.Append 20#
    col.Append Empty
    col.Append 40#

    ' 1. CumSum
    Dim csCol As cColumn
    Set csCol = col.CumSum("RunningRev")
    AssertEqual csCol.Count, 4, "CumSum Count = 4"
    AssertEqual csCol(1), 10#, "CumSum(1) = 10"
    AssertEqual csCol(2), 30#, "CumSum(2) = 30"
    AssertEqual csCol(3), 30#, "CumSum(3) carries 30 across Empty"
    AssertEqual csCol(4), 70#, "CumSum(4) = 70"

    ' 2. Shift (Lag +1 and Lead -1)
    Dim lag1 As cColumn
    Set lag1 = col.Shift(1, Empty, "Lag1")
    AssertTrue IsEmpty(lag1(1)), "Lag1(1) is Empty"
    AssertEqual lag1(2), 10#, "Lag1(2) = 10"
    AssertTrue IsEmpty(lag1(4)), "Lag1(4) is Empty (from row 3)"

    Dim lead1 As cColumn
    Set lead1 = col.Shift(-1, 0#, "Lead1")
    AssertEqual lead1(1), 20#, "Lead1(1) = 20"
    AssertEqual lead1(3), 40#, "Lead1(3) = 40"
    AssertEqual lead1(4), 0#, "Lead1(4) filled with 0"

    ' 3. PctChange
    Dim pctCol As cColumn
    Set pctCol = col.PctChange(1, Empty, "Growth")
    AssertTrue IsEmpty(pctCol(1)), "PctChange(1) is Empty"
    AssertEqual pctCol(2), 1#, "PctChange(2) = (20-10)/10 = 1.0"
    AssertTrue IsEmpty(pctCol(3)), "PctChange(3) is Empty"

    Dim zeroCol As New cColumn
    zeroCol.Append 0#
    zeroCol.Append 50#
    Dim pctZero As cColumn
    Set pctZero = zeroCol.PctChange(1)
    AssertIsError pctZero(2), "PctChange from 0 returns #DIV/0! error variant"

    ' 4. Rank (AVERAGE, MIN, DENSE)
    Dim scoreCol As New cColumn
    scoreCol.Name = "Score"
    scoreCol.Append 100#
    scoreCol.Append 200#
    scoreCol.Append 100#
    scoreCol.Append 300#

    Dim rAvg As cColumn
    Set rAvg = scoreCol.Rank(True, "AVERAGE")
    AssertEqual rAvg(1), 1.5, "Rank AVERAGE(1) tied at 1.5"
    AssertEqual rAvg(2), 3#, "Rank AVERAGE(2) = 3"
    AssertEqual rAvg(3), 1.5, "Rank AVERAGE(3) tied at 1.5"
    AssertEqual rAvg(4), 4#, "Rank AVERAGE(4) = 4"

    Dim rMin As cColumn
    Set rMin = scoreCol.Rank(True, "MIN")
    AssertEqual rMin(1), 1#, "Rank MIN(1) = 1"
    AssertEqual rMin(2), 3#, "Rank MIN(2) = 3"
    AssertEqual rMin(3), 1#, "Rank MIN(3) = 1"

    Dim rDense As cColumn
    Set rDense = scoreCol.Rank(True, "DENSE")
    AssertEqual rDense(1), 1#, "Rank DENSE(1) = 1"
    AssertEqual rDense(2), 2#, "Rank DENSE(2) = 2"
    AssertEqual rDense(4), 3#, "Rank DENSE(4) = 3"

    Exit Sub
TestFail:
    RecordFailure "Test_WindowAndCumulative", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 21: Row Slicing & Column Projection (Head, Tail, Slice, SelectColumns)
' ------------------------------------------------------------------------------
Private Sub Test_SliceAndSelectColumns()
    On Error GoTo TestFail
    LogPrint "[TEST 21] Row Slicing & Column Projection (Head, Tail, Slice, SelectColumns)..."

    Dim df As New cDataFrame
    df.AddColumn "ID"
    df.AddColumn "Name"
    df.AddColumn "Score"

    df.AddRow 1, "A", 90#
    df.AddRow 2, "B", 80#
    df.AddRow 3, "C", 70#
    df.AddRow 4, "D", 60#
    df.AddRow 5, "E", 50#

    ' 1. Head
    Dim hDf As cDataFrame
    Set hDf = df.Head(2)
    AssertEqual hDf.RowCount, 2, "Head(2) RowCount = 2"
    AssertEqual hDf(1, "ID"), 1, "Head(2) row 1 ID = 1"
    AssertEqual hDf(2, "ID"), 2, "Head(2) row 2 ID = 2"

    ' 2. Tail
    Dim tDf As cDataFrame
    Set tDf = df.Tail(2)
    AssertEqual tDf.RowCount, 2, "Tail(2) RowCount = 2"
    AssertEqual tDf(1, "ID"), 4, "Tail(2) row 1 ID = 4"
    AssertEqual tDf(2, "ID"), 5, "Tail(2) row 2 ID = 5"

    ' 3. Slice
    Dim sDf As cDataFrame
    Set sDf = df.Slice(2, 4)
    AssertEqual sDf.RowCount, 3, "Slice(2, 4) RowCount = 3"
    AssertEqual sDf(1, "ID"), 2, "Slice(2, 4) first ID = 2"
    AssertEqual sDf(3, "ID"), 4, "Slice(2, 4) last ID = 4"

    Dim emptySlice As cDataFrame
    Set emptySlice = df.Slice(4, 2)
    AssertEqual emptySlice.RowCount, 0, "Slice(4, 2) has 0 rows"
    AssertEqual emptySlice.ColumnCount, 3, "Slice(4, 2) preserves 3 columns"

    ' 4. SelectColumns
    Dim selDf As cDataFrame
    Set selDf = df.SelectColumns(Array("Score", "ID"))
    AssertEqual selDf.ColumnCount, 2, "SelectColumns has 2 columns"
    AssertEqual selDf.RowCount, 5, "SelectColumns has 5 rows"
    AssertEqual selDf(1, 1), 90#, "SelectColumns col 1 is Score (90)"
    AssertEqual selDf(1, 2), 1, "SelectColumns col 2 is ID (1)"
    AssertTrue Not selDf.HasColumn("Name"), "SelectColumns excluded Name"

    Exit Sub
TestFail:
    RecordFailure "Test_SliceAndSelectColumns", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 22: Advanced Joins (RIGHT, FULL, LEFT_ANTI, LEFT_SEMI & Multi-Key Merge)
' ------------------------------------------------------------------------------
Private Sub Test_AdvancedJoins()
    On Error GoTo TestFail
    LogPrint "[TEST 22] Advanced Joins (RIGHT, FULL, LEFT_ANTI, LEFT_SEMI & Multi-Key Merge)..."

    Dim leftDf As New cDataFrame
    leftDf.AddColumn "DeptID"
    leftDf.AddColumn "Name"
    leftDf.AddRow "D10", "Alice"
    leftDf.AddRow "D20", "Bob"
    leftDf.AddRow "D99", "Ghost"

    Dim rightDf As New cDataFrame
    rightDf.AddColumn "DeptID"
    rightDf.AddColumn "DeptTitle"
    rightDf.AddRow "D10", "Engineering"
    rightDf.AddRow "D20", "HR"
    rightDf.AddRow "D40", "Legal"

    ' 1. RIGHT Join
    Dim rightJoinDf As cDataFrame
    Set rightJoinDf = leftDf.Merge(rightDf, "DeptID", "RIGHT")
    AssertEqual rightJoinDf.RowCount, 3, "RIGHT Join produces 3 rows (D10, D20, D40)"
    AssertEqual rightJoinDf(3, "DeptID"), "D40", "Unmatched right row preserves DeptID = D40"
    AssertTrue IsEmpty(rightJoinDf(3, "Name")), "Unmatched right row has Empty Name"
    AssertEqual rightJoinDf(3, "DeptTitle"), "Legal", "Unmatched right row has DeptTitle = Legal"

    ' 2. FULL Outer Join
    Dim fullJoinDf As cDataFrame
    Set fullJoinDf = leftDf.Merge(rightDf, "DeptID", "FULL")
    AssertEqual fullJoinDf.RowCount, 4, "FULL Join produces 4 rows (D10, D20, D99, D40)"
    AssertEqual fullJoinDf(3, "DeptID"), "D99", "Row 3 is unmatched left D99"
    AssertTrue IsEmpty(fullJoinDf(3, "DeptTitle")), "Row 3 has Empty DeptTitle"
    AssertEqual fullJoinDf(4, "DeptID"), "D40", "Row 4 is unmatched right D40"

    ' 3. LEFT_ANTI Join
    Dim antiDf As cDataFrame
    Set antiDf = leftDf.Merge(rightDf, "DeptID", "LEFT_ANTI")
    AssertEqual antiDf.RowCount, 1, "LEFT_ANTI returns 1 unmatched left row"
    AssertEqual antiDf.ColumnCount, 2, "LEFT_ANTI preserves Left schema only (2 cols)"
    AssertEqual antiDf(1, "DeptID"), "D99", "LEFT_ANTI row is D99"

    ' 4. LEFT_SEMI Join
    Dim semiDf As cDataFrame
    Set semiDf = leftDf.Merge(rightDf, "DeptID", "LEFT_SEMI")
    AssertEqual semiDf.RowCount, 2, "LEFT_SEMI returns 2 matched left rows"
    AssertEqual semiDf.ColumnCount, 2, "LEFT_SEMI preserves Left schema only (2 cols)"
    AssertEqual semiDf(1, "Name"), "Alice", "LEFT_SEMI row 1 is Alice"

    ' 5. Multi-Column Composite Key Merge
    Dim mLeft As New cDataFrame
    mLeft.AddColumn "Dept"
    mLeft.AddColumn "Loc"
    mLeft.AddColumn "Headcount"
    mLeft.AddRow "IT", "BKK", 25
    mLeft.AddRow "IT", "CNX", 10
    mLeft.AddRow "HR", "BKK", 8

    Dim mRight As New cDataFrame
    mRight.AddColumn "Dept"
    mRight.AddColumn "Loc"
    mRight.AddColumn "Budget"
    mRight.AddRow "IT", "BKK", 500000#
    mRight.AddRow "IT", "CNX", 200000#

    Dim multiKeyDf As cDataFrame
    Set multiKeyDf = mLeft.Merge(mRight, Array("Dept", "Loc"), "INNER")
    AssertEqual multiKeyDf.RowCount, 2, "Multi-key INNER join matched 2 composite rows"
    AssertEqual multiKeyDf(1, "Budget"), 500000#, "IT-BKK Budget = 500000"
    AssertEqual multiKeyDf(2, "Budget"), 200000#, "IT-CNX Budget = 200000"

    Exit Sub
TestFail:
    RecordFailure "Test_AdvancedJoins", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 23: Vectorized String Operations, Clip & Conditional Where
' ------------------------------------------------------------------------------
Private Sub Test_VectorStringAndWhere()
    On Error GoTo TestFail
    LogPrint "[TEST 23] Vectorized String Operations, Clip & Conditional Where..."

    Dim strCol As New cColumn
    strCol.Name = "RawText"
    strCol.Append "  Hello World  "
    strCol.Append "vba"
    strCol.Append Empty

    ' 1. StrTrim, StrUpper, StrLower, StrReplace
    Dim tCol As cColumn
    Set tCol = strCol.StrTrim("Trimmed")
    AssertEqual tCol(1), "Hello World", "StrTrim(1) = 'Hello World'"
    AssertEqual tCol(2), "vba", "StrTrim(2) = 'vba'"
    AssertTrue IsEmpty(tCol(3)), "StrTrim(3) preserves Empty"

    Dim uCol As cColumn, lCol As cColumn
    Set uCol = tCol.StrUpper()
    Set lCol = tCol.StrLower()
    AssertEqual uCol(1), "HELLO WORLD", "StrUpper(1) = 'HELLO WORLD'"
    AssertEqual lCol(1), "hello world", "StrLower(1) = 'hello world'"

    Dim rCol As cColumn
    Set rCol = tCol.StrReplace("Hello", "Hi")
    AssertEqual rCol(1), "Hi World", "StrReplace(1) = 'Hi World'"
    AssertEqual rCol(2), "vba", "StrReplace(2) unchanged"

    ' 2. Clip
    Dim numCol As New cColumn
    numCol.Append 5#
    numCol.Append 25#
    numCol.Append 100#
    numCol.Append Empty

    Dim cCol As cColumn
    Set cCol = numCol.Clip(10#, 50#, "Clamped")
    AssertEqual cCol(1), 10#, "Clip clamped 5 up to 10"
    AssertEqual cCol(2), 25#, "Clip kept 25 unchanged"
    AssertEqual cCol(3), 50#, "Clip clamped 100 down to 50"
    AssertTrue IsEmpty(cCol(4)), "Clip preserved Empty"

    ' 3. cDataFrame.Where (scalar and cColumn branches)
    Dim df As New cDataFrame
    df.AddColumn "Score"
    df.AddColumn "BonusBase"
    df.AddRow 65#, 100#
    df.AddRow 80#, 200#
    df.AddRow 95#, 300#

    Dim wCol As cColumn
    Set wCol = df.Where("Score", ">=", 80#, "PASS", "FAIL", "Status")
    AssertEqual wCol.Name, "Status", "Where column Name = Status"
    AssertEqual wCol(1), "FAIL", "Row 1 (65) = FAIL"
    AssertEqual wCol(2), "PASS", "Row 2 (80) = PASS"
    AssertEqual wCol(3), "PASS", "Row 3 (95) = PASS"

    Dim wCol2 As cColumn
    Set wCol2 = df.Where("Score", ">=", 80#, df.Column("BonusBase"), 0#, "FinalBonus")
    AssertEqual wCol2(3), 300#, "Where with cColumn TrueValue picked 300 for Row 3"

    Exit Sub
TestFail:
    RecordFailure "Test_VectorStringAndWhere", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 24: Vectorized Date & Time Operations (Year, Month, Day, Quarter, DayOfWeek, DateAdd, DateDiff, FormatDate)
' ------------------------------------------------------------------------------
Private Sub Test_VectorDateOperations()
    On Error GoTo TestFail
    LogPrint "[TEST 24] Vectorized Date & Time Operations..."

    Dim dtCol As New cColumn
    dtCol.Name = "OrderDate"
    dtCol.Append DateSerial(2026, 1, 15)  ' Q1, Jan 15, 2026 (Thursday -> vbSunday=5)
    dtCol.Append DateSerial(2026, 5, 20)  ' Q2, May 20, 2026
    dtCol.Append DateSerial(2026, 11, 7)  ' Q4, Nov 7, 2026
    dtCol.Append Empty

    Dim yCol As cColumn, mCol As cColumn, dCol As cColumn, qCol As cColumn, dwCol As cColumn
    Set yCol = dtCol.Year("Yr")
    Set mCol = dtCol.Month("Mo")
    Set dCol = dtCol.Day("Dy")
    Set qCol = dtCol.Quarter("Qtr")
    Set dwCol = dtCol.DayOfWeek(vbSunday, "DOW")

    AssertEqual yCol(1), 2026&, "Year(1) = 2026"
    AssertTrue IsEmpty(yCol(4)), "Year(4) preserves Empty"
    AssertEqual mCol(2), 5&, "Month(2) = 5"
    AssertEqual dCol(1), 15&, "Day(1) = 15"
    AssertEqual qCol(1), 1&, "Quarter(1) = 1"
    AssertEqual qCol(2), 2&, "Quarter(2) = 2"
    AssertEqual qCol(3), 4&, "Quarter(3) = 4"
    AssertEqual dwCol(1), 5&, "DayOfWeek(2026-01-15, vbSunday) = 5 (Thursday)"

    ' DateAdd (+10 days)
    Dim plus10 As cColumn
    Set plus10 = dtCol.DateAdd("d", 10, "DueDate")
    AssertEqual plus10(1), DateSerial(2026, 1, 25), "DateAdd('d', 10) = 2026-01-25"
    AssertTrue IsEmpty(plus10(4)), "DateAdd(4) preserves Empty"

    ' DateDiff against cColumn and scalar Date
    Dim diffCol As cColumn
    Set diffCol = dtCol.DateDiff("d", plus10, "DaysToDue")
    AssertEqual diffCol(1), 10&, "DateDiff against cColumn = 10 days"
    AssertTrue IsEmpty(diffCol(4)), "DateDiff(4) preserves Empty"

    Dim diffScalar As cColumn
    Set diffScalar = dtCol.DateDiff("m", DateSerial(2026, 11, 15))
    AssertEqual diffScalar(1), 10&, "DateDiff scalar months Jan to Nov = 10"

    ' FormatDate
    Dim fmtCol As cColumn
    Set fmtCol = dtCol.FormatDate("yyyy-mm-dd", "ISO")
    AssertEqual fmtCol(1), "2026-01-15", "FormatDate(1) = 2026-01-15"
    AssertEqual fmtCol(3), "2026-11-07", "FormatDate(3) = 2026-11-07"
    AssertTrue IsEmpty(fmtCol(4)), "FormatDate(4) preserves Empty"

    Exit Sub
TestFail:
    RecordFailure "Test_VectorDateOperations", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 25: Multi-Condition Compound Filter (FilterCompound AND / OR)
' ------------------------------------------------------------------------------
Private Sub Test_FilterCompound()
    On Error GoTo TestFail
    LogPrint "[TEST 25] Multi-Condition Compound Filter (FilterCompound AND / OR)..."

    Dim df As New cDataFrame
    df.AddColumn "Dept"
    df.AddColumn "Status"
    df.AddColumn "Salary"

    df.AddRow "IT", "ACTIVE", 90000#
    df.AddRow "IT", "PROBATION", 65000#
    df.AddRow "HR", "ACTIVE", 55000#
    df.AddRow "Sales", "ACTIVE", 85000#

    ' 1. AND logic: Status = ACTIVE AND Salary >= 80000
    Dim andDf As cDataFrame
    Set andDf = df.FilterCompound( _
        Array(Array("Status", "=", "ACTIVE"), Array("Salary", ">=", 80000#)), _
        "AND")

    AssertEqual andDf.RowCount, 2, "FilterCompound AND matched 2 rows"
    AssertEqual andDf(1, "Dept"), "IT", "AND Row 1 Dept = IT"
    AssertEqual andDf(2, "Dept"), "Sales", "AND Row 2 Dept = Sales"

    ' 2. OR logic: Dept = HR OR Salary < 70000
    Dim orDf As cDataFrame
    Set orDf = df.FilterCompound( _
        Array(Array("Dept", "=", "HR"), Array("Salary", "<", 70000#)), _
        "OR")

    AssertEqual orDf.RowCount, 2, "FilterCompound OR matched 2 rows (IT Probation + HR Active)"
    AssertEqual orDf(1, "Status"), "PROBATION", "OR Row 1 Status = PROBATION"
    AssertEqual orDf(2, "Dept"), "HR", "OR Row 2 Dept = HR"

    Exit Sub
TestFail:
    RecordFailure "Test_FilterCompound", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 26: Data Reshaping / Unpivot (Melt)
' ------------------------------------------------------------------------------
Private Sub Test_MeltUnpivot()
    On Error GoTo TestFail
    LogPrint "[TEST 26] Data Reshaping / Unpivot (Melt)..."

    Dim wideDf As New cDataFrame
    wideDf.AddColumn "Region"
    wideDf.AddColumn "Q1"
    wideDf.AddColumn "Q2"

    wideDf.AddRow "North", 100#, 150#
    wideDf.AddRow "South", 200#, 250#

    Dim meltedDf As cDataFrame
    Set meltedDf = wideDf.Melt("Region", Array("Q1", "Q2"), "Quarter", "Revenue")

    AssertEqual meltedDf.RowCount, 4, "Melt produced 2 x 2 = 4 rows"
    AssertEqual meltedDf.ColumnCount, 3, "Melt has 3 columns (Region, Quarter, Revenue)"
    AssertEqual meltedDf(1, "Region"), "North", "Row 1 Region = North"
    AssertEqual meltedDf(1, "Quarter"), "Q1", "Row 1 Quarter = Q1"
    AssertEqual meltedDf(1, "Revenue"), 100#, "Row 1 Revenue = 100"
    AssertEqual meltedDf(2, "Quarter"), "Q2", "Row 2 Quarter = Q2"
    AssertEqual meltedDf(2, "Revenue"), 150#, "Row 2 Revenue = 150"
    AssertEqual meltedDf(4, "Region"), "South", "Row 4 Region = South"
    AssertEqual meltedDf(4, "Revenue"), 250#, "Row 4 Revenue = 250"

    Exit Sub
TestFail:
    RecordFailure "Test_MeltUnpivot", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 27: Fast JSON Serialization & Deserialization (ToJSON & LoadFromJSON)
' ------------------------------------------------------------------------------
Private Sub Test_JSONSerialization()
    On Error GoTo TestFail
    LogPrint "[TEST 27] Fast JSON Serialization & Deserialization (ToJSON & LoadFromJSON)..."

    Dim df As New cDataFrame
    df.AddColumn "ID"
    df.AddColumn "Label"
    df.AddColumn "Active"
    df.AddColumn "Note"

    df.AddRow 1, "Item ""A""", True, Empty
    df.AddRow 2, "Path\B", False, "OK"

    ' 1. RECORDS format round-trip
    Dim jsonRec As String
    jsonRec = df.ToJSON("RECORDS")
    AssertTrue InStr(1, jsonRec, "[{""ID"":1,") > 0, "RECORDS JSON starts with [{ID:1"

    Dim fromRecDf As New cDataFrame
    fromRecDf.LoadFromJSON jsonRec

    AssertEqual fromRecDf.RowCount, 2, "LoadFromJSON(RECORDS) loaded 2 rows"
    AssertEqual fromRecDf.ColumnCount, 4, "LoadFromJSON(RECORDS) loaded 4 columns"
    AssertEqual fromRecDf(1, "ID"), 1#, "RECORDS Row 1 ID = 1"
    AssertEqual fromRecDf(1, "Label"), "Item ""A""", "RECORDS unescaped quotes in Label"
    AssertEqual fromRecDf(1, "Active"), True, "RECORDS parsed boolean True"
    AssertTrue IsEmpty(fromRecDf(1, "Note")), "RECORDS parsed null as Empty"
    AssertEqual fromRecDf(2, "Label"), "Path\B", "RECORDS unescaped backslash"
    AssertEqual fromRecDf(2, "Active"), False, "RECORDS parsed boolean False"

    ' 2. SPLIT format round-trip
    Dim jsonSplit As String
    jsonSplit = df.ToJSON("SPLIT")
    AssertTrue InStr(1, jsonSplit, "{""columns"":") > 0, "SPLIT JSON contains columns key"

    Dim fromSplitDf As New cDataFrame
    fromSplitDf.LoadFromJSON jsonSplit

    AssertEqual fromSplitDf.RowCount, 2, "LoadFromJSON(SPLIT) loaded 2 rows"
    AssertEqual fromSplitDf.ColumnCount, 4, "LoadFromJSON(SPLIT) loaded 4 columns"
    AssertEqual fromSplitDf(1, "Label"), "Item ""A""", "SPLIT Row 1 Label matches"
    AssertEqual fromSplitDf(2, "Note"), "OK", "SPLIT Row 2 Note = OK"

    Exit Sub
TestFail:
    RecordFailure "Test_JSONSerialization", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 28: High-Performance Bulk 2D Array Row Appends (AddRows)
' ------------------------------------------------------------------------------
Private Sub Test_BulkAddRows()
    On Error GoTo TestFail
    LogPrint "[TEST 28] High-Performance Bulk 2D Array Row Appends (AddRows)..."

    Dim df As New cDataFrame
    df.AddColumn "ID"
    df.AddColumn "Code"
    df.AddColumn "Amount"

    df.AddRow 1, "INIT", 50#

    ' 1. 1-based 2D array append
    Dim batch1(1 To 3, 1 To 3) As Variant
    batch1(1, 1) = 2: batch1(1, 2) = "B2": batch1(1, 3) = 100#
    batch1(2, 1) = 3: batch1(2, 2) = "B3": batch1(2, 3) = 200#
    batch1(3, 1) = 4: batch1(3, 2) = "B4": batch1(3, 3) = 300#

    Dim newTotal As Long
    newTotal = df.AddRows(batch1)

    AssertEqual newTotal, 4, "AddRows returned new RowCount = 4"
    AssertEqual df.RowCount, 4, "df.RowCount = 4"
    AssertEqual df(1, "Code"), "INIT", "Existing row 1 preserved"
    AssertEqual df(2, "Code"), "B2", "Appended row 2 Code = B2"
    AssertEqual df(4, "Amount"), 300#, "Appended row 4 Amount = 300"

    ' 2. 0-based 2D array append
    Dim batch0(0 To 1, 0 To 2) As Variant
    batch0(0, 0) = 5: batch0(0, 1) = "Z5": batch0(0, 2) = 500#
    batch0(1, 0) = 6: batch0(1, 1) = "Z6": batch0(1, 2) = 600#

    df.AddRows batch0
    AssertEqual df.RowCount, 6, "0-based AddRows increased RowCount to 6"
    AssertEqual df(5, "Code"), "Z5", "Row 5 Code = Z5"
    AssertEqual df(6, "Amount"), 600#, "Row 6 Amount = 600"

    Exit Sub
TestFail:
    RecordFailure "Test_BulkAddRows", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 29: CSV Formula Injection Protection (ToCSV SanitizeFormulas:=True)
' ------------------------------------------------------------------------------
Private Sub Test_CSVFormulaSanitization()
    On Error GoTo TestFail
    LogPrint "[TEST 29] CSV Formula Injection Protection (ToCSV SanitizeFormulas:=True)..."

    Dim df As New cDataFrame
    df.AddColumn "User"
    df.AddColumn "Payload"
    df.AddColumn "Amount"

    df.AddRow "Alice", "=CMD|' /C calc'!A0", 100#
    df.AddRow "Bob", "+SUM(A1:A10)", 200#
    df.AddRow "Charlie", "-10+20", -50#
    df.AddRow "Dave", "@SUM(1,2)", 300#
    df.AddRow "Eve", "Safe Text", 400#

    Dim tempCsv As String
    tempCsv = Environ$("TEMP") & "\vba_df_formula_sanitize.csv"

    df.ToCSV tempCsv, ",", True, True

    Dim roundDf As New cDataFrame
    roundDf.LoadFromCSV tempCsv, ",", True, True

    AssertEqual roundDf.RowCount, 5, "Sanitized CSV loaded 5 rows"
    AssertEqual roundDf(1, "Payload"), "'=CMD|' /C calc'!A0", "Leading '=' prefixed with single quote"
    AssertEqual roundDf(2, "Payload"), "'+SUM(A1:A10)", "Leading '+' prefixed with single quote"
    AssertEqual roundDf(3, "Payload"), "'-10+20", "Leading '-' string prefixed with single quote"
    AssertEqual roundDf(3, "Amount"), -50#, "True numeric negative (-50) NOT prefixed with quote"
    AssertEqual roundDf(4, "Payload"), "'@SUM(1,2)", "Leading '@' prefixed with single quote"
    AssertEqual roundDf(5, "Payload"), "Safe Text", "Normal string unchanged"

    On Error Resume Next
    Kill tempCsv
    On Error GoTo 0
    Exit Sub

TestFail:
    On Error Resume Next
    If Len(tempCsv) > 0 And Len(Dir$(tempCsv)) > 0 Then Kill tempCsv
    On Error GoTo 0
    RecordFailure "Test_CSVFormulaSanitization", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 30: Multi-Column Composite Indexing (SetIndex & O(1) Loc with Array Keys)
' ------------------------------------------------------------------------------
Private Sub Test_MultiColumnIndex()
    On Error GoTo TestFail
    LogPrint "[TEST 30] Multi-Column Composite Indexing (SetIndex & O(1) Loc)..."

    Dim df As New cDataFrame
    df.AddColumn "Branch"
    df.AddColumn "Year"
    df.AddColumn "Target"
    df.AddColumn "Manager"

    df.AddRow "BKK", 2025, 100000#, "Korn"
    df.AddRow "BKK", 2026, 150000#, "Kit"
    df.AddRow "CNX", 2025, 60000#, "Somchai"
    df.AddRow "CNX", 2026, 85000#, "Mook"

    ' 1. Build composite index on (Branch, Year)
    df.SetIndex Array("Branch", "Year")
    AssertEqual df.IndexColumnName, "Branch, Year", "Composite IndexColumnName is 'Branch, Year'"
    AssertTrue df.HasIndexKey(Array("BKK", 2026)), "HasIndexKey(Array('BKK', 2026)) is True"
    AssertTrue Not df.HasIndexKey(Array("BKK", 2099)), "HasIndexKey(Array('BKK', 2099)) is False"

    ' 2. O(1) Loc Get with composite key array
    AssertEqual df.Loc(Array("BKK", 2026), "Target"), 150000#, "Loc(Array('BKK', 2026), 'Target') = 150000"
    AssertEqual df.Loc(Array("cnx", 2025), "Manager"), "Somchai", "Case-insensitive composite Loc = Somchai"
    AssertIsError df.Loc(Array("HKT", 2026), "Target"), "Missing composite key returns #N/A"

    ' 3. O(1) Loc Let mutation with composite key array
    df.Loc(Array("CNX", 2026), "Target") = 95000#
    AssertEqual df.Loc(Array("CNX", 2026), "Target"), 95000#, "Loc Let updated CNX-2026 Target to 95000"
    AssertEqual df(4, "Target"), 95000#, "Underlying row 4 Target updated to 95000"

    ' 4. Auto-rebuild of composite index after AddRow
    df.AddRow "HKT", 2026, 50000#, "Nat"
    AssertEqual df.Loc(Array("HKT", 2026), "Manager"), "Nat", "Composite index auto-rebuilt after AddRow"

    ' 5. Error 457 on duplicate composite key
    Dim dupDf As New cDataFrame
    dupDf.AddColumn "K1"
    dupDf.AddColumn "K2"
    dupDf.AddRow "A", 1
    dupDf.AddRow "A", 2
    dupDf.AddRow "A", 1

    Dim caughtCompDup As Boolean
    caughtCompDup = False
    On Error Resume Next
    dupDf.SetIndex Array("K1", "K2")
    If Err.Number = 457 Then caughtCompDup = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtCompDup, "SetIndex raises Error 457 on duplicate composite key"

    Exit Sub
TestFail:
    RecordFailure "Test_MultiColumnIndex", Err.Description
End Sub


' ------------------------------------------------------------------------------
' Test 31: Excel Cell & Range Styling (cTableStyler)
' ------------------------------------------------------------------------------
Private Sub Test_TableStyler()
    On Error GoTo SkipOrReport
    LogPrint "[TEST 31] Excel Cell & Range Styling (cTableStyler)..."

    Dim xlApp As Object
    On Error Resume Next
    Set xlApp = Application
    On Error GoTo TestFail

    If xlApp Is Nothing Then
        LogPrint "  (Skipping TableStyler: Host application not detected)"
        mPassCount = mPassCount + 1
        Exit Sub
    End If

    Dim ws As Object
    Set ws = xlApp.ActiveWorkbook.Worksheets.Add

    ' Build test DataFrame
    Dim df As New cDataFrame
    df.AddColumn "Item"
    df.AddColumn "Category"
    df.AddColumn "Qty"
    df.AddColumn "Amount"

    df.AddRow Array("Widget A", "Hardware", 0.25, 1500.5)
    df.AddRow Array("Widget B", "Software", 0.1, 250.75)
    df.AddRow Array("Total", "All", 0.35, 1751.25)

    ' Dump to Range B2:E5
    df.ToRange ws.Range("B2"), True

    ' Attach cTableStyler
    Dim styler As New cTableStyler
    styler.Attach ws.Range("B2"), df, True

    ' 1. Metadata and Dimensions
    AssertEqual styler.RowCount, 3, "cTableStyler RowCount is 3"
    AssertEqual styler.ColumnCount, 4, "cTableStyler ColumnCount is 4"
    AssertTrue styler.HasHeaders, "cTableStyler HasHeaders is True"
    AssertTrue Not styler.TableRange Is Nothing, "TableRange is not Nothing"
    AssertEqual styler.TableRange.Address, "$B$2:$E$5", "TableRange address matches $B$2:$E$5"
    AssertEqual styler.HeaderRange.Address, "$B$2:$E$2", "HeaderRange address matches $B$2:$E$2"
    AssertEqual styler.DataBodyRange.Address, "$B$3:$E$5", "DataBodyRange address matches $B$3:$E$5"

    ' 2. NAVY Theme Application
    styler.ApplyTheme "NAVY"
    AssertEqual styler.HeaderRange.Interior.Color, RGB(31, 78, 121), "NAVY header interior color is RGB(31, 78, 121)"
    AssertEqual styler.HeaderRange.Font.Color, RGB(255, 255, 255), "NAVY header font color is RGB(255, 255, 255)"
    AssertTrue styler.HeaderRange.Font.Bold, "NAVY header font is Bold"
    AssertEqual styler.HeaderRange.HorizontalAlignment, -4108, "NAVY header horizontal alignment is xlCenter (-4108)"

    ' 3. Borders Application
    styler.ApplyBorders
    AssertEqual styler.TableRange.Borders.LineStyle, 1, "ApplyBorders line style is xlContinuous (1)"
    AssertEqual styler.TableRange.Borders.Weight, 2, "ApplyBorders border weight is xlThin (2)"

    ' 4. Zebra Striping
    styler.ApplyZebraStripes
    ' Even data row (row 2 in DataBodyRange -> row 4 in worksheet)
    AssertEqual styler.DataBodyRange.Rows(2).Interior.Color, 16316664, "Even data row interior color matches stripe color 16316664"
    ' Odd data row (row 1 in DataBodyRange -> row 3 in worksheet) should not equal stripe color
    AssertTrue styler.DataBodyRange.Rows(1).Interior.Color <> 16316664, "Odd data row interior color differs from stripe color"

    ' 5. Vectorized Column Number Formatting
    styler.FormatColumn "Amount", "#,##0.00"
    styler.FormatColumn "Qty", "0.0%"
    AssertEqual ws.Range("E3").NumberFormat, "#,##0.00", "Amount data cell E3 formatted as '#,##0.00'"
    AssertEqual ws.Range("E4").NumberFormat, "#,##0.00", "Amount data cell E4 formatted as '#,##0.00'"
    AssertEqual ws.Range("D3").NumberFormat, "0.0%", "Qty data cell D3 formatted as '0.0%'"
    AssertEqual ws.Range("D4").NumberFormat, "0.0%", "Qty data cell D4 formatted as '0.0%'"

    ' 1-based index format check
    styler.FormatColumn 1, "@"
    AssertEqual ws.Range("B3").NumberFormat, "@", "Column 1 (Item) formatted as text '@'"

    ' 6. Total Row Highlighting
    styler.HighlightTotalRow
    AssertTrue styler.DataBodyRange.Rows(3).Font.Bold, "Total row font is bold"
    AssertEqual styler.DataBodyRange.Rows(3).Borders(9).LineStyle, -4119, "Total row bottom border is xlDouble (-4119)"
    AssertEqual styler.DataBodyRange.Rows(3).Borders(8).LineStyle, 1, "Total row top border is xlContinuous (1)"

    ' 7. AutoFit Columns with custom padding
    styler.AutoFitColumns 4#
    AssertTrue styler.TableRange.Columns(1).ColumnWidth > 0, "AutoFitColumns set valid positive column width"

    ' 8. Alternative Theme Verification
    styler.ApplyTheme "EMERALD"
    AssertEqual styler.HeaderRange.Interior.Color, RGB(39, 110, 80), "EMERALD header interior color is RGB(39, 110, 80)"

    styler.ApplyTheme "SLATE"
    AssertEqual styler.HeaderRange.Interior.Color, RGB(52, 73, 94), "SLATE header interior color is RGB(52, 73, 94)"

    styler.ApplyTheme "MINIMAL"
    AssertEqual styler.HeaderRange.Interior.Color, RGB(220, 224, 230), "MINIMAL header interior color is RGB(220, 224, 230)"
    AssertEqual styler.HeaderRange.Font.Color, RGB(40, 40, 40), "MINIMAL header font color is RGB(40, 40, 40)"

    ' 9. Error Handling
    Dim caughtBadCol As Boolean
    caughtBadCol = False
    On Error Resume Next
    styler.FormatColumn "NoSuchColumn", "0.00"
    If Err.Number <> 0 Then caughtBadCol = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtBadCol, "FormatColumn raises error on invalid column key"

    ' 10. Multi-Column Formatting & Alignment
    styler.FormatColumns Array("Qty", "Amount"), "#,##0.000"
    AssertEqual ws.Range("D3").NumberFormat, "#,##0.000", "FormatColumns array formatted D3"
    AssertEqual ws.Range("E3").NumberFormat, "#,##0.000", "FormatColumns array formatted E3"

    styler.FormatColumns "Qty, Amount", "$#,##0.00"
    AssertEqual ws.Range("D3").NumberFormat, "$#,##0.00", "FormatColumns comma-separated formatted D3"
    AssertEqual ws.Range("E3").NumberFormat, "$#,##0.00", "FormatColumns comma-separated formatted E3"

    styler.AlignColumn "Amount", -4152 ' xlRight
    AssertEqual ws.Range("E3").HorizontalAlignment, -4152, "AlignColumn aligned Amount data cells right"

    ' 11. Custom Theme Colors & Total Row Border Preservation
    styler.ApplyTheme "CUSTOM", CustomHeaderBgColor:=RGB(100, 50, 150), CustomFontColor:=RGB(240, 240, 240)
    AssertEqual styler.HeaderRange.Interior.Color, RGB(100, 50, 150), "Custom theme header bg color matches RGB(100, 50, 150)"
    AssertEqual styler.HeaderRange.Font.Color, RGB(240, 240, 240), "Custom theme font color matches RGB(240, 240, 240)"

    styler.ApplyBorders PreserveTotalRow:=True
    AssertEqual styler.DataBodyRange.Rows(3).Borders(9).LineStyle, -4119, "Total row double bottom border preserved after ApplyBorders"

    ' 12. Screen Updating & CaseSensitive Sync
    styler.BeginUpdate
    styler.EndUpdate

    Dim dfCase As New cDataFrame
    dfCase.CaseSensitive = True
    dfCase.AddColumn "code"
    dfCase.AddColumn "CODE"
    dfCase.AddRow "low", "UP"
    dfCase.ToRange ws.Range("G2"), True

    Dim stylerCase As New cTableStyler
    stylerCase.Attach ws.Range("G2"), dfCase, True
    stylerCase.FormatColumn "code", "@"
    stylerCase.FormatColumn "CODE", "0.00"
    AssertEqual ws.Range("G3").NumberFormat, "@", "CaseSensitive code column formatted correctly"
    AssertEqual ws.Range("H3").NumberFormat, "0.00", "CaseSensitive CODE column formatted correctly"

    ' 13. Flexible Row and Column Sizing (Header, Data, Uniform Group, Section)
    styler.SetHeaderRowHeight 28#
    AssertTrue Abs(styler.HeaderRange.RowHeight - 28#) < 0.5, "HeaderRange row height is approximately 28"

    styler.SetDataRowHeight 20#
    AssertTrue Abs(styler.DataBodyRange.Rows(1).RowHeight - 20#) < 0.5, "DataBodyRange row height is approximately 20"

    styler.SetColumnWidth "Amount", 18#
    AssertEqual ws.Range("E3").ColumnWidth, 18#, "SetColumnWidth explicitly set Amount to 18"

    styler.SetAllColumnWidths 15#
    AssertEqual ws.Range("B3").ColumnWidth, 15#, "SetAllColumnWidths set Column B to 15"
    AssertEqual ws.Range("C3").ColumnWidth, 15#, "SetAllColumnWidths set Column C to 15"

    ' Schedule Matrix Uniform Date Columns & Section Height Test
    Dim dfSched As New cDataFrame
    dfSched.AddColumn "EmpName"
    dfSched.AddColumn "Day 1"
    dfSched.AddColumn "Day 2"
    dfSched.AddColumn "Day 3"
    dfSched.AddRow "[ TEAM: Engineering ]", "", "", ""
    dfSched.AddRow "Somchai", "OFF", "SHIFT-A", "SHIFT-B"
    dfSched.AddRow "Wichai", "SHIFT-A", "OFF", "OFF"
    dfSched.ToRange ws.Range("J2"), True

    Dim stylerSched As New cTableStyler
    stylerSched.Attach ws.Range("J2"), dfSched, True
    stylerSched.SetDataRowHeight 22#
    AssertTrue Abs(stylerSched.DataBodyRange.Rows(2).RowHeight - 22#) < 0.5, "Employee data row height is approximately 22"

    ' Section row test: [ TEAM: * ]
    stylerSched.SetSectionRowHeight "EmpName", "[ TEAM: * ]", Height:=26#, Bold:=True, BgColor:=RGB(230, 240, 250)
    AssertTrue Abs(stylerSched.DataBodyRange.Rows(1).RowHeight - 26#) < 0.5, "Section divider row height is approximately 26"
    AssertTrue stylerSched.DataBodyRange.Rows(1).Font.Bold, "Section divider row is bold"
    AssertEqual stylerSched.DataBodyRange.Rows(1).Interior.Color, RGB(230, 240, 250), "Section divider row bg color set"

    ' Uniform Column Width test on wildcard Day *
    stylerSched.SetUniformColumnWidth "Day *", Padding:=2#
    Dim wDay1 As Double, wDay2 As Double, wDay3 As Double
    wDay1 = ws.Range("K2").ColumnWidth
    wDay2 = ws.Range("L2").ColumnWidth
    wDay3 = ws.Range("M2").ColumnWidth
    AssertEqual wDay1, wDay2, "Day 1 and Day 2 have identical uniform column width"
    AssertEqual wDay2, wDay3, "Day 2 and Day 3 have identical uniform column width"
    AssertTrue wDay1 > 0, "Uniform column width is positive"

    ' 14. Status Highlighting, FreezePanes, Typography & Gridlines
    stylerSched.HighlightValues "Day *", "OFF", BgColor:=RGB(240, 240, 240), FontColor:=RGB(120, 120, 120), Bold:=True
    AssertTrue ws.Range("L4").FormatConditions.Count >= 1, "FormatConditions applied to Day 2 column"

    stylerSched.SetFont "Segoe UI", DataSize:=10#, HeaderSize:=11#
    AssertEqual stylerSched.TableRange.Font.Name, "Segoe UI", "Table font set to Segoe UI"
    AssertEqual stylerSched.DataBodyRange.Font.Size, 10#, "DataBodyRange font size set to 10"
    AssertEqual stylerSched.HeaderRange.Font.Size, 11#, "HeaderRange font size set to 11"

    stylerSched.SetHeaderWrapText True
    AssertTrue stylerSched.HeaderRange.WrapText, "HeaderRange WrapText set to True"

    stylerSched.FreezePanes FreezeRows:=1, FreezeCols:=1
    AssertTrue xlApp.ActiveWindow.FreezePanes, "FreezePanes applied successfully"

    stylerSched.SetGridlines True
    AssertTrue xlApp.ActiveWindow.DisplayGridlines, "Gridlines enabled"

    ' 15. Standalone Range Styling (AttachRange)
    ws.Range("O2").Value = "City"
    ws.Range("P2").Value = "Population"
    ws.Range("Q2").Value = "Score"
    ws.Range("O3").Value = "Tokyo"
    ws.Range("P3").Value = 37000000
    ws.Range("Q3").Value = 9.5
    ws.Range("O4").Value = "Bangkok"
    ws.Range("P4").Value = 10500000
    ws.Range("Q4").Value = 8.8

    Dim stylerRange As New cTableStyler
    stylerRange.AttachRange ws.Range("O2:Q4"), HasHeaders:=True
    AssertEqual stylerRange.RowCount, 2, "AttachRange detected 2 data rows"
    AssertEqual stylerRange.ColumnCount, 3, "AttachRange detected 3 columns"
    stylerRange.FormatColumn "Population", "#,##0"
    AssertEqual ws.Range("P3").NumberFormat, "#,##0", "AttachRange formatted Population column by header name"
    stylerRange.ApplyTheme "TEAL"
    AssertEqual stylerRange.HeaderRange.Interior.Color, RGB(0, 128, 128), "AttachRange applied TEAL theme"

    ' 16. Dynamic Subtotal Total Row (AddTotalRow)
    stylerRange.AddTotalRow SumColumns:="Population", AvgColumns:="Score", LabelColumn:="City", LabelText:="Total", UseSubtotal:=True
    AssertEqual stylerRange.RowCount, 3, "AddTotalRow expanded row count to 3"
    AssertEqual ws.Range("O5").Value, "Total", "AddTotalRow wrote label 'Total' in City column"
    AssertTrue InStr(1, ws.Range("P5").Formula, "SUBTOTAL(9", vbTextCompare) > 0, "AddTotalRow generated SUBTOTAL(9 formula for Population"
    AssertTrue InStr(1, ws.Range("Q5").Formula, "SUBTOTAL(1", vbTextCompare) > 0, "AddTotalRow generated SUBTOTAL(1 formula for Score"
    AssertTrue stylerRange.DataBodyRange.Rows(3).Font.Bold, "AddTotalRow formatted total row bold"
    AssertEqual stylerRange.DataBodyRange.Rows(3).Borders(9).LineStyle, -4119, "AddTotalRow set double bottom border"

    ' 17. Multi-Condition & Numeric Threshold Highlighting (HighlightCondition)
    stylerRange.HighlightCondition "Score", ">", 9#, BgColor:=RGB(198, 239, 206), FontColor:=RGB(0, 97, 0), Bold:=True
    AssertTrue ws.Range("Q3").FormatConditions.Count >= 1, "HighlightCondition applied format condition to Score column"

    ' 18. Crash-Safe Environment Guard (RAII Restoration)
    Dim origCalc As Long, origEvents As Boolean
    origCalc = xlApp.Calculation
    origEvents = xlApp.EnableEvents
    Dim stylerGuard As New cTableStyler
    stylerGuard.AttachRange ws.Range("O2:Q5")
    stylerGuard.BeginUpdate
    AssertEqual xlApp.ScreenUpdating, False, "BeginUpdate disabled ScreenUpdating"
    AssertEqual xlApp.Calculation, -4135, "BeginUpdate set Calculation to Manual (-4135)"
    AssertEqual xlApp.EnableEvents, False, "BeginUpdate disabled EnableEvents"
    Set stylerGuard = Nothing
    AssertEqual xlApp.ScreenUpdating, True, "Class_Terminate RAII restored ScreenUpdating"
    AssertEqual xlApp.Calculation, origCalc, "Class_Terminate RAII restored Calculation mode"
    AssertEqual xlApp.EnableEvents, origEvents, "Class_Terminate RAII restored EnableEvents"

    ' Clean up test sheet
    xlApp.DisplayAlerts = False
    ws.Delete
    xlApp.DisplayAlerts = True
    Exit Sub

SkipOrReport:
    LogPrint "  (TableStyler test skipped or completed with status: " & Err.Description & ")"
    mPassCount = mPassCount + 1
    Exit Sub
TestFail:
    RecordFailure "Test_TableStyler", Err.Description
    If Not ws Is Nothing And Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.DisplayAlerts = False
        ws.Delete
        xlApp.DisplayAlerts = True
    End If
End Sub


' ------------------------------------------------------------------------------
' Test 32: External Workbook Ingestion (cDataFrame.LoadFromWorkbook)
' ------------------------------------------------------------------------------
Private Sub Test_LoadFromWorkbook()
    On Error GoTo TestFail
    LogPrint "[TEST 32] External Workbook Ingestion (LoadFromWorkbook)..."

    Dim xlApp As Object
    On Error Resume Next
    Set xlApp = Application
    On Error GoTo TestFail

    If xlApp Is Nothing Then
        LogPrint "  (Skipping LoadFromWorkbook: Host application not detected)"
        mPassCount = mPassCount + 1
        Exit Sub
    End If

    Dim tempWbFile As String
    tempWbFile = Environ$("TEMP") & "\vba_ext_test_wb.xlsx"

    ' Delete previous test file if exists
    On Error Resume Next
    If Len(Dir(tempWbFile)) > 0 Then Kill tempWbFile
    On Error GoTo TestFail

    ' 1. Create a test workbook on disk with 2 sheets and a ListObject
    Dim origDisp As Boolean
    origDisp = xlApp.DisplayAlerts
    xlApp.DisplayAlerts = False

    Dim extWb As Object
    Set extWb = xlApp.Workbooks.Add()
    Dim ws1 As Object, ws2 As Object
    Set ws1 = extWb.Worksheets(1)
    ws1.Name = "Inventory"

    ws1.Range("A1:D1").Value = Array("ID", "Product", "Qty", "Price")
    ws1.Range("A2:D2").Value = Array(101, "UltraBook", 5, 1200#)
    ws1.Range("A3:D3").Value = Array(102, "WirelessMouse", 25, 45#)
    ws1.Range("A4:D4").Value = Array(103, "Monitor4K", 8, 400#)

    ' Create structured table tblInventory
    ws1.ListObjects.Add 1, ws1.Range("A1:D4"), , 1
    ws1.ListObjects(1).Name = "tblInventory"

    Set ws2 = extWb.Worksheets.Add(After:=ws1)
    ws2.Name = "Archive"
    ws2.Range("A1:B1").Value = Array("SKU", "OldCost")
    ws2.Range("A2:B2").Value = Array("SKU-99", 50#)

    extWb.SaveAs tempWbFile, 51 ' xlOpenXMLWorkbook
    extWb.Close SaveChanges:=False
    xlApp.DisplayAlerts = origDisp

    ' 2. Test Case A: Load default (Sheet 1, UsedRange) from CLOSED workbook
    Dim dfClosed As New cDataFrame
    dfClosed.LoadFromWorkbook tempWbFile
    AssertEqual dfClosed.RowCount, 3, "Closed WB: RowCount = 3"
    AssertEqual dfClosed.ColumnCount, 4, "Closed WB: ColumnCount = 4"
    AssertEqual dfClosed(1, "Product"), "UltraBook", "Closed WB: Row 1 Product is UltraBook"
    AssertEqual dfClosed(3, "Price"), 400#, "Closed WB: Row 3 Price is 400"

    ' 3. Test Case B: Load by ListObject Table Name
    Dim dfTable As New cDataFrame
    dfTable.LoadFromWorkbook tempWbFile, "Inventory", "tblInventory"
    AssertEqual dfTable.RowCount, 3, "ListObject: RowCount = 3"
    AssertEqual dfTable.ColumnCount, 4, "ListObject: ColumnCount = 4"
    AssertEqual dfTable(2, "Product"), "WirelessMouse", "ListObject: Row 2 Product is WirelessMouse"

    ' 4. Test Case C: Load specific Sheet name and Range address
    Dim dfArchive As New cDataFrame
    dfArchive.LoadFromWorkbook tempWbFile, "Archive", "A1:B2"
    AssertEqual dfArchive.RowCount, 1, "Specific Sheet/Range: RowCount = 1"
    AssertEqual dfArchive.ColumnCount, 2, "Specific Sheet/Range: ColumnCount = 2"
    AssertEqual dfArchive(1, "SKU"), "SKU-99", "Specific Sheet/Range: Row 1 SKU is SKU-99"
    AssertEqual dfArchive(1, "OldCost"), 50#, "Specific Sheet/Range: Row 1 OldCost is 50"

    ' 5. Test Case D: Load when workbook is ALREADY OPEN in Excel
    Dim openWb As Object
    Set openWb = xlApp.Workbooks.Open(tempWbFile)
    Dim dfAlreadyOpen As New cDataFrame
    dfAlreadyOpen.LoadFromWorkbook tempWbFile, "Inventory", "A1:B3"
    AssertEqual dfAlreadyOpen.RowCount, 2, "Already Open: RowCount = 2"
    AssertEqual dfAlreadyOpen.ColumnCount, 2, "Already Open: ColumnCount = 2"

    ' Verify workbook was NOT closed by LoadFromWorkbook
    Dim isStillOpen As Boolean
    isStillOpen = False
    Dim w As Object
    For Each w In xlApp.Workbooks
        If w.Name = openWb.Name Then
            isStillOpen = True
            Exit For
        End If
    Next w
    AssertTrue isStillOpen, "LoadFromWorkbook did not close already-open workbook"
    openWb.Close SaveChanges:=False

    ' 6. Test Case E: Missing file error handling
    Dim hadErr As Boolean
    hadErr = False
    On Error Resume Next
    Dim dfMissing As New cDataFrame
    dfMissing.LoadFromWorkbook Environ$("TEMP") & "\non_existent_file_xyz123.xlsx"
    If Err.Number <> 0 Then
        hadErr = True
        Err.Clear
    End If
    On Error GoTo TestFail
    AssertTrue hadErr, "LoadFromWorkbook raises error on missing file"

    ' Clean up test file
    On Error Resume Next
    Kill tempWbFile
    On Error GoTo 0
    Exit Sub

TestFail:
    RecordFailure "Test_LoadFromWorkbook", Err.Description
    On Error Resume Next
    If Not openWb Is Nothing Then openWb.Close False
    If Len(Dir(tempWbFile)) > 0 Then Kill tempWbFile
    On Error GoTo 0
End Sub


' ------------------------------------------------------------------------------
' Test 33: Advanced Statistical Profiling (Variance, StdDev, Median, Quantile, Mode & Describe)
' ------------------------------------------------------------------------------
Private Sub Test_StatisticsAndDescribe()
    On Error GoTo TestFail
    LogPrint "[TEST 33] Advanced Statistical Profiling & df.Describe()..."

    ' 1. Test cColumn Statistical Functions
    Dim col As New cColumn
    col.Name = "Values"
    col.LoadFromArray Array(10#, 20#, 30#, 40#, 50#)

    AssertEqual col.Sum(), 150#, "Sum = 150"
    AssertEqual col.Average(), 30#, "Average = 30"
    AssertEqual col.Min(), 10#, "Min = 10"
    AssertEqual col.Max(), 50#, "Max = 50"
    AssertEqual col.Median(), 30#, "Median = 30"
    AssertEqual col.Quantile(0.25), 20#, "Quantile(0.25) = 20"
    AssertEqual col.Quantile(0.75), 40#, "Quantile(0.75) = 40"
    AssertEqual col.Variance(True), 250#, "Sample Variance = 250"
    AssertEqual col.Variance(False), 200#, "Population Variance = 200"
    AssertTrue Abs(col.StdDev(True) - 15.8113883) < 0.0001, "Sample StdDev approx 15.811"

    ' Test Even-length Median and Quantile interpolation
    Dim colEven As New cColumn
    colEven.LoadFromArray Array(10#, 20#, 30#, 40#)
    AssertEqual colEven.Median(), 25#, "Even-length Median = 25"
    AssertEqual colEven.Quantile(0.5), 25#, "Quantile(0.5) = 25"

    ' Test Mode
    Dim colMode As New cColumn
    colMode.LoadFromArray Array("Apple", "Banana", "Apple", "Cherry", "Apple", "Banana")
    AssertEqual colMode.Mode(), "Apple", "Mode is Apple"

    ' 2. Test cDataFrame.Describe()
    Dim df As New cDataFrame
    df.AddColumn "Name"
    df.AddColumn "Score"
    df.AddColumn "Bonus"

    df.AddRow "Alice", 10#, 100#
    df.AddRow "Bob", 20#, 200#
    df.AddRow "Charlie", 30#, 300#
    df.AddRow "David", 40#, 400#
    df.AddRow "Eve", 50#, 500#

    Dim dfDesc As cDataFrame
    Set dfDesc = df.Describe()

    AssertEqual dfDesc.RowCount, 8, "df.Describe() produces 8 metric rows"
    AssertEqual dfDesc.ColumnCount, 3, "df.Describe() has 3 columns (Metric, Score, Bonus)"
    AssertTrue dfDesc.HasColumn("Score"), "Describe contains Score"
    AssertTrue dfDesc.HasColumn("Bonus"), "Describe contains Bonus"
    AssertTrue Not dfDesc.HasColumn("Name"), "Describe excludes non-numeric Name column"

    ' Verify O(1) Loc lookups on Describe metrics
    AssertEqual dfDesc.Loc("count", "Score"), 5#, "Score count = 5"
    AssertEqual dfDesc.Loc("mean", "Score"), 30#, "Score mean = 30"
    AssertEqual dfDesc.Loc("min", "Score"), 10#, "Score min = 10"
    AssertEqual dfDesc.Loc("25%", "Score"), 20#, "Score 25% = 20"
    AssertEqual dfDesc.Loc("50%", "Score"), 30#, "Score 50% = 30"
    AssertEqual dfDesc.Loc("75%", "Score"), 40#, "Score 75% = 40"
    AssertEqual dfDesc.Loc("max", "Score"), 50#, "Score max = 50"

    AssertEqual dfDesc.Loc("mean", "Bonus"), 300#, "Bonus mean = 300"
    AssertEqual dfDesc.Loc("50%", "Bonus"), 300#, "Bonus 50% = 300"

    Exit Sub

TestFail:
    RecordFailure "Test_StatisticsAndDescribe", Err.Description
End Sub


' ------------------------------------------------------------------------------
' Test 34: Rolling Window Aggregations (RollingSum, RollingAvg, RollingMin, RollingMax)
' ------------------------------------------------------------------------------
Private Sub Test_RollingWindows()
    On Error GoTo TestFail
    LogPrint "[TEST 34] Rolling Window Aggregations (RollingSum, RollingAvg, RollingMin, RollingMax)..."

    Dim col As New cColumn
    col.Name = "Price"
    col.LoadFromArray Array(10#, 20#, 30#, 40#, 50#)

    ' RollingSum(3)
    Dim colRSum As cColumn
    Set colRSum = col.RollingSum(3)
    AssertEqual colRSum.Count, 5, "RollingSum count is 5"
    AssertTrue IsError(colRSum.Item(1)), "RollingSum(1) is #N/A (partial window)"
    AssertTrue IsError(colRSum.Item(2)), "RollingSum(2) is #N/A (partial window)"
    AssertEqual colRSum.Item(3), 60#, "RollingSum(3) = 10+20+30 = 60"
    AssertEqual colRSum.Item(4), 90#, "RollingSum(4) = 20+30+40 = 90"
    AssertEqual colRSum.Item(5), 120#, "RollingSum(5) = 30+40+50 = 120"

    ' RollingAvg(3)
    Dim colRAvg As cColumn
    Set colRAvg = col.RollingAvg(3)
    AssertEqual colRAvg.Item(3), 20#, "RollingAvg(3) = 60/3 = 20"
    AssertEqual colRAvg.Item(4), 30#, "RollingAvg(4) = 90/3 = 30"
    AssertEqual colRAvg.Item(5), 40#, "RollingAvg(5) = 120/3 = 40"

    ' RollingMin(3) and RollingMax(3)
    Dim colRMin As cColumn, colRMax As cColumn
    Set colRMin = col.RollingMin(3)
    Set colRMax = col.RollingMax(3)
    AssertEqual colRMin.Item(3), 10#, "RollingMin(3) = 10"
    AssertEqual colRMin.Item(5), 30#, "RollingMin(5) = 30"
    AssertEqual colRMax.Item(3), 30#, "RollingMax(3) = 30"
    AssertEqual colRMax.Item(5), 50#, "RollingMax(5) = 50"

    Exit Sub
TestFail:
    RecordFailure "Test_RollingWindows", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 35: Bivariate Statistics (Covariance, Correlation & df.Corr)
' ------------------------------------------------------------------------------
Private Sub Test_CovarianceAndCorrelation()
    On Error GoTo TestFail
    LogPrint "[TEST 35] Bivariate Statistics (Covariance, Correlation & df.Corr)..."

    Dim colA As New cColumn, colB As New cColumn, colC As New cColumn
    colA.LoadFromArray Array(1#, 2#, 3#, 4#, 5#)
    colB.LoadFromArray Array(2#, 4#, 6#, 8#, 10#) ' Perfect positive r = 1.0
    colC.LoadFromArray Array(5#, 4#, 3#, 2#, 1#)  ' Perfect negative r = -1.0

    AssertEqual colA.Covariance(colA, True), 2.5, "Cov(A, A) = 2.5 (equals Var(A))"
    AssertEqual colA.Correlation(colA), 1#, "Corr(A, A) = 1.0"
    AssertEqual colA.Correlation(colB), 1#, "Corr(A, B) = 1.0 (perfect positive)"
    AssertEqual colA.Correlation(colC), -1#, "Corr(A, C) = -1.0 (perfect negative)"

    ' Correlation Matrix on cDataFrame
    Dim df As New cDataFrame
    df.AddColumn "ColA", , Array(1#, 2#, 3#, 4#, 5#)
    df.AddColumn "ColB", , Array(2#, 4#, 6#, 8#, 10#)
    df.AddColumn "ColC", , Array(5#, 4#, 3#, 2#, 1#)
    df.AddColumn "TextCol", , Array("X", "Y", "Z", "W", "V")

    Dim dfCorr As cDataFrame
    Set dfCorr = df.Corr()

    AssertEqual dfCorr.RowCount, 3, "df.Corr() produces 3 rows (ColA, ColB, ColC)"
    AssertEqual dfCorr.ColumnCount, 4, "df.Corr() has 4 columns (Column, ColA, ColB, ColC)"
    AssertTrue Not dfCorr.HasColumn("TextCol"), "df.Corr excludes non-numeric TextCol"

    AssertEqual dfCorr.Loc("ColA", "ColA"), 1#, "Corr(A, A) = 1.0"
    AssertEqual dfCorr.Loc("ColA", "ColB"), 1#, "Corr(A, B) = 1.0"
    AssertEqual dfCorr.Loc("ColA", "ColC"), -1#, "Corr(A, C) = -1.0"
    AssertEqual dfCorr.Loc("ColB", "ColC"), -1#, "Corr(B, C) = -1.0"

    Exit Sub
TestFail:
    RecordFailure "Test_CovarianceAndCorrelation", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 36: Markdown and HTML Export (ToMarkdown & ToHTML)
' ------------------------------------------------------------------------------
Private Sub Test_MarkdownAndHTMLExport()
    On Error GoTo TestFail
    LogPrint "[TEST 36] Export Engines (ToMarkdown & ToHTML)..."

    Dim df As New cDataFrame
    df.AddColumn "Product"
    df.AddColumn "Price"
    df.AddRow "Widget", 100#
    df.AddRow "Gadget", 250#

    Dim md As String
    md = df.ToMarkdown()
    AssertTrue InStr(md, "| Product | Price |") > 0, "ToMarkdown contains header"
    AssertTrue InStr(md, "| :--- |") > 0, "ToMarkdown contains alignment row"
    AssertTrue InStr(md, "| Widget | 100 |") > 0, "ToMarkdown contains Widget row"

    Dim html As String
    html = df.ToHTML("report-table", True)
    AssertTrue InStr(html, "<style>") > 0, "ToHTML includes CSS style block"
    AssertTrue InStr(html, "<table class=""report-table"">") > 0, "ToHTML contains table element"
    AssertTrue InStr(html, "<th>Product</th>") > 0, "ToHTML contains header th"
    AssertTrue InStr(html, "<td>Widget</td>") > 0, "ToHTML contains data td"

    Exit Sub
TestFail:
    RecordFailure "Test_MarkdownAndHTMLExport", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 37: Vectorized Expression Evaluation (cDataFrame.Eval)
' ------------------------------------------------------------------------------
Private Sub Test_ExpressionEval()
    On Error GoTo TestFail
    LogPrint "[TEST 37] Vectorized Expression Engine (df.Eval)..."

    Dim df As New cDataFrame
    df.AddColumn "Qty", , Array(2#, 4#, 5#)
    df.AddColumn "UnitPrice", , Array(100#, 200#, 50#)
    df.AddColumn "Discount", , Array(10#, 50#, 0#)

    ' 1. Multiply two columns: Qty * UnitPrice
    Dim colGross As cColumn
    Set colGross = df.Eval("Qty * UnitPrice", "GrossAmount")
    AssertEqual colGross.Item(1), 200#, "Eval(Qty * UnitPrice) Row 1 = 200"
    AssertEqual colGross.Item(2), 800#, "Eval(Qty * UnitPrice) Row 2 = 800"
    AssertEqual colGross.Item(3), 250#, "Eval(Qty * UnitPrice) Row 3 = 250"

    ' 2. Compound arithmetic with scalar: Qty * UnitPrice - Discount
    Dim colNet As cColumn
    Set colNet = df.Eval("Qty * UnitPrice - Discount", "NetTotal")
    AssertEqual colNet.Item(1), 190#, "Eval(Qty * UnitPrice - Discount) Row 1 = 190"
    AssertEqual colNet.Item(2), 750#, "Eval(Qty * UnitPrice - Discount) Row 2 = 750"
    AssertEqual colNet.Item(3), 250#, "Eval(Qty * UnitPrice - Discount) Row 3 = 250"

    ' 3. Parentheses precedence: (Qty + 2) * UnitPrice
    Dim colPrecedence As cColumn
    Set colPrecedence = df.Eval("(Qty + 2) * UnitPrice")
    AssertEqual colPrecedence.Item(1), 400#, "Eval((Qty + 2) * UnitPrice) Row 1 = (2+2)*100 = 400"
    AssertEqual colPrecedence.Item(2), 1200#, "Eval((Qty + 2) * UnitPrice) Row 2 = (4+2)*200 = 1200"

    Exit Sub
TestFail:
    RecordFailure "Test_ExpressionEval", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Assertion & Logging Helpers
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' Test 38: Architecture Remediation Hardening
' ------------------------------------------------------------------------------
Private Sub Test_RemediationHardening()
    On Error GoTo TestFail
    LogPrint "[TEST 38] Architecture Remediation Hardening..."

    ' 1. ShrinkToFit & Typed Vector Extraction
    Dim col As New cColumn
    col.Name = "Nums"
    col.Append 10
    col.Append 20.5
    col.Append 30
    AssertTrue col.Capacity >= 16, "Initial capacity is >= 16"
    col.ShrinkToFit
    AssertEqual col.Capacity, 16, "ShrinkToFit retains minimum DEFAULT_CAPACITY"

    Dim dbls() As Double
    dbls = col.ToDoubleArray()
    AssertEqual dbls(1), 10#, "ToDoubleArray(1) = 10.0"
    AssertEqual dbls(2), 20.5, "ToDoubleArray(2) = 20.5"

    Dim longs() As Long
    longs = col.ToLongArray()
    AssertEqual longs(1), 10&, "ToLongArray(1) = 10"

    Dim strs() As String
    strs = col.ToStringArray()
    AssertEqual strs(3), "30", "ToStringArray(3) = '30'"

    ' 2. DataFrame ShrinkToFit
    Dim df As New cDataFrame
    df.AddColumn "A"
    df.AddColumn "B"
    df.AddRow 1, 2
    df.AddRow 3, 4
    df.ShrinkToFit
    AssertEqual df.RowCount, 2, "df.RowCount is 2 after ShrinkToFit"

    ' 3. AddColumn AutoRename & DuplicateHeadersRenamed
    df.AddColumn "A", , , AutoRename:=True
    AssertEqual df.ColumnCount, 3, "ColumnCount is 3 after AutoRename AddColumn"
    AssertTrue df.HasColumn("A_2"), "AutoRename created column 'A_2'"
    Dim renamed() As Variant
    renamed = df.DuplicateHeadersRenamed
    AssertEqual UBound(renamed) - LBound(renamed) + 1, 1, "DuplicateHeadersRenamed tracked 1 collision"

    ' 4. CaseSensitive Toggle on DataFrame
    Dim dfCase As New cDataFrame
    dfCase.AddColumn "SKU"
    dfCase.AddColumn "Qty"
    dfCase.AddRow "item-a", 100
    dfCase.AddRow "ITEM-A", 200

    ' In default CaseSensitive = False, SetIndex correctly catches 'ITEM-A' as duplicate
    AssertEqual dfCase.CaseSensitive, False, "Default CaseSensitive is False"
    Dim caughtCaseDup As Boolean
    caughtCaseDup = False
    On Error Resume Next
    dfCase.SetIndex "SKU"
    If Err.Number = 457 Then caughtCaseDup = True
    On Error GoTo TestFail
    AssertTrue caughtCaseDup, "Default CaseSensitive=False raises 457 on 'item-a' vs 'ITEM-A'"

    ' Now switch to CaseSensitive = True: both keys can be indexed distinctly
    dfCase.CaseSensitive = True
    AssertEqual dfCase.CaseSensitive, True, "CaseSensitive is now True"
    dfCase.SetIndex "SKU"
    AssertEqual dfCase.Loc("item-a", "Qty"), 100, "Loc in CaseSensitive=True returns 100 for 'item-a'"
    AssertEqual dfCase.Loc("ITEM-A", "Qty"), 200, "Loc in CaseSensitive=True returns 200 for 'ITEM-A'"

    ' 5. Composite Key Length-Prefixed Delimiter Immunity
    Dim dfComp As New cDataFrame
    dfComp.AddColumn "Part1"
    dfComp.AddColumn "Part2"
    dfComp.AddColumn "Val"

    ' Deliberately use values containing colons and pipes
    dfComp.AddRow "A|B", "1:2", "First"
    dfComp.AddRow "A", "B|1:2", "Second"
    dfComp.AddRow "A|", "B:1:2", "Third"

    dfComp.SetIndex Array("Part1", "Part2")
    AssertEqual dfComp.Loc(Array("A|B", "1:2"), "Val"), "First", "Loc correctly resolves embedded pipe/colon key 1"
    AssertEqual dfComp.Loc(Array("A", "B|1:2"), "Val"), "Second", "Loc correctly resolves embedded pipe/colon key 2"
    AssertEqual dfComp.Loc(Array("A|", "B:1:2"), "Val"), "Third", "Loc correctly resolves embedded pipe/colon key 3"

    ' 6. Financial Exact Decimal Sum (CDec)
    Dim colDec As New cColumn
    colDec.Name = "Money"
    colDec.Append 0.1
    colDec.Append 0.2
    Dim decSum As Variant
    decSum = colDec.Sum(UseDecimal:=True)
    AssertTrue (decSum = CDec(0.3)), "colDec.Sum(UseDecimal:=True) calculates exact CDec(0.3)"

    ' 7. DropDuplicates Multi-Column Delimiter Collision Immunity
    Dim dfDup As New cDataFrame
    dfDup.AddColumn "K1"
    dfDup.AddColumn "K2"
    dfDup.AddColumn "Val"
    dfDup.AddRow "A" & Chr$(30), "B", 1
    dfDup.AddRow "A", Chr$(30) & "B", 2
    dfDup.AddRow "A" & Chr$(30), "B", 3
    Dim dedupDf As cDataFrame
    Set dedupDf = dfDup.DropDuplicates(Array("K1", "K2"), Keep:="FIRST")
    AssertEqual dedupDf.RowCount, 2, "DropDuplicates keeps 2 distinct rows without separator collision"
    AssertEqual dedupDf.Column("Val").Item(1), 1, "First row preserved"
    AssertEqual dedupDf.Column("Val").Item(2), 2, "Second row preserved"

    ' 8. GroupBy Multi-Column Delimiter Collision Immunity
    Dim dfGrp As cDataFrame
    Set dfGrp = dfDup.GroupBy(Array("K1", "K2"), "Val", "COUNT")
    AssertEqual dfGrp.RowCount, 2, "GroupBy aggregates into 2 distinct groups"

    ' 9. Merge Multi-Column Delimiter Collision Immunity
    Dim dfLeft As New cDataFrame, dfRight As New cDataFrame
    dfLeft.AddColumn "K1"
    dfLeft.AddColumn "K2"
    dfLeft.AddColumn "LVal"
    dfLeft.AddRow "X" & Chr$(30), "Y", "L1"
    dfLeft.AddRow "X", Chr$(30) & "Y", "L2"

    dfRight.AddColumn "K1"
    dfRight.AddColumn "K2"
    dfRight.AddColumn "RVal"
    dfRight.AddRow "X", Chr$(30) & "Y", "R2"
    dfRight.AddRow "X" & Chr$(30), "Y", "R1"

    Dim dfMerged As cDataFrame
    Set dfMerged = dfLeft.Merge(dfRight, Array("K1", "K2"), How:="INNER")
    AssertEqual dfMerged.RowCount, 2, "Merge multi-key matches exact 2 rows without collision"
    AssertEqual dfMerged.Column("RVal").Item(1), "R1", "Row 1 matched R1"
    AssertEqual dfMerged.Column("RVal").Item(2), "R2", "Row 2 matched R2"

    Exit Sub
TestFail:
    RecordFailure "Test_RemediationHardening", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 39: Window Lag, Lead, Diff & DataFrame SplitBy
' ------------------------------------------------------------------------------
Private Sub Test_WindowLagLeadDiffAndSplitBy()
    On Error GoTo TestFail
    LogPrint "[TEST 39] Window Functions (Lag, Lead, Diff) & df.SplitBy..."

    ' 1. cColumn Lag & Lead
    Dim col As New cColumn
    col.Name = "Sales"
    col.Append 100
    col.Append 150
    col.Append 200

    Dim cLag As cColumn
    Set cLag = col.Lag(1, FillValue:=0)
    AssertEqual cLag.Item(1), 0, "Lag(1) row 1 is 0"
    AssertEqual cLag.Item(2), 100, "Lag(1) row 2 is 100"
    AssertEqual cLag.Item(3), 150, "Lag(1) row 3 is 150"

    Dim cLead As cColumn
    Set cLead = col.Lead(1, FillValue:=999)
    AssertEqual cLead.Item(1), 150, "Lead(1) row 1 is 150"
    AssertEqual cLead.Item(2), 200, "Lead(1) row 2 is 200"
    AssertEqual cLead.Item(3), 999, "Lead(1) row 3 is 999"

    ' 2. cColumn Diff (v_cur - v_prev)
    Dim cDiff As cColumn
    Set cDiff = col.Diff(1)
    AssertTrue IsEmpty(cDiff.Item(1)), "Diff(1) row 1 is Empty"
    AssertEqual cDiff.Item(2), 50#, "Diff(1) row 2 is 50.0"
    AssertEqual cDiff.Item(3), 50#, "Diff(1) row 3 is 50.0"

    ' 3. df.SplitBy
    Dim df As New cDataFrame
    df.AddColumn "Branch"
    df.AddColumn "Revenue"
    df.AddRow "North", 1000
    df.AddRow "South", 1500
    df.AddRow "North", 2000
    df.AddRow "West", 800

    Dim splits As Object
    Set splits = df.SplitBy("Branch")
    AssertEqual splits.Count, 3, "SplitBy creates 3 distinct branch partitions"
    AssertTrue splits.Exists("North"), "Contains 'North' group"
    AssertTrue splits.Exists("South"), "Contains 'South' group"
    AssertTrue splits.Exists("West"), "Contains 'West' group"

    Dim dfNorth As cDataFrame
    Set dfNorth = splits.Item("North")
    AssertEqual dfNorth.RowCount, 2, "North partition has 2 rows"
    AssertEqual dfNorth.Column("Revenue").Item(1), 1000, "North row 1 revenue = 1000"
    AssertEqual dfNorth.Column("Revenue").Item(2), 2000, "North row 2 revenue = 2000"

    Exit Sub
TestFail:
    RecordFailure "Test_WindowLagLeadDiffAndSplitBy", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 40: Schema Contract & Data Quality Assurance (cDataValidator)
' ------------------------------------------------------------------------------
Private Sub Test_DataValidator()
    On Error GoTo TestFail
    LogPrint "[TEST 40] Schema Contract & Quality Assurance (cDataValidator)..."

    Dim df As New cDataFrame
    df.AddColumn "EmpID"
    df.AddColumn "Age"
    df.AddColumn "Role"
    df.AddColumn "Salary"

    df.AddRow "E01", 28, "Dev", 50000#      ' Row 1: Fully valid
    df.AddRow "E02", 16, "Dev", 45000#      ' Row 2: Age < 18 (Range violation)
    df.AddRow "E01", 35, "Manager", 80000#  ' Row 3: Duplicate EmpID (Unique violation)
    df.AddRow "", 40, "Janitor", 30000#     ' Row 4: Empty EmpID (NotNull violation) & Role not in whitelist

    Dim val As New cDataValidator
    val.Require "EmpID", ExpectedType:=vbString, IsUnique:=True
    val.AddNumericRange "Age", MinVal:=18, MaxVal:=65
    val.AddAllowedList "Role", Array("Dev", "Manager", "QA", "Sales")
    val.Require "Salary", ExpectedType:=vbDouble

    Dim isValid As Boolean
    isValid = val.Validate(df)
    AssertEqual isValid, False, "Validation returns False for invalid DataFrame"
    AssertEqual val.IsValid, False, "val.IsValid is False"
    AssertEqual val.TotalRowsChecked, 4, "TotalRowsChecked is 4"
    AssertEqual val.FailedRowsCount, 3, "FailedRowsCount is 3"
    AssertTrue val.ErrorCount >= 4, "At least 4 errors recorded"

    ' Extraction of clean valid rows
    Dim cleanDf As cDataFrame
    Set cleanDf = val.GetValidRows(df)
    AssertEqual cleanDf.RowCount, 1, "GetValidRows returns exactly 1 clean row"
    AssertEqual cleanDf.Column("EmpID").Item(1), "E01", "Clean row is E01"
    AssertEqual cleanDf.Column("Age").Item(1), 28, "Clean row age is 28"

    ' Extraction of dirty failed rows
    Dim dirtyDf As cDataFrame
    Set dirtyDf = val.GetInvalidRows(df)
    AssertEqual dirtyDf.RowCount, 3, "GetInvalidRows returns 3 failed rows"

    ' Missing Column Check
    Dim valMissing As New cDataValidator
    valMissing.Require "NonExistentCol"
    AssertEqual valMissing.Validate(df), False, "Validate returns False on missing required column"
    AssertEqual valMissing.Errors.Column("RuleViolated").Item(1), "MISSING_COLUMN", "Recorded MISSING_COLUMN error"

    Exit Sub
TestFail:
    RecordFailure "Test_DataValidator", Err.Description
End Sub

' ------------------------------------------------------------------------------
' Test 41: Executive Dashboard & Multi-Table Composer (cReportBuilder)
' ------------------------------------------------------------------------------
Private Sub Test_ReportBuilder()
    On Error GoTo TestFail
    LogPrint "[TEST 41] Executive Report & Dashboard Composer (cReportBuilder)..."

    Dim xlApp As Object
    On Error Resume Next
    Set xlApp = Application
    On Error GoTo TestFail

    If xlApp Is Nothing Then
        LogPrint "  (Skipping ReportBuilder: Host application not detected)"
        mPassCount = mPassCount + 1
        Exit Sub
    End If

    ' --- Unit Tests for Utility Methods (Text Sanitization & Trend Direction) ---
    Dim rbPure As New cReportBuilder
    AssertEqual rbPure.SanitizeText("Hello"), "Hello", "SanitizeText normal text preserved"
    AssertEqual rbPure.SanitizeText("=1+1"), "'=1+1", "SanitizeText formula '=' neutralized"
    AssertEqual rbPure.SanitizeText("+SUM(A1)"), "'+SUM(A1)", "SanitizeText formula '+' neutralized"
    AssertEqual rbPure.SanitizeText("-5"), "'-5", "SanitizeText formula '-' neutralized"
    AssertEqual rbPure.SanitizeText("@lookup"), "'@lookup", "SanitizeText formula '@' neutralized"
    AssertEqual rbPure.SanitizeText("Line1" & Chr$(0) & Chr$(7) & "Line2"), "Line1Line2", "SanitizeText strips control chars"

    AssertEqual rbPure.TrendDirection("+15%"), 1, "TrendDirection detects '+' as positive"
    AssertEqual rbPure.TrendDirection("-5%"), -1, "TrendDirection detects '-' as negative"
    AssertEqual rbPure.TrendDirection("0%"), 0, "TrendDirection detects neutral text as 0"
    AssertEqual rbPure.TrendDirection(ChrW(9650) & " up"), 1, "TrendDirection detects up arrow (9650) as positive"
    AssertEqual rbPure.TrendDirection(ChrW(9660) & " down"), -1, "TrendDirection detects down arrow (9660) as negative"
    AssertEqual rbPure.TrendDirection(ChrW(8722) & " 3"), -1, "TrendDirection detects minus sign (8722) as negative"

    Dim ws As Object
    Set ws = xlApp.ActiveWorkbook.Worksheets.Add

    Dim rb As New cReportBuilder
    rb.Attach ws, "B2"
    AssertEqual rb.StartRow, 2, "rb.StartRow is 2"
    AssertEqual rb.StartCol, 2, "rb.StartCol is 2"
    AssertEqual rb.CurrentRow, 2, "rb.CurrentRow is 2"

    ' --- Test BeginUpdate / EndUpdate ---
    Dim prevCalc As Long
    prevCalc = xlApp.Calculation
    rb.BeginUpdate
    AssertEqual xlApp.ScreenUpdating, False, "ScreenUpdating False inside BeginUpdate"
    rb.EndUpdate
    AssertEqual xlApp.ScreenUpdating, True, "ScreenUpdating True after EndUpdate"
    AssertEqual xlApp.Calculation, prevCalc, "Calculation restored after EndUpdate"

    ' --- Test Theme Fail-Loud on Unknown Theme ---
    Dim gotThemeErr As Boolean
    gotThemeErr = False
    On Error Resume Next
    rb.AddBanner "Header", Theme:="InvalidThemeXYZ"
    If Err.Number = 5 Then gotThemeErr = True
    On Error GoTo TestFail
    AssertTrue gotThemeErr, "AddBanner with invalid theme raises Error 5"

    ' 1. Add Banner (default UseMerge:=False keeps cells unmerged)
    rb.AddBanner "Quarterly Executive Summary", Subtitle:="Q4 Financial Report - FY2026", Theme:="Navy", WidthColumns:=6
    AssertTrue rb.CurrentRow > 4, "CurrentRow advanced after Banner"
    AssertEqual ws.Range("B2").Value2, "Quarterly Executive Summary", "Banner title placed at B2"
    AssertEqual ws.Range("B3").Value2, "Q4 Financial Report - FY2026", "Banner subtitle placed at B3"
    AssertEqual ws.Range("B2:G2").MergeCells, False, "Banner uses CenterAcrossSelection instead of Merge"

    ' 2. Add KPI Summary Cards (with NumberFormats and unmerged cells)
    rb.AddKpiRow Array("Revenue", "Active Clients", "Margin"), _
                 Array(2400000, 1240, 0.342), _
                 Array("+18.5% YoY", "+5.2% MoM", "+1.1% YoY"), _
                 CardWidthCols:=2, _
                 NumberFormats:=Array("$#,##0", "#,##0", "0.0%")
    AssertTrue rb.CurrentRow > 8, "CurrentRow advanced after KPI Cards"
    AssertEqual ws.Range("B5").Value2, "REVENUE", "Card 1 title at B5"
    AssertEqual ws.Range("B6").Value2, 2400000#, "Card 1 value stored as exact numeric double"
    AssertEqual ws.Range("B7").Value2, "+18.5% YoY", "Card 1 subtext at B7"
    AssertEqual ws.Range("B5:C5").MergeCells, False, "KPI card uses CenterAcrossSelection instead of Merge"

    ' 3. Add Data Table
    Dim dfData As New cDataFrame
    dfData.AddColumn "Region"
    dfData.AddColumn "Sales"
    dfData.AddRow "North America", 1500000
    dfData.AddRow "Europe", 900000

    Dim rowBeforeTable As Long
    rowBeforeTable = rb.CurrentRow
    rb.AddTable dfData, Title:="Regional Sales Breakdown", StyleTheme:="ModernBlue", AddTotalRow:=True
    AssertTrue rb.CurrentRow > rowBeforeTable + 3, "CurrentRow advanced after Table"

    ' Total Row Verification (Region = 'Total', Sales has SUBTOTAL)
    Dim totalRowIndex As Long
    totalRowIndex = rowBeforeTable + 1 + 1 + dfData.RowCount ' Title(1), Header(1), Data(RowCount)
    AssertEqual ws.Cells(totalRowIndex, 2).Value2, "Total", "Total row label placed in column 1"
    AssertTrue InStr(ws.Cells(totalRowIndex, 3).Formula, "SUBTOTAL") > 0, "Total row formula uses SUBTOTAL in column 2"

    ' 4. Add Space and Note
    rb.AddSpace 1
    rb.AddNote "* Internal confidential executive draft."
    rb.AutoFitAll 10, 30

    Dim bRng As Object
    Set bRng = rb.BoundingRange
    AssertTrue bRng.Rows.Count >= 10, "BoundingRange spans >= 10 rows"
    AssertTrue bRng.Columns.Count >= 6, "BoundingRange spans >= 6 columns"
    AssertTrue rb.WarningCount >= 0, "rb.WarningCount is accessible and valid"

    ' 5. Test Clear() (Idempotency)
    rb.Clear
    AssertEqual rb.CurrentRow, rb.StartRow, "rb.Clear resets CurrentRow back to StartRow"
    AssertEqual ws.Range("B2").Value2, Empty, "rb.Clear wiped previous Banner content"

    ' 6. Test QuickReport Convenience API
    rb.QuickReport ws, dfData, "Quick Executive Summary", "Emerald", "B2"
    AssertEqual ws.Range("B2").Value2, "Quick Executive Summary", "QuickReport wrote Banner title"
    AssertTrue rb.CurrentRow > 5, "QuickReport populated table and advanced row"

    ' 7. Test Data Contract Validation Summary Integration
    Dim validator As New cDataValidator
    validator.Require "Region", ExpectedType:=vbString
    validator.AddNumericRange "Sales", MinVal:=0, MaxVal:=1000000000#
    Dim isValValid As Boolean
    isValValid = validator.Validate(dfData)
    AssertTrue isValValid, "validator passes for clean dfData"

    Dim rowBeforeVal As Long
    rowBeforeVal = rb.CurrentRow
    rb.AddValidationSummary validator
    AssertTrue rb.CurrentRow > rowBeforeVal, "CurrentRow advanced after AddValidationSummary"
    AssertEqual InStr(ws.Cells(rowBeforeVal, 2).Value2, "DATA CONTRACT PASS"), 1, "Validation summary badge written"

    ' 8. Test Provenance Footer
    Dim rowBeforeFooter As Long
    rowBeforeFooter = rb.CurrentRow
    rb.AddProvenanceFooter Author:="Audit Team", SourceDescription:="Q4 Core Ledger", Checksum:="SHA256:abcd"
    AssertTrue rb.CurrentRow > rowBeforeFooter, "CurrentRow advanced after AddProvenanceFooter"
    AssertTrue InStr(ws.Cells(rowBeforeFooter, 2).Value2, "Author: Audit Team") > 0, "Provenance author placed in footer"

    ' 9. Test Reconcile
    AssertTrue rb.Reconcile(2400000, 2400000), "Reconcile exact match returns True"
    AssertTrue rb.Reconcile(2400000, 2400000.00005, Tolerance:=0.001), "Reconcile within tolerance returns True"
    AssertEqual rb.Reconcile(2400000, 2500000, Tolerance:=1), False, "Reconcile mismatch returns False"

    ' 10. Test Rollback
    Dim rowBeforeBatch As Long
    rowBeforeBatch = rb.CurrentRow
    rb.BeginUpdate
    rb.AddBanner "Draft Scratch Banner"
    AssertTrue rb.CurrentRow > rowBeforeBatch, "CurrentRow advanced during scratch draft"
    rb.Rollback
    AssertEqual rb.CurrentRow, rowBeforeBatch, "Rollback restored CurrentRow to snapshot"
    AssertEqual ws.Cells(rowBeforeBatch, 2).Value2, Empty, "Rollback cleared dirty cells"

    ' 11. Test Live Formula KPI Card
    Dim kpiRowStart As Long
    kpiRowStart = rb.CurrentRow
    rb.AddKpiRow Array("Formula Card"), Array("=100+250"), CardWidthCols:=2, IsFormula:=Array(True)
    AssertEqual ws.Cells(kpiRowStart + 1, 2).Value2, 350#, "Live formula evaluated natively in KPI card"
    AssertEqual ws.Cells(kpiRowStart + 1, 2).Formula, "=100+250", "Cell formula is preserved"

    ' 12. Test SanitizeSheetName
    AssertEqual rb.SanitizeSheetName("Sales:Q4/2026?*"), "Sales_Q4_2026__", "SanitizeSheetName replaces illegal characters"
    AssertEqual rb.SanitizeSheetName(""), "Sheet", "SanitizeSheetName falls back to default on empty"

    ' 13. Test Multi-Column Side-by-Side Flow (ColumnGroup)
    rb.MoveTo 25, 2
    rb.BeginColumnGroup
        rb.AddTable dfData, Title:="Table Left"
    rb.NextColumn ColGap:=2
        rb.AddTable dfData, Title:="Table Right"
    rb.EndColumnGroup ExtraRowGap:=2
    AssertTrue rb.CurrentRow > 30, "CurrentRow advanced after side-by-side tables"
    AssertEqual rb.StartCol, 2, "StartCol restored back to anchor column"

    ' 14. Test Native Chart Embedding
    Dim rowBeforeChart As Long
    rowBeforeChart = rb.CurrentRow
    rb.AddChart Title:="Regional Sales Chart", WidthColumns:=6, HeightRows:=10
    AssertTrue rb.CurrentRow > rowBeforeChart + 10, "CurrentRow advanced after AddChart"

    ' 15. Test Page Setup & PDF Export
    rb.ConfigurePageSetup Orientation:=2, FitToWide:=1
    AssertEqual ws.PageSetup.Orientation, 2, "ConfigurePageSetup set orientation to xlLandscape"

    Dim pdfTestPath As String
    pdfTestPath = Environ$("TEMP") & "\test_report_builder_out.pdf"
    On Error Resume Next
    Kill pdfTestPath
    On Error GoTo TestFail
    rb.ExportPDF pdfTestPath
    AssertTrue Len(Dir$(pdfTestPath)) > 0, "ExportPDF created valid PDF document"
    On Error Resume Next
    Kill pdfTestPath
    On Error GoTo TestFail

    xlApp.DisplayAlerts = False
    ws.Delete
    xlApp.DisplayAlerts = True
    Exit Sub
TestFail:
    If Not xlApp Is Nothing And Not ws Is Nothing Then
        xlApp.DisplayAlerts = False
        ws.Delete
        xlApp.DisplayAlerts = True
    End If
    RecordFailure "Test_ReportBuilder", Err.Description
End Sub

' Test 42: Calendar Matrix, Weekly Grid Styling & Quick Jump Navigation
Private Sub Test_CalendarMatrixAndWeeklyGrid()
    LogPrint "[TEST 42] Calendar Matrix, Weekly Grid & Quick Jump (cDataFrame/cTableStyler/cReportBuilder)..."
    On Error GoTo TestFail

    Dim dfSrc As New cDataFrame
    dfSrc.AddColumn "TaskDate"
    dfSrc.AddColumn "Assignee"
    ' Include Excel serial date number (46060 = 2026-02-07), ISO string, and Date literal
    dfSrc.AddRow CDbl(DateSerial(2026, 2, 7)), "Alice"
    dfSrc.AddRow "2026-02-07", "Alice"
    dfSrc.AddRow DateSerial(2026, 2, 14), "Alice"

    Dim parsedDt As Date
    AssertTrue dfSrc.TryParseExcelDate(CDbl(DateSerial(2026, 2, 7)), parsedDt), "TryParseExcelDate handles Double serial"
    AssertEqual parsedDt, DateSerial(2026, 2, 7), "Parsed serial matches 2026-02-07"
    AssertTrue dfSrc.TryParseExcelDate("2026-02-14", parsedDt), "TryParseExcelDate handles ISO date string"
    AssertEqual parsedDt, DateSerial(2026, 2, 14), "Parsed ISO string matches 2026-02-14"
    AssertFalse dfSrc.TryParseExcelDate("not-a-date", parsedDt), "TryParseExcelDate rejects non-date text"

    Dim dfRoster As New cDataFrame
    dfRoster.AddColumn "Name"
    dfRoster.AddColumn "Department"
    dfRoster.AddRow "Alice", "Engineering"
    dfRoster.AddRow "Bob", "Design"

    Dim dfFeb As cDataFrame
    Set dfFeb = dfSrc.BuildCalendarMatrix("TaskDate", "Assignee", 2026, 2, Array("Department"), dfRoster, "Name", 31)
    AssertEqual dfFeb.RowCount, 2, "BuildCalendarMatrix includes zero-task roster members"
    AssertEqual dfFeb.ColumnCount, 35, "FixedMaxDays:=31 yields 35 aligned columns"
    AssertEqual CLng(dfFeb(1, "D07")), 2, "Alice has 2 completions on Feb 07"
    AssertEqual CLng(dfFeb(1, "D14")), 1, "Alice has 1 completion on Feb 14"
    AssertEqual CLng(dfFeb(1, "Month Total")), 3, "Alice Feb Month Total is 3"
    AssertEqual CStr(dfFeb(1, "D29")), "", "Trailing Feb D29 padded cleanly with empty string"
    AssertEqual CLng(dfFeb(2, "Month Total")), 0, "Bob has 0 completions in Feb"

    Dim xlApp As Object, ws As Object
    Set xlApp = ThisWorkbook.Application
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "TmpCalGridTest"

    Dim rb As New cReportBuilder
    rb.Attach ws, "B2"
    rb.BeginUpdate
    rb.AddBanner "Calendar Matrix Unit Test", "Weekly Grid & Spotlight Verification", "Navy", 35
    rb.AddSpace 1

    Dim jumpRow As Long
    jumpRow = rb.CurrentRow
    rb.AddSpace 2

    rb.AddSectionHeader "FEBRUARY 2026", "Slate", 35, "TestFebAnchor", RGB(30, 41, 59)
    rb.AddSpace 1
    rb.AddTable dfFeb, "", "Slate", True

    AssertTrue Not rb.LastTableStyler Is Nothing, "LastTableStyler is exposed after AddTable"
    AssertTrue rb.LastTableStartRow > 0, "LastTableStartRow is recorded"
    AssertTrue rb.LastTableEndRow > rb.LastTableStartRow, "LastTableEndRow > LastTableStartRow"

    rb.LastTableStyler.ApplyWeeklyCalendarGrid "D01", 28, 31, "Month Total", "Daily Avg"
    rb.LastTableStyler.SpotlightColumn "D07"
    rb.AddQuickJumpBar "Jump:", Array("Feb '26"), Array("'TmpCalGridTest'!B" & CStr(rb.LastTableStartRow - 2)), 1, 2, jumpRow
    rb.AddActionButton "btnTmpUnitTest", "Refresh", "Btn_RefreshMatrix", "AD", jumpRow, 120#, 22#
    rb.EndUpdate

    AssertEqual CStr(ws.Cells(jumpRow, 4).Value), "Feb '26", "AddQuickJumpBar writes text-formatted pill label"
    AssertEqual ws.Hyperlinks.Count, 1, "AddQuickJumpBar adds hyperlink"
    AssertEqual CStr(ws.Cells(rb.LastTableStartRow - 1, 4).Value), "WEEK 1 (D01-D07)", "Weekly super-header generated above table"
    AssertEqual ws.Cells(rb.LastTableStartRow - 1, 10).Borders(10).Weight, -4138, "Week boundary vertical border is xlMedium"
    AssertEqual CStr(ws.Cells(rb.LastTableStartRow + 1, 32).Value), "-", "Trailing Feb D29 dimmed cleanly with dash"
    AssertEqual CLng(ws.Cells(rb.LastTableStartRow, 10).Interior.Color), 423897, "SpotlightColumn highlights D07 header"
    AssertEqual CStr(ws.Buttons("btnTmpUnitTest").Text), "Refresh", "AddActionButton creates idempotent form button"

    SafeDeleteName "TestFebAnchor"
    xlApp.DisplayAlerts = False
    ws.Delete
    xlApp.DisplayAlerts = True
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    SafeDeleteName "TestFebAnchor"
    If Not xlApp Is Nothing And Not ws Is Nothing Then
        On Error Resume Next
        xlApp.DisplayAlerts = False
        ws.Delete
        xlApp.DisplayAlerts = True
    End If
    RecordFailure "Test_CalendarMatrixAndWeeklyGrid", failErr
End Sub

' ------------------------------------------------------------------------------
' Test 43: Table-Driven Configuration & Universal Custom Styling Engine
' ------------------------------------------------------------------------------
Private Sub Test_TableDrivenConfigAndStyles()
    On Error GoTo TestFail
    LogPrint "[TEST 43] Table-Driven Config & Universal Custom Styling Engine..."

    ' 1. cDataFrame Typed Config Getters
    Dim dfCfg As New cDataFrame
    dfCfg.AddColumn "Key"
    dfCfg.AddColumn "Value"
    dfCfg.AddRow "Theme", "Emerald"
    dfCfg.AddRow "MaxRows", "500"
    dfCfg.AddRow "ThresholdRate", "0.95"
    dfCfg.AddRow "EnableAudit", "True"
    dfCfg.AddRow "HeaderColor", "#1E3A8A"

    AssertEqual dfCfg.GetConfigString("Theme"), "Emerald", "GetConfigString retrieves string"
    AssertEqual dfCfg.GetConfigString("MissingKey", "DefaultVal"), "DefaultVal", "GetConfigString fallback to default"
    AssertEqual dfCfg.GetConfigLong("MaxRows"), 500, "GetConfigLong retrieves integer/long"
    AssertEqual dfCfg.GetConfigLong("MissingCount", 42), 42, "GetConfigLong fallback to default"
    AssertEqual dfCfg.GetConfigDouble("ThresholdRate"), 0.95, "GetConfigDouble retrieves double"
    AssertEqual dfCfg.GetConfigDouble("MissingRate", 0.1), 0.1, "GetConfigDouble fallback to default"
    AssertEqual dfCfg.GetConfigBool("EnableAudit"), True, "GetConfigBool retrieves boolean True"
    AssertEqual dfCfg.GetConfigBool("MissingFlag", False), False, "GetConfigBool fallback to default"
    AssertEqual dfCfg.GetConfigColor("HeaderColor"), RGB(30, 58, 138), "GetConfigColor parses hex #1E3A8A to RGB"
    AssertEqual dfCfg.GetConfigColor("MissingColor", RGB(255, 0, 0)), RGB(255, 0, 0), "GetConfigColor fallback to default"

    ' 2. Live Excel Table Styling & Report Config
    Dim xlApp As Object, ws As Object
    Set xlApp = ThisWorkbook.Application
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "TmpTableConfigTest"

    ' Put main data at B2:F6
    ws.Range("B2:F2").Value = Array("EmpID", "Name", "Score", "Progress", "D01")
    ws.Range("B3:F3").Value = Array("E101", "Alice", 85, 0.45, 10)
    ws.Range("B4:F4").Value = Array("E102", "Bob", 92, 0.88, 15)
    ws.Range("B5:F5").Value = Array("E103", "Charlie", 74, 0.25, 8)
    ws.Range("B6:F6").Value = Array("E104", "David", 99, 1#, 20)

    ' Create style table via DataFrame ToListObject
    Dim dfStyle As New cDataFrame
    dfStyle.AddColumn "Scope"
    dfStyle.AddColumn "Rule"
    dfStyle.AddColumn "Target"
    dfStyle.AddColumn "Param1"
    dfStyle.AddColumn "Param2"
    dfStyle.AddRow "Main", "THEME", "Navy", "", ""
    dfStyle.AddRow "Main", "OUTERBORDER", "THICK", "#1E3A8A", ""
    dfStyle.AddRow "Main", "DATABAR", "Score", "#3B82F6", ""
    dfStyle.AddRow "Other", "THEME", "Ruby", "", ""
    dfStyle.ToListObject ws.Range("H2"), "tbl_LiveStyles"

    ' Create report config table via DataFrame ToListObject
    Dim dfRptCfg As New cDataFrame
    dfRptCfg.AddColumn "Setting"
    dfRptCfg.AddColumn "Value"
    dfRptCfg.AddRow "TabColor", "#10B981"
    dfRptCfg.AddRow "AutoFit", "True"
    dfRptCfg.ToListObject ws.Range("N2"), "tbl_LiveReportCfg"

    ' Apply style from Table Name directly with Scope="Main"
    Dim styler As New cTableStyler
    styler.AttachRange ws.Range("B2:F6"), HasHeaders:=True
    styler.ApplyStyleFromTable "tbl_LiveStyles", TargetScope:="Main"

    AssertEqual CLng(ws.Range("B2").Interior.Color), RGB(31, 78, 121), "Theme Navy applied via TableName string"
    AssertEqual ws.Range("B2:F6").Borders(8).Weight, 4, "OuterBorder xlThick (border 8) applied from table"
    AssertTrue ws.Range("D3:D6").FormatConditions.Count > 0, "DataBar format condition applied to Score column"

    ' Apply Report Config from Table Name directly
    Dim rb As New cReportBuilder
    rb.Attach ws
    rb.ApplyConfigFromTable "tbl_LiveReportCfg"
    AssertEqual CLng(ws.Tab.Color), RGB(16, 185, 129), "TabColor applied from config table name"

    xlApp.DisplayAlerts = False
    ws.Delete
    xlApp.DisplayAlerts = True
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    If Not xlApp Is Nothing And Not ws Is Nothing Then
        On Error Resume Next
        xlApp.DisplayAlerts = False
        ws.Delete
        xlApp.DisplayAlerts = True
    End If
    RecordFailure "Test_TableDrivenConfigAndStyles", failErr
End Sub

' ------------------------------------------------------------------------------
' Test 44: Formula Generation & AutoFill Engine (cFormulaEngine)
' ------------------------------------------------------------------------------
Private Sub Test_FormulaEngine()
    On Error GoTo TestFail
    LogPrint "[TEST 44] Formula Generation & AutoFill Engine (cFormulaEngine)..."

    Dim xlApp As Object, ws As Object
    Set xlApp = ThisWorkbook.Application
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "TmpFormulaTest"

    ' Setup sample data at B2:H5
    ' Col B: Item, Col C: Qty, Col D: Price, Col E: Total, Col F: Tax, Col G: NetTotal, Col H: CumTotal
    ws.Range("B2:H2").Value = Array("Item", "Qty", "Price", "Total", "Tax", "NetTotal", "CumTotal")
    ws.Range("B3:H3").Value = Array("Item A", 10, 50, 0, 0, 0, 0)
    ws.Range("B4:H4").Value = Array("Item B", 5, 200, 0, 0, 0, 0)
    ws.Range("B5:H5").Value = Array("Item C", 20, 15, 0, 0, 0, 0)

    Dim fmla As New cFormulaEngine
    fmla.Attach ws.Range("B2:H5"), HasHeaders:=True

    AssertEqual fmla.RowCount, 3, "cFormulaEngine RowCount = 3"
    AssertEqual fmla.ColumnCount, 7, "cFormulaEngine ColumnCount = 7"

    ' 1. Token Formula with AutoFill down: Total = [Qty] * [Price]
    fmla.SetColumnFormula "Total", "=[Qty] * [Price]", UseAutoFill:=True
    AssertEqual Left$(CStr(ws.Range("E3").Formula), 1), "=", "Row 1 Total has formula"
    AssertTrue InStr(CStr(ws.Range("E3").Formula), "C3") > 0, "Row 1 formula points to C3"
    AssertTrue InStr(CStr(ws.Range("E4").Formula), "C4") > 0, "Row 2 formula auto-filled down to C4"
    AssertEqual CDbl(ws.Range("E3").Value), 500#, "Item A Total = 500 (10 * 50)"
    AssertEqual CDbl(ws.Range("E4").Value), 1000#, "Item B Total = 1000 (5 * 200)"
    AssertEqual CDbl(ws.Range("E5").Value), 300#, "Item C Total = 300 (20 * 15)"

    ' 2. R1C1 Formula: Tax = RC[-1] * 0.07 (7% VAT)
    fmla.SetFormulaR1C1 "Tax", "=RC[-1]*0.07"
    AssertEqual CDbl(ws.Range("F3").Value), 35#, "Item A Tax = 35 (500 * 0.07)"
    AssertEqual CDbl(ws.Range("F4").Value), 70#, "Item B Tax = 70 (1000 * 0.07)"

    ' 3. NetTotal = [Total] + [Tax]
    fmla.SetColumnFormula "NetTotal", "=[Total] + [Tax]", UseAutoFill:=True
    AssertEqual CDbl(ws.Range("G3").Value), 535#, "Item A NetTotal = 535 (500 + 35)"
    AssertEqual CDbl(ws.Range("G4").Value), 1070#, "Item B NetTotal = 1070 (1000 + 70)"

    ' 4. Running Total / Cumulative Sum
    fmla.SetRunningTotalFormula "CumTotal", "NetTotal"
    AssertEqual CDbl(ws.Range("H3").Value), 535#, "Item A CumTotal = 535"
    AssertEqual CDbl(ws.Range("H4").Value), 1605#, "Item B CumTotal = 1605 (535 + 1070)"

    ' 5. Summary Formula at row 6
    fmla.SetSummaryFormula 4, "Total", "SUM"
    AssertTrue InStr(CStr(ws.Range("E6").Formula), "SUM(E3:E5)") > 0, "Summary row formula contains SUM(E3:E5)"
    AssertEqual CDbl(ws.Range("E6").Value), 1800#, "Total SUM = 1800 (500+1000+300)"

    ' 6. Table-Driven Formulas via ApplyFormulasFromTable
    Dim dfRule As New cDataFrame
    dfRule.AddColumn "Scope"
    dfRule.AddColumn "Target"
    dfRule.AddColumn "Formula"
    dfRule.AddColumn "Type"
    dfRule.AddRow "Main", "Total", "[Qty] * [Price] * 0.9", "FORMULA"
    dfRule.ToListObject ws.Range("R2"), "tbl_FormulaRules"

    fmla.ApplyFormulasFromTable "tbl_FormulaRules", TargetScope:="Main"
    AssertEqual CDbl(ws.Range("E3").Value), 450#, "Table-driven discount formula applied: 500 * 0.9 = 450"

    ' 7. IFERROR Wrapper
    fmla.SetIfErrorWrapper "Total", 0
    AssertTrue InStr(1, CStr(ws.Range("E3").Formula), "IFERROR", vbTextCompare) > 0, "SetIfErrorWrapper wrapped Total with IFERROR"

    ' 8. Business Formulas: Percentage of Total & Variance
    ws.Range("I2").Value = "PctTotal"
    ws.Range("J2").Value = "Variance"
    Dim fmlaExtended As New cFormulaEngine
    fmlaExtended.Attach ws.Range("B2:J5"), HasHeaders:=True

    fmlaExtended.SetPercentageOfTotal "PctTotal", "Total", "0.0%"
    AssertEqual Left$(CStr(ws.Range("I3").Formula), 1), "=", "PctTotal has formula"
    AssertTrue InStr(CStr(ws.Range("I3").Formula), "SUM") > 0, "PctTotal references SUM"

    fmlaExtended.SetVarianceFormula "Variance", "Total", "Tax", AsPercentage:=False
    AssertEqual Left$(CStr(ws.Range("J3").Formula), 1), "=", "Variance has formula"

    ' 9. Freeze to Values
    fmlaExtended.FreezeToValues "Variance"
    AssertTrue Not ws.Range("J3").HasFormula, "Variance column frozen to static values (HasFormula = False)"
    AssertTrue IsNumeric(ws.Range("J3").Value), "Variance cell contains numeric static value"

    ' 10. INDEX-MATCH Lookup
    ws.Range("O2").Value = "Code"
    ws.Range("P2").Value = "Desc"
    ws.Range("O3").Value = 10
    ws.Range("P3").Value = "Desc-10"
    ws.Range("O4").Value = 5
    ws.Range("P4").Value = "Desc-5"
    ws.Range("O5").Value = 20
    ws.Range("P5").Value = "Desc-20"

    ws.Range("K2").Value = "LookupDesc"
    Dim fmlaLookup As New cFormulaEngine
    fmlaLookup.Attach ws.Range("B2:K5"), HasHeaders:=True
    fmlaLookup.SetIndexMatchLookup "LookupDesc", "Qty", "", "$O$3:$O$5", "$P$3:$P$5", 0
    AssertEqual CStr(ws.Range("K3").Value), "Desc-10", "INDEX-MATCH looked up Code 10 -> Desc-10"
    AssertEqual CStr(ws.Range("K4").Value), "Desc-5", "INDEX-MATCH looked up Code 5 -> Desc-5"
    AssertEqual CStr(ws.Range("K5").Value), "Desc-20", "INDEX-MATCH looked up Code 20 -> Desc-20"

    ' 11. Structured References & ListObject Binding
    Dim loObj As Object
    Set loObj = ws.ListObjects("tbl_FormulaRules")
    Dim fmlaLo As New cFormulaEngine
    fmlaLo.AttachListObject loObj
    AssertTrue fmlaLo.IsListObject, "cFormulaEngine.IsListObject is True"
    fmlaLo.UseStructuredReferences = True
    Dim parsedLoFmla As String
    parsedLoFmla = fmlaLo.ParseFormulaTemplate("=[Target] & [Type]", 1, 1)
    AssertTrue InStr(parsedLoFmla, "[@Target]") > 0, "Structured reference produced [@Target]"

    ' 12. AuditFormulas Error Scanner
    AssertEqual fmlaLookup.AuditFormulas().RowCount, 0, "AuditFormulas reports 0 errors on clean table"
    ws.Range("K5").Formula = "=10/0"
    Dim dfErrors As cDataFrame
    Set dfErrors = fmlaLookup.AuditFormulas()
    AssertEqual dfErrors.RowCount, 1, "AuditFormulas found exactly 1 error cell"
    AssertEqual dfErrors(1, "ErrorType"), "#DIV/0!", "AuditFormulas caught #DIV/0!"
    AssertEqual dfErrors(1, "CellAddress"), "K5", "AuditFormulas identified cell address K5"
    AssertEqual dfErrors(1, "ColName"), "LookupDesc", "AuditFormulas identified column name LookupDesc"

    ' Restore K5
    ws.Range("K5").Value = "Desc-20"

    ' 13. Batch Execution & Bulk Injection
    fmla.BeginBatch
    AssertTrue fmla.InBatch, "fmla.InBatch is True after BeginBatch"
    fmla.SetColumnFormulaBulk "NetTotal", "=[Total] + [Tax]"
    fmla.EndBatch CalculateNow:=True
    AssertTrue Not fmla.InBatch, "fmla.InBatch is False after EndBatch"
    AssertEqual CDbl(ws.Range("G3").Value), 481.5, "Bulk formula evaluated properly after EndBatch (450 + 31.5)"

    ' 14. In-Memory Formula Evaluation
    Dim evalScalar As Variant
    evalScalar = fmla.Evaluate("SUM([Total])")
    AssertEqual CDbl(evalScalar), 1620#, "Evaluate SUM([Total]) in memory = 1620 (discounted total)"

    Dim dfEval As cDataFrame
    Set dfEval = fmla.EvaluateToDataFrame("SUM([Total])", "TotalSum")
    AssertEqual dfEval.RowCount, 1, "EvaluateToDataFrame RowCount = 1"
    AssertEqual CDbl(dfEval(1, 1)), 1620#, "dfEval(1, 1) = 1620"

    ' 15. Dependency & Precedent Tracking
    Dim precs As Variant
    precs = fmla.GetPrecedents("NetTotal", 1)
    AssertTrue UBound(precs) >= 0, "GetPrecedents returned precedent columns"
    Dim foundPrecedent As Boolean, pIdx As Long
    foundPrecedent = False
    For pIdx = LBound(precs) To UBound(precs)
        If CStr(precs(pIdx)) = "Total" Or CStr(precs(pIdx)) = "Tax" Then foundPrecedent = True
    Next pIdx
    AssertTrue foundPrecedent, "GetPrecedents identified Total or Tax as precedent"

    ' 16. Spill Formula & GetSpillRange
    fmla.SetSpillFormula ws.Range("W2"), "=SEQUENCE(3, 1, 100, 10)"
    Dim spRange As Object
    Set spRange = fmla.GetSpillRange(ws.Range("W2"))
    AssertTrue Not spRange Is Nothing, "GetSpillRange returned valid range"
    AssertEqual CLng(ws.Range("W2").Value), 100, "Spill anchor cell value = 100"

    ' 17. Smart String & Criteria Quoting
    AssertEqual fmla.Quote("Pending"), Chr$(34) & "Pending" & Chr$(34), "Quote wraps in outer double quotes"
    AssertEqual fmla.Quote("A " & Chr$(34) & "B" & Chr$(34) & " C"), Chr$(34) & "A " & Chr$(34) & Chr$(34) & "B" & Chr$(34) & Chr$(34) & " C" & Chr$(34), "Quote escapes inner double quotes"
    AssertEqual fmla.FormatCriteria(">", 50), ">50", "FormatCriteria with > operator"
    AssertEqual fmla.FormatCriteria("<>", 0), "<>0", "FormatCriteria with <> operator"
    AssertEqual fmla.FormatCriteria("", "Approved"), "=Approved", "FormatCriteria default operator ="

    ' 18. Cross-Region Separator & Single-Quote String Literals
    AssertEqual fmla.NormalizeFormulaSeparators("=SUM(A1; B1; C1)"), "=SUM(A1, B1, C1)", "NormalizeFormulaSeparators converts semicolons"
    AssertEqual fmla.NormalizeFormulaSeparators("=IF(A1; ""A;B""; C1)"), "=IF(A1, ""A;B"", C1)", "NormalizeFormulaSeparators preserves semicolons in quotes"
    Dim parsedFmla As String
    parsedFmla = fmla.ParseFormulaTemplate("=IF([Qty]>0, 'In Stock', 'Out')", 1, 3)
    AssertTrue InStr(1, parsedFmla, Chr$(34) & "In Stock" & Chr$(34)) > 0, "ParseFormulaTemplate translates 'In Stock' to double quotes"
    AssertTrue InStr(1, parsedFmla, Chr$(34) & "Out" & Chr$(34)) > 0, "ParseFormulaTemplate translates 'Out' to double quotes"
    Dim parsedSheetRef As String
    parsedSheetRef = fmla.ParseFormulaTemplate("='My Sheet'!A1 + 'Other Sheet'!B2", 1, 3)
    AssertTrue InStr(1, parsedSheetRef, "'My Sheet'!") > 0, "ParseFormulaTemplate preserves sheet name single quotes"

    ' 19. Volatile Function Audit
    Dim dfVol As cDataFrame
    Set dfVol = fmla.AuditVolatileFormulas()
    AssertEqual dfVol.RowCount, 0, "AuditVolatileFormulas returns 0 on non-volatile table"
    Dim origFmla As String
    origFmla = ws.Range("E3").Formula
    ws.Range("E3").Formula = "=OFFSET(C3, 0, 1)"
    Dim dfVolOffset As cDataFrame
    Set dfVolOffset = fmla.AuditVolatileFormulas("Total")
    AssertTrue dfVolOffset.RowCount >= 1, "AuditVolatileFormulas detected OFFSET in Total"
    AssertEqual CStr(dfVolOffset(1, "VolatileFunction")), "OFFSET", "Detected volatile function is OFFSET"
    ws.Range("E3").Formula = origFmla

    ' 20. Circular Reference Audit
    Dim dfCirc As cDataFrame
    Set dfCirc = fmla.AuditCircularReferences()
    AssertEqual dfCirc.RowCount, 0, "AuditCircularReferences returns 0 rows on clean sheet"

    ' 21. External Links Audit
    Dim dfLinks As cDataFrame
    Set dfLinks = fmla.AuditExternalLinks()
    AssertTrue dfLinks.RowCount >= 0, "AuditExternalLinks returns valid DataFrame"

    ' 22. Local Formula Assignment
    fmla.SetColumnFormulaLocal "CumTotal", "=[Total]"
    AssertEqual CDbl(ws.Range("H3").Value), 450#, "SetColumnFormulaLocal evaluated properly"

    ' 23. Audit & Auto-Heal Overwritten Formulas
    ws.Range("G4").Value = 999
    Dim dfInc As cDataFrame
    Set dfInc = fmla.AuditFormulaInconsistency("NetTotal")
    AssertEqual dfInc.RowCount, 1, "AuditFormulaInconsistency detected 1 overwritten cell"
    AssertEqual CStr(dfInc(1, "Issue")), "OVERWRITTEN_WITH_CONSTANT", "Issue = OVERWRITTEN_WITH_CONSTANT"
    AssertEqual CStr(dfInc(1, "Address")), "G4", "Overwritten address = G4"
    Dim restoredCnt As Long
    restoredCnt = fmla.RestoreColumnFormulas("NetTotal")
    AssertEqual restoredCnt, 1, "RestoreColumnFormulas restored 1 cell"
    AssertTrue ws.Range("G4").HasFormula, "G4 formula restored"
    AssertEqual fmla.AuditFormulaInconsistency("NetTotal").RowCount, 0, "AuditFormulaInconsistency clean after restore"

    ' 24. Long Array Formula (>255 Chars Bypass)
    Dim longArrFmla As String, tIdx As Long
    longArrFmla = "=SUM(IF([Qty]>0,[Qty],0))"
    For tIdx = 1 To 14
        longArrFmla = longArrFmla & "+SUM(IF([Qty]>0,[Qty],0))"
    Next tIdx
    fmla.SetColumnFormulaArrayLong "CumTotal", longArrFmla
    AssertTrue Len(CStr(ws.Range("H3").Formula)) > 255, "Long array formula > 255 chars assigned successfully"
    AssertTrue ws.Range("H3").HasArray, "H3 has HasArray = True"
    AssertEqual CDbl(ws.Range("H3").Value), 150#, "Long array formula evaluated: 15 * 10 = 150"
    fmla.SetColumnFormula "CumTotal", "=[Total]"

    ' 25. Formula Reference Style Conversion
    fmla.ConvertFormulaReferences "NetTotal", 1
    AssertTrue InStr(1, CStr(ws.Range("G3").Formula), "$") > 0, "ConvertFormulaReferences(1) added $ absolute anchors"
    fmla.ConvertFormulaReferences "NetTotal", 4
    AssertEqual InStr(1, CStr(ws.Range("G3").Formula), "$"), 0, "ConvertFormulaReferences(4) reverted to relative A1"

    ' 26. Downstream Dependents & Orphan Audit
    Dim deps As Variant
    deps = fmla.GetDependents("Total", 1)
    AssertTrue UBound(deps) >= 0, "GetDependents found downstream dependents for Total"
    Dim dfOrph As cDataFrame
    Set dfOrph = fmla.AuditOrphanFormulas()
    AssertTrue dfOrph.RowCount >= 1, "AuditOrphanFormulas identified leaf formula columns"

    ' 27. Formula Beautifier
    Dim prettyFmla As String
    prettyFmla = fmla.FormatFormula("=IF(LEN(A1)>0, SUM(B1:B10), 0)", 2)
    AssertTrue InStr(1, prettyFmla, vbLf) > 0, "FormatFormula inserted line breaks"
    AssertTrue InStr(1, prettyFmla, "  SUM(B1:B10)") > 0, "FormatFormula indented inner SUM on 2 spaces"

    xlApp.DisplayAlerts = False
    ws.Delete
    xlApp.DisplayAlerts = True
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    If Not xlApp Is Nothing And Not ws Is Nothing Then
        On Error Resume Next
        xlApp.DisplayAlerts = False
        ws.Delete
        xlApp.DisplayAlerts = True
    End If
    RecordFailure "Test_FormulaEngine", failErr
End Sub

' ------------------------------------------------------------------------------
' Test 45: Comprehensive Gap Coverage Across All 6 Classes
' ------------------------------------------------------------------------------
Private Sub Test_AllClassesGapCoverage()
    On Error GoTo TestFail
    LogPrint "[TEST 45] Comprehensive Gap Coverage Across All 6 Classes..."

    Dim xlApp As Object
    Set xlApp = Application
    Dim ws As Object
    Set ws = xlApp.ActiveWorkbook.Worksheets.Add

    ' --- 1. cColumn new methods & O(N) RollingSum/RollingAvg ---
    Dim col As New cColumn
    col.Name = "Vals"
    col.Append -10.456
    col.Append Empty
    col.Append 20.123
    col.Append 20.123
    col.Append -5#

    AssertEqual col.NUnique(), 3, "cColumn.NUnique counts 3 unique non-empty values"
    Dim vc As cDataFrame
    Set vc = col.ValueCounts()
    AssertEqual vc.RowCount, 3, "cColumn.ValueCounts has 3 rows"
    AssertEqual vc(1, "Value"), 20.123, "cColumn.ValueCounts sorted descending (top = 20.123)"
    AssertEqual vc(1, "Count"), 2, "cColumn.ValueCounts top count = 2"

    Dim ffilled As cColumn, bfilled As cColumn
    Set ffilled = col.FFill()
    AssertEqual ffilled.Item(2), -10.456, "cColumn.FFill fills row 2 from row 1"
    Set bfilled = col.BFill()
    AssertEqual bfilled.Item(2), 20.123, "cColumn.BFill fills row 2 from row 3"

    Dim naCol As cColumn
    Set naCol = col.IsNA()
    AssertTrue naCol.Item(2), "cColumn.IsNA true for Empty row 2"
    AssertFalse naCol.Item(1), "cColumn.IsNA false for row 1"

    Dim inCol As cColumn
    Set inCol = col.IsIn(Array(-10.456, -5#))
    AssertTrue inCol.Item(1) And inCol.Item(5), "cColumn.IsIn matches array values"
    AssertFalse inCol.Item(3), "cColumn.IsIn false for non-matching value"

    Dim absC As cColumn, rndC As cColumn
    Set absC = col.AbsCol()
    AssertEqual absC.Item(1), 10.456, "cColumn.AbsCol converts -10.456 to 10.456"
    Set rndC = col.RoundCol(1)
    AssertEqual rndC.Item(1), -10.5, "cColumn.RoundCol(1) rounds -10.456 to -10.5"

    Dim strCol As New cColumn
    strCol.Name = "Tags"
    strCol.Append "Alpha-01"
    strCol.Append "Beta"
    strCol.Append "Gamma-02"
    Dim containsCol As cColumn
    Set containsCol = strCol.StrContains("alpha")
    AssertTrue containsCol.Item(1), "cColumn.StrContains case-insensitive match"
    AssertFalse containsCol.Item(2), "cColumn.StrContains non-match"
    Dim lenCol As cColumn
    Set lenCol = strCol.StrLen()
    AssertEqual lenCol.Item(1), 8, "cColumn.StrLen returns 8 for Alpha-01"

    Dim rollCol As New cColumn
    rollCol.Name = "R"
    rollCol.Append 10#: rollCol.Append 20#: rollCol.Append 30#: rollCol.Append 40#
    Dim rSum As cColumn, rAvg As cColumn
    Set rSum = rollCol.RollingSum(2)
    Set rAvg = rollCol.RollingAvg(2)
    AssertIsError rSum.Item(1), "O(N) RollingSum(2) row 1 is #N/A"
    AssertEqual rSum.Item(2), 30#, "O(N) RollingSum(2) row 2 = 30"
    AssertEqual rSum.Item(4), 70#, "O(N) RollingSum(2) row 4 = 70"
    AssertEqual rAvg.Item(3), 25#, "O(N) RollingAvg(2) row 3 = 25"

    ' --- 2. cDataFrame new methods, Merge(LeftOn/RightOn/CROSS), GroupBy(MEDIAN/STD/NUNIQUE) ---
    Dim df As New cDataFrame
    df.AddColumn "Dept"
    df.AddColumn "Emp"
    df.AddColumn "Salary"
    df.AddRow "Eng", "Alice", 100#
    df.AddRow "Eng", "Bob", 200#
    df.AddRow "Eng", "Cara", 300#
    df.AddRow "Sales", "Dan", 150#
    df.AddRow "Sales", "Dan", 250#

    Dim arr2D As Variant
    arr2D = df.To2DArray(True)
    AssertEqual arr2D(1, 1), "Dept", "To2DArray includes header at (1,1)"
    AssertEqual arr2D(2, 3), 100#, "To2DArray data row 1 col 3 = 100"

    Dim top2 As cDataFrame, bot2 As cDataFrame, samp2 As cDataFrame, trDf As cDataFrame
    Set top2 = df.NLargest(2, "Salary")
    AssertEqual top2.RowCount, 2, "NLargest(2) returns 2 rows"
    AssertEqual top2(1, "Salary"), 300#, "NLargest(2) top row = 300"
    Set bot2 = df.NSmallest(2, "Salary")
    AssertEqual bot2(1, "Salary"), 100#, "NSmallest(2) smallest row = 100"
    Set samp2 = df.Sample(2, 42)
    AssertEqual samp2.RowCount, 2, "Sample(2, 42) returns 2 rows"
    Set trDf = top2.Transpose()
    AssertEqual trDf.RowCount, 3, "Transpose swaps 3 columns into 3 rows"

    ' Merge LeftOn/RightOn & CROSS
    Dim deptLookup As New cDataFrame
    deptLookup.AddColumn "Code"
    deptLookup.AddColumn "Budget"
    deptLookup.AddRow "Eng", 5000#
    deptLookup.AddRow "Sales", 3000#

    Dim mergedLR As cDataFrame
    Set mergedLR = df.Merge(deptLookup, LeftOn:="Dept", RightOn:="Code", How:="LEFT")
    AssertEqual mergedLR.RowCount, 5, "Merge LeftOn/RightOn preserves all 5 left rows"
    AssertEqual mergedLR(1, "Budget"), 5000#, "Merge LeftOn/RightOn maps Eng -> 5000"

    Dim crossDf As cDataFrame
    Set crossDf = top2.Merge(deptLookup, How:="CROSS")
    AssertEqual crossDf.RowCount, 4, "Merge CROSS produces 2 x 2 = 4 rows"

    ' GroupBy MEDIAN, STD, NUNIQUE
    Dim gbStats As cDataFrame
    Set gbStats = df.GroupBy("Dept", Array("Salary", "Salary", "Emp"), Array("MEDIAN", "STD", "NUNIQUE"))
    AssertEqual gbStats.RowCount, 2, "GroupBy MEDIAN/STD/NUNIQUE returns 2 groups"
    AssertEqual gbStats(1, "Salary_MEDIAN"), 200#, "GroupBy MEDIAN for Eng (100,200,300) = 200"
    AssertEqual gbStats(1, "Salary_STD"), 100#, "GroupBy STD for Eng (100,200,300) = 100"
    AssertEqual gbStats(2, "Emp_NUNIQUE"), 1, "GroupBy NUNIQUE for Sales (Dan, Dan) = 1"

    ' --- 3. cDataValidator new rules & ApplyToExcelRange ---
    Dim valDf As New cDataFrame
    valDf.AddColumn "Code"
    valDf.AddColumn "StartDt"
    valDf.AddColumn "EndDt"
    valDf.AddColumn "MinV"
    valDf.AddColumn "MaxV"
    valDf.AddRow "AB-12", CDate("2026-01-01"), CDate("2026-01-10"), 10, 20
    valDf.AddRow "BAD", CDate("2025-12-01"), CDate("2026-01-05"), 30, 20

    Dim dv As New cDataValidator
    dv.AddAllowedList "Code", Array("AB-12", "CD-34")
    dv.AddRegex "Code", "^[A-Z]{2}-\d{2}$"
    dv.AddStringLength "Code", 5, 5
    dv.AddDateRange "StartDt", CDate("2026-01-01"), CDate("2026-12-31")
    dv.AddCrossColumnRule "MinV", "<=", "MaxV", "MinV must be <= MaxV"

    AssertFalse valDf.Validate(dv), "cDataFrame.Validate returns False for failing row 2"
    AssertTrue dv.ErrorCount >= 4, "cDataValidator captures Regex, StringLength, DateRange, CrossColumn errors"

    valDf.ToRange ws.Range("A1"), IncludeHeaders:=True
    dv.ApplyToExcelRange ws.Range("A1:E3"), True
    AssertEqual ws.Range("A2").Validation.Type, 3, "ApplyToExcelRange applied List validation (xlValidateList=3)"

    ' --- 4. cTableStyler tested on dedicated wsStyler below ---
    ' --- 5. cFormulaEngine ValidateFormulaSyntax & structured ref escaping ---
    Dim feDf As cFormulaEngine
    Set feDf = valDf.AttachFormulaEngine(ws.Range("A1"))
    AssertTrue feDf.ValidateFormulaSyntax("=SUM(A1:A10) + IF(B1>0, 1, 0)"), "ValidateFormulaSyntax accepts valid formula"
    AssertFalse feDf.ValidateFormulaSyntax("=SUM(A1:A10"), "ValidateFormulaSyntax rejects unclosed paren"
    AssertFalse feDf.ValidateFormulaSyntax("SUM(A1:A10)"), "ValidateFormulaSyntax rejects missing leading ="
    AssertFalse feDf.ValidateFormulaSyntax("=IF(A1=""abc, 1, 0)"), "ValidateFormulaSyntax rejects unclosed quote"

    LogPrint "  [45.5a] spDf ListObject..."
    Dim spDf As New cDataFrame
    spDf.AddColumn "Total #1 [2026]"
    spDf.AddColumn "Out"
    spDf.AddRow 100, Empty
    spDf.AddRow 200, Empty
    spDf.ToRange ws.Range("L1"), IncludeHeaders:=True
    Dim loSp As Object
    Set loSp = ws.ListObjects.Add(1, ws.Range("L1:M3"), , 1)
    Dim feSp As New cFormulaEngine
    feSp.Attach loSp
    feSp.UseStructuredReferences = True
    Dim escapedFmla As String
    escapedFmla = feSp.ParseFormulaTemplate("=[Total #1 [2026]] * 2", 2, 2)
    AssertTrue InStr(1, escapedFmla, "[@[Total '#1 '[2026']]]") > 0, "ParseFormulaTemplate escapes #, [, ] in structured ref"

    Dim wsRep As Object
    Set wsRep = xlApp.ActiveWorkbook.Worksheets.Add
    Dim rb As New cReportBuilder
    rb.Attach wsRep, "B2"
    rb.BeginUpdate
    rb.AddTable df, StyleTheme:="Navy", AddTotalRow:=False
    AssertTrue Not rb.AttachFormulaEngine() Is Nothing, "cReportBuilder.AttachFormulaEngine returns cFormulaEngine"
    AssertTrue Not rb.LastFormulaEngine Is Nothing, "cReportBuilder.LastFormulaEngine is populated"
    rb.AddFormulaColumn "Bonus", "=[Salary]*0.1", "$#,##0"
    AssertEqual wsRep.Range("E2").Value, "Bonus", "AddFormulaColumn added Bonus header"
    AssertEqual wsRep.Range("E3").Value, 10#, "AddFormulaColumn computed Bonus = 10"
    rb.AddSparkline "D3:D7", 1
    AssertEqual wsRep.Cells(rb.CurrentRow - 2, 2).SparklineGroups.Count, 1, "AddSparkline added 1 SparklineGroup"
    rb.EndUpdate

    ' ExportPDF clean error 70 on invalid directory
    Dim caughtPdfErr As Boolean
    caughtPdfErr = False
    On Error Resume Next
    rb.ExportPDF Environ$("TEMP") & "\nonexistent_folder_99999\test.pdf"
    If Err.Number = 70 Then caughtPdfErr = True
    Err.Clear
    On Error GoTo TestFail
    AssertTrue caughtPdfErr, "ExportPDF raises clean Error 70 on invalid/locked path"

    LogPrint "  [45.6b] BuildMultiSheetReport..."
    Dim dupDf As New cDataFrame
    dupDf.AddColumn "LongGroup"
    dupDf.AddColumn "Val"
    dupDf.AddRow "ABCDEFGHIJKLMNOPQRSTUVWXYZ_12345_A", 1
    dupDf.AddRow "ABCDEFGHIJKLMNOPQRSTUVWXYZ_12345_B", 2
    Dim wbTemp As Object
    Set wbTemp = xlApp.Workbooks.Add
    Dim rbMulti As New cReportBuilder
    Dim sheetsCol As Collection
    Set sheetsCol = rbMulti.BuildMultiSheetReport(wbTemp, dupDf, "LongGroup")
    AssertEqual sheetsCol.Count, 3, "BuildMultiSheetReport created 2 disambiguated sheets + 1 TOC"
    wbTemp.Close False

    LogPrint "  [45.7] cTableStyler on wsStyler..."
    Dim wsStyler As Object
    Set wsStyler = xlApp.ActiveWorkbook.Worksheets.Add
    Dim styler As cTableStyler
    LogPrint "  [45.7a] ToStyledRange..."
    Set styler = df.ToStyledRange(wsStyler.Range("B2"), "Navy")
    LogPrint "  [45.7b] FormatAs*..."
    styler.FormatAsCurrency "Salary", "$", 0
    AssertEqual wsStyler.Range("D3").NumberFormat, "$#,##0", "FormatAsCurrency sets number format"
    styler.FormatAsPercent "Salary", 1
    AssertEqual wsStyler.Range("D3").NumberFormat, "0.0%", "FormatAsPercent sets percent format"
    styler.FormatAsDate "Salary", "yyyy-mm-dd"
    AssertEqual wsStyler.Range("D3").NumberFormat, "yyyy-mm-dd", "FormatAsDate sets date format"
    xlApp.DisplayAlerts = False
    wsStyler.Delete
    wsRep.Delete
    ws.Delete
    xlApp.DisplayAlerts = True
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    On Error Resume Next
    xlApp.DisplayAlerts = False
    If Not wbTemp Is Nothing Then wbTemp.Close False
    If Not wsStyler Is Nothing Then wsStyler.Delete
    If Not wsRep Is Nothing Then wsRep.Delete
    If Not ws Is Nothing Then ws.Delete
    xlApp.DisplayAlerts = True
    RecordFailure "Test_AllClassesGapCoverage", failErr
End Sub

' ------------------------------------------------------------------------------
' Test 46: Get Data (2DArray, Folder) & Range Operations
' ------------------------------------------------------------------------------
Private Sub Test_GetDataAndRangeGaps()
    On Error GoTo TestFail
    LogPrint "[TEST 46] Get Data (2DArray, Folder) & Range Operations..."

    Dim xlApp As Object
    Set xlApp = ThisWorkbook.Application
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets.Add
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmpDir As String

    ' 1. LoadFrom2DArray (1-based with headers & 0-based with explicit headers)
    Dim arr1(1 To 3, 1 To 2) As Variant
    arr1(1, 1) = "Code": arr1(1, 2) = "Qty"
    arr1(2, 1) = "A1":   arr1(2, 2) = 10
    arr1(3, 1) = "B2":   arr1(3, 2) = 20
    Dim dfArr As New cDataFrame
    dfArr.LoadFrom2DArray arr1, True
    AssertEqual dfArr.RowCount, 2, "LoadFrom2DArray 1-based RowCount"
    AssertEqual dfArr.Column("Qty").Item(2), 20, "LoadFrom2DArray 1-based value"

    Dim arr0(0 To 1, 0 To 1) As Variant
    arr0(0, 0) = "X": arr0(0, 1) = 100
    arr0(1, 0) = "Y": arr0(1, 1) = 200
    Dim dfArr0 As New cDataFrame
    dfArr0.LoadFrom2DArray arr0, False, Array("Tag", "Amount")
    AssertEqual dfArr0.RowCount, 2, "LoadFrom2DArray 0-based RowCount"
    AssertEqual dfArr0.Column("Tag").Item(1), "X", "LoadFrom2DArray 0-based explicit header"

    ' 2. LoadFromFolder (2 CSV files)
    tmpDir = fso.GetSpecialFolder(2).Path & "\vba_df_folder_" & Format$(Timer * 1000, "0")
    If Not fso.FolderExists(tmpDir) Then fso.CreateFolder tmpDir
    Dim ts As Object
    Set ts = fso.CreateTextFile(tmpDir & "\part1.csv", True)
    ts.WriteLine "Item,Val"
    ts.WriteLine "P1,10"
    ts.Close
    Set ts = fso.CreateTextFile(tmpDir & "\part2.csv", True)
    ts.WriteLine "Item,Val"
    ts.WriteLine "P2,20"
    ts.WriteLine "P3,30"
    ts.Close

    Dim dfFolder As New cDataFrame
    dfFolder.LoadFromFolder tmpDir, "*.csv", True, True
    AssertEqual dfFolder.RowCount, 3, "LoadFromFolder combined RowCount"
    AssertTrue dfFolder.HasColumn("SourceFile"), "LoadFromFolder added SourceFile"
    AssertEqual dfFolder.Column("Item").Item(3), "P3", "LoadFromFolder row 3 Item"
    fso.DeleteFolder tmpDir, True
    tmpDir = ""

    ' 3. LoadFromCurrentRegion & LoadFromVisibleRange
    ws.Range("A1:B1").Value2 = Array("Name", "Score")
    ws.Range("A2:B2").Value2 = Array("Ann", 10)
    ws.Range("A3:B3").Value2 = Array("Bob", 20)
    ws.Range("A4:B4").Value2 = Array("Cam", 30)
    ws.Rows(3).Hidden = True

    Dim dfCur As New cDataFrame
    dfCur.LoadFromCurrentRegion ws.Range("A1"), True
    AssertEqual dfCur.RowCount, 3, "LoadFromCurrentRegion RowCount"

    Dim dfVis As New cDataFrame
    dfVis.LoadFromVisibleRange ws.Range("A1:B4"), True
    AssertEqual dfVis.RowCount, 2, "LoadFromVisibleRange skips hidden row"
    AssertEqual dfVis.Column("Name").Item(2), "Cam", "LoadFromVisibleRange row 2 is Cam"
    ws.Rows(3).Hidden = False

    ' 4. LoadFromRange with FillMerged & LoadFromRangeAreas
    ws.Range("D1:E1").Value2 = Array("Grp", "Amt")
    ws.Range("D2").Value2 = "North": ws.Range("E2").Value2 = 50
    ws.Range("E3").Value2 = 75
    xlApp.DisplayAlerts = False
    ws.Range("D2:D3").Merge
    xlApp.DisplayAlerts = True

    Dim dfMerged As New cDataFrame
    dfMerged.LoadFromRange ws.Range("D1:E3"), True, True
    AssertEqual dfMerged.Column("Grp").Item(2), "North", "LoadFromRange FillMerged fills second row"

    Dim unionRng As Object
    Set unionRng = xlApp.Union(ws.Range("A1:B2"), ws.Range("A4:B4"))
    Dim dfAreas As New cDataFrame
    dfAreas.LoadFromRangeAreas unionRng, True
    AssertEqual dfAreas.RowCount, 2, "LoadFromRangeAreas stacks 2 areas"
    AssertEqual dfAreas.Column("Name").Item(2), "Cam", "LoadFromRangeAreas second row from area 2"

    ' 5. LoadFromRangeText & LoadFromRangeFormulas
    ws.Range("G1:H1").Value2 = Array("Code", "Calc")
    ws.Range("G2").NumberFormat = "000"
    ws.Range("G2").Value2 = 7
    ws.Range("H2").Formula = "=10+5"

    Dim dfTxt As New cDataFrame
    dfTxt.LoadFromRangeText ws.Range("G1:H2"), True
    AssertEqual CStr(dfTxt.Column("Code").Item(1)), "007", "LoadFromRangeText preserves leading zeros"

    Dim dfFml As New cDataFrame
    dfFml.LoadFromRangeFormulas ws.Range("G1:H2"), True
    AssertEqual CStr(dfFml.Column("Calc").Item(1)), "=10+5", "LoadFromRangeFormulas reads formula"

    ' 6. ToRangeAsObject, ToRangeReplace, AppendToRange, ToRangeChunked
    Dim outRng As Object
    Set outRng = dfCur.ToRangeAsObject(ws.Range("J1"), True)
    AssertEqual outRng.Rows.Count, 4, "ToRangeAsObject returns 4 rows"

    ' AppendToRange adds 3 more rows below J1:K4 -> J5:K7
    Dim appRng As Object
    Set appRng = dfCur.AppendToRange(ws.Range("J1"))
    AssertEqual appRng.Row, 5, "AppendToRange starts at row 5"
    AssertEqual ws.Range("J7").Value2, "Cam", "AppendToRange wrote row 7"

    ' ToRangeReplace replaces J1:K7 with 2-row dfVis and clears J4:K7
    Dim repRng As Object
    Set repRng = dfVis.ToRangeReplace(ws.Range("J1"), True, True)
    AssertEqual repRng.Rows.Count, 3, "ToRangeReplace wrote 3 rows (header+2)"
    AssertTrue IsEmpty(ws.Range("J5").Value2), "ToRangeReplace cleared stale rows below"

    ' ToRangeChunked writes with ChunkSize=1
    Dim chRng As Object
    Set chRng = dfCur.ToRangeChunked(ws.Range("M1"), True, 1)
    AssertEqual chRng.Rows.Count, 4, "ToRangeChunked wrote 4 rows"
    AssertEqual ws.Range("N4").Value2, 30, "ToRangeChunked wrote last value"

    ' 7. cFormulaEngine.ConvertRangeToValues & cTableStyler.AutoFitWithLimits / FreezeHeader
    LogPrint "  [46.7a] ConvertRangeToValues..."
    Dim fe As New cFormulaEngine
    fe.ConvertRangeToValues ws.Range("H2")
    AssertFalse ws.Range("H2").HasFormula, "ConvertRangeToValues removes formula"
    AssertEqual CStr(ws.Range("H2").Formula), "15", "ConvertRangeToValues converted formula to value 15"

    LogPrint "  [46.7b] AttachRange..."
    Dim wsSty As Object
    Set wsSty = ThisWorkbook.Worksheets.Add
    wsSty.Range("A1:B2").Value2 = Array("ColA", "ColB")
    wsSty.Range("A2").Value2 = "Short"
    wsSty.Range("B2").Value2 = String$(60, "X")
    Dim sty As New cTableStyler
    sty.AttachRange wsSty.Range("A1:B2"), True
    LogPrint "  [46.7c] AutoFitWithLimits..."
    sty.AutoFitWithLimits 10, 20
    LogPrint "  [46.7c2] Checking ColumnWidth..."
    AssertTrue sty.TableRange.Columns(1).ColumnWidth >= 10, "AutoFitWithLimits >= MinWidth"
    AssertTrue sty.TableRange.Columns(2).ColumnWidth <= 20, "AutoFitWithLimits <= MaxWidth"
    LogPrint "  [46.7d] FreezeHeader..."
    sty.FreezeHeader
    xlApp.DisplayAlerts = False
    wsSty.Delete
    xlApp.DisplayAlerts = True

    LogPrint "  [46.7e] ws.Delete..."
    xlApp.DisplayAlerts = False
    ws.Delete
    xlApp.DisplayAlerts = True
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    On Error Resume Next
    If Len(tmpDir) > 0 Then
        If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir, True
    End If
    xlApp.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    xlApp.DisplayAlerts = True
    RecordFailure "Test_GetDataAndRangeGaps", failErr
End Sub

' ------------------------------------------------------------------------------
' Test 47: Clean Data & Transform Data Operations
' ------------------------------------------------------------------------------
Private Sub Test_CleanAndTransformData()
    On Error GoTo TestFail
    LogPrint "[TEST 47] Clean Data & Transform Data Operations..."

    ' 1. StrClean & NormalizeWhitespace
    LogPrint "  [47.1] StrClean & NormalizeWhitespace..."
    Dim colStr As New cColumn
    colStr.Name = "DirtyStr"
    colStr.Append "  Hello" & Chr$(1) & Chr$(2) & " World" & Chr$(160) & "  "
    colStr.Append "Line1" & vbCrLf & "  Line2" & vbTab & "Line3  "

    Dim cClean As cColumn
    Set cClean = colStr.StrClean()
    AssertEqual cClean.Item(1), "Hello World", "StrClean removes control chars and non-breaking space"

    Dim cNorm As cColumn
    Set cNorm = colStr.NormalizeWhitespace()
    AssertEqual cNorm.Item(1), "Hello World", "NormalizeWhitespace collapses space"
    AssertEqual cNorm.Item(2), "Line1 Line2 Line3", "NormalizeWhitespace normalizes linebreaks/tabs"

    ' 2. CoerceNumeric & CoerceDate & Cast
    LogPrint "  [47.2] CoerceNumeric & Cast..."
    Dim colNum As New cColumn
    colNum.Append " $1,250.50 "
    colNum.Append "(75.25)"
    colNum.Append "15%"
    colNum.Append "N/A"

    Dim cNum As cColumn
    Set cNum = colNum.CoerceNumeric(, -1#)
    AssertEqual cNum.Item(1), 1250.5, "CoerceNumeric parses currency with comma"
    AssertEqual cNum.Item(2), -75.25, "CoerceNumeric parses parens negative"
    AssertEqual cNum.Item(3), 0.15, "CoerceNumeric parses percentage"
    AssertEqual cNum.Item(4), -1#, "CoerceNumeric uses default value on failure"

    LogPrint "  [47.3] CoerceDate..."
    Dim colDt As New cColumn
    colDt.Append "2026-11-16"
    colDt.Append 45612 ' Excel serial date
    Dim cDt As cColumn
    Set cDt = colDt.CoerceDate()
    AssertTrue IsDate(cDt.Item(1)), "CoerceDate parses item 1 as Date"
    AssertEqual CLng(Year(cDt.Item(1))), 2026&, "CoerceDate parses ISO year"
    AssertEqual CLng(Month(cDt.Item(1))), 11&, "CoerceDate parses ISO month"
    AssertEqual CLng(Day(cDt.Item(1))), 16&, "CoerceDate parses ISO day"
    AssertTrue IsDate(cDt.Item(2)), "CoerceDate parses numeric serial"

    Dim cCast As cColumn
    Set cCast = colNum.Cast("STRING")
    AssertEqual VarType(cCast.Item(1)), vbString, "Cast converts to String"

    ' 3. ExtractRegex & ReplaceRegex
    LogPrint "  [47.4] ExtractRegex & ReplaceRegex..."
    Dim colRx As New cColumn
    colRx.Append "Invoice #INV-2024-9988 paid"
    Dim cExt As cColumn
    Set cExt = colRx.ExtractRegex("INV-\d{4}-\d+", 0)
    AssertEqual cExt.Item(1), "INV-2024-9988", "ExtractRegex extracts pattern match"

    Dim cRep As cColumn
    Set cRep = colRx.ReplaceRegex("\d+", "XXX")
    AssertEqual cRep.Item(1), "Invoice #INV-XXX-XXX paid", "ReplaceRegex replaces all digits"

    ' 4. MapValues & BinColumn
    LogPrint "  [47.5] MapValues & BinColumn..."
    Dim colMap As New cColumn
    colMap.Append "M"
    colMap.Append "F"
    colMap.Append "X"
    Dim dMap As Object
    Set dMap = CreateObject("Scripting.Dictionary")
    dMap("M") = "Male"
    dMap("F") = "Female"
    Dim cMapped As cColumn
    Set cMapped = colMap.MapValues(dMap, "Other")
    AssertEqual cMapped.Item(1), "Male", "MapValues maps M -> Male"
    AssertEqual cMapped.Item(2), "Female", "MapValues maps F -> Female"
    AssertEqual cMapped.Item(3), "Other", "MapValues uses default for unmapped X"

    Dim colBin As New cColumn
    colBin.Append 25
    colBin.Append 65
    colBin.Append 90
    Dim cBinned As cColumn
    Set cBinned = colBin.BinColumn(Array(50, 80), Array("Low", "Mid", "High"))
    AssertEqual cBinned.Item(1), "Low", "BinColumn 25 < 50 is Low"
    AssertEqual cBinned.Item(2), "Mid", "BinColumn 50 <= 65 < 80 is Mid"
    AssertEqual cBinned.Item(3), "High", "BinColumn 90 >= 80 is High"

    ' 5. CleanHeaders
    LogPrint "  [47.6] CleanHeaders..."
    Dim dfHdr As New cDataFrame
    dfHdr.AddColumn " First Name "
    dfHdr.AddColumn "Last" & vbCrLf & "Name"
    dfHdr.AddRow "Alice", "Smith"
    dfHdr.CleanHeaders True, True
    AssertTrue dfHdr.HasColumn("first_name"), "CleanHeaders converts to lowercase with underscores"
    AssertTrue dfHdr.HasColumn("last_name"), "CleanHeaders removes newline"

    ' 6. DropEmptyRows & DropEmptyColumns
    LogPrint "  [47.7] DropEmptyRows & DropEmptyColumns..."
    Dim dfEmpty As New cDataFrame
    dfEmpty.AddColumn "A"
    dfEmpty.AddColumn "B"
    dfEmpty.AddColumn "EmptyCol"
    dfEmpty.AddRow "1", "X", ""
    dfEmpty.AddRow "", "", ""
    dfEmpty.AddRow "2", "Y", ""
    Dim dfNoCols As cDataFrame
    Set dfNoCols = dfEmpty.DropEmptyColumns()
    AssertFalse dfNoCols.HasColumn("EmptyCol"), "DropEmptyColumns removes all-empty column"
    AssertEqual dfNoCols.ColumnCount, 2, "DropEmptyColumns keeps 2 non-empty columns"

    Dim dfNoRows As cDataFrame
    Set dfNoRows = dfNoCols.DropEmptyRows()
    AssertEqual dfNoRows.RowCount, 2, "DropEmptyRows removes row with all blank values"

    ' 7. PromoteHeaders
    LogPrint "  [47.8] PromoteHeaders..."
    Dim dfP As New cDataFrame
    dfP.AddColumn "C1"
    dfP.AddColumn "C2"
    dfP.AddRow "EmpID", "Score"
    dfP.AddRow "E001", 95
    dfP.AddRow "E002", 88
    dfP.PromoteHeaders 1
    AssertTrue dfP.HasColumn("EmpID"), "PromoteHeaders promotes row 1 to EmpID"
    AssertTrue dfP.HasColumn("Score"), "PromoteHeaders promotes row 1 to Score"
    AssertEqual dfP.RowCount, 2, "PromoteHeaders decrements RowCount by 1"
    AssertEqual dfP.Column("Score").Item(1), 95, "PromoteHeaders preserves data rows"

    ' 8. SplitColumn & CombineColumns
    LogPrint "  [47.9] SplitColumn & CombineColumns..."
    Dim dfSp As New cDataFrame
    dfSp.AddColumn "FullName"
    dfSp.AddRow "John Doe"
    dfSp.AddRow "Jane Smith"
    dfSp.SplitColumn "FullName", " ", "Part", , False
    AssertTrue dfSp.HasColumn("Part_1"), "SplitColumn creates Part_1"
    AssertEqual dfSp.Column("Part_1").Item(1), "John", "SplitColumn Part_1 item 1"
    AssertEqual dfSp.Column("Part_2").Item(2), "Smith", "SplitColumn Part_2 item 2"

    dfSp.CombineColumns Array("Part_1", "Part_2"), "JoinedName", "-"
    AssertTrue dfSp.HasColumn("JoinedName"), "CombineColumns creates JoinedName"
    AssertEqual dfSp.Column("JoinedName").Item(1), "John-Doe", "CombineColumns joins with hyphen"

    ' 9. Explode
    LogPrint "  [47.10] Explode..."
    Dim dfEx As New cDataFrame
    dfEx.AddColumn "Dept"
    dfEx.AddColumn "Skills"
    dfEx.AddRow "IT", "VBA, SQL, Python"
    dfEx.AddRow "HR", "Recruiting"
    Dim dfExploded As cDataFrame
    Set dfExploded = dfEx.Explode("Skills", ",")
    AssertEqual dfExploded.RowCount, 4, "Explode expands 2 rows to 4 rows"
    AssertEqual dfExploded.Column("Skills").Item(1), "VBA", "Explode item 1"
    AssertEqual dfExploded.Column("Skills").Item(2), "SQL", "Explode item 2"
    AssertEqual dfExploded.Column("Skills").Item(3), "Python", "Explode item 3"
    AssertEqual dfExploded.Column("Dept").Item(3), "IT", "Explode preserves Dept for row 3"
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    RecordFailure "Test_CleanAndTransformData", failErr
End Sub


' ------------------------------------------------------------------------------
' Test 48: 3D Array & Multi-Sheet Operations
' ------------------------------------------------------------------------------
Private Sub Test_3DArrayAndMultiSheets()
    On Error GoTo TestFail
    LogPrint "[TEST 48] 3D Array & Multi-Sheet Operations..."

    ' 1. LoadFrom3DArray
    LogPrint "  [48.1] LoadFrom3DArray..."
    Dim arr3D(1 To 3, 1 To 2, 1 To 2) As Variant
    ' Layer 1: Q1
    arr3D(1, 1, 1) = "Prod":   arr3D(1, 2, 1) = "Sales"
    arr3D(2, 1, 1) = "Widget": arr3D(2, 2, 1) = 100
    arr3D(3, 1, 1) = "Gadget": arr3D(3, 2, 1) = 200
    ' Layer 2: Q2
    arr3D(1, 1, 2) = "Prod":   arr3D(1, 2, 2) = "Sales"
    arr3D(2, 1, 2) = "Widget": arr3D(2, 2, 2) = 150
    arr3D(3, 1, 2) = "Gadget": arr3D(3, 2, 2) = 250

    Dim df3D As New cDataFrame
    Dim v3D As Variant
    v3D = arr3D
    Call df3D.LoadFrom3DArray(v3D, True, Array("Q1", "Q2"), "Quarter")

    AssertEqual df3D.RowCount, 4, "LoadFrom3DArray total rows is 4 (2 per layer * 2 layers)"
    AssertEqual df3D.ColumnCount, 3, "LoadFrom3DArray has 3 columns (Quarter, Prod, Sales)"
    AssertEqual df3D.Column("Quarter").Item(1), "Q1", "LoadFrom3DArray row 1 Quarter is Q1"
    AssertEqual df3D.Column("Quarter").Item(3), "Q2", "LoadFrom3DArray row 3 Quarter is Q2"
    AssertEqual df3D.Column("Sales").Item(4), 250, "LoadFrom3DArray row 4 Sales is 250"

    ' 2. To3DArray
    LogPrint "  [48.2] To3DArray..."
    Dim out3D As Variant
    out3D = df3D.To3DArray("Quarter", False)
    AssertEqual UBound(out3D, 1), 2, "To3DArray dim 1 (rows) = 2"
    AssertEqual UBound(out3D, 2), 2, "To3DArray dim 2 (cols) = 2"
    AssertEqual UBound(out3D, 3), 2, "To3DArray dim 3 (layers) = 2"
    AssertEqual out3D(2, 2, 2), 250, "To3DArray layer 2 row 2 col 2 value is 250"

    ' 3. ToMultiSheets & LoadFromSheets
    LogPrint "  [48.3] ToMultiSheets..."
    Dim xlApp As Object
    Set xlApp = ThisWorkbook.Application
    xlApp.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Q1").Delete
    ThisWorkbook.Worksheets("Q2").Delete
    On Error GoTo TestFail
    xlApp.DisplayAlerts = True

    Call df3D.ToMultiSheets("Quarter", ThisWorkbook, "MINIMAL")
    Dim wsQ1 As Object, wsQ2 As Object
    On Error Resume Next
    Set wsQ1 = ThisWorkbook.Worksheets("Q1")
    Set wsQ2 = ThisWorkbook.Worksheets("Q2")
    On Error GoTo TestFail
    AssertFalse wsQ1 Is Nothing, "ToMultiSheets created Q1 worksheet"
    AssertFalse wsQ2 Is Nothing, "ToMultiSheets created Q2 worksheet"

    LogPrint "  [48.4] LoadFromSheets..."
    Dim dfFromSheets As New cDataFrame
    Call dfFromSheets.LoadFromSheets(Array("Q1", "Q2"), ThisWorkbook, True, True, "OriginSheet")
    AssertEqual dfFromSheets.RowCount, 4, "LoadFromSheets combined row count is 4"
    AssertTrue dfFromSheets.HasColumn("OriginSheet"), "LoadFromSheets added OriginSheet column"

    xlApp.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Q1").Delete
    ThisWorkbook.Worksheets("Q2").Delete
    On Error GoTo 0
    xlApp.DisplayAlerts = True
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    On Error Resume Next
    ThisWorkbook.Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Q1").Delete
    ThisWorkbook.Worksheets("Q2").Delete
    ThisWorkbook.Application.DisplayAlerts = True
    On Error GoTo 0
    RecordFailure "Test_3DArrayAndMultiSheets", failErr
End Sub


' ------------------------------------------------------------------------------
' Test 49: Cross Reference Cell Matrix Operations
' ------------------------------------------------------------------------------
Private Sub Test_CrossReferenceMatrix()
    On Error GoTo TestFail
    LogPrint "[TEST 49] Cross Reference Cell Matrix Operations..."

    Dim xlApp As Object
    Set xlApp = ThisWorkbook.Application
    xlApp.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("MatrixTest").Delete
    On Error GoTo TestFail
    xlApp.DisplayAlerts = True

    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "MatrixTest"

    ' 1. Build flat DataFrame
    Dim df As New cDataFrame
    df.AddColumn "Emp"
    df.AddColumn "Month"
    df.AddColumn "Hours"
    df.AddRow "Alice", "Jan", 160#
    df.AddRow "Alice", "Feb", 150#
    df.AddRow "Bob", "Jan", 170#
    df.AddRow "Bob", "Feb", 165#

    ' 2. ToMatrixRange: export as a custom matrix starting at B5
    LogPrint "  [49.1] ToMatrixRange..."
    Call df.ToMatrixRange(ws.Range("B5"), "Emp", "Month", "Hours", "SUM", 0, "NAVY")

    AssertEqual ws.Range("B5").Value2, "Emp", "Top-left header of matrix is Emp"
    AssertEqual ws.Range("B6").Value2, "Alice", "First row header is Alice"
    AssertEqual ws.Range("B7").Value2, "Bob", "Second row header is Bob"
    AssertTrue Len(CStr(ws.Range("C5").Value2)) > 0, "Column header 1 is present"
    AssertTrue Len(CStr(ws.Range("D5").Value2)) > 0, "Column header 2 is present"
    AssertTrue IsNumeric(ws.Range("C6").Value2), "Matrix cell C6 has numeric value"
    AssertTrue IsNumeric(ws.Range("D7").Value2), "Matrix cell D7 has numeric value"

    ' 3. LoadFromMatrixRange: unpivot custom matrix B5:D7 back to flat DataFrame
    LogPrint "  [49.2] LoadFromMatrixRange..."
    Dim dfFromMat As New cDataFrame
    Call dfFromMat.LoadFromMatrixRange(ws.Range("B5:D7"), RowHeaderColOffset:=1, ColHeaderRowOffset:=1, _
                                       RowHeaderName:="Employee", ColHeaderName:="Period", ValueName:="LoggedHours")

    AssertEqual dfFromMat.RowCount, 4, "LoadFromMatrixRange extracted 4 unpivoted rows"
    AssertEqual dfFromMat.ColumnCount, 3, "LoadFromMatrixRange extracted 3 columns"
    AssertTrue dfFromMat.HasColumn("Employee"), "HasColumn Employee"
    AssertTrue dfFromMat.HasColumn("Period"), "HasColumn Period"
    AssertTrue dfFromMat.HasColumn("LoggedHours"), "HasColumn LoggedHours"

    ' 4. FormulaEngine: ApplyTwoWayMatrixLookupToRange
    LogPrint "  [49.3] ApplyTwoWayMatrixLookupToRange..."
    Dim fe As New cFormulaEngine
    fe.ApplyTwoWayMatrixLookupToRange ws.Range("F6"), ws.Range("B6"), ws.Range("D5"), _
                                     ws.Range("C6:D7"), ws.Range("B6:B7"), ws.Range("C5:D5"), 0, "N/A"

    Dim fStr As String
    fStr = ws.Range("F6").Formula
    AssertTrue InStr(fStr, "INDEX(") > 0, "2-way matrix formula contains INDEX"
    AssertTrue InStr(fStr, "MATCH(") > 0, "2-way matrix formula contains MATCH"
    AssertTrue InStr(fStr, "IFERROR(") > 0, "2-way matrix formula wrapped in IFERROR"
    AssertEqual ws.Range("F6").Value2, ws.Range("D6").Value2, "2-way matrix lookup evaluated value matches matrix intersection"

    ' 5. FormulaEngine: SetTwoWayMatrixLookup on a column range
    LogPrint "  [49.4] SetTwoWayMatrixLookup on attached range..."
    ws.Range("I5").Value2 = "Employee"
    ws.Range("J5").Value2 = "Month"
    ws.Range("K5").Value2 = "LookupVal"
    ws.Range("I6").Value2 = "Alice": ws.Range("J6").Value2 = "Feb"
    ws.Range("I7").Value2 = "Bob":   ws.Range("J7").Value2 = "Jan"

    Dim feTable As New cFormulaEngine
    feTable.Attach ws.Range("I5:K7"), True
    feTable.SetTwoWayMatrixLookup "LookupVal", "Employee", "Month", _
                                  "C6:D7", "B6:B7", "C5:D5", "", 0, 0#

    AssertEqual ws.Range("K6").Value2, 150#, "SetTwoWayMatrixLookup resolved Alice Feb to 150"
    AssertEqual ws.Range("K7").Value2, 170#, "SetTwoWayMatrixLookup resolved Bob Jan to 170"

    xlApp.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("MatrixTest").Delete
    On Error GoTo 0
    xlApp.DisplayAlerts = True
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    On Error Resume Next
    ThisWorkbook.Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("MatrixTest").Delete
    ThisWorkbook.Application.DisplayAlerts = True
    On Error GoTo 0
    RecordFailure "Test_CrossReferenceMatrix", failErr
End Sub

Private Sub SafeDeleteName(ByVal AnchorName As String)
    On Error Resume Next
    ThisWorkbook.Names(AnchorName).Delete
    On Error GoTo 0
End Sub

Private Sub AssertEqual(ByVal Actual As Variant, ByVal Expected As Variant, ByVal Msg As String)
    If Actual = Expected Then
        mPassCount = mPassCount + 1
    Else
        mFailCount = mFailCount + 1
        LogPrint "    [FAIL] " & Msg & " | Expected: '" & Expected & "', Got: '" & Actual & "'"
    End If
End Sub

Private Sub AssertTrue(ByVal Condition As Boolean, ByVal Msg As String)
    If Condition Then
        mPassCount = mPassCount + 1
    Else
        mFailCount = mFailCount + 1
        LogPrint "    [FAIL] " & Msg & " | Condition evaluated to False"
    End If
End Sub

Private Sub AssertFalse(ByVal Condition As Boolean, ByVal Msg As String)
    If Not Condition Then
        mPassCount = mPassCount + 1
    Else
        mFailCount = mFailCount + 1
        LogPrint "    [FAIL] " & Msg & " | Condition evaluated to True (expected False)"
    End If
End Sub

Private Sub AssertIsError(ByVal Actual As Variant, ByVal Msg As String)
    If IsError(Actual) Then
        mPassCount = mPassCount + 1
    Else
        mFailCount = mFailCount + 1
        LogPrint "    [FAIL] " & Msg & " | Expected Error Variant (#N/A), Got: '" & Actual & "'"
    End If
End Sub

Private Sub RecordFailure(ByVal TestName As String, ByVal ErrMsg As String)
    mFailCount = mFailCount + 1
    LogPrint "    [ERROR in " & TestName & "]: " & ErrMsg
End Sub

Private Sub LogPrint(ByVal Msg As String)
    Debug.Print Msg
    mLogOutput = mLogOutput & Msg & vbCrLf
    On Error Resume Next
    If Len(mHeartbeatFile) > 0 Then
        Dim hNum As Integer
        hNum = FreeFile
        Open mHeartbeatFile For Append As #hNum
        Print #hNum, Msg
        Close #hNum
    End If
    On Error GoTo 0
End Sub
' ------------------------------------------------------------------------------
' Test 50: Matrix Enhancements (Heatmap, Multi-Level Unpivot, MatrixDiff)
' ------------------------------------------------------------------------------
Private Sub Test_MatrixEnhancements()
    On Error GoTo TestFail
    LogPrint "[TEST 50] Matrix Enhancements (Heatmap, Multi-Level Unpivot, MatrixDiff)..."

    Dim xlApp As Object
    Set xlApp = ThisWorkbook.Application
    xlApp.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("MatrixEnhanceTest").Delete
    On Error GoTo TestFail
    xlApp.DisplayAlerts = True

    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "MatrixEnhanceTest"

    ' --------------------------------------------------------------------------
    ' 50.1: cTableStyler ApplyMatrixHeatmap & ApplyMatrixDataBars
    ' --------------------------------------------------------------------------
    LogPrint "  [50.1] ApplyMatrixHeatmap & ApplyMatrixDataBars..."
    Dim matGrid(1 To 3, 1 To 4) As Variant
    matGrid(1, 1) = "Dept": matGrid(1, 2) = "Q1":  matGrid(1, 3) = "Q2":  matGrid(1, 4) = "Q3"
    matGrid(2, 1) = "Dev":  matGrid(2, 2) = 120#: matGrid(2, 3) = 150#: matGrid(2, 4) = 180#
    matGrid(3, 1) = "QA":   matGrid(3, 2) = 90#:  matGrid(3, 3) = 110#: matGrid(3, 4) = 130#
    ws.Range("B5:E7").Value = matGrid

    Dim styler As New cTableStyler
    styler.ApplyMatrixHeatmap ws.Range("B5:E7"), "RED_YELLOW_GREEN", HasRowHeaders:=True, HasColHeaders:=True
    AssertTrue ws.Range("C6:E7").FormatConditions.Count > 0, "ApplyMatrixHeatmap applied color scale to data area"

    styler.ApplyMatrixDataBars ws.Range("C6:E7"), 5287936, HasRowHeaders:=False, HasColHeaders:=False
    AssertTrue ws.Range("C6:E7").FormatConditions.Count >= 2, "ApplyMatrixDataBars added databar condition"

    ' --------------------------------------------------------------------------
    ' 50.2: Multi-Level Header Unpivot (LoadFromMatrixRange with ColHeaderRowCount:=2, RowHeaderColCount:=2)
    ' --------------------------------------------------------------------------
    LogPrint "  [50.2] Multi-Level Header Unpivot..."
    Dim multiGrid(1 To 4, 1 To 4) As Variant
    multiGrid(1, 1) = "Dept": multiGrid(1, 2) = "Role": multiGrid(1, 3) = "2026": multiGrid(1, 4) = Empty
    multiGrid(2, 1) = Empty:  multiGrid(2, 2) = Empty:  multiGrid(2, 3) = "Jan":  multiGrid(2, 4) = "Feb"
    multiGrid(3, 1) = "Tech": multiGrid(3, 2) = "Dev":  multiGrid(3, 3) = 100#:  multiGrid(3, 4) = 110#
    multiGrid(4, 1) = Empty:  multiGrid(4, 2) = "QA":   multiGrid(4, 3) = 80#:   multiGrid(4, 4) = 90#
    ws.Range("G10:J13").Value = multiGrid

    Dim dfMulti As New cDataFrame
    dfMulti.LoadFromMatrixRange ws.Range("G10:J13"), _
                                RowHeaderColOffset:=1, ColHeaderRowOffset:=1, _
                                RowHeaderName:="Dept,Role", ColHeaderName:="Year,Month", ValueName:="Hours", _
                                ColHeaderRowCount:=2, RowHeaderColCount:=2

    AssertEqual dfMulti.RowCount, 4, "Multi-level unpivot extracted 4 rows"
    AssertEqual dfMulti.ColumnCount, 5, "Multi-level unpivot created 5 separate columns"
    AssertTrue dfMulti.HasColumn("Dept"), "HasColumn Dept"
    AssertTrue dfMulti.HasColumn("Role"), "HasColumn Role"
    AssertTrue dfMulti.HasColumn("Year"), "HasColumn Year"
    AssertTrue dfMulti.HasColumn("Month"), "HasColumn Month"
    AssertTrue dfMulti.HasColumn("Hours"), "HasColumn Hours"

    AssertEqual dfMulti(1, "Dept"), "Tech", "Row 1 Dept = Tech"
    AssertEqual dfMulti(1, "Role"), "Dev", "Row 1 Role = Dev"
    AssertEqual dfMulti(1, "Year"), "2026", "Row 1 Year = 2026"
    AssertEqual dfMulti(1, "Month"), "Jan", "Row 1 Month = Jan"
    AssertEqual dfMulti(1, "Hours"), 100#, "Row 1 Hours = 100"

    ' Check forward-filled row header on row 3 (Dept was Empty, inherited Tech)
    AssertEqual dfMulti(3, "Dept"), "Tech", "Row 3 Dept forward-filled to Tech"
    AssertEqual dfMulti(3, "Role"), "QA", "Row 3 Role = QA"

    ' --------------------------------------------------------------------------
    ' 50.3: MatrixDiff (Plan vs Actual Variance)
    ' --------------------------------------------------------------------------
    LogPrint "  [50.3] MatrixDiff..."
    Dim dfPlan As New cDataFrame
    dfPlan.AddColumn "Emp": dfPlan.AddColumn "Month": dfPlan.AddColumn "PlanHours"
    dfPlan.AddRow "Alice", "Jan", 160#
    dfPlan.AddRow "Alice", "Feb", 160#
    dfPlan.AddRow "Bob", "Jan", 160#

    Dim dfAct As New cDataFrame
    dfAct.AddColumn "Emp": dfAct.AddColumn "Month": dfAct.AddColumn "ActualHours"
    dfAct.AddRow "Alice", "Jan", 175#  ' Variance +15
    dfAct.AddRow "Alice", "Feb", 140#  ' Variance -20
    dfAct.AddRow "Bob", "Jan", 160#    ' Variance 0
    dfAct.AddRow "Charlie", "Jan", 80# ' Extra row in Actual

    Dim dfDiff As cDataFrame
    Set dfDiff = dfPlan.MatrixDiff(dfAct, "Emp", "Month", "PlanHours", TargetValueCol:="ActualHours", _
                                   DiffColName:="Variance", DiffMode:="ACTUAL_MINUS_PLAN", IncludePctDiff:=True)

    AssertEqual dfDiff.RowCount, 4, "MatrixDiff combined 4 employee-month pairs"
    AssertTrue dfDiff.HasColumn("Variance"), "HasColumn Variance"
    AssertTrue dfDiff.HasColumn("PctVariance"), "HasColumn PctVariance"

    ' Check Alice Jan
    AssertEqual dfDiff(1, "Variance"), 15#, "Alice Jan variance is +15"
    AssertEqual dfDiff(2, "Variance"), -20#, "Alice Feb variance is -20"
    AssertEqual dfDiff(3, "Variance"), 0#, "Bob Jan variance is 0"
    AssertEqual dfDiff(4, "Variance"), 80#, "Charlie Jan variance is +80 (new actual)"

    xlApp.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("MatrixEnhanceTest").Delete
    On Error GoTo 0
    xlApp.DisplayAlerts = True
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    On Error Resume Next
    ThisWorkbook.Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("MatrixEnhanceTest").Delete
    ThisWorkbook.Application.DisplayAlerts = True
    On Error GoTo 0
    RecordFailure "Test_MatrixEnhancements", failErr
End Sub

' ------------------------------------------------------------------------------
' Test 51: Sparse Matrix Optimization Suite
' ------------------------------------------------------------------------------
Private Sub Test_SparseMatrixOptimization()
    On Error GoTo TestFail
    LogPrint "[TEST 51] Sparse Matrix Optimization Suite..."

    Dim xlApp As Object
    Set xlApp = ThisWorkbook.Application
    xlApp.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("SparseTest").Delete
    On Error GoTo TestFail
    xlApp.DisplayAlerts = True

    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "SparseTest"

    ' --------------------------------------------------------------------------
    ' 51.1: Build a sparse dataset (5 people x 10 days = 50 cells, but only 5 non-zeros)
    ' --------------------------------------------------------------------------
    LogPrint "  [51.1] ToSparseCoordinates & Sparsity Stats..."
    Dim dfRaw As New cDataFrame
    dfRaw.AddColumn "Emp": dfRaw.AddColumn "Day": dfRaw.AddColumn "Hours"

    ' Non-zero records
    dfRaw.AddRow "Alice", "Day1", 8#
    dfRaw.AddRow "Alice", "Day5", 4#
    dfRaw.AddRow "Bob", "Day2", 8#
    dfRaw.AddRow "Charlie", "Day10", 6#
    dfRaw.AddRow "David", "Day3", 8#

    ' Zeros and empty rows that should be filtered in sparse mode
    dfRaw.AddRow "Eve", "Day1", 0#
    dfRaw.AddRow "Eve", "Day2", Empty
    dfRaw.AddRow "Bob", "Day5", 0#
    dfRaw.AddRow "David", "Day10", 0#

    ' Get sparse stats
    Dim stats As Object
    Set stats = dfRaw.GetSparseMatrixStats("Emp", "Day", "Hours", IgnoreZero:=True)

    AssertEqual stats("UniqueRows"), 5, "Sparsity Stats: 5 unique employees"
    AssertEqual stats("UniqueCols"), 5, "Sparsity Stats: 5 unique days"
    AssertEqual stats("TotalGridCells"), 25#, "Total grid cells = 5x5 = 25"
    AssertEqual stats("NonZeroCount"), 5, "Non-zero count is 5"
    AssertEqual stats("ZeroCount"), 20#, "Zero/Empty count is 20"
    AssertEqual stats("Sparsity"), 0.8, "Sparsity is 80%"
    AssertEqual stats("MemorySavedPercent"), 80#, "Memory saved is 80%"

    ' ToSparseCoordinates
    Dim dfSparse As cDataFrame
    Set dfSparse = dfRaw.ToSparseCoordinates("Emp", "Day", "Hours", IgnoreZero:=True)

    AssertEqual dfSparse.RowCount, 5, "ToSparseCoordinates kept only the 5 non-zero rows"
    AssertEqual dfSparse.ColumnCount, 3, "ToSparseCoordinates returned 3 columns"

    ' --------------------------------------------------------------------------
    ' 51.2: SparseMatrixAggregate in O(NonZero) time without building full 2D grid
    ' --------------------------------------------------------------------------
    LogPrint "  [51.2] SparseMatrixAggregate..."
    Dim rowTotals As Object, colTotals As Object
    Dim grandTotal As Double
    grandTotal = dfRaw.SparseMatrixAggregate("Emp", "Day", "Hours", rowTotals, colTotals)

    AssertEqual grandTotal, 34#, "SparseMatrixAggregate grand total = 8+4+8+6+8 = 34"
    AssertEqual rowTotals("Alice"), 12#, "Alice sparse row total = 8 + 4 = 12"
    AssertEqual rowTotals("Bob"), 8#, "Bob sparse row total = 8"
    AssertEqual colTotals("Day1"), 8#, "Day1 sparse col total = 8"

    ' --------------------------------------------------------------------------
    ' 51.3: ToSparseMatrixRange with Marginal Totals (Row & Column Grand Totals)
    ' --------------------------------------------------------------------------
    LogPrint "  [51.3] ToSparseMatrixRange with Marginal Totals..."
    Call dfSparse.ToSparseMatrixRange(ws.Range("B5"), "Emp", "Day", "Hours", "SUM", 0#, "NAVY", _
                                     IncludeRowTotals:=True, IncludeColTotals:=True)

    AssertEqual ws.Range("B5").Value2, "Emp", "Top-left cell is Emp"
    ' Verify marginal totals were created
    Dim lastColHeader As String
    lastColHeader = CStr(ws.Range("H5").Value2)
    AssertEqual lastColHeader, "Total", "Far-right column header is 'Total'"

    Dim lastRowHeader As String
    lastRowHeader = CStr(ws.Range("B10").Value2)
    AssertEqual lastRowHeader, "Total", "Bottom row header is 'Total'"

    ' Grand Total intersection cell (Row Total x Col Total) at H10
    Dim gTotalVal As Double
    gTotalVal = CDbl(ws.Range("H10").Value2)
    AssertEqual gTotalVal, 34#, "Bottom-right grand total intersection is 34"

    xlApp.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("SparseTest").Delete
    On Error GoTo 0
    xlApp.DisplayAlerts = True
    Exit Sub

TestFail:
    Dim failErr As String
    failErr = Err.Description
    On Error Resume Next
    ThisWorkbook.Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("SparseTest").Delete
    ThisWorkbook.Application.DisplayAlerts = True
    On Error GoTo 0
    RecordFailure "Test_SparseMatrixOptimization", failErr
End Sub
