VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFormulaEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' Class: cFormulaEngine
' Description: Dynamic Excel Formula Generator, Template Parser & AutoFill Engine.
'              Injects formulas into cells or columns using A1, R1C1, or Named
'              Column tokens (e.g. "[Qty] * [Price]"), with native Excel AutoFill
'              expansion (xlFillDefault, xlFillSeries, xlFillCopy), Cross-Sheet
'              Lookups, Summary Row aggregation, and Table-Driven batch configuration.
' Author: Korn (AI Backend Developer, Technodevmoo)
' Architecture: Kit (Tech Lead)
' ==============================================================================

' Excel AutoFill Constants (late-bound)
Private Const XL_FILL_DEFAULT As Long = 0
Private Const XL_FILL_COPY As Long = 1
Private Const XL_FILL_SERIES As Long = 2
Private Const XL_FILL_FORMATS As Long = 3
Private Const XL_FILL_VALUES As Long = 4

Private mWorksheet As Object
Private mTableRange As Object
Private mHeaderRange As Object
Private mDataBodyRange As Object
Private mTopLeftCell As Object
Private mListObject As Object
Private mRowCount As Long
Private mColumnCount As Long
Private mHasHeaders As Boolean
Private mUseStructuredReferences As Boolean
Private mInBatch As Boolean
Private mPrevCalculation As Long
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents As Boolean
Private mColIndexMap As Object ' Scripting.Dictionary: Name/Key -> 1-based column index

Private Sub Class_Initialize()
    mRowCount = 0
    mColumnCount = 0
    mHasHeaders = True
    mUseStructuredReferences = False
    mInBatch = False
    Set mListObject = Nothing
    Set mColIndexMap = CreateObject("Scripting.Dictionary")
    mColIndexMap.CompareMode = 1 ' TextCompare (case-insensitive)
End Sub

Private Sub Class_Terminate()
    If mInBatch Then
        On Error Resume Next
        Application.Calculation = mPrevCalculation
        Application.ScreenUpdating = mPrevScreenUpdating
        Application.EnableEvents = mPrevEnableEvents
        mInBatch = False
        On Error GoTo 0
    End If
    Set mWorksheet = Nothing
    Set mTableRange = Nothing
    Set mHeaderRange = Nothing
    Set mDataBodyRange = Nothing
    Set mTopLeftCell = Nothing
    Set mListObject = Nothing
    Set mColIndexMap = Nothing
End Sub

' ------------------------------------------------------------------------------
' Attachment & Range Binding
' ------------------------------------------------------------------------------

' Attaches the formula engine to an existing worksheet Range or ListObject.
Public Function Attach(ByVal TargetRangeOrSheet As Object, _
                       Optional ByVal HasHeaders As Boolean = True) As cFormulaEngine
    If TargetRangeOrSheet Is Nothing Then
        Err.Raise 91, "cFormulaEngine.Attach", "TargetRangeOrSheet cannot be Nothing."
    End If

    mColIndexMap.RemoveAll
    mHasHeaders = HasHeaders

    On Error Resume Next
    Dim testLo As Object
    Set testLo = TargetRangeOrSheet.ListColumns
    On Error GoTo 0

    If Not testLo Is Nothing Then
        Set Attach = AttachListObject(TargetRangeOrSheet)
        Exit Function
    End If

    Dim rng As Object
    Set rng = TargetRangeOrSheet

    Set mListObject = Nothing
    Set mTableRange = rng
    Set mWorksheet = rng.Worksheet
    Set mTopLeftCell = rng.Cells(1, 1)
    mColumnCount = rng.Columns.Count

    If mHasHeaders Then
        If rng.Rows.Count < 2 Then
            mRowCount = 0
            Set mHeaderRange = rng.Rows(1)
            Set mDataBodyRange = Nothing
        Else
            mRowCount = rng.Rows.Count - 1
            Set mHeaderRange = rng.Rows(1)
            Set mDataBodyRange = rng.Offset(1, 0).Resize(mRowCount, mColumnCount)
        End If

        ' Index column headers
        Dim c As Long, hName As String
        For c = 1 To mColumnCount
            hName = Trim$(CStr(mHeaderRange.Cells(1, c).Value2))
            If Len(hName) = 0 Then hName = Trim$(CStr(mHeaderRange.Cells(1, c).Value))
            If Len(hName) > 0 Then
                If Not mColIndexMap.Exists(hName) Then mColIndexMap.Add hName, c
            End If
            ' Index positional column key "C1", "C2"...
            If Not mColIndexMap.Exists("C" & CStr(c)) Then mColIndexMap.Add "C" & CStr(c), c
        Next c
    Else
        mRowCount = rng.Rows.Count
        Set mHeaderRange = Nothing
        Set mDataBodyRange = rng
        For c = 1 To mColumnCount
            mColIndexMap.Add "C" & CStr(c), c
        Next c
    End If

    Set Attach = Me
End Function

Public Function AttachTo(ByVal TargetRangeOrSheet As Object, _
                         Optional ByVal HasHeaders As Boolean = True) As cFormulaEngine
    Set AttachTo = Attach(TargetRangeOrSheet, HasHeaders)
End Function

' Attaches directly to an Excel structured Table (ListObject).
Public Function AttachListObject(ByVal TableObj As Object) As cFormulaEngine
    If TableObj Is Nothing Then
        Err.Raise 91, "cFormulaEngine.AttachListObject", "TableObj cannot be Nothing."
    End If

    mColIndexMap.RemoveAll
    mHasHeaders = True

    Set mListObject = TableObj
    Set mTableRange = TableObj.Range
    Set mWorksheet = TableObj.Range.Worksheet
    Set mTopLeftCell = TableObj.Range.Cells(1, 1)
    Set mHeaderRange = TableObj.HeaderRowRange
    Set mDataBodyRange = TableObj.DataBodyRange

    mColumnCount = TableObj.ListColumns.Count
    If mDataBodyRange Is Nothing Then
        mRowCount = 0
    Else
        mRowCount = TableObj.ListRows.Count
    End If

    Dim colObj As Object, idx As Long
    idx = 1
    For Each colObj In TableObj.ListColumns
        Dim colName As String
        colName = Trim$(colObj.Name)
        If Len(colName) > 0 And Not mColIndexMap.Exists(colName) Then
            mColIndexMap.Add colName, idx
        End If
        If Not mColIndexMap.Exists("C" & CStr(idx)) Then
            mColIndexMap.Add "C" & CStr(idx), idx
        End If
        idx = idx + 1
    Next colObj

    Set AttachListObject = Me
End Function

' Attaches using cDataFrame metadata and column names.
Public Function AttachDataFrame(ByVal TargetTopLeft As Object, _
                                ByVal Df As cDataFrame, _
                                Optional ByVal HasHeaders As Boolean = True) As cFormulaEngine
    If TargetTopLeft Is Nothing Then
        Err.Raise 91, "cFormulaEngine.AttachDataFrame", "TargetTopLeft cannot be Nothing."
    End If
    If Df Is Nothing Then
        Err.Raise 91, "cFormulaEngine.AttachDataFrame", "Df cannot be Nothing."
    End If

    Dim totalRows As Long
    totalRows = Df.RowCount + IIf(HasHeaders, 1, 0)
    If totalRows < 1 Then totalRows = 1

    Dim targetRng As Object
    Set targetRng = TargetTopLeft.Cells(1, 1).Resize(totalRows, Df.ColumnCount)
    Attach targetRng, HasHeaders

    ' Ensure all column names from Df are mapped
    Dim names() As String
    names = Df.ColumnNames()
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Not mColIndexMap.Exists(names(i)) Then
            mColIndexMap.Add names(i), (i - LBound(names) + 1)
        End If
    Next i

    Set AttachDataFrame = Me
End Function

' ------------------------------------------------------------------------------
' Formula Injection & AutoFill Operations
' ------------------------------------------------------------------------------

' Applies a formula to an entire column.
' FormulaTemplate supports:
'   1. Named Column Tokens: "[Qty] * [UnitPrice]" or "[Amount] - [Discount]"
'   2. Row Tokens: "{row}" -> current absolute row number (e.g. "=D{row}+E{row}")
'   3. Boundary Tokens: "{first_row}", "{last_row}" -> boundary row numbers
'   4. Col Letter Tokens: "{col}" -> letter of target column
'   5. Direct R1C1 notation (if IsR1C1:=True or template starts with "=RC")
' If UseAutoFill is True, writes to top cell and performs native Excel AutoFill down to mRowCount.
' If UseAutoFill is False, performs instant bulk Range.Formula assignment across the column (10-50x faster).
Public Function SetColumnFormula(ByVal ColumnKey As Variant, _
                                 ByVal FormulaTemplate As String, _
                                 Optional ByVal UseAutoFill As Boolean = True, _
                                 Optional ByVal FillType As Long = 0, _
                                 Optional ByVal IsR1C1 As Boolean = False, _
                                 Optional ByVal PreferFormula2 As Boolean = False) As cFormulaEngine
    CheckAttached "SetColumnFormula"
    If mRowCount = 0 Then
        Set SetColumnFormula = Me
        Exit Function
    End If

    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)

    Dim colDataRng As Object
    Set colDataRng = mDataBodyRange.Columns(cIdx)

    Dim firstDataRow As Long
    firstDataRow = mDataBodyRange.Row

    Dim cleanTmpl As String
    cleanTmpl = Trim$(FormulaTemplate)
    If Left$(cleanTmpl, 1) = "'" Then cleanTmpl = Trim$(Mid$(cleanTmpl, 2))
    If Left$(cleanTmpl, 1) <> "=" Then cleanTmpl = "=" & cleanTmpl
    cleanTmpl = NormalizeFormulaSeparators(cleanTmpl)

    Dim detectR1C1 As Boolean
    detectR1C1 = IsR1C1 Or (InStr(1, cleanTmpl, "=RC", vbTextCompare) = 1) Or (InStr(1, cleanTmpl, "R[", vbTextCompare) > 0)

    If detectR1C1 Then
        colDataRng.FormulaR1C1 = cleanTmpl
    Else
        ' Parse template for first data row
        Dim parsedFirstFormula As String
        parsedFirstFormula = ParseFormulaTemplate(cleanTmpl, cIdx, firstDataRow)

        Dim firstCell As Object
        Set firstCell = colDataRng.Cells(1, 1)
        SafeAssignFormula firstCell, parsedFirstFormula, PreferFormula2

        If mRowCount > 1 Then
            If UseAutoFill Then
                On Error Resume Next
                firstCell.AutoFill Destination:=colDataRng, Type:=FillType
                If Err.Number <> 0 Then
                    Err.Clear
                    ' Fallback to bulk formula assignment if AutoFill is restricted
                    SafeAssignFormula colDataRng, parsedFirstFormula, PreferFormula2
                End If
                On Error GoTo 0
            Else
                SafeAssignFormula colDataRng, parsedFirstFormula, PreferFormula2
            End If
        End If
    End If

    Set SetColumnFormula = Me
End Function

' Applies formula across an entire column using instant bulk Range.Formula write (10-50x faster than AutoFill).
Public Function SetColumnFormulaBulk(ByVal ColumnKey As Variant, _
                                     ByVal FormulaTemplate As String, _
                                     Optional ByVal PreferFormula2 As Boolean = False) As cFormulaEngine
    Set SetColumnFormulaBulk = SetColumnFormula(ColumnKey, FormulaTemplate, _
                                                UseAutoFill:=False, _
                                                PreferFormula2:=PreferFormula2)
End Function

' Sets column formula using localized formula syntax (.FormulaLocal, e.g. German/French localized functions)
Public Function SetColumnFormulaLocal(ByVal ColumnKey As Variant, ByVal FormulaLocalTemplate As String) As cFormulaEngine
    CheckAttached "SetColumnFormulaLocal"
    If mRowCount = 0 Then
        Set SetColumnFormulaLocal = Me
        Exit Function
    End If

    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)

    Dim colDataRng As Object
    Set colDataRng = mDataBodyRange.Columns(cIdx)

    Dim firstDataRow As Long
    firstDataRow = mDataBodyRange.Row

    Dim cleanTmpl As String
    cleanTmpl = Trim$(FormulaLocalTemplate)
    If Left$(cleanTmpl, 1) = "'" Then cleanTmpl = Trim$(Mid$(cleanTmpl, 2))
    If Left$(cleanTmpl, 1) <> "=" Then cleanTmpl = "=" & cleanTmpl

    Dim parsedFormula As String
    parsedFormula = ParseFormulaTemplate(cleanTmpl, cIdx, firstDataRow)

    colDataRng.FormulaLocal = parsedFormula
    Set SetColumnFormulaLocal = Me
End Function

' Injects a modern Dynamic Array formula (spill range, e.g. =UNIQUE(), =FILTER(), =SORT()) into an anchor cell.
Public Function SetSpillFormula(ByVal AnchorCellOrColKey As Variant, _
                                ByVal FormulaTemplate As String, _
                                Optional ByVal PreferFormula2 As Boolean = True) As cFormulaEngine
    CheckAttached "SetSpillFormula"

    Dim anchorCell As Object
    Dim absRow As Long, cIdx As Long

    If IsObject(AnchorCellOrColKey) Then
        Set anchorCell = AnchorCellOrColKey
        absRow = anchorCell.Row
        cIdx = 1
    Else
        cIdx = ResolveColIndex(AnchorCellOrColKey)
        Set anchorCell = mDataBodyRange.Cells(1, cIdx)
        absRow = anchorCell.Row
    End If

    Dim cleanTmpl As String
    cleanTmpl = Trim$(FormulaTemplate)
    If Left$(cleanTmpl, 1) = "'" Then cleanTmpl = Trim$(Mid$(cleanTmpl, 2))
    If Left$(cleanTmpl, 1) <> "=" Then cleanTmpl = "=" & cleanTmpl

    Dim parsedFormula As String
    parsedFormula = ParseFormulaTemplate(cleanTmpl, cIdx, absRow)

    SafeAssignFormula anchorCell, parsedFormula, PreferFormula2
    On Error Resume Next
    anchorCell.Calculate
    On Error GoTo 0

    Set SetSpillFormula = Me
End Function

' Retrieves the dynamic Spill Range originating from an anchor cell, or the anchor cell itself if not spilling.
Public Function GetSpillRange(ByVal AnchorCellOrColKey As Variant) As Object
    CheckAttached "GetSpillRange"
    Dim anchorCell As Object
    If IsObject(AnchorCellOrColKey) Then
        Set anchorCell = AnchorCellOrColKey
    Else
        Dim cIdx As Long
        cIdx = ResolveColIndex(AnchorCellOrColKey)
        Set anchorCell = mDataBodyRange.Cells(1, cIdx)
    End If

    Dim spillRng As Object
    On Error Resume Next
    Set spillRng = anchorCell.SpillingToRange
    On Error GoTo 0

    If spillRng Is Nothing Then
        Set GetSpillRange = anchorCell
    Else
        Set GetSpillRange = spillRng
    End If
End Function

' Sets a formula directly using R1C1 notation across an entire column.
Public Function SetFormulaR1C1(ByVal ColumnKey As Variant, ByVal R1C1Formula As String) As cFormulaEngine
    Set SetFormulaR1C1 = SetColumnFormula(ColumnKey, R1C1Formula, UseAutoFill:=False, IsR1C1:=True)
End Function

' Sets a formula on a single specific cell in the table data body (1-based RowIndex).
Public Function SetCellFormula(ByVal RowIndex As Long, _
                               ByVal ColumnKey As Variant, _
                               ByVal FormulaTemplate As String, _
                               Optional ByVal IsR1C1 As Boolean = False) As cFormulaEngine
    CheckAttached "SetCellFormula"
    If RowIndex < 1 Or RowIndex > mRowCount Then
        Err.Raise 9, "cFormulaEngine.SetCellFormula", "RowIndex out of range: " & RowIndex & " (Valid: 1 To " & mRowCount & ")"
    End If

    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)

    Dim targetCell As Object
    Set targetCell = mDataBodyRange.Cells(RowIndex, cIdx)

    Dim absRow As Long
    absRow = targetCell.Row

    Dim cleanTmpl As String
    cleanTmpl = Trim$(FormulaTemplate)
    If Left$(cleanTmpl, 1) = "'" Then cleanTmpl = Trim$(Mid$(cleanTmpl, 2))
    If Left$(cleanTmpl, 1) <> "=" Then cleanTmpl = "=" & cleanTmpl

    If IsR1C1 Or InStr(cleanTmpl, "=RC") = 1 Then
        targetCell.FormulaR1C1 = cleanTmpl
    Else
        targetCell.Formula = ParseFormulaTemplate(cleanTmpl, cIdx, absRow)
    End If

    Set SetCellFormula = Me
End Function

' AutoFills an existing formula from the first cell of ColumnKey down across all rows.
Public Function AutoFillDown(ByVal ColumnKey As Variant, _
                             Optional ByVal FillType As Long = 0) As cFormulaEngine
    CheckAttached "AutoFillDown"
    If mRowCount <= 1 Then
        Set AutoFillDown = Me
        Exit Function
    End If

    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)

    Dim colDataRng As Object
    Set colDataRng = mDataBodyRange.Columns(cIdx)

    Dim topCell As Object
    Set topCell = colDataRng.Cells(1, 1)

    topCell.AutoFill Destination:=colDataRng, Type:=FillType

    Set AutoFillDown = Me
End Function

' AutoFills a source range into a destination range.
Public Function AutoFillRange(ByVal SourceRange As Object, _
                              ByVal DestinationRange As Object, _
                              Optional ByVal FillType As Long = 0) As cFormulaEngine
    If SourceRange Is Nothing Or DestinationRange Is Nothing Then
        Err.Raise 91, "cFormulaEngine.AutoFillRange", "SourceRange and DestinationRange cannot be Nothing."
    End If

    SourceRange.AutoFill Destination:=DestinationRange, Type:=FillType
    Set AutoFillRange = Me
End Function

' ------------------------------------------------------------------------------
' High-Level Formula Generators (Summary, VLOOKUP, Running Total)
' ------------------------------------------------------------------------------

' Injects an aggregation formula (SUM, AVERAGE, MIN, MAX, COUNT, SUBTOTAL) referencing a column.
' TargetCellOrRow: Range cell, or row offset relative to table (e.g. mRowCount + 1).
Public Function SetSummaryFormula(ByVal TargetCellOrRow As Variant, _
                                  ByVal TargetColKey As Variant, _
                                  ByVal AggregationType As String, _
                                  Optional ByVal SourceColKey As Variant = Empty) As cFormulaEngine
    CheckAttached "SetSummaryFormula"

    Dim tColIdx As Long
    tColIdx = ResolveColIndex(TargetColKey)

    Dim sColIdx As Long
    If IsEmpty(SourceColKey) Then
        sColIdx = tColIdx
    Else
        sColIdx = ResolveColIndex(SourceColKey)
    End If

    Dim sColLetter As String
    sColLetter = ColIndexToLetter(mDataBodyRange.Column + sColIdx - 1)

    Dim rStart As Long, rEnd As Long
    rStart = mDataBodyRange.Row
    rEnd = mDataBodyRange.Row + mRowCount - 1

    Dim rangeRef As String
    rangeRef = sColLetter & CStr(rStart) & ":" & sColLetter & CStr(rEnd)

    Dim formulaStr As String
    Select Case UCase$(Trim$(AggregationType))
        Case "SUM"
            formulaStr = "=SUM(" & rangeRef & ")"
        Case "AVG", "AVERAGE"
            formulaStr = "=AVERAGE(" & rangeRef & ")"
        Case "MIN"
            formulaStr = "=MIN(" & rangeRef & ")"
        Case "MAX"
            formulaStr = "=MAX(" & rangeRef & ")"
        Case "COUNT"
            formulaStr = "=COUNT(" & rangeRef & ")"
        Case "COUNTA"
            formulaStr = "=COUNTA(" & rangeRef & ")"
        Case "SUBTOTAL", "SUBTOTAL_SUM"
            formulaStr = "=SUBTOTAL(9," & rangeRef & ")"
        Case "SUBTOTAL_AVG", "SUBTOTAL_AVERAGE"
            formulaStr = "=SUBTOTAL(1," & rangeRef & ")"
        Case Else
            formulaStr = "=" & Trim$(AggregationType) & "(" & rangeRef & ")"
    End Select

    Dim destCell As Object
    If IsObject(TargetCellOrRow) Then
        Set destCell = TargetCellOrRow
    ElseIf IsNumeric(TargetCellOrRow) Then
        Dim rowNum As Long
        rowNum = CLng(TargetCellOrRow)
        If rowNum <= mRowCount Then
            Set destCell = mDataBodyRange.Cells(rowNum, tColIdx)
        Else
            Set destCell = mWorksheet.Cells(mDataBodyRange.Row + rowNum - 1, mDataBodyRange.Column + tColIdx - 1)
        End If
    Else
        Set destCell = mWorksheet.Range(CStr(TargetCellOrRow))
    End If

    destCell.Formula = formulaStr
    Set SetSummaryFormula = Me
End Function

' Injects a cross-sheet VLOOKUP formula across an entire column.
Public Function SetCrossSheetLookup(ByVal TargetColKey As Variant, _
                                    ByVal KeyColKey As Variant, _
                                    ByVal LookupSheetName As String, _
                                    ByVal LookupRangeAddress As String, _
                                    ByVal ReturnColIndex As Long, _
                                    Optional ByVal ExactMatch As Boolean = True) As cFormulaEngine
    CheckAttached "SetCrossSheetLookup"

    Dim kColIdx As Long
    kColIdx = ResolveColIndex(KeyColKey)
    Dim kColLetter As String
    kColLetter = ColIndexToLetter(mDataBodyRange.Column + kColIdx - 1)

    Dim sheetRef As String
    If InStr(LookupSheetName, " ") > 0 Or InStr(LookupSheetName, "-") > 0 Then
        sheetRef = "'" & Replace(LookupSheetName, "'", "''") & "'!"
    Else
        sheetRef = LookupSheetName & "!"
    End If

    Dim tmpl As String
    tmpl = "=VLOOKUP(" & kColLetter & "{row}, " & sheetRef & LookupRangeAddress & ", " & _
           CStr(ReturnColIndex) & ", " & IIf(ExactMatch, "FALSE", "TRUE") & ")"

    Set SetCrossSheetLookup = SetColumnFormula(TargetColKey, tmpl, UseAutoFill:=True)
End Function

' Injects a cumulative running sum column (e.g. D3:D{row}).
Public Function SetRunningTotalFormula(ByVal TargetColKey As Variant, _
                                      ByVal SourceColKey As Variant) As cFormulaEngine
    CheckAttached "SetRunningTotalFormula"

    Dim sColIdx As Long
    sColIdx = ResolveColIndex(SourceColKey)
    Dim sColLetter As String
    sColLetter = ColIndexToLetter(mDataBodyRange.Column + sColIdx - 1)

    Dim firstRow As Long
    firstRow = mDataBodyRange.Row

    ' Formula: =SUM(D$3:D3)
    Dim tmpl As String
    tmpl = "=SUM(" & sColLetter & "$" & CStr(firstRow) & ":" & sColLetter & "{row})"

    Set SetRunningTotalFormula = SetColumnFormula(TargetColKey, tmpl, UseAutoFill:=True)
End Function

' Freezes formulas into static values across the entire table or a specific column.
' Replaces Range.Formula with Range.Value to prevent re-calculation.
Public Function FreezeToValues(Optional ByVal ColumnKey As Variant = Empty) As cFormulaEngine
    CheckAttached "FreezeToValues"
    If mRowCount = 0 Then
        Set FreezeToValues = Me
        Exit Function
    End If

    If IsEmpty(ColumnKey) Then
        mDataBodyRange.Value = mDataBodyRange.Value
    Else
        Dim cIdx As Long
        cIdx = ResolveColIndex(ColumnKey)
        Dim targetColRng As Object
        Set targetColRng = mDataBodyRange.Columns(cIdx)
        targetColRng.Value = targetColRng.Value
    End If

    Set FreezeToValues = Me
End Function

' Wraps existing formula in ColumnKey with =IFERROR(..., FallbackValue).
Public Function SetIfErrorWrapper(ByVal ColumnKey As Variant, _
                                  Optional ByVal FallbackValue As Variant = 0) As cFormulaEngine
    CheckAttached "SetIfErrorWrapper"
    If mRowCount = 0 Then
        Set SetIfErrorWrapper = Me
        Exit Function
    End If

    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)
    Dim colDataRng As Object
    Set colDataRng = mDataBodyRange.Columns(cIdx)

    Dim topCell As Object
    Set topCell = colDataRng.Cells(1, 1)

    Dim curFmla As String
    curFmla = Trim$(CStr(topCell.Formula))
    If Len(curFmla) = 0 Then
        Err.Raise 5, "cFormulaEngine.SetIfErrorWrapper", "Column '" & CStr(ColumnKey) & "' has no formula to wrap."
    End If

    If Left$(curFmla, 1) = "=" Then curFmla = Mid$(curFmla, 2)
    If Left$(UCase$(curFmla), 8) = "IFERROR(" Then
        ' Already wrapped
        Set SetIfErrorWrapper = Me
        Exit Function
    End If

    Dim strFallback As String
    If VarType(FallbackValue) = vbString Then
        strFallback = """" & Replace(CStr(FallbackValue), """", """""") & """"
    ElseIf IsEmpty(FallbackValue) Then
        strFallback = """"""
    Else
        strFallback = CStr(FallbackValue)
    End If

    Dim wrappedFmla As String
    wrappedFmla = "=IFERROR(" & curFmla & ", " & strFallback & ")"
    topCell.Formula = wrappedFmla

    If mRowCount > 1 Then
        topCell.AutoFill Destination:=colDataRng, Type:=XL_FILL_DEFAULT
    End If

    Set SetIfErrorWrapper = Me
End Function

' Injects an INDEX-MATCH lookup formula across an entire column.
Public Function SetIndexMatchLookup(ByVal TargetColKey As Variant, _
                                    ByVal MatchValColKey As Variant, _
                                    ByVal LookupSheetName As String, _
                                    ByVal LookupRangeAddress As String, _
                                    ByVal ReturnRangeAddress As String, _
                                    Optional ByVal MatchType As Long = 0) As cFormulaEngine
    CheckAttached "SetIndexMatchLookup"

    Dim kColIdx As Long
    kColIdx = ResolveColIndex(MatchValColKey)
    Dim kColLetter As String
    kColLetter = ColIndexToLetter(mDataBodyRange.Column + kColIdx - 1)

    Dim sheetRef As String
    If Len(LookupSheetName) > 0 Then
        If InStr(LookupSheetName, " ") > 0 Or InStr(LookupSheetName, "-") > 0 Then
            sheetRef = "'" & Replace(LookupSheetName, "'", "''") & "'!"
        Else
            sheetRef = LookupSheetName & "!"
        End If
    Else
        sheetRef = ""
    End If

    Dim tmpl As String
    tmpl = "=INDEX(" & sheetRef & ReturnRangeAddress & ", MATCH(" & kColLetter & "{row}, " & _
           sheetRef & LookupRangeAddress & ", " & CStr(MatchType) & "))"

    Set SetIndexMatchLookup = SetColumnFormula(TargetColKey, tmpl, UseAutoFill:=True)
End Function

' Injects an XLOOKUP formula across an entire column.
Public Function SetXLookup(ByVal TargetColKey As Variant, _
                           ByVal LookupValColKey As Variant, _
                           ByVal LookupSheetName As String, _
                           ByVal LookupRangeAddress As String, _
                           ByVal ReturnRangeAddress As String, _
                           Optional ByVal IfNotFound As String = "", _
                           Optional ByVal MatchMode As Long = 0) As cFormulaEngine
    CheckAttached "SetXLookup"

    Dim kColIdx As Long
    kColIdx = ResolveColIndex(LookupValColKey)
    Dim kColLetter As String
    kColLetter = ColIndexToLetter(mDataBodyRange.Column + kColIdx - 1)

    Dim sheetRef As String
    If Len(LookupSheetName) > 0 Then
        If InStr(LookupSheetName, " ") > 0 Or InStr(LookupSheetName, "-") > 0 Then
            sheetRef = "'" & Replace(LookupSheetName, "'", "''") & "'!"
        Else
            sheetRef = LookupSheetName & "!"
        End If
    Else
        sheetRef = ""
    End If

    Dim tmpl As String
    tmpl = "=XLOOKUP(" & kColLetter & "{row}, " & sheetRef & LookupRangeAddress & ", " & _
           sheetRef & ReturnRangeAddress

    If Len(IfNotFound) > 0 Or MatchMode <> 0 Then
        tmpl = tmpl & ", """ & Replace(IfNotFound, """", """""") & """"
    End If
    If MatchMode <> 0 Then
        tmpl = tmpl & ", " & CStr(MatchMode)
    End If
    tmpl = tmpl & ")"

    Set SetXLookup = SetColumnFormula(TargetColKey, tmpl, UseAutoFill:=True)
End Function

' Injects a Percentage of Total formula (=SourceCell / SUM(SourceCol$first:SourceCol$last))
Public Function SetPercentageOfTotal(ByVal TargetColKey As Variant, _
                                     ByVal SourceColKey As Variant, _
                                     Optional ByVal NumberFormatStr As String = "0.0%") As cFormulaEngine
    CheckAttached "SetPercentageOfTotal"

    Dim sColIdx As Long
    sColIdx = ResolveColIndex(SourceColKey)
    Dim sColLetter As String
    sColLetter = ColIndexToLetter(mDataBodyRange.Column + sColIdx - 1)

    Dim rFirst As Long, rLast As Long
    rFirst = mDataBodyRange.Row
    rLast = mDataBodyRange.Row + mRowCount - 1

    Dim tmpl As String
    tmpl = "=IFERROR(" & sColLetter & "{row} / SUM(" & sColLetter & "$" & CStr(rFirst) & ":" & sColLetter & "$" & CStr(rLast) & "), 0)"

    Set SetPercentageOfTotal = SetColumnFormula(TargetColKey, tmpl, UseAutoFill:=True)

    If Len(NumberFormatStr) > 0 Then
        Dim tColIdx As Long
        tColIdx = ResolveColIndex(TargetColKey)
        mDataBodyRange.Columns(tColIdx).NumberFormat = NumberFormatStr
    End If
End Function

' Injects an Actual vs Budget Variance formula across a column.
Public Function SetVarianceFormula(ByVal TargetColKey As Variant, _
                                   ByVal ActualColKey As Variant, _
                                   ByVal BudgetColKey As Variant, _
                                   Optional ByVal AsPercentage As Boolean = False, _
                                   Optional ByVal NumberFormatStr As String = "") As cFormulaEngine
    CheckAttached "SetVarianceFormula"

    Dim aColIdx As Long, bColIdx As Long
    aColIdx = ResolveColIndex(ActualColKey)
    bColIdx = ResolveColIndex(BudgetColKey)

    Dim aColLetter As String, bColLetter As String
    aColLetter = ColIndexToLetter(mDataBodyRange.Column + aColIdx - 1)
    bColLetter = ColIndexToLetter(mDataBodyRange.Column + bColIdx - 1)

    Dim tmpl As String
    If AsPercentage Then
        tmpl = "=IFERROR((" & aColLetter & "{row} - " & bColLetter & "{row}) / " & bColLetter & "{row}, 0)"
    Else
        tmpl = "=" & aColLetter & "{row} - " & bColLetter & "{row}"
    End If

    Set SetVarianceFormula = SetColumnFormula(TargetColKey, tmpl, UseAutoFill:=True)

    Dim fmt As String
    fmt = NumberFormatStr
    If Len(fmt) = 0 And AsPercentage Then fmt = "0.0%"

    If Len(fmt) > 0 Then
        Dim tColIdx As Long
        tColIdx = ResolveColIndex(TargetColKey)
        mDataBodyRange.Columns(tColIdx).NumberFormat = fmt
    End If
End Function

' Scans the data range for cells containing Excel formula errors (#DIV/0!, #REF!, #N/A, #VALUE!, etc.)
' Returns a cDataFrame with: Row, AbsRow, Col, ColName, CellAddress, Formula, ErrorType
Public Function AuditFormulas() As cDataFrame
    Dim auditDf As New cDataFrame
    auditDf.AddColumn "Row"
    auditDf.AddColumn "AbsRow"
    auditDf.AddColumn "Col"
    auditDf.AddColumn "ColName"
    auditDf.AddColumn "CellAddress"
    auditDf.AddColumn "Formula"
    auditDf.AddColumn "ErrorType"

    If mDataBodyRange Is Nothing Or mRowCount = 0 Then
        Set AuditFormulas = auditDf
        Exit Function
    End If

    ' Pre-fetch column names
    Dim colNames() As String
    ReDim colNames(1 To mColumnCount)
    Dim c As Long, k As Variant
    For c = 1 To mColumnCount
        colNames(c) = "C" & CStr(c)
    Next c
    For Each k In mColIndexMap.Keys
        Dim cPos As Long
        cPos = mColIndexMap.Item(k)
        If cPos >= 1 And cPos <= mColumnCount Then
            If Left$(CStr(k), 1) <> "C" Or Not IsNumeric(Mid$(CStr(k), 2)) Then
                colNames(cPos) = CStr(k)
            End If
        End If
    Next k

    Dim r As Long, cellObj As Object, cellVal As Variant
    For r = 1 To mRowCount
        For c = 1 To mColumnCount
            Set cellObj = mDataBodyRange.Cells(r, c)
            If cellObj.HasFormula Then
                cellVal = cellObj.Value
                If IsError(cellVal) Then
                    Dim errStr As String
                    errStr = ""
                    On Error Resume Next
                    errStr = CStr(cellObj.Text)
                    On Error GoTo 0
                    If Len(errStr) = 0 Then
                        Dim errNum As Long
                        errNum = CLng(cellVal)
                        Select Case errNum
                            Case 2007: errStr = "#DIV/0!"
                            Case 2023: errStr = "#REF!"
                            Case 2042: errStr = "#N/A"
                            Case 2015: errStr = "#VALUE!"
                            Case 2029: errStr = "#NAME?"
                            Case 2036: errStr = "#NUM!"
                            Case 2000: errStr = "#NULL!"
                            Case Else: errStr = "#ERROR"
                        End Select
                    End If
                    auditDf.AddRow r, cellObj.Row, c, colNames(c), cellObj.Address(False, False), CStr(cellObj.Formula), errStr
                End If
            End If
        Next c
    Next r

    Set AuditFormulas = auditDf
End Function

' Audits the worksheet for circular references using Excel's built-in circular calculation detection.
' Returns a cDataFrame with details of any circular reference loops found.
Public Function AuditCircularReferences() As cDataFrame
    On Error GoTo CatchErr
    Dim circDf As New cDataFrame
    circDf.AddColumn "Address"
    circDf.AddColumn "Row"
    circDf.AddColumn "Column"
    circDf.AddColumn "Formula"
    circDf.AddColumn "Severity"
    circDf.AddColumn "Suggestion"

    If mDataBodyRange Is Nothing Then
        Set AuditCircularReferences = circDf
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = mDataBodyRange.Worksheet
    Dim circCell As Range
    Set circCell = ws.CircularReference

    If Not circCell Is Nothing Then
        Dim fText As String
        fText = circCell.Formula

        circDf.AddRow circCell.Address(False, False), _
                      circCell.Row, _
                      circCell.Column, _
                      fText, _
                      "CRITICAL", _
                      "Circular reference detected. Break circular loop by eliminating self/mutual reference."
    End If

    Set AuditCircularReferences = circDf
    Exit Function

CatchErr:
    Set AuditCircularReferences = circDf
End Function

' Audits the attached range or worksheet for volatile functions (OFFSET, INDIRECT, TODAY, NOW, RAND, etc.).
' Volatile functions recalculate on every worksheet change, severely degrading performance in large models.
Public Function AuditVolatileFormulas(Optional ScopeRange As Variant) As cDataFrame
    On Error GoTo CatchErr
    Dim volDf As New cDataFrame
    volDf.AddColumn "Address"
    volDf.AddColumn "Column"
    volDf.AddColumn "Formula"
    volDf.AddColumn "VolatileFunction"
    volDf.AddColumn "Severity"
    volDf.AddColumn "Suggestion"

    If mDataBodyRange Is Nothing Then
        Set AuditVolatileFormulas = volDf
        Exit Function
    End If

    Dim targetRng As Range
    If IsMissing(ScopeRange) Then
        Set targetRng = mDataBodyRange
    ElseIf IsEmpty(ScopeRange) Then
        Set targetRng = mDataBodyRange
    ElseIf IsObject(ScopeRange) Then
        If TypeOf ScopeRange Is Range Then
            Set targetRng = ScopeRange
        End If
    Else
        Dim cIdx As Long
        cIdx = ResolveColIndex(ScopeRange)
        If cIdx > 0 And cIdx <= mColumnCount Then
            Set targetRng = mDataBodyRange.Columns(cIdx)
        Else
            Set targetRng = mDataBodyRange
        End If
    End If

    If targetRng Is Nothing Then
        Set AuditVolatileFormulas = volDf
        Exit Function
    End If

    Dim cellObj As Range
    Dim fText As String, uText As String
    Dim volName As String, sugg As String, sev As String

    For Each cellObj In targetRng.Cells
        fText = ""
        If cellObj.HasFormula Then
            fText = cellObj.Formula
        End If

        If Len(fText) > 0 Then
            uText = UCase$(fText)
            volName = ""
            sugg = ""
            sev = "WARNING"

            If InStr(1, uText, "OFFSET(") > 0 Then
                volName = "OFFSET"
                sugg = "Replace OFFSET with non-volatile INDEX(:INDEX) range pattern."
            ElseIf InStr(1, uText, "INDIRECT(") > 0 Then
                volName = "INDIRECT"
                sugg = "Replace INDIRECT with structured table references or CHOOSE."
            ElseIf InStr(1, uText, "TODAY(") > 0 Then
                volName = "TODAY"
                sugg = "Store date in a single parameter cell or freeze to static values."
                sev = "INFO"
            ElseIf InStr(1, uText, "NOW(") > 0 Then
                volName = "NOW"
                sugg = "Store timestamp in a single parameter cell or freeze to static values."
                sev = "INFO"
            ElseIf InStr(1, uText, "RAND(") > 0 Or InStr(1, uText, "RANDBETWEEN(") > 0 Or InStr(1, uText, "RANDARRAY(") > 0 Then
                volName = "RAND"
                sugg = "Freeze random values using FreezeToValues to avoid constant recalculations."
            ElseIf InStr(1, uText, "INFO(") > 0 Then
                volName = "INFO"
                sugg = "Query environment info once via VBA instead of repeating volatile cell formulas."
                sev = "INFO"
            End If

            If Len(volName) > 0 Then
                Dim colLtr As String
                colLtr = ColIndexToLetter(cellObj.Column)
                volDf.AddRow cellObj.Address(False, False), colLtr, fText, volName, sev, sugg
            End If
        End If
    Next cellObj

    Set AuditVolatileFormulas = volDf
    Exit Function

CatchErr:
    Set AuditVolatileFormulas = volDf
End Function

' Audits external workbook links (LinkSources) in the parent workbook.
' Detects broken links, missing files, and orphaned remote references.
Public Function AuditExternalLinks() As cDataFrame
    On Error GoTo CatchErr
    Dim extDf As New cDataFrame
    extDf.AddColumn "LinkIndex"
    extDf.AddColumn "SourcePath"
    extDf.AddColumn "FileName"
    extDf.AddColumn "Status"
    extDf.AddColumn "Severity"
    extDf.AddColumn "Suggestion"

    If mDataBodyRange Is Nothing Then
        Set AuditExternalLinks = extDf
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = mDataBodyRange.Worksheet.Parent
    Dim links As Variant
    links = wb.LinkSources(1) ' 1 = xlExcelLinks

    If IsEmpty(links) Or IsNull(links) Then
        Set AuditExternalLinks = extDf
        Exit Function
    End If

    Dim i As Long
    Dim pth As String, fName As String, stat As String, sev As String, sugg As String
    For i = LBound(links) To UBound(links)
        pth = CStr(links(i))
        fName = Mid$(pth, InStrRev(pth, "\") + 1)
        stat = "OK"
        sev = "INFO"
        sugg = "External link active and accessible."

        If Left$(pth, 2) = "\\" Or Mid$(pth, 2, 2) = ":\" Then
            Dim fExist As String
            fExist = ""
            On Error Resume Next
            fExist = Dir$(pth)
            On Error GoTo CatchErr
            If Len(fExist) = 0 Then
                stat = "BROKEN"
                sev = "HIGH"
                sugg = "File not found at specified local/network path. Use RepairExternalLink or BreakExternalLinks."
            End If
        End If

        extDf.AddRow i, pth, fName, stat, sev, sugg
    Next i

    Set AuditExternalLinks = extDf
    Exit Function

CatchErr:
    Set AuditExternalLinks = extDf
End Function

Public Function RepairExternalLink(ByVal OldPath As String, ByVal NewPath As String) As Boolean
    If mDataBodyRange Is Nothing Then Exit Function
    On Error Resume Next
    Dim wb As Object
    Set wb = mDataBodyRange.Worksheet.Parent
    wb.ChangeLink OldPath, NewPath, 1 ' 1 = xlExcelLinks
    RepairExternalLink = (Err.Number = 0)
    On Error GoTo 0
End Function

' Converts external linked formulas to their current static values, severing remote workbook dependencies.
Public Function BreakExternalLinks(Optional ByVal SpecificPath As String = "") As Long
    If mDataBodyRange Is Nothing Then Exit Function
    Dim wb As Object
    Set wb = mDataBodyRange.Worksheet.Parent

    Dim links As Variant
    On Error Resume Next
    links = wb.LinkSources(1)
    On Error GoTo 0

    If IsEmpty(links) Or IsNull(links) Then Exit Function

    Dim i As Long, brokenCount As Long
    brokenCount = 0

    For i = LBound(links) To UBound(links)
        Dim pth As String
        pth = CStr(links(i))
        If Len(SpecificPath) = 0 Or StrComp(pth, SpecificPath, vbTextCompare) = 0 Then
            On Error Resume Next
            wb.BreakLink pth, 1 ' 1 = xlExcelLinks
            If Err.Number = 0 Then brokenCount = brokenCount + 1
            On Error GoTo 0
        End If
    Next i

    BreakExternalLinks = brokenCount
End Function

' Quotes a string for use in Excel formulas or VBA strings without double-quote madness
' e.g. Quote("Pending") -> """Pending"""
Public Function Quote(ByVal Text As String) As String
    Quote = Chr$(34) & Replace(Text, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

' Formats an operator and value into standard Excel criteria format for COUNTIF, SUMIF, etc.
' e.g. FormatCriteria(">", 50) -> ">50"
' e.g. FormatCriteria("<>", 0) -> "<>0"
' e.g. FormatCriteria("=", "Active") -> "=Active"
Public Function FormatCriteria(ByVal OperatorStr As String, ByVal Value As Variant) As String
    Dim op As String
    op = Trim$(OperatorStr)
    If Len(op) = 0 Then op = "="

    If IsDate(Value) Then
        FormatCriteria = op & Format$(Value, "yyyy-mm-dd")
    Else
        FormatCriteria = op & CStr(Value)
    End If
End Function

' Bridges regional formula separator differences (; vs ,)
' Replaces semicolons outside of quotes with commas for English Excel engine (.Formula)
Public Function NormalizeFormulaSeparators(ByVal FormulaStr As String) As String
    If InStr(1, FormulaStr, ";") = 0 Then
        NormalizeFormulaSeparators = FormulaStr
        Exit Function
    End If

    Dim i As Long, n As Long
    n = Len(FormulaStr)
    Dim ch As String
    Dim inDoubleQuotes As Boolean
    Dim inSingleQuotes As Boolean
    inDoubleQuotes = False
    inSingleQuotes = False
    Dim result As String
    result = ""

    For i = 1 To n
        ch = Mid$(FormulaStr, i, 1)
        If ch = Chr$(34) And Not inSingleQuotes Then
            inDoubleQuotes = Not inDoubleQuotes
            result = result & ch
        ElseIf ch = "'" And Not inDoubleQuotes Then
            inSingleQuotes = Not inSingleQuotes
            result = result & ch
        ElseIf ch = ";" And Not inDoubleQuotes And Not inSingleQuotes Then
            result = result & ","
        Else
            result = result & ch
        End If
    Next i

    NormalizeFormulaSeparators = result
End Function

' Traces column and cell precedents for a formula cell.
' Returns an array of referenced column names (e.g. Array("Qty", "Price")).
Public Function GetPrecedents(ByVal ColumnKey As Variant, Optional ByVal RowIndex As Long = 1) As Variant
    CheckAttached "GetPrecedents"
    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)
    Dim cellObj As Object
    Set cellObj = mDataBodyRange.Cells(RowIndex, cIdx)

    Dim fmlaStr As String
    fmlaStr = Trim$(CStr(cellObj.Formula))
    If Len(fmlaStr) = 0 Then
        GetPrecedents = Array()
        Exit Function
    End If

    Dim foundDict As Object
    Set foundDict = CreateObject("Scripting.Dictionary")
    foundDict.CompareMode = 1

    ' 1. Check named columns in mColIndexMap
    Dim k As Variant
    For Each k In mColIndexMap.Keys
        Dim cName As String
        cName = CStr(k)
        Dim colPos As Long
        colPos = mColIndexMap.Item(k)
        Dim cLetter As String
        cLetter = ColIndexToLetter(mDataBodyRange.Column + colPos - 1)

        If InStr(1, fmlaStr, "[" & cName & "]", vbTextCompare) > 0 Or _
           InStr(1, fmlaStr, "[@" & cName & "]", vbTextCompare) > 0 Or _
           InStr(1, fmlaStr, cLetter & CStr(cellObj.Row), vbTextCompare) > 0 Or _
           InStr(1, fmlaStr, cLetter & "$" & CStr(cellObj.Row), vbTextCompare) > 0 Then
            If Left$(cName, 1) <> "C" Or Not IsNumeric(Mid$(cName, 2)) Then
                If Not foundDict.Exists(cName) Then foundDict.Add cName, colPos
            End If
        End If
    Next k

    ' 2. Check native DirectPrecedents on same worksheet
    On Error Resume Next
    Dim precRange As Object
    Set precRange = cellObj.DirectPrecedents
    On Error GoTo 0
    If Not precRange Is Nothing Then
        Dim precArea As Object, pCell As Object
        For Each precArea In precRange.Areas
            For Each pCell In precArea.Cells
                Dim pCol As Long
                pCol = pCell.Column - mDataBodyRange.Column + 1
                If pCol >= 1 And pCol <= mColumnCount Then
                    For Each k In mColIndexMap.Keys
                        If mColIndexMap.Item(k) = pCol Then
                            If Left$(CStr(k), 1) <> "C" Or Not IsNumeric(Mid$(CStr(k), 2)) Then
                                If Not foundDict.Exists(CStr(k)) Then foundDict.Add CStr(k), pCol
                            End If
                        End If
                    Next k
                End If
            Next pCell
        Next precArea
    End If

    If foundDict.Count = 0 Then
        GetPrecedents = Array()
    Else
        GetPrecedents = foundDict.Keys
    End If
End Function

' Evaluates an Excel formula expression in memory (optionally resolving column tokens [ColName] to ranges).
' Returns scalar value, 1D array, or 2D array computed by Excel calculation engine without modifying grid.
Public Function Evaluate(ByVal Expression As String) As Variant
    CheckAttached "Evaluate"
    Dim parsedExpr As String
    parsedExpr = Trim$(Expression)
    If Left$(parsedExpr, 1) = "=" Then parsedExpr = Mid$(parsedExpr, 2)

    Dim rFirst As Long, rLast As Long
    rFirst = mDataBodyRange.Row
    rLast = mDataBodyRange.Row + mRowCount - 1

    Dim k As Variant
    For Each k In mColIndexMap.Keys
        Dim cName As String
        cName = CStr(k)
        Dim tokenToFind As String
        tokenToFind = "[" & cName & "]"
        If InStr(1, parsedExpr, tokenToFind, vbTextCompare) > 0 Then
            Dim cIdx As Long
            cIdx = mColIndexMap.Item(k)
            Dim colLetter As String
            colLetter = ColIndexToLetter(mDataBodyRange.Column + cIdx - 1)
            Dim colRangeRef As String
            colRangeRef = colLetter & CStr(rFirst) & ":" & colLetter & CStr(rLast)
            parsedExpr = Replace(parsedExpr, tokenToFind, colRangeRef, , , vbTextCompare)
        End If
    Next k

    parsedExpr = Replace(parsedExpr, "{first_row}", CStr(rFirst), , , vbTextCompare)
    parsedExpr = Replace(parsedExpr, "{last_row}", CStr(rLast), , , vbTextCompare)

    Dim evalResult As Variant
    If Not mWorksheet Is Nothing Then
        evalResult = mWorksheet.Evaluate(parsedExpr)
    Else
        evalResult = Application.Evaluate(parsedExpr)
    End If

    If IsObject(evalResult) Then
        Set Evaluate = evalResult
    Else
        Evaluate = evalResult
    End If
End Function

' Evaluates a formula in memory and packs the output into a cDataFrame.
Public Function EvaluateToDataFrame(ByVal Expression As String, _
                                    Optional ByVal ResultColName As String = "Result") As cDataFrame
    Dim df As New cDataFrame
    Dim res As Variant
    res = Me.Evaluate(Expression)

    If IsArray(res) Then
        Dim cDim As Long
        cDim = 0
        On Error Resume Next
        cDim = UBound(res, 2) - LBound(res, 2) + 1
        On Error GoTo 0

        If cDim > 1 Then
            Dim c As Long
            For c = 1 To cDim
                df.AddColumn ResultColName & "_" & CStr(c)
            Next c
            df.AddRows res
        Else
            df.AddColumn ResultColName
            Dim r As Long
            For r = LBound(res, 1) To UBound(res, 1)
                df.AddRow res(r, LBound(res, 2))
            Next r
        End If
    Else
        df.AddColumn ResultColName
        df.AddRow res
    End If

    Set EvaluateToDataFrame = df
End Function

' Begins a high-performance batch session: disables ScreenUpdating, Events, and sets Calculation=Manual.
Public Function BeginBatch() As cFormulaEngine
    If Not mInBatch Then
        On Error Resume Next
        mPrevCalculation = Application.Calculation
        mPrevScreenUpdating = Application.ScreenUpdating
        mPrevEnableEvents = Application.EnableEvents
        Application.Calculation = -4135 ' xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        mInBatch = True
        On Error GoTo 0
    End If
    Set BeginBatch = Me
End Function

' Ends batch session: recalculates table range and restores Calculation/ScreenUpdating/Events safely.
Public Function EndBatch(Optional ByVal CalculateNow As Boolean = True) As cFormulaEngine
    If mInBatch Then
        On Error Resume Next
        If CalculateNow Then
            If Not mTableRange Is Nothing Then
                mTableRange.Calculate
            Else
                Application.Calculate
            End If
        End If
        Application.Calculation = mPrevCalculation
        Application.ScreenUpdating = mPrevScreenUpdating
        Application.EnableEvents = mPrevEnableEvents
        mInBatch = False
        On Error GoTo 0
    End If
    Set EndBatch = Me
End Function

Public Property Get InBatch() As Boolean
    InBatch = mInBatch
End Property

Private Sub SafeAssignFormula(ByVal TargetRange As Object, _
                              ByVal FormulaStr As String, _
                              Optional ByVal PreferFormula2 As Boolean = False)
    If PreferFormula2 Then
        On Error Resume Next
        TargetRange.Formula2 = FormulaStr
        If Err.Number <> 0 Then
            Err.Clear
            TargetRange.Formula = FormulaStr
        End If
        On Error GoTo 0
    Else
        TargetRange.Formula = FormulaStr
    End If
End Sub

' ------------------------------------------------------------------------------
' Table-Driven Formula Engine
' ------------------------------------------------------------------------------

' Applies formula rules loaded from an Excel Table (ListObject or Table Name), Range, or cDataFrame.
' Supported columns (case-insensitive):
'   - Target / Column / ColumnKey: target column to receive formula
'   - Formula / Template / Value / Expression: the formula string
'   - Type / Mode / Kind: FORMULA, R1C1, SUMMARY, RUNNINGTOTAL
'   - SourceCol / Source: source column for summary/running totals
'   - Scope / Section: optional scope tag filter
'   - Enabled / Active: boolean flag
Public Function ApplyFormulasFromTable(ByVal FormulaTableOrDf As Variant, _
                                       Optional ByVal TargetScope As String = "") As cFormulaEngine
    CheckAttached "ApplyFormulasFromTable"
    If mRowCount = 0 Then
        Set ApplyFormulasFromTable = Me
        Exit Function
    End If

    Dim cfgDf As cDataFrame
    Set cfgDf = ResolveConfigDf(FormulaTableOrDf)
    If cfgDf.RowCount = 0 Then
        Set ApplyFormulasFromTable = Me
        Exit Function
    End If

    Dim targetCol As String, fmlaCol As String, typeCol As String
    Dim srcCol As String, enabledCol As String, scopeCol As String

    targetCol = FindColName(cfgDf, Array("Target", "Column", "ColumnKey", "Columns", "TargetCol"))
    If Len(targetCol) = 0 Then
        Err.Raise 5, "cFormulaEngine.ApplyFormulasFromTable", "Formula table requires a 'Target' or 'Column' column."
    End If

    fmlaCol = FindColName(cfgDf, Array("Formula", "Template", "Value", "Expression", "RuleValue"))
    typeCol = FindColName(cfgDf, Array("Type", "Mode", "Kind", "FormulaType"))
    srcCol = FindColName(cfgDf, Array("SourceCol", "Source", "Param1"))
    enabledCol = FindColName(cfgDf, Array("Enabled", "Active", "IsEnabled"))
    scopeCol = FindColName(cfgDf, Array("Scope", "Section", "TargetTable"))

    Dim r As Long
    For r = 1 To cfgDf.RowCount
        If Len(enabledCol) > 0 Then
            If Not ParseBool(cfgDf.Item(r, enabledCol), True) Then GoTo NextFormulaRow
        End If

        If Len(scopeCol) > 0 And Len(TargetScope) > 0 Then
            Dim rowScope As String
            rowScope = UCase$(Trim$(CStr(cfgDf.Item(r, scopeCol))))
            If Len(rowScope) > 0 And rowScope <> "ALL" And rowScope <> "*" Then
                If StrComp(rowScope, UCase$(Trim$(TargetScope)), vbTextCompare) <> 0 Then
                    GoTo NextFormulaRow
                End If
            End If
        End If

        Dim tgtKey As String, fmlaStr As String, modeStr As String, sourceParam As String
        tgtKey = Trim$(CStr(cfgDf.Item(r, targetCol)))
        If Len(tgtKey) = 0 Then GoTo NextFormulaRow

        If Len(fmlaCol) > 0 Then fmlaStr = Trim$(CStr(cfgDf.Item(r, fmlaCol))) Else fmlaStr = ""
        If Len(typeCol) > 0 Then modeStr = UCase$(Trim$(CStr(cfgDf.Item(r, typeCol)))) Else modeStr = "FORMULA"
        If Len(srcCol) > 0 Then sourceParam = Trim$(CStr(cfgDf.Item(r, srcCol))) Else sourceParam = ""

        Select Case modeStr
            Case "R1C1"
                Me.SetFormulaR1C1 tgtKey, fmlaStr
            Case "RUNNINGTOTAL", "CUMSUM"
                Me.SetRunningTotalFormula tgtKey, IIf(Len(sourceParam) > 0, sourceParam, fmlaStr)
            Case "SUMMARY", "TOTAL"
                Me.SetSummaryFormula mRowCount + 1, tgtKey, IIf(Len(fmlaStr) > 0, fmlaStr, "SUM"), sourceParam
            Case "PCT_TOTAL", "PERCENTAGE"
                Me.SetPercentageOfTotal tgtKey, IIf(Len(sourceParam) > 0, sourceParam, fmlaStr)
            Case "VARIANCE"
                Dim vParts() As String
                vParts = Split(fmlaStr, ",")
                If UBound(vParts) >= 1 Then
                    Me.SetVarianceFormula tgtKey, Trim$(vParts(0)), Trim$(vParts(1))
                Else
                    Me.SetVarianceFormula tgtKey, fmlaStr, sourceParam
                End If
            Case "IFERROR"
                Me.SetIfErrorWrapper tgtKey, IIf(Len(fmlaStr) > 0, fmlaStr, 0)
            Case "FREEZE"
                Me.FreezeToValues tgtKey
            Case "SPILL"
                Me.SetSpillFormula tgtKey, fmlaStr, PreferFormula2:=True
            Case "BULK"
                Me.SetColumnFormulaBulk tgtKey, fmlaStr
            Case Else
                ' Standard Formula (supports token replacement and AutoFill)
                If Len(fmlaStr) > 0 Then
                    Me.SetColumnFormula tgtKey, fmlaStr, UseAutoFill:=True
                End If
        End Select

NextFormulaRow:
    Next r

    Set ApplyFormulasFromTable = Me
End Function

' ------------------------------------------------------------------------------
' Template Parsing & Column Resolution Helpers
' ------------------------------------------------------------------------------

' Replaces tokens in a formula template:
'   [ColName] -> Relative Column Letter + Row (e.g. B3)
'   {row} -> Current absolute row number
'   {first_row} -> Data body top row number
'   {last_row} -> Data body bottom row number
'   {col} -> Target column letter
Public Function ParseFormulaTemplate(ByVal Template As String, _
                                     ByVal TargetColIdx As Long, _
                                     ByVal CurrentRow As Long) As String
    Dim result As String
    result = TranslateSingleQuoteLiterals(Template)

    Dim firstRow As Long, lastRow As Long
    firstRow = mDataBodyRange.Row
    lastRow = mDataBodyRange.Row + mRowCount - 1

    Dim targetAbsCol As Long
    targetAbsCol = mDataBodyRange.Column + TargetColIdx - 1
    Dim targetColLetter As String
    targetColLetter = ColIndexToLetter(targetAbsCol)

    result = Replace(result, "{row}", CStr(CurrentRow), , , vbTextCompare)
    result = Replace(result, "[row]", CStr(CurrentRow), , , vbTextCompare)
    result = Replace(result, "{first_row}", CStr(firstRow), , , vbTextCompare)
    result = Replace(result, "{last_row}", CStr(lastRow), , , vbTextCompare)
    result = Replace(result, "{col}", targetColLetter, , , vbTextCompare)

    ' Replace Named Column Tokens: [ColumnName] or [@ColumnName]
    Dim colKey As Variant
    For Each colKey In mColIndexMap.Keys
        Dim cName As String
        cName = CStr(colKey)
        Dim tokenToFind As String
        tokenToFind = "[" & cName & "]"
        Dim atTokenToFind As String
        atTokenToFind = "[@" & cName & "]"

        If mUseStructuredReferences And Not (mListObject Is Nothing) Then
            Dim structRef As String
            structRef = "[@" & cName & "]"
            If InStr(1, result, tokenToFind, vbTextCompare) > 0 Then
                result = Replace(result, tokenToFind, structRef, , , vbTextCompare)
            End If
        Else
            Dim cIdx As Long
            cIdx = mColIndexMap.Item(colKey)
            Dim colLetter As String
            colLetter = ColIndexToLetter(mDataBodyRange.Column + cIdx - 1)

            If InStr(1, result, tokenToFind, vbTextCompare) > 0 Then
                result = Replace(result, tokenToFind, colLetter & CStr(CurrentRow), , , vbTextCompare)
            End If
            If InStr(1, result, atTokenToFind, vbTextCompare) > 0 Then
                result = Replace(result, atTokenToFind, colLetter & CStr(CurrentRow), , , vbTextCompare)
            End If
        End If
    Next colKey

    ParseFormulaTemplate = result
End Function

Private Function TranslateSingleQuoteLiterals(ByVal Template As String) As String
    If InStr(1, Template, "'") = 0 Then
        TranslateSingleQuoteLiterals = Template
        Exit Function
    End If

    Dim n As Long
    n = Len(Template)
    Dim result As String
    result = ""
    Dim i As Long
    i = 1

    Do While i <= n
        Dim ch As String
        ch = Mid$(Template, i, 1)
        If ch = "'" Then
            Dim closePos As Long
            closePos = InStr(i + 1, Template, "'")
            If closePos > 0 Then
                Dim nextIdx As Long
                nextIdx = closePos + 1
                Do While nextIdx <= n
                    If Mid$(Template, nextIdx, 1) <> " " Then Exit Do
                    nextIdx = nextIdx + 1
                Loop

                Dim isSheetRef As Boolean
                isSheetRef = False
                If nextIdx <= n Then
                    If Mid$(Template, nextIdx, 1) = "!" Then
                        isSheetRef = True
                    End If
                End If

                If isSheetRef Then
                    result = result & Mid$(Template, i, closePos - i + 1)
                    i = closePos + 1
                Else
                    Dim literalText As String
                    literalText = Mid$(Template, i + 1, closePos - i - 1)
                    literalText = Replace(literalText, Chr$(34), Chr$(34) & Chr$(34))
                    result = result & Chr$(34) & literalText & Chr$(34)
                    i = closePos + 1
                End If
            Else
                result = result & ch
                i = i + 1
            End If
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    TranslateSingleQuoteLiterals = result
End Function

Public Function ColIndexToLetter(ByVal ColIndex As Long) As String
    Dim colNum As Long
    colNum = ColIndex
    Dim colLetter As String
    colLetter = ""
    Dim modNum As Long

    Do While colNum > 0
        modNum = (colNum - 1) Mod 26
        colLetter = Chr$(65 + modNum) & colLetter
        colNum = (colNum - modNum) \ 26
    Loop

    ColIndexToLetter = colLetter
End Function

Public Function ResolveColIndex(ByVal ColumnKey As Variant) As Long
    If VarType(ColumnKey) = vbString Then
        Dim strKey As String
        strKey = Trim$(CStr(ColumnKey))
        If mColIndexMap.Exists(strKey) Then
            ResolveColIndex = mColIndexMap.Item(strKey)
        ElseIf Left$(UCase$(strKey), 1) = "C" And IsNumeric(Mid$(strKey, 2)) Then
            ResolveColIndex = CLng(Mid$(strKey, 2))
        Else
            Err.Raise 5, "cFormulaEngine", "Column '" & strKey & "' not found in table."
        End If
    ElseIf IsNumeric(ColumnKey) Then
        Dim numIdx As Long
        numIdx = CLng(ColumnKey)
        If numIdx < 1 Or numIdx > mColumnCount Then
            Err.Raise 9, "cFormulaEngine", "Column index out of range: " & numIdx & " (Valid: 1 To " & mColumnCount & ")"
        End If
        ResolveColIndex = numIdx
    Else
        Err.Raise 13, "cFormulaEngine", "Invalid ColumnKey type: " & TypeName(ColumnKey)
    End If
End Function

' ------------------------------------------------------------------------------
' Property Getters
' ------------------------------------------------------------------------------

Public Property Get RowCount() As Long
    RowCount = mRowCount
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = mColumnCount
End Property

Public Property Get HasHeaders() As Boolean
    HasHeaders = mHasHeaders
End Property

Public Property Get TableRange() As Object
    Set TableRange = mTableRange
End Property

Public Property Get DataBodyRange() As Object
    Set DataBodyRange = mDataBodyRange
End Property

Public Property Get HeaderRange() As Object
    Set HeaderRange = mHeaderRange
End Property

Public Property Get Worksheet() As Object
    Set Worksheet = mWorksheet
End Property

Public Property Get ListObject() As Object
    Set ListObject = mListObject
End Property

Public Property Get IsListObject() As Boolean
    IsListObject = Not (mListObject Is Nothing)
End Property

Public Property Get UseStructuredReferences() As Boolean
    UseStructuredReferences = mUseStructuredReferences
End Property

Public Property Let UseStructuredReferences(ByVal Val As Boolean)
    mUseStructuredReferences = Val
End Property

' ------------------------------------------------------------------------------
' Private Resolvers
' ------------------------------------------------------------------------------

Private Sub CheckAttached(ByVal MethodName As String)
    If mTableRange Is Nothing Or mWorksheet Is Nothing Then
        Err.Raise 91, "cFormulaEngine." & MethodName, "cFormulaEngine is not attached. Call Attach() or AttachTo() first."
    End If
End Sub

Private Function ResolveConfigDf(ByVal Source As Variant) As cDataFrame
    Dim df As New cDataFrame

    If VarType(Source) = vbString Then
        Dim strName As String
        strName = Trim$(CStr(Source))
        If Len(strName) = 0 Then
            Err.Raise 5, "cFormulaEngine.ResolveConfigDf", "TableName string cannot be blank."
        End If
        If Not mWorksheet Is Nothing Then
            df.LoadFromTableName mWorksheet, strName
        Else
            df.LoadFromTableName Application.ActiveWorkbook, strName
        End If
        Set ResolveConfigDf = df
        Exit Function
    End If

    If Not IsObject(Source) Then
        Err.Raise 5, "cFormulaEngine.ResolveConfigDf", "FormulaTableOrDf must be Table Name string, cDataFrame, ListObject, or Range."
    End If
    If Source Is Nothing Then
        Err.Raise 91, "cFormulaEngine.ResolveConfigDf", "FormulaTableOrDf cannot be Nothing."
    End If

    If TypeOf Source Is cDataFrame Then
        Set ResolveConfigDf = Source
        Exit Function
    End If

    On Error Resume Next
    Dim loTest As Object
    Set loTest = Source.ListColumns
    On Error GoTo 0
    If Not loTest Is Nothing Then
        df.LoadFromListObject Source
    Else
        df.LoadFromRange Source, True
    End If
    Set ResolveConfigDf = df
End Function

Private Function FindColName(ByVal Df As cDataFrame, ByVal Candidates As Variant) As String
    Dim names() As String
    names = Df.ColumnNames()
    Dim i As Long, j As Long
    For j = LBound(Candidates) To UBound(Candidates)
        For i = LBound(names) To UBound(names)
            If StrComp(Trim$(names(i)), CStr(Candidates(j)), vbTextCompare) = 0 Then
                FindColName = names(i)
                Exit Function
            End If
        Next i
    Next j
End Function

Private Function ParseBool(ByVal Val As Variant, Optional ByVal DefaultVal As Boolean = True) As Boolean
    If IsEmpty(Val) Or IsNull(Val) Then
        ParseBool = DefaultVal
        Exit Function
    End If
    If VarType(Val) = vbBoolean Then
        ParseBool = CBool(Val)
        Exit Function
    End If
    Dim s As String
    s = UCase$(Trim$(CStr(Val)))
    If Len(s) = 0 Then
        ParseBool = DefaultVal
    ElseIf s = "TRUE" Or s = "YES" Or s = "1" Or s = "Y" Or s = "ON" Then
        ParseBool = True
    ElseIf s = "FALSE" Or s = "NO" Or s = "0" Or s = "N" Or s = "OFF" Then
        ParseBool = False
    Else
        ParseBool = DefaultVal
    End If
End Function