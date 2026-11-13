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
Private mRowCount As Long
Private mColumnCount As Long
Private mHasHeaders As Boolean
Private mColIndexMap As Object ' Scripting.Dictionary: Name/Key -> 1-based column index

Private Sub Class_Initialize()
    mRowCount = 0
    mColumnCount = 0
    mHasHeaders = True
    Set mColIndexMap = CreateObject("Scripting.Dictionary")
    mColIndexMap.CompareMode = 1 ' TextCompare (case-insensitive)
End Sub

Private Sub Class_Terminate()
    Set mWorksheet = Nothing
    Set mTableRange = Nothing
    Set mHeaderRange = Nothing
    Set mDataBodyRange = Nothing
    Set mTopLeftCell = Nothing
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
' If UseAutoFill is False, performs instant bulk Range.Formula assignment across the column.
Public Function SetColumnFormula(ByVal ColumnKey As Variant, _
                                 ByVal FormulaTemplate As String, _
                                 Optional ByVal UseAutoFill As Boolean = True, _
                                 Optional ByVal FillType As Long = 0, _
                                 Optional ByVal IsR1C1 As Boolean = False) As cFormulaEngine
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
        firstCell.Formula = parsedFirstFormula

        If mRowCount > 1 Then
            If UseAutoFill Then
                On Error Resume Next
                firstCell.AutoFill Destination:=colDataRng, Type:=FillType
                If Err.Number <> 0 Then
                    Err.Clear
                    ' Fallback to bulk formula assignment if AutoFill is restricted
                    colDataRng.Formula = parsedFirstFormula
                End If
                On Error GoTo 0
            Else
                colDataRng.Formula = parsedFirstFormula
            End If
        End If
    End If

    Set SetColumnFormula = Me
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
    result = Template

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

    ' Replace Named Column Tokens: [ColumnName] with relative cell reference
    Dim colKey As Variant
    For Each colKey In mColIndexMap.Keys
        Dim cName As String
        cName = CStr(colKey)
        Dim tokenToFind As String
        tokenToFind = "[" & cName & "]"
        If InStr(1, result, tokenToFind, vbTextCompare) > 0 Then
            Dim cIdx As Long
            cIdx = mColIndexMap.Item(colKey)
            Dim colLetter As String
            colLetter = ColIndexToLetter(mDataBodyRange.Column + cIdx - 1)
            result = Replace(result, tokenToFind, colLetter & CStr(CurrentRow), , , vbTextCompare)
        End If
    Next colKey

    ParseFormulaTemplate = result
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
