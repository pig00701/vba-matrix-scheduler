VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTableStyler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' Class: cTableStyler
' Description: High-Performance Fluent Excel Cell & Range Styling Engine for cDataFrame.
'              Decouples presentation/formatting concerns from the core in-memory
'              data structure. Provides single-COM-call header theming, border
'              application, vectorized column number formatting, multi-column
'              formatting, batch-union zebra striping, total row highlighting,
'              flexible row/column sizing (uniform group widths, wildcard pattern
'              matching, section row height control), conditional status highlighting,
'              freeze panes, typography/font control, and screen updating management.
' Author: Korn (AI Backend Developer, Technodevmoo)
' Architecture: Kit (Tech Lead)
' ==============================================================================

Private mTopLeftCell As Object         ' Excel.Range (top-left anchor cell)
Private mWorksheet As Object           ' Excel.Worksheet parent
Private mRowCount As Long              ' Number of data rows (excluding header)
Private mColumnCount As Long           ' Number of columns
Private mHasHeaders As Boolean         ' True if header row is present
Private mColIndexMap As Object         ' Scripting.Dictionary (ColName -> 1-based ColIndex)
Private mTableRange As Object          ' Excel.Range spanning entire styled table
Private mHeaderRange As Object         ' Excel.Range spanning header row (or Nothing)
Private mDataBodyRange As Object       ' Excel.Range spanning data body rows (or Nothing)
Private mTotalRowHighlighted As Boolean' Track if total row has been highlighted
Private mInBatchOperation As Boolean   ' Track screen updating & environment suppression
Private mPrevScreenUpdating As Boolean ' Stored ScreenUpdating state
Private mPrevCalculation As Long       ' Stored Calculation state
Private mPrevEnableEvents As Boolean   ' Stored EnableEvents state
Private mPrevDisplayAlerts As Boolean  ' Stored DisplayAlerts state

' ------------------------------------------------------------------------------
' Lifecycle
' ------------------------------------------------------------------------------

Private Sub Class_Initialize()
    Set mColIndexMap = CreateObject("Scripting.Dictionary")
    mColIndexMap.CompareMode = 1      ' vbTextCompare (default)
    mRowCount = 0
    mColumnCount = 0
    mHasHeaders = True
    mTotalRowHighlighted = False
    mInBatchOperation = False
    mPrevScreenUpdating = True
    mPrevCalculation = -4105 ' xlCalculationAutomatic
    mPrevEnableEvents = True
    mPrevDisplayAlerts = True
    Set mTopLeftCell = Nothing
    Set mWorksheet = Nothing
    Set mTableRange = Nothing
    Set mHeaderRange = Nothing
    Set mDataBodyRange = Nothing
End Sub

Private Sub Class_Terminate()
    If mInBatchOperation Then
        EndUpdate
    End If
    Set mColIndexMap = Nothing
    Set mTopLeftCell = Nothing
    Set mWorksheet = Nothing
    Set mTableRange = Nothing
    Set mHeaderRange = Nothing
    Set mDataBodyRange = Nothing
End Sub

' ------------------------------------------------------------------------------
' Public Properties (Range References & Dimensions)
' ------------------------------------------------------------------------------

Public Property Get TableRange() As Object
    Set TableRange = mTableRange
End Property

Public Property Get HeaderRange() As Object
    Set HeaderRange = mHeaderRange
End Property

Public Property Get DataBodyRange() As Object
    Set DataBodyRange = mDataBodyRange
End Property

Public Property Get TopLeftCell() As Object
    Set TopLeftCell = mTopLeftCell
End Property

Public Property Get Worksheet() As Object
    Set Worksheet = mWorksheet
End Property

Public Property Get RowCount() As Long
    RowCount = mRowCount
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = mColumnCount
End Property

Public Property Get HasHeaders() As Boolean
    HasHeaders = mHasHeaders
End Property

' ------------------------------------------------------------------------------
' Fluent Attachment
' ------------------------------------------------------------------------------

Public Function Attach(ByVal TargetRangeOrTopLeft As Object, _
                       ByVal SourceDataFrame As cDataFrame, _
                       Optional ByVal HasHeaders As Boolean = True) As cTableStyler
    If TargetRangeOrTopLeft Is Nothing Then
        Err.Raise 91, "cTableStyler.Attach", "TargetRangeOrTopLeft cannot be Nothing."
    End If
    If SourceDataFrame Is Nothing Then
        Err.Raise 91, "cTableStyler.Attach", "SourceDataFrame cannot be Nothing."
    End If

    Set mTopLeftCell = TargetRangeOrTopLeft.Cells(1, 1)
    Set mWorksheet = mTopLeftCell.Worksheet
    If mWorksheet.ProtectContents Then
        Err.Raise 1004, "cTableStyler.Attach", _
            "Cannot attach styler: Worksheet '" & mWorksheet.Name & "' is protected."
    End If
    mRowCount = SourceDataFrame.RowCount
    mColumnCount = SourceDataFrame.ColumnCount
    mHasHeaders = HasHeaders
    mTotalRowHighlighted = False

    mColIndexMap.RemoveAll
    ' Synchronize CompareMode with DataFrame CaseSensitive setting
    mColIndexMap.CompareMode = IIf(SourceDataFrame.CaseSensitive, 0, 1)

    Dim colNames() As String
    colNames = SourceDataFrame.ColumnNames()
    Dim i As Long
    If mColumnCount > 0 Then
        For i = LBound(colNames) To UBound(colNames)
            mColIndexMap.Item(colNames(i)) = (i - LBound(colNames) + 1)
        Next i
    End If

    Dim totalRows As Long
    totalRows = mRowCount + IIf(mHasHeaders, 1, 0)

    If mColumnCount > 0 And totalRows > 0 Then
        Set mTableRange = mTopLeftCell.Resize(totalRows, mColumnCount)
        If mHasHeaders Then
            Set mHeaderRange = mTopLeftCell.Resize(1, mColumnCount)
            If mRowCount > 0 Then
                Set mDataBodyRange = mTopLeftCell.Offset(1, 0).Resize(mRowCount, mColumnCount)
            Else
                Set mDataBodyRange = Nothing
            End If
        Else
            Set mHeaderRange = Nothing
            If mRowCount > 0 Then
                Set mDataBodyRange = mTopLeftCell.Resize(mRowCount, mColumnCount)
            Else
                Set mDataBodyRange = Nothing
            End If
        End If
    Else
        Set mTableRange = Nothing
        Set mHeaderRange = Nothing
        Set mDataBodyRange = Nothing
    End If

    Set Attach = Me
End Function

' Binds styler to an existing Excel Range (worksheet range or table) without requiring a cDataFrame instance.
' Automatically inspects header names, row/column counts, and populates column name map.
Public Function AttachRange(ByVal SourceRange As Object, _
                            Optional ByVal HasHeaders As Boolean = True) As cTableStyler
    If SourceRange Is Nothing Then
        Err.Raise 91, "cTableStyler.AttachRange", "SourceRange cannot be Nothing."
    End If

    Dim target As Object
    Set target = SourceRange

    ' If a single cell is passed, expand to CurrentRegion if it has content
    If target.Rows.Count = 1 And target.Columns.Count = 1 Then
        If Not IsEmpty(target.Value2) Then
            Set target = target.CurrentRegion
        End If
    End If

    Set mTopLeftCell = target.Cells(1, 1)
    Set mWorksheet = mTopLeftCell.Worksheet
    If mWorksheet.ProtectContents Then
        Err.Raise 1004, "cTableStyler.AttachRange", _
            "Cannot attach styler: Worksheet '" & mWorksheet.Name & "' is protected."
    End If
    mHasHeaders = HasHeaders
    mTotalRowHighlighted = False

    Dim totalRows As Long, totalCols As Long
    totalRows = target.Rows.Count
    totalCols = target.Columns.Count
    mColumnCount = totalCols

    If mHasHeaders Then
        mRowCount = totalRows - 1
        If mRowCount < 0 Then mRowCount = 0
    Else
        mRowCount = totalRows
    End If

    mColIndexMap.RemoveAll
    mColIndexMap.CompareMode = 1 ' vbTextCompare (case-insensitive)

    If mColumnCount > 0 And totalRows > 0 Then
        Set mTableRange = target
        Dim c As Long
        If mHasHeaders Then
            Set mHeaderRange = target.Rows(1)
            If mRowCount > 0 Then
                Set mDataBodyRange = target.Offset(1, 0).Resize(mRowCount, mColumnCount)
            Else
                Set mDataBodyRange = Nothing
            End If

            ' Read header row values into mColIndexMap safely
            Dim headerVals As Variant
            headerVals = mHeaderRange.Value2
            Dim hName As String, rawH As Variant
            For c = 1 To mColumnCount
                If mColumnCount = 1 Then
                    rawH = headerVals
                Else
                    rawH = headerVals(1, c)
                End If
                If IsError(rawH) Then
                    hName = "Col" & CStr(c)
                ElseIf IsNull(rawH) Or IsEmpty(rawH) Then
                    hName = "Col" & CStr(c)
                Else
                    hName = Trim$(CStr(rawH))
                    If Len(hName) = 0 Then hName = "Col" & CStr(c)
                End If
                If Not mColIndexMap.Exists(hName) Then
                    mColIndexMap.Item(hName) = c
                End If
            Next c
        Else
            Set mHeaderRange = Nothing
            If mRowCount > 0 Then
                Set mDataBodyRange = target
            Else
                Set mDataBodyRange = Nothing
            End If
            For c = 1 To mColumnCount
                mColIndexMap.Item("Col" & CStr(c)) = c
            Next c
        End If
    Else
        Set mTableRange = Nothing
        Set mHeaderRange = Nothing
        Set mDataBodyRange = Nothing
    End If

    Set AttachRange = Me
End Function

' Binds styler directly to an existing Excel ListObject (Table).
Public Function AttachListObject(ByVal TableObj As Object) As cTableStyler
    If TableObj Is Nothing Then
        Err.Raise 91, "cTableStyler.AttachListObject", "TableObj cannot be Nothing."
    End If
    Dim hasHdrs As Boolean
    hasHdrs = True
    On Error Resume Next
    hasHdrs = TableObj.ShowHeaders
    On Error GoTo 0
    AttachRange TableObj.Range, HasHeaders:=hasHdrs
    Set AttachListObject = Me
End Function

' ------------------------------------------------------------------------------
' Screen Updating Optimization Helpers
' ------------------------------------------------------------------------------

' Suppresses Excel screen updates, events, and automatic calculation for batch styling operations.
' Crash-safe RAII design: Class_Terminate automatically invokes EndUpdate to guarantee state restoration.
Public Function BeginUpdate() As cTableStyler
    If Not mInBatchOperation And Not mWorksheet Is Nothing Then
        On Error Resume Next
        Dim appObj As Object
        Set appObj = mWorksheet.Application
        mPrevScreenUpdating = appObj.ScreenUpdating
        mPrevCalculation = appObj.Calculation
        mPrevEnableEvents = appObj.EnableEvents
        mPrevDisplayAlerts = appObj.DisplayAlerts

        appObj.ScreenUpdating = False
        appObj.Calculation = -4135 ' xlCalculationManual
        appObj.EnableEvents = False
        mInBatchOperation = True
        On Error GoTo 0
    End If
    Set BeginUpdate = Me
End Function

' Restores previous Excel screen updating, calculation mode, and events state
Public Function EndUpdate() As cTableStyler
    If mInBatchOperation And Not mWorksheet Is Nothing Then
        On Error Resume Next
        Dim appObj As Object
        Set appObj = mWorksheet.Application
        appObj.Calculation = mPrevCalculation
        appObj.EnableEvents = mPrevEnableEvents
        appObj.DisplayAlerts = mPrevDisplayAlerts
        appObj.ScreenUpdating = mPrevScreenUpdating
        mInBatchOperation = False
        On Error GoTo 0
    End If
    Set EndUpdate = Me
End Function

' ------------------------------------------------------------------------------
' Fluent Styling Methods
' ------------------------------------------------------------------------------

' Applies pre-configured theme colors or custom RGB values to the header row in 1 COM call.
Public Function ApplyTheme(Optional ByVal ThemeName As String = "NAVY", _
                           Optional ByVal CustomHeaderBgColor As Long = -1, _
                           Optional ByVal CustomFontColor As Long = -1, _
                           Optional ByVal Alignment As Long = -4108) As cTableStyler
    If mHeaderRange Is Nothing Then
        Set ApplyTheme = Me
        Exit Function
    End If

    Dim headerBgColor As Long
    Dim fontColor As Long
    Dim isBold As Boolean

    isBold = True

    If CustomHeaderBgColor >= 0 Then
        headerBgColor = CustomHeaderBgColor
        If CustomFontColor >= 0 Then
            fontColor = CustomFontColor
        Else
            fontColor = RGB(255, 255, 255)
        End If
    Else
        Select Case UCase$(Trim$(ThemeName))
            Case "NAVY", "CORPORATE_NAVY"
                headerBgColor = RGB(31, 78, 121)
                fontColor = RGB(255, 255, 255)
            Case "SLATE", "MODERN_CHARCOAL"
                headerBgColor = RGB(52, 73, 94)
                fontColor = RGB(255, 255, 255)
            Case "EMERALD", "GREEN"
                headerBgColor = RGB(39, 110, 80)
                fontColor = RGB(255, 255, 255)
            Case "MINIMAL", "CLASSIC_GRAY"
                headerBgColor = RGB(220, 224, 230)
                fontColor = RGB(40, 40, 40)
            Case "TEAL"
                headerBgColor = RGB(0, 128, 128)
                fontColor = RGB(255, 255, 255)
            Case "BURGUNDY"
                headerBgColor = RGB(128, 0, 32)
                fontColor = RGB(255, 255, 255)
            Case "INDIGO"
                headerBgColor = RGB(63, 81, 181)
                fontColor = RGB(255, 255, 255)
            Case "DARK_GRAY"
                headerBgColor = RGB(64, 64, 64)
                fontColor = RGB(255, 255, 255)
            Case Else
                headerBgColor = RGB(31, 78, 121)
                fontColor = RGB(255, 255, 255)
        End Select
    End If

    With mHeaderRange
        .Interior.Color = headerBgColor
        .Font.Color = fontColor
        .Font.Bold = isBold
        .HorizontalAlignment = Alignment
    End With

    Set ApplyTheme = Me
End Function

' Sets NumberFormat on data cells for a given column (or 1D array/pattern of columns) in a single COM call.
Public Function FormatColumn(ByVal ColumnKey As Variant, ByVal NumberFormatString As String) As cTableStyler
    If mDataBodyRange Is Nothing Or mRowCount = 0 Then
        Set FormatColumn = Me
        Exit Function
    End If

    If IsArray(ColumnKey) Then
        FormatColumns ColumnKey, NumberFormatString
        Set FormatColumn = Me
        Exit Function
    End If

    Dim colIndices() As Long
    colIndices = ResolveColIndices(ColumnKey)
    If UBound(colIndices) >= LBound(colIndices) Then
        Dim i As Long
        For i = LBound(colIndices) To UBound(colIndices)
            Dim colIdx As Long
            colIdx = colIndices(i)
            If colIdx >= 1 And colIdx <= mColumnCount Then
                mDataBodyRange.Columns(colIdx).NumberFormat = NumberFormatString
            End If
        Next i
    End If

    Set FormatColumn = Me
End Function

' Sets NumberFormat on data cells for multiple columns (array, comma-delimited string, or wildcard).
Public Function FormatColumns(ByVal ColumnKeys As Variant, ByVal NumberFormatString As String) As cTableStyler
    If mDataBodyRange Is Nothing Or mRowCount = 0 Then
        Set FormatColumns = Me
        Exit Function
    End If

    Dim colIndices() As Long
    colIndices = ResolveColIndices(ColumnKeys)
    If UBound(colIndices) >= LBound(colIndices) Then
        Dim i As Long
        For i = LBound(colIndices) To UBound(colIndices)
            Dim colIdx As Long
            colIdx = colIndices(i)
            If colIdx >= 1 And colIdx <= mColumnCount Then
                mDataBodyRange.Columns(colIdx).NumberFormat = NumberFormatString
            End If
        Next i
    End If

    Set FormatColumns = Me
End Function

' Sets HorizontalAlignment on data cells (or entire column including header) for a given column or group.
Public Function AlignColumn(ByVal ColumnKeyOrPattern As Variant, ByVal Alignment As Long, _
                            Optional ByVal IncludeHeader As Boolean = False) As cTableStyler
    If mTableRange Is Nothing Then
        Set AlignColumn = Me
        Exit Function
    End If

    Dim colIndices() As Long
    colIndices = ResolveColIndices(ColumnKeyOrPattern)
    If UBound(colIndices) >= LBound(colIndices) Then
        Dim i As Long
        For i = LBound(colIndices) To UBound(colIndices)
            Dim colIdx As Long
            colIdx = colIndices(i)
            If colIdx >= 1 And colIdx <= mColumnCount Then
                If IncludeHeader Then
                    mTableRange.Columns(colIdx).HorizontalAlignment = Alignment
                ElseIf Not mDataBodyRange Is Nothing Then
                    mDataBodyRange.Columns(colIdx).HorizontalAlignment = Alignment
                End If
            End If
        Next i
    End If

    Set AlignColumn = Me
End Function

' Applies thin continuous borders across the entire table range.
' If PreserveTotalRow is True and HighlightTotalRow was already called, it restores total row styling.
Public Function ApplyBorders(Optional ByVal BorderColor As Long = 14277081, _
                             Optional ByVal PreserveTotalRow As Boolean = True) As cTableStyler
    If mTableRange Is Nothing Then
        Set ApplyBorders = Me
        Exit Function
    End If

    Const xlContinuous As Long = 1
    Const xlThin As Long = 2

    With mTableRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = BorderColor
    End With

    If PreserveTotalRow And mTotalRowHighlighted Then
        HighlightTotalRow TopBorder:=True, DoubleBottomBorder:=True
    End If

    Set ApplyBorders = Me
End Function

' Applies alternating light background fill to even data rows.
' Uses batch-union to color even rows in minimal COM roundtrips (100x faster than row-by-row loops).
Public Function ApplyZebraStripes(Optional ByVal StripeColor As Long = 16316664) As cTableStyler
    If mDataBodyRange Is Nothing Or mRowCount < 2 Then
        Set ApplyZebraStripes = Me
        Exit Function
    End If

    Dim appObj As Object
    Set appObj = mWorksheet.Application

    Dim unionRng As Object
    Dim r As Long, batchCount As Long

    ' Batch Union in chunks of 30 to avoid Excel COM Union stack limits on massive tables
    For r = 2 To mRowCount Step 2
        If unionRng Is Nothing Then
            Set unionRng = mDataBodyRange.Rows(r)
        Else
            Set unionRng = appObj.Union(unionRng, mDataBodyRange.Rows(r))
        End If
        batchCount = batchCount + 1

        If batchCount >= 30 Then
            unionRng.Interior.Color = StripeColor
            Set unionRng = Nothing
            batchCount = 0
        End If
    Next r

    If Not unionRng Is Nothing Then
        unionRng.Interior.Color = StripeColor
        Set unionRng = Nothing
    End If

    Set ApplyZebraStripes = Me
End Function

' Highlights the final data row as a Total/Summary row with bold font and double bottom border.
Public Function HighlightTotalRow(Optional ByVal TopBorder As Boolean = True, _
                                  Optional ByVal DoubleBottomBorder As Boolean = True) As cTableStyler
    If mDataBodyRange Is Nothing Or mRowCount = 0 Then
        Set HighlightTotalRow = Me
        Exit Function
    End If

    Dim lastRowRng As Object
    Set lastRowRng = mDataBodyRange.Rows(mRowCount)
    lastRowRng.Font.Bold = True

    Const xlEdgeTop As Long = 8
    Const xlEdgeBottom As Long = 9
    Const xlContinuous As Long = 1
    Const xlDouble As Long = -4119
    Const xlThin As Long = 2

    If TopBorder Then
        With lastRowRng.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End If

    If DoubleBottomBorder Then
        With lastRowRng.Borders(xlEdgeBottom)
            .LineStyle = xlDouble
        End With
    End If

    mTotalRowHighlighted = True

    Set HighlightTotalRow = Me
End Function

' Dynamically writes a summary Total Row below the data body with live Excel formulas (SUBTOTAL or SUM/AVERAGE/COUNTA),
' copies column number formats, expands table dimensions, and formats with bold font and accounting double bottom border.
Public Function AddTotalRow(Optional ByVal SumColumns As Variant, _
                            Optional ByVal AvgColumns As Variant, _
                            Optional ByVal CountColumns As Variant, _
                            Optional ByVal LabelColumn As Variant = 1, _
                            Optional ByVal LabelText As String = "Total", _
                            Optional ByVal UseSubtotal As Boolean = True, _
                            Optional ByVal TopBorder As Boolean = True, _
                            Optional ByVal DoubleBottomBorder As Boolean = True) As cTableStyler
    If mTableRange Is Nothing Or mRowCount = 0 Or mDataBodyRange Is Nothing Then
        Set AddTotalRow = Me
        Exit Function
    End If

    ' Locate target row directly below current table range
    Dim totalRowRng As Object
    Set totalRowRng = mTableRange.Rows(mTableRange.Rows.Count + 1)

    ' 1. Write summary label
    If Not IsMissing(LabelColumn) And Not IsEmpty(LabelColumn) Then
        Dim labelColIdx As Long
        labelColIdx = ResolveColIndex(LabelColumn)
        If labelColIdx >= 1 And labelColIdx <= mColumnCount Then
            totalRowRng.Cells(1, labelColIdx).Value2 = LabelText
        End If
    End If

    Dim i As Long, cIdx As Long, colAddr As String

    ' 2. Process SUM / SUBTOTAL(9)
    If Not IsMissing(SumColumns) And Not IsEmpty(SumColumns) Then
        Dim sumIndices() As Long
        sumIndices = ResolveColIndices(SumColumns)
        If UBound(sumIndices) >= LBound(sumIndices) Then
            For i = LBound(sumIndices) To UBound(sumIndices)
                cIdx = sumIndices(i)
                If cIdx >= 1 And cIdx <= mColumnCount Then
                    colAddr = mDataBodyRange.Columns(cIdx).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    If UseSubtotal Then
                        totalRowRng.Cells(1, cIdx).Formula = "=SUBTOTAL(9," & colAddr & ")"
                    Else
                        totalRowRng.Cells(1, cIdx).Formula = "=SUM(" & colAddr & ")"
                    End If
                    On Error Resume Next
                    totalRowRng.Cells(1, cIdx).NumberFormat = mDataBodyRange.Columns(cIdx).Cells(1, 1).NumberFormat
                    On Error GoTo 0
                End If
            Next i
        End If
    End If

    ' 3. Process AVERAGE / SUBTOTAL(1)
    If Not IsMissing(AvgColumns) And Not IsEmpty(AvgColumns) Then
        Dim avgIndices() As Long
        avgIndices = ResolveColIndices(AvgColumns)
        If UBound(avgIndices) >= LBound(avgIndices) Then
            For i = LBound(avgIndices) To UBound(avgIndices)
                cIdx = avgIndices(i)
                If cIdx >= 1 And cIdx <= mColumnCount Then
                    colAddr = mDataBodyRange.Columns(cIdx).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    If UseSubtotal Then
                        totalRowRng.Cells(1, cIdx).Formula = "=SUBTOTAL(1," & colAddr & ")"
                    Else
                        totalRowRng.Cells(1, cIdx).Formula = "=AVERAGE(" & colAddr & ")"
                    End If
                    On Error Resume Next
                    totalRowRng.Cells(1, cIdx).NumberFormat = mDataBodyRange.Columns(cIdx).Cells(1, 1).NumberFormat
                    On Error GoTo 0
                End If
            Next i
        End If
    End If

    ' 4. Process COUNTA / SUBTOTAL(3)
    If Not IsMissing(CountColumns) And Not IsEmpty(CountColumns) Then
        Dim cntIndices() As Long
        cntIndices = ResolveColIndices(CountColumns)
        If UBound(cntIndices) >= LBound(cntIndices) Then
            For i = LBound(cntIndices) To UBound(cntIndices)
                cIdx = cntIndices(i)
                If cIdx >= 1 And cIdx <= mColumnCount Then
                    colAddr = mDataBodyRange.Columns(cIdx).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    If UseSubtotal Then
                        totalRowRng.Cells(1, cIdx).Formula = "=SUBTOTAL(3," & colAddr & ")"
                    Else
                        totalRowRng.Cells(1, cIdx).Formula = "=COUNTA(" & colAddr & ")"
                    End If
                End If
            Next i
        End If
    End If

    ' 5. Expand internal table range to include new total row
    mRowCount = mRowCount + 1
    Dim totalRows As Long
    totalRows = mRowCount + IIf(mHasHeaders, 1, 0)
    Set mTableRange = mTopLeftCell.Resize(totalRows, mColumnCount)
    Set mDataBodyRange = mTopLeftCell.Offset(IIf(mHasHeaders, 1, 0), 0).Resize(mRowCount, mColumnCount)

    ' 6. Apply total row styling
    HighlightTotalRow TopBorder:=TopBorder, DoubleBottomBorder:=DoubleBottomBorder

    Set AddTotalRow = Me
End Function

' ------------------------------------------------------------------------------
' Flexible Row & Column Sizing Methods
' ------------------------------------------------------------------------------

' Auto-fits all table columns and applies extra width padding.
Public Function AutoFitColumns(Optional ByVal Padding As Double = 3#) As cTableStyler
    If mTableRange Is Nothing Then
        Set AutoFitColumns = Me
        Exit Function
    End If

    On Error Resume Next
    mTableRange.EntireColumn.AutoFit

    If Padding > 0# Then
        Dim c As Long
        For c = 1 To mColumnCount
            mTableRange.Columns(c).ColumnWidth = mTableRange.Columns(c).ColumnWidth + Padding
        Next c
    End If
    On Error GoTo 0

    Set AutoFitColumns = Me
End Function

' Auto-fits row heights across the table (useful after text wrapping).
Public Function AutoFitRows() As cTableStyler
    If Not mTableRange Is Nothing Then
        On Error Resume Next
        mTableRange.EntireRow.AutoFit
        On Error GoTo 0
    End If
    Set AutoFitRows = Me
End Function

' Sets ColumnWidth for specific columns by name, 1-based index, array, comma-list, or wildcard pattern ("Day *").
Public Function SetColumnWidth(ByVal ColumnKeyOrPattern As Variant, ByVal Width As Double) As cTableStyler
    If mTableRange Is Nothing Or Width <= 0# Then
        Set SetColumnWidth = Me
        Exit Function
    End If

    Dim colIndices() As Long
    colIndices = ResolveColIndices(ColumnKeyOrPattern)
    If UBound(colIndices) >= LBound(colIndices) Then
        Dim i As Long
        For i = LBound(colIndices) To UBound(colIndices)
            Dim colIdx As Long
            colIdx = colIndices(i)
            If colIdx >= 1 And colIdx <= mColumnCount Then
                mTableRange.Columns(colIdx).ColumnWidth = Width
            End If
        Next i
    End If

    Set SetColumnWidth = Me
End Function

' Auto-fits columns in a matching group (e.g. "Day *", Array(...)), measures the maximum width,
' and sets ALL columns in that group to that exact same uniform width (+ optional padding).
Public Function SetUniformColumnWidth(ByVal ColumnKeyOrPattern As Variant, _
                                     Optional ByVal Padding As Double = 1.5#, _
                                     Optional ByVal MinWidth As Double = 0#) As cTableStyler
    If mTableRange Is Nothing Then
        Set SetUniformColumnWidth = Me
        Exit Function
    End If

    Dim colIndices() As Long
    colIndices = ResolveColIndices(ColumnKeyOrPattern)
    If UBound(colIndices) < LBound(colIndices) Then
        Set SetUniformColumnWidth = Me
        Exit Function
    End If

    ' 1. AutoFit only the columns in this group to calculate natural content widths
    Dim i As Long, colIdx As Long
    Dim maxW As Double
    maxW = MinWidth

    On Error Resume Next
    For i = LBound(colIndices) To UBound(colIndices)
        colIdx = colIndices(i)
        If colIdx >= 1 And colIdx <= mColumnCount Then
            mTableRange.Columns(colIdx).EntireColumn.AutoFit
            Dim curW As Double
            curW = mTableRange.Columns(colIdx).ColumnWidth
            If curW > maxW Then maxW = curW
        End If
    Next i
    On Error GoTo 0

    ' 2. Apply uniform width across the entire group
    Dim targetW As Double
    targetW = maxW + Padding
    For i = LBound(colIndices) To UBound(colIndices)
        colIdx = colIndices(i)
        If colIdx >= 1 And colIdx <= mColumnCount Then
            mTableRange.Columns(colIdx).ColumnWidth = targetW
        End If
    Next i

    Set SetUniformColumnWidth = Me
End Function

' Sets a uniform column width across all columns in the table in 1 COM call.
Public Function SetAllColumnWidths(ByVal Width As Double) As cTableStyler
    If Not mTableRange Is Nothing And Width > 0# Then
        mTableRange.Columns.ColumnWidth = Width
    End If
    Set SetAllColumnWidths = Me
End Function

' Hides or unhides specific table columns by key or pattern (e.g. "D06,D07" or "*Notes*").
Public Function HideColumn(ByVal ColumnKeyOrPattern As Variant, _
                           Optional ByVal IsHidden As Boolean = True) As cTableStyler
    If mTableRange Is Nothing Or mColumnCount = 0 Then
        Set HideColumn = Me
        Exit Function
    End If

    Dim colIndices() As Long
    colIndices = ResolveColIndices(ColumnKeyOrPattern)
    Dim i As Long, cIdx As Long
    For i = LBound(colIndices) To UBound(colIndices)
        cIdx = colIndices(i)
        If cIdx >= 1 And cIdx <= mColumnCount Then
            mTableRange.Columns(cIdx).EntireColumn.Hidden = IsHidden
        End If
    Next i

    Set HideColumn = Me
End Function

' Highlights entire row when cell condition in ColumnKey is satisfied.
Public Function HighlightRowCondition(ByVal ColumnKey As Variant, _
                                      ByVal Op As String, _
                                      ByVal Val As Variant, _
                                      Optional ByVal FillColor As Long = 13561798, _
                                      Optional ByVal FontColor As Long = 1875718, _
                                      Optional ByVal Bold As Boolean = False) As cTableStyler
    If mDataBodyRange Is Nothing Or mColumnCount = 0 Then
        Set HighlightRowCondition = Me
        Exit Function
    End If

    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)
    If cIdx < 1 Or cIdx > mColumnCount Then
        Set HighlightRowCondition = Me
        Exit Function
    End If

    Dim colLetter As String
    colLetter = ColLetterFromIndex(mTopLeftCell.Column + cIdx - 1)
    Dim firstRow As Long
    firstRow = mDataBodyRange.Row

    Dim fmla As String
    Dim valStr As String
    If IsNumeric(Val) Then
        valStr = CStr(Val)
    Else
        valStr = """" & Replace(CStr(Val), """", """""") & """"
    End If

    Select Case UCase$(Trim$(Op))
        Case "=", "==", "EQUALS": fmla = "=$" & colLetter & firstRow & "=" & valStr
        Case "<>", "!=", "NOT_EQUALS": fmla = "=$" & colLetter & firstRow & "<>" & valStr
        Case ">", "GREATER": fmla = "=$" & colLetter & firstRow & ">" & valStr
        Case ">=", "GREATER_OR_EQUAL": fmla = "=$" & colLetter & firstRow & ">=" & valStr
        Case "<", "LESS": fmla = "=$" & colLetter & firstRow & "<" & valStr
        Case "<=", "LESS_OR_EQUAL": fmla = "=$" & colLetter & firstRow & "<=" & valStr
        Case Else: fmla = "=$" & colLetter & firstRow & "=" & valStr
    End Select

    On Error Resume Next
    Dim fc As Object
    ' xlExpression = 2
    Set fc = mDataBodyRange.FormatConditions.Add(2, , fmla)
    If Not fc Is Nothing Then
        If FillColor >= 0 Then fc.Interior.Color = FillColor
        If FontColor >= 0 Then fc.Font.Color = FontColor
        If Bold Then fc.Font.Bold = True
    End If
    On Error GoTo 0

    Set HighlightRowCondition = Me
End Function

' Adds an Excel In-Cell Data Bar to a column.
Public Function ApplyDataBars(ByVal ColumnKey As Variant, _
                              Optional ByVal BarColor As Long = 15773696) As cTableStyler
    If mDataBodyRange Is Nothing Or mColumnCount = 0 Then
        Set ApplyDataBars = Me
        Exit Function
    End If

    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)
    If cIdx < 1 Or cIdx > mColumnCount Then
        Set ApplyDataBars = Me
        Exit Function
    End If

    Dim colRng As Object
    Set colRng = mDataBodyRange.Columns(cIdx)

    On Error Resume Next
    Dim db As Object
    ' xlDatabar = 4
    Set db = colRng.FormatConditions.AddDatabar
    If Not db Is Nothing Then
        db.BarColor.Color = BarColor
    End If
    On Error GoTo 0

    Set ApplyDataBars = Me
End Function

' Adds a 2-color gradient Color Scale across a column (e.g. Min to Max heatmap).
Public Function ApplyColorScale(ByVal ColumnKey As Variant, _
                                Optional ByVal MinColor As Long = 16777215, _
                                Optional ByVal MaxColor As Long = 13561798) As cTableStyler
    If mDataBodyRange Is Nothing Or mColumnCount = 0 Then
        Set ApplyColorScale = Me
        Exit Function
    End If

    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)
    If cIdx < 1 Or cIdx > mColumnCount Then
        Set ApplyColorScale = Me
        Exit Function
    End If

    Dim colRng As Object
    Set colRng = mDataBodyRange.Columns(cIdx)

    On Error Resume Next
    Dim cs As Object
    ' xlColorScale = 3
    Set cs = colRng.FormatConditions.AddColorScale(2)
    If Not cs Is Nothing Then
        If MinColor >= 0 Then cs.ColorScaleCriteria(1).FormatColor.Color = MinColor
        If MaxColor >= 0 Then cs.ColorScaleCriteria(2).FormatColor.Color = MaxColor
    End If
    On Error GoTo 0

    Set ApplyColorScale = Me
End Function

' Sets row height specifically on the header row in 1 COM call.
Public Function SetHeaderRowHeight(ByVal Height As Double) As cTableStyler
    If Not mHeaderRange Is Nothing And Height > 0# Then
        mHeaderRange.RowHeight = Height
    End If
    Set SetHeaderRowHeight = Me
End Function

' Sets row height uniformly across all data rows (employee records) in 1 COM call.
Public Function SetDataRowHeight(ByVal Height As Double) As cTableStyler
    If Not mDataBodyRange Is Nothing And Height > 0# Then
        mDataBodyRange.RowHeight = Height
    End If
    Set SetDataRowHeight = Me
End Function

' Identifies section divider rows in a column by pattern (e.g. "[ TEAM: * ]" or "[ * ]")
' and sets their row height and optional styling (e.g. larger height or background for section dividers).
Public Function SetSectionRowHeight(ByVal ColumnKey As Variant, _
                                    ByVal MatchPattern As String, _
                                    ByVal Height As Double, _
                                    Optional ByVal Bold As Boolean = True, _
                                    Optional ByVal BgColor As Long = -1, _
                                    Optional ByVal FontColor As Long = -1) As cTableStyler
    If mDataBodyRange Is Nothing Or mRowCount = 0 Or Height <= 0# Then
        Set SetSectionRowHeight = Me
        Exit Function
    End If

    Dim colIdx As Long
    colIdx = ResolveColIndex(ColumnKey)
    If colIdx < 1 Or colIdx > mColumnCount Then
        Err.Raise 9, "cTableStyler.SetSectionRowHeight", "Column index out of range: " & CStr(ColumnKey)
    End If

    ' Read entire column values in 1 COM call
    Dim colVals As Variant
    colVals = mDataBodyRange.Columns(colIdx).Value2

    Dim r As Long, rowText As String
    For r = 1 To mRowCount
        If mRowCount = 1 Then
            rowText = CStr(colVals)
        Else
            rowText = CStr(colVals(r, 1))
        End If

        Dim isMatch As Boolean
        isMatch = (rowText Like MatchPattern)
        If Not isMatch And (InStr(MatchPattern, "[") > 0 Or InStr(MatchPattern, "]") > 0) Then
            Dim safePat As String
            safePat = Replace(MatchPattern, "[", Chr$(1))
            safePat = Replace(safePat, "]", "[]]")
            safePat = Replace(safePat, Chr$(1), "[[]")
            isMatch = (rowText Like safePat)
        End If
        If Not isMatch And InStr(MatchPattern, "*") = 0 And InStr(MatchPattern, "?") = 0 Then
            isMatch = (InStr(1, rowText, MatchPattern, vbTextCompare) > 0)
        End If

        If isMatch Then
            Dim rowRng As Object
            Set rowRng = mDataBodyRange.Rows(r)
            rowRng.RowHeight = Height
            If Bold Then rowRng.Font.Bold = True
            If BgColor >= 0 Then rowRng.Interior.Color = BgColor
            If FontColor >= 0 Then rowRng.Font.Color = FontColor
        End If
    Next r

    Set SetSectionRowHeight = Me
End Function

' Sets WrapText property for specific columns (array, pattern, or single column).
Public Function SetWrapText(ByVal ColumnKeyOrPattern As Variant, Optional ByVal Wrap As Boolean = True) As cTableStyler
    If mTableRange Is Nothing Then
        Set SetWrapText = Me
        Exit Function
    End If

    Dim colIndices() As Long
    colIndices = ResolveColIndices(ColumnKeyOrPattern)
    If UBound(colIndices) >= LBound(colIndices) Then
        Dim i As Long
        For i = LBound(colIndices) To UBound(colIndices)
            Dim colIdx As Long
            colIdx = colIndices(i)
            If colIdx >= 1 And colIdx <= mColumnCount Then
                If Not mDataBodyRange Is Nothing Then
                    mDataBodyRange.Columns(colIdx).WrapText = Wrap
                End If
            End If
        Next i
    End If

    Set SetWrapText = Me
End Function

' Sets WrapText on the header row to allow multi-line column titles.
Public Function SetHeaderWrapText(Optional ByVal Wrap As Boolean = True) As cTableStyler
    If Not mHeaderRange Is Nothing Then
        mHeaderRange.WrapText = Wrap
    End If
    Set SetHeaderWrapText = Me
End Function

' ------------------------------------------------------------------------------
' Conditional Status Highlighting, Freeze Panes & Typography
' ------------------------------------------------------------------------------

' Applies conditional formatting rule on specified column(s) matching a status or KPI value.
' Supports exact text/numeric comparison. ColumnKeyOrPattern can be name, index, array, or wildcard ("Day *").
Public Function HighlightValues(ByVal ColumnKeyOrPattern As Variant, _
                                ByVal ValueToMatch As Variant, _
                                ByVal BgColor As Long, _
                                Optional ByVal FontColor As Long = -1, _
                                Optional ByVal Bold As Boolean = False) As cTableStyler
    If mDataBodyRange Is Nothing Or mRowCount = 0 Then
        Set HighlightValues = Me
        Exit Function
    End If

    Dim colIndices() As Long
    colIndices = ResolveColIndices(ColumnKeyOrPattern)
    If UBound(colIndices) < LBound(colIndices) Then
        Set HighlightValues = Me
        Exit Function
    End If

    Const xlCellValue As Long = 1
    Const xlEqual As Long = 3

    Dim formulaStr As String
    If VarType(ValueToMatch) = vbString Then
        formulaStr = Chr$(34) & CStr(ValueToMatch) & Chr$(34)
    Else
        formulaStr = CStr(ValueToMatch)
    End If

    Dim i As Long, colIdx As Long
    For i = LBound(colIndices) To UBound(colIndices)
        colIdx = colIndices(i)
        If colIdx >= 1 And colIdx <= mColumnCount Then
            Dim colRng As Object
            Set colRng = mDataBodyRange.Columns(colIdx)

            Dim fc As Object
            Set fc = colRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=formulaStr)
            fc.Interior.Color = BgColor
            If FontColor >= 0 Then fc.Font.Color = FontColor
            If Bold Then fc.Font.Bold = True
        End If
    Next i

    Set HighlightValues = Me
End Function

' Applies conditional formatting based on comparison operators (>, <, >=, <=, =, <>, BETWEEN, NOT_BETWEEN).
' Supports numeric thresholds and range boundaries across single columns, comma-delimited strings, or wildcards.
Public Function HighlightCondition(ByVal ColumnKeyOrPattern As Variant, _
                                   ByVal Operator As Variant, _
                                   ByVal Formula1 As Variant, _
                                   ByVal BgColor As Long, _
                                   Optional ByVal FontColor As Long = -1, _
                                   Optional ByVal Bold As Boolean = False, _
                                   Optional ByVal Formula2 As Variant) As cTableStyler
    If mDataBodyRange Is Nothing Or mRowCount = 0 Then
        Set HighlightCondition = Me
        Exit Function
    End If

    Dim colIndices() As Long
    colIndices = ResolveColIndices(ColumnKeyOrPattern)
    If UBound(colIndices) < LBound(colIndices) Then
        Set HighlightCondition = Me
        Exit Function
    End If

    Const xlCellValue As Long = 1
    Const xlBetween As Long = 1
    Const xlNotBetween As Long = 2
    Const xlEqual As Long = 3
    Const xlNotEqual As Long = 4
    Const xlGreater As Long = 5
    Const xlLess As Long = 6
    Const xlGreaterEqual As Long = 7
    Const xlLessEqual As Long = 8

    Dim opCode As Long
    If IsNumeric(Operator) Then
        opCode = CLng(Operator)
    Else
        Select Case UCase$(Trim$(CStr(Operator)))
            Case ">", "GT", "GREATER": opCode = xlGreater
            Case "<", "LT", "LESS": opCode = xlLess
            Case ">=", "GTE", "GE", "GREATER_EQUAL": opCode = xlGreaterEqual
            Case "<=", "LTE", "LE", "LESS_EQUAL": opCode = xlLessEqual
            Case "=", "==", "EQ", "EQUAL": opCode = xlEqual
            Case "<>", "!=", "NE", "NOT_EQUAL": opCode = xlNotEqual
            Case "BETWEEN": opCode = xlBetween
            Case "NOT_BETWEEN", "NOT BETWEEN": opCode = xlNotBetween
            Case Else: opCode = xlEqual
        End Select
    End If

    Dim f1Str As String, f2Str As String
    If VarType(Formula1) = vbString And Not IsNumeric(Formula1) Then
        f1Str = Chr$(34) & CStr(Formula1) & Chr$(34)
    Else
        f1Str = CStr(Formula1)
    End If

    If Not IsMissing(Formula2) And Not IsEmpty(Formula2) Then
        If VarType(Formula2) = vbString And Not IsNumeric(Formula2) Then
            f2Str = Chr$(34) & CStr(Formula2) & Chr$(34)
        Else
            f2Str = CStr(Formula2)
        End If
    End If

    Dim i As Long, cIdx As Long
    For i = LBound(colIndices) To UBound(colIndices)
        cIdx = colIndices(i)
        If cIdx >= 1 And cIdx <= mColumnCount Then
            Dim colRng As Object
            Set colRng = mDataBodyRange.Columns(cIdx)
            Dim fc As Object
            If opCode = xlBetween Or opCode = xlNotBetween Then
                Set fc = colRng.FormatConditions.Add(Type:=xlCellValue, Operator:=opCode, Formula1:=f1Str, Formula2:=f2Str)
            Else
                Set fc = colRng.FormatConditions.Add(Type:=xlCellValue, Operator:=opCode, Formula1:=f1Str)
            End If
            fc.Interior.Color = BgColor
            If FontColor >= 0 Then fc.Font.Color = FontColor
            If Bold Then fc.Font.Bold = True
        End If
    Next i

    Set HighlightCondition = Me
End Function

' Freezes worksheet panes at the boundary of the table (header row and leading columns).
Public Function FreezePanes(Optional ByVal FreezeRows As Long = 1, _
                            Optional ByVal FreezeCols As Long = 0) As cTableStyler
    If mTopLeftCell Is Nothing Then
        Set FreezePanes = Me
        Exit Function
    End If

    On Error Resume Next
    Dim appObj As Object
    Set appObj = mWorksheet.Application
    mWorksheet.Activate
    appObj.ActiveWindow.FreezePanes = False

    Dim splitRow As Long, splitCol As Long
    splitRow = mTopLeftCell.Row + FreezeRows
    splitCol = mTopLeftCell.Column + FreezeCols

    mWorksheet.Cells(splitRow, splitCol).Select
    appObj.ActiveWindow.FreezePanes = True
    On Error GoTo 0

    Set FreezePanes = Me
End Function

' Sets font name and sizes across the entire table, data body, and header row.
Public Function SetFont(Optional ByVal FontName As String = "Segoe UI", _
                        Optional ByVal DataSize As Double = 10#, _
                        Optional ByVal HeaderSize As Double = 11#) As cTableStyler
    If Not mTableRange Is Nothing And Len(Trim$(FontName)) > 0 Then
        mTableRange.Font.Name = FontName
    End If
    If Not mDataBodyRange Is Nothing And DataSize > 0# Then
        mDataBodyRange.Font.Size = DataSize
    End If
    If Not mHeaderRange Is Nothing And HeaderSize > 0# Then
        mHeaderRange.Font.Size = HeaderSize
    End If
    Set SetFont = Me
End Function

' Toggles worksheet gridlines visibility.
Public Function SetGridlines(Optional ByVal Visible As Boolean = True) As cTableStyler
    If Not mWorksheet Is Nothing Then
        On Error Resume Next
        mWorksheet.Activate
        mWorksheet.Application.ActiveWindow.DisplayGridlines = Visible
        On Error GoTo 0
    End If
    Set SetGridlines = Me
End Function

' Renders super-header group banners (using xlCenterAcrossSelection) above table columns.
' GroupSpecs: array of Array(Label, StartColKey, EndColKey [, BgColor]).
Public Function ApplyColumnGroupBanner(ByVal GroupSpecs As Variant, _
                                       Optional ByVal HeaderRowOffset As Long = -1, _
                                       Optional ByVal RowHeight As Double = 20#, _
                                       Optional ByVal FontColor As Long = 16777215) As cTableStyler
    If mTableRange Is Nothing Or mColumnCount = 0 Or Not IsArray(GroupSpecs) Then
        Set ApplyColumnGroupBanner = Me
        Exit Function
    End If

    Dim bannerRow As Long
    bannerRow = mTopLeftCell.Row + HeaderRowOffset
    If bannerRow < 1 Then
        Set ApplyColumnGroupBanner = Me
        Exit Function
    End If

    If RowHeight > 0# Then mWorksheet.Rows(bannerRow).RowHeight = RowHeight

    Dim i As Long, idxNum As Long
    idxNum = 0
    For i = LBound(GroupSpecs) To UBound(GroupSpecs)
        Dim spec As Variant
        spec = GroupSpecs(i)
        If IsArray(spec) Then
            idxNum = idxNum + 1
            Dim lb As Long, ub As Long
            lb = LBound(spec)
            ub = UBound(spec)
            Dim lbl As String, cStart As Long, cEnd As Long, bgClr As Long
            lbl = CStr(spec(lb))
            cStart = ResolveColIndex(spec(lb + 1))
            cEnd = ResolveColIndex(spec(lb + 2))
            If ub >= lb + 3 Then
                bgClr = CLng(spec(lb + 3))
            Else
                bgClr = IIf((idxNum Mod 2) = 1, RGB(30, 41, 59), RGB(51, 65, 85))
            End If

            If cStart >= 1 And cEnd >= cStart And cEnd <= mColumnCount Then
                Dim absColStart As Long, absColEnd As Long
                absColStart = mTopLeftCell.Column + cStart - 1
                absColEnd = mTopLeftCell.Column + cEnd - 1
                Dim grpRng As Object
                Set grpRng = mWorksheet.Range(mWorksheet.Cells(bannerRow, absColStart), mWorksheet.Cells(bannerRow, absColEnd))
                mWorksheet.Cells(bannerRow, absColStart).Value = lbl
                grpRng.HorizontalAlignment = IIf(absColEnd > absColStart, 7, -4108) ' xlCenterAcrossSelection / xlCenter
                grpRng.VerticalAlignment = -4108
                grpRng.Font.Name = "Segoe UI"
                grpRng.Font.Size = 9
                grpRng.Font.Bold = True
                grpRng.Font.Color = FontColor
                grpRng.Interior.Color = bgClr
                ApplyBoxBorders grpRng, 2, RGB(148, 163, 184)
            End If
        End If
    Next i

    Set ApplyColumnGroupBanner = Me
End Function

' Applies vertical xlMedium divider borders across column groups and subtle alternating fill on even groups.
' GroupSpecs: array of Array(LabelOrUnused, StartColKey, EndColKey [, BgColor]).
Public Function ApplyColumnGroupGrid(ByVal GroupSpecs As Variant, _
                                     Optional ByVal SuperHeaderRowOffset As Long = -1, _
                                     Optional ByVal DividerColor As Long = 6903111, _
                                     Optional ByVal EvenGroupFillColor As Long = 16381425, _
                                     Optional ByVal IncludeLeftEdge As Boolean = True) As cTableStyler
    If mTableRange Is Nothing Or mColumnCount = 0 Or Not IsArray(GroupSpecs) Then
        Set ApplyColumnGroupGrid = Me
        Exit Function
    End If

    Const xlContinuous As Long = 1
    Const xlMedium As Long = -4138
    Const xlEdgeLeft As Long = 7
    Const xlEdgeRight As Long = 10

    Dim topRow As Long, bottomRow As Long
    topRow = mTopLeftCell.Row + SuperHeaderRowOffset
    If topRow < 1 Then topRow = mTopLeftCell.Row
    bottomRow = mTopLeftCell.Row + IIf(mHasHeaders, 1, 0) + mRowCount - 1
    Dim pureRows As Long
    pureRows = PureDataRowCount()

    Dim i As Long, grpNum As Long
    grpNum = 0
    For i = LBound(GroupSpecs) To UBound(GroupSpecs)
        Dim spec As Variant
        spec = GroupSpecs(i)
        If IsArray(spec) Then
            grpNum = grpNum + 1
            Dim lb As Long
            lb = LBound(spec)
            Dim cStart As Long, cEnd As Long
            cStart = ResolveColIndex(spec(lb + 1))
            cEnd = ResolveColIndex(spec(lb + 2))
            If cStart >= 1 And cEnd >= cStart And cEnd <= mColumnCount Then
                Dim absColStart As Long, absColEnd As Long
                absColStart = mTopLeftCell.Column + cStart - 1
                absColEnd = mTopLeftCell.Column + cEnd - 1

                If grpNum = 1 And IncludeLeftEdge Then
                    Dim leftRng As Object
                    Set leftRng = mWorksheet.Range(mWorksheet.Cells(topRow, absColStart), mWorksheet.Cells(bottomRow, absColStart))
                    With leftRng.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                        .Color = DividerColor
                    End With
                End If

                If (grpNum Mod 2) = 0 And pureRows > 0 And Not mDataBodyRange Is Nothing Then
                    Dim bodyStartRow As Long
                    bodyStartRow = mDataBodyRange.Row
                    Dim fillRng As Object
                    Set fillRng = mWorksheet.Range(mWorksheet.Cells(bodyStartRow, absColStart), mWorksheet.Cells(bodyStartRow + pureRows - 1, absColEnd))
                    fillRng.Interior.Color = EvenGroupFillColor
                End If

                Dim rightRng As Object
                Set rightRng = mWorksheet.Range(mWorksheet.Cells(topRow, absColEnd), mWorksheet.Cells(bottomRow, absColEnd))
                With rightRng.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = DividerColor
                End With
            End If
        End If
    Next i

    Set ApplyColumnGroupGrid = Me
End Function

' Convenience preset for monthly daily calendar grids (WEEK 1..W5 + SUMMARY, weekly borders, active highlights, unused day dimming).
Public Function ApplyWeeklyCalendarGrid(ByVal FirstDayColKey As Variant, _
                                        ByVal DaysInMonth As Long, _
                                        Optional ByVal MaxDayCols As Long = 31, _
                                        Optional ByVal SummaryStartColKey As Variant = Empty, _
                                        Optional ByVal SummaryEndColKey As Variant = Empty, _
                                        Optional ByVal SuperHeaderRowOffset As Long = -1) As cTableStyler
    If mTableRange Is Nothing Or mColumnCount = 0 Or DaysInMonth < 1 Then
        Set ApplyWeeklyCalendarGrid = Me
        Exit Function
    End If

    Dim firstDayIdx As Long, lastActiveDayIdx As Long
    firstDayIdx = ResolveColIndex(FirstDayColKey)
    lastActiveDayIdx = firstDayIdx + DaysInMonth - 1
    If lastActiveDayIdx > mColumnCount Then lastActiveDayIdx = mColumnCount

    Dim pureRows As Long
    pureRows = PureDataRowCount()

    ' Center daily values and apply >=1 emerald highlight on active day columns
    If pureRows > 0 And Not mDataBodyRange Is Nothing Then
        Dim activeDataRng As Object
        Set activeDataRng = mWorksheet.Range(mDataBodyRange.Cells(1, firstDayIdx), mDataBodyRange.Cells(pureRows, lastActiveDayIdx))
        activeDataRng.HorizontalAlignment = -4108
        activeDataRng.NumberFormat = "#,##0;-#,##0;""-"""

        On Error Resume Next
        activeDataRng.FormatConditions.Delete
        Dim cf As Object
        Set cf = activeDataRng.FormatConditions.Add(1, 5, "1") ' xlCellValue, xlGreaterEqual
        If Not cf Is Nothing Then
            cf.Interior.Color = RGB(220, 252, 231)
            cf.Font.Color = RGB(22, 101, 52)
            cf.Font.Bold = True
        End If
        On Error GoTo 0
    End If

    ' Center total row daily values
    If mTotalRowHighlighted And Not mDataBodyRange Is Nothing Then
        Dim totalDayRng As Object
        Set totalDayRng = mWorksheet.Range(mDataBodyRange.Cells(mRowCount, firstDayIdx), mDataBodyRange.Cells(mRowCount, lastActiveDayIdx))
        totalDayRng.HorizontalAlignment = -4108
        totalDayRng.NumberFormat = "#,##0"
    End If

    ' Build week group descriptors (W1..W5)
    Dim weekSpecs() As Variant
    Dim specCount As Long
    specCount = 0
    Dim w As Long, wkStartDay As Long, wkEndDay As Long
    For w = 1 To 5
        wkStartDay = (w - 1) * 7 + 1
        If wkStartDay <= DaysInMonth Then
            wkEndDay = w * 7
            If wkEndDay > DaysInMonth Then wkEndDay = DaysInMonth

            Dim wkLabel As String
            If (wkEndDay - wkStartDay + 1) >= 4 Then
                wkLabel = "WEEK " & CStr(w) & " (D" & Format$(wkStartDay, "00") & "-D" & Format$(wkEndDay, "00") & ")"
            ElseIf (wkEndDay - wkStartDay + 1) >= 2 Then
                wkLabel = "W" & CStr(w) & " (" & Format$(wkStartDay, "00") & "-" & Format$(wkEndDay, "00") & ")"
            Else
                wkLabel = "W" & CStr(w)
            End If

            specCount = specCount + 1
            ReDim Preserve weekSpecs(0 To specCount - 1)
            weekSpecs(specCount - 1) = Array(wkLabel, firstDayIdx + wkStartDay - 1, firstDayIdx + wkEndDay - 1, _
                                             IIf((w Mod 2) = 1, RGB(30, 41, 59), RGB(51, 65, 85)))
        End If
    Next w

    If specCount > 0 Then
        ApplyColumnGroupBanner weekSpecs, SuperHeaderRowOffset, 20#, RGB(255, 255, 255)
        ApplyColumnGroupGrid weekSpecs, SuperHeaderRowOffset, RGB(71, 85, 105), RGB(241, 245, 249), True
    End If

    ' Dim trailing padded day columns if month has fewer days than MaxDayCols
    If DaysInMonth < MaxDayCols Then
        Dim padStartIdx As Long, padEndIdx As Long
        padStartIdx = firstDayIdx + DaysInMonth
        padEndIdx = firstDayIdx + MaxDayCols - 1
        If padEndIdx > mColumnCount Then padEndIdx = mColumnCount
        If padEndIdx >= padStartIdx Then
            Dim padDividerRng As Object
            Dim topPadRow As Long, botPadRow As Long
            topPadRow = mTopLeftCell.Row + SuperHeaderRowOffset
            If topPadRow < 1 Then topPadRow = mTopLeftCell.Row
            botPadRow = mTopLeftCell.Row + IIf(mHasHeaders, 1, 0) + mRowCount - 1

            DimColumns padStartIdx, padEndIdx, "-", RGB(226, 232, 240), RGB(148, 163, 184), SuperHeaderRowOffset
            Set padDividerRng = mWorksheet.Range(mWorksheet.Cells(topPadRow, mTopLeftCell.Column + padEndIdx - 1), _
                                                 mWorksheet.Cells(botPadRow, mTopLeftCell.Column + padEndIdx - 1))
            With padDividerRng.Borders(10) ' xlEdgeRight
                .LineStyle = 1
                .Weight = -4138 ' xlMedium
                .Color = RGB(71, 85, 105)
            End With
        End If
    End If

    ' Format optional Summary super-header & columns
    If Not IsMissing(SummaryStartColKey) And Not IsEmpty(SummaryStartColKey) Then
        Dim sumStartIdx As Long, sumEndIdx As Long
        sumStartIdx = ResolveColIndex(SummaryStartColKey)
        If IsMissing(SummaryEndColKey) Or IsEmpty(SummaryEndColKey) Then
            sumEndIdx = sumStartIdx
        Else
            sumEndIdx = ResolveColIndex(SummaryEndColKey)
        End If
        Dim sumSpec(0 To 0) As Variant
        sumSpec(0) = Array("SUMMARY", sumStartIdx, sumEndIdx, RGB(15, 23, 42))
        ApplyColumnGroupBanner sumSpec, SuperHeaderRowOffset, 20#, RGB(255, 255, 255)

        If Not mDataBodyRange Is Nothing Then
            mDataBodyRange.Columns(sumStartIdx).NumberFormat = "#,##0"
            mDataBodyRange.Columns(sumStartIdx).Font.Bold = True
            If sumEndIdx > sumStartIdx Then
                mDataBodyRange.Columns(sumEndIdx).NumberFormat = "0.0"
            End If
        End If
    End If

    Set ApplyWeeklyCalendarGrid = Me
End Function

' Spotlights a specific column (e.g., Today) with distinct header fill, body tint, and medium vertical borders.
Public Function SpotlightColumn(ByVal ColumnKey As Variant, _
                                Optional ByVal HeaderBgColor As Long = 423897, _
                                Optional ByVal BodyBgColor As Long = 13104382, _
                                Optional ByVal BorderColor As Long = 423897, _
                                Optional ByVal HeaderFontColor As Long = 16777215) As cTableStyler
    If mTableRange Is Nothing Or mColumnCount = 0 Then
        Set SpotlightColumn = Me
        Exit Function
    End If

    Dim cIdx As Long
    cIdx = ResolveColIndex(ColumnKey)
    If cIdx < 1 Or cIdx > mColumnCount Then
        Set SpotlightColumn = Me
        Exit Function
    End If

    If Not mHeaderRange Is Nothing Then
        With mHeaderRange.Cells(1, cIdx)
            .Interior.Color = HeaderBgColor
            .Font.Color = HeaderFontColor
            .Font.Bold = True
        End With
    End If

    Dim pureRows As Long
    pureRows = PureDataRowCount()
    If pureRows > 0 And Not mDataBodyRange Is Nothing Then
        Dim bodyRng As Object
        Set bodyRng = mWorksheet.Range(mDataBodyRange.Cells(1, cIdx), mDataBodyRange.Cells(pureRows, cIdx))
        bodyRng.Interior.Color = BodyBgColor
    End If

    Dim fullColRng As Object
    Set fullColRng = mTableRange.Columns(cIdx)
    With fullColRng.Borders(7) ' xlEdgeLeft
        .LineStyle = 1
        .Weight = -4138 ' xlMedium
        .Color = BorderColor
    End With
    With fullColRng.Borders(10) ' xlEdgeRight
        .LineStyle = 1
        .Weight = -4138 ' xlMedium
        .Color = BorderColor
    End With

    Set SpotlightColumn = Me
End Function

' Dims trailing or inactive columns (across optional super-header, header, data, and total row).
Public Function DimColumns(ByVal FromColKey As Variant, _
                           ByVal ToColKey As Variant, _
                           Optional ByVal PlaceholderText As String = "-", _
                           Optional ByVal BgColor As Long = 15788258, _
                           Optional ByVal FontColor As Long = 12100498, _
                           Optional ByVal SuperHeaderRowOffset As Long = -1) As cTableStyler
    If mTableRange Is Nothing Or mColumnCount = 0 Then
        Set DimColumns = Me
        Exit Function
    End If

    Dim cFrom As Long, cTo As Long
    cFrom = ResolveColIndex(FromColKey)
    cTo = ResolveColIndex(ToColKey)
    If cFrom < 1 Or cTo < cFrom Or cTo > mColumnCount Then
        Set DimColumns = Me
        Exit Function
    End If

    Dim absFrom As Long, absTo As Long
    absFrom = mTopLeftCell.Column + cFrom - 1
    absTo = mTopLeftCell.Column + cTo - 1

    Dim topRow As Long, bottomRow As Long
    topRow = mTopLeftCell.Row + SuperHeaderRowOffset
    If topRow < 1 Then topRow = mTopLeftCell.Row
    bottomRow = mTopLeftCell.Row + IIf(mHasHeaders, 1, 0) + mRowCount - 1

    Dim dimRng As Object
    Set dimRng = mWorksheet.Range(mWorksheet.Cells(topRow, absFrom), mWorksheet.Cells(bottomRow, absTo))
    dimRng.Interior.Color = BgColor
    dimRng.Font.Color = FontColor
    dimRng.HorizontalAlignment = -4108

    If Len(PlaceholderText) > 0 And Not mDataBodyRange Is Nothing Then
        mWorksheet.Range(mDataBodyRange.Cells(1, cFrom), mDataBodyRange.Cells(mRowCount, cTo)).Value = PlaceholderText
    End If

    Set DimColumns = Me
End Function

' Applies styling rules read from an Excel Table (ListObject or Table Name), Range, or cDataFrame.
' Supported columns (case-insensitive): RuleType/Action/Type, Target/Column/ColumnKey,
' Value/Param/Format, Operator/Op, FillColor/BgColor, FontColor/TextColor, Scope/Section, Enabled.
Public Function ApplyStyleFromTable(ByVal StyleTableOrDf As Variant, _
                                     Optional ByVal TargetScope As String = "") As cTableStyler
    If mTableRange Is Nothing Or mColumnCount = 0 Then
        Set ApplyStyleFromTable = Me
        Exit Function
    End If

    Dim cfgDf As cDataFrame
    Set cfgDf = ResolveStyleConfigDf(StyleTableOrDf)
    If cfgDf.RowCount = 0 Then
        Set ApplyStyleFromTable = Me
        Exit Function
    End If

    Dim ruleCol As String, targetCol As String, valCol As String
    Dim opCol As String, fillCol As String, fontCol As String, enabledCol As String, scopeCol As String
    ruleCol = FindStyleColName(cfgDf, Array("RuleType", "Action", "Type", "Rule"))
    If Len(ruleCol) = 0 Then
        Err.Raise 5, "cTableStyler.ApplyStyleFromTable", "Style table requires a 'RuleType' (or 'Action'/'Type') column."
    End If

    targetCol = FindStyleColName(cfgDf, Array("Target", "Column", "ColumnKey", "Columns"))
    valCol = FindStyleColName(cfgDf, Array("Value", "Param", "Param1", "Format", "StyleValue"))
    opCol = FindStyleColName(cfgDf, Array("Operator", "Op", "Condition", "Param2"))
    fillCol = FindStyleColName(cfgDf, Array("FillColor", "BgColor", "Background", "Fill"))
    fontCol = FindStyleColName(cfgDf, Array("FontColor", "TextColor", "Color"))
    enabledCol = FindStyleColName(cfgDf, Array("Enabled", "Active", "IsEnabled"))
    scopeCol = FindStyleColName(cfgDf, Array("Scope", "Section", "TargetTable", "TableScope"))

    Dim r As Long
    For r = 1 To cfgDf.RowCount
        If Len(enabledCol) > 0 Then
            If Not ParseStyleBool(cfgDf.Item(r, enabledCol), True) Then GoTo NextStyleRow
        End If

        ' Filter by Scope if specified
        If Len(scopeCol) > 0 And Len(TargetScope) > 0 Then
            Dim rowScope As String
            rowScope = UCase$(Trim$(CStr(cfgDf.Item(r, scopeCol))))
            If Len(rowScope) > 0 And rowScope <> "ALL" And rowScope <> "*" Then
                If StrComp(rowScope, UCase$(Trim$(TargetScope)), vbTextCompare) <> 0 Then
                    GoTo NextStyleRow
                End If
            End If
        End If

        Dim ruleType As String
        ruleType = UCase$(Trim$(CStr(cfgDf.Item(r, ruleCol))))
        If Len(ruleType) = 0 Then GoTo NextStyleRow

        Dim targetVal As String, paramVal As String, opVal As String
        If Len(targetCol) > 0 Then targetVal = Trim$(CStr(cfgDf.Item(r, targetCol))) Else targetVal = ""
        If Len(valCol) > 0 Then paramVal = Trim$(CStr(cfgDf.Item(r, valCol))) Else paramVal = ""
        If Len(opCol) > 0 Then opVal = Trim$(CStr(cfgDf.Item(r, opCol))) Else opVal = ""

        Dim fillClr As Long, fontClr As Long
        If Len(fillCol) > 0 Then fillClr = ParseStyleColor(cfgDf.Item(r, fillCol), -1) Else fillClr = -1
        If Len(fontCol) > 0 Then fontClr = ParseStyleColor(cfgDf.Item(r, fontCol), -1) Else fontClr = -1

        Select Case ruleType
            Case "THEME"
                Dim thmName As String
                thmName = IIf(Len(paramVal) > 0, paramVal, targetVal)
                If Len(thmName) > 0 Then Me.ApplyTheme thmName
            Case "FONT"
                Dim fntName As String
                fntName = IIf(Len(paramVal) > 0, paramVal, targetVal)
                If Len(fntName) > 0 Then Me.SetFont fntName
            Case "ZEBRA", "ZEBRASTRIPES"
                If fillClr >= 0 Then Me.ApplyZebraStripes fillClr Else Me.ApplyZebraStripes
            Case "BORDERS"
                If fillClr >= 0 Then Me.ApplyBorders fillClr Else Me.ApplyBorders
            Case "OUTERBORDER", "BOXBORDER"
                Dim boxClr As Long
                boxClr = IIf(fillClr >= 0, fillClr, IIf(fontClr >= 0, fontClr, 14277081))
                Dim boxWt As Long
                Dim bVal As String
                bVal = UCase$(Trim$(IIf(Len(targetVal) > 0 And Not IsNumeric(targetVal), targetVal, paramVal)))
                Select Case bVal
                    Case "THICK": boxWt = 4& ' xlThick
                    Case "MEDIUM": boxWt = -4138 ' xlMedium
                    Case "HAIRLINE": boxWt = 1& ' xlHairline
                    Case Else
                        If IsNumeric(bVal) Then boxWt = CLng(bVal) Else boxWt = 2& ' xlThin
                End Select
                ApplyBoxBorders mTableRange, boxWt, boxClr
            Case "FREEZE", "FREEZEPANES"
                Dim frzR As Long, frzC As Long
                frzR = IIf(IsNumeric(paramVal), CLng(paramVal), 1&)
                frzC = IIf(IsNumeric(opVal), CLng(opVal), 0&)
                Me.FreezePanes frzR, frzC
            Case "GRIDLINES"
                Me.SetGridlines ParseStyleBool(paramVal, True)
            Case "FORMAT", "NUMBERFORMAT"
                If Len(targetVal) > 0 And Len(paramVal) > 0 Then Me.FormatColumns targetVal, paramVal
            Case "ALIGN", "ALIGNMENT"
                If Len(targetVal) > 0 And Len(paramVal) > 0 Then Me.AlignColumn targetVal, ParseAlignCode(paramVal)
            Case "WIDTH", "COLWIDTH"
                If Len(targetVal) > 0 And IsNumeric(paramVal) Then Me.SetColumnWidth targetVal, CDbl(paramVal)
            Case "UNIFORMWIDTH"
                If Len(targetVal) > 0 And IsNumeric(paramVal) Then Me.SetUniformColumnWidth targetVal, CDbl(paramVal)
            Case "WRAP", "WRAPTEXT"
                If Len(targetVal) > 0 Then Me.SetWrapText targetVal, ParseStyleBool(paramVal, True)
            Case "HEADERHEIGHT"
                If IsNumeric(paramVal) Then Me.SetHeaderRowHeight CDbl(paramVal)
            Case "ROWHEIGHT", "DATAROWHEIGHT"
                If IsNumeric(paramVal) Then Me.SetDataRowHeight CDbl(paramVal)
            Case "HIDE", "HIDDEN", "HIDECOLUMN"
                If Len(targetVal) > 0 Then Me.HideColumn targetVal, True
            Case "SHOW", "UNHIDE"
                If Len(targetVal) > 0 Then Me.HideColumn targetVal, False
            Case "HIGHLIGHT", "HIGHLIGHTVALUES"
                If Len(targetVal) > 0 Then
                    Me.HighlightValues targetVal, paramVal, _
                        IIf(fillClr >= 0, fillClr, 13561798), _
                        IIf(fontClr >= 0, fontClr, 1875718)
                End If
            Case "HIGHLIGHTCOND", "HIGHLIGHTCONDITION"
                If Len(targetVal) > 0 And Len(opVal) > 0 Then
                    Me.HighlightCondition targetVal, opVal, paramVal, _
                        IIf(fillClr >= 0, fillClr, 13561798), _
                        IIf(fontClr >= 0, fontClr, 1875718)
                End If
            Case "HIGHLIGHTROW"
                If Len(targetVal) > 0 And Len(opVal) > 0 Then
                    Me.HighlightRowCondition targetVal, opVal, paramVal, _
                        IIf(fillClr >= 0, fillClr, 13561798), _
                        IIf(fontClr >= 0, fontClr, 1875718)
                End If
            Case "COLORSCALE"
                If Len(targetVal) > 0 Then
                    Me.ApplyColorScale targetVal, fillClr, fontClr
                End If
            Case "DATABAR", "DATABARS"
                If Len(targetVal) > 0 Then
                    Me.ApplyDataBars targetVal, IIf(fillClr >= 0, fillClr, RGB(99, 102, 241))
                End If
            Case "SPOTLIGHT"
                If Len(targetVal) > 0 Then
                    Me.SpotlightColumn targetVal, _
                        IIf(fillClr >= 0, fillClr, 16772566), _
                        IIf(fontClr >= 0, fontClr, 9453854)
                End If
            Case "DIM", "DIMCOLUMNS"
                If Len(targetVal) > 0 And Len(opVal) > 0 Then
                    Me.DimColumns targetVal, opVal, IIf(Len(paramVal) > 0, paramVal, "-"), _
                        IIf(fillClr >= 0, fillClr, 15788258), _
                        IIf(fontClr >= 0, fontClr, 12100498)
                End If
            Case "WEEKLYGRID", "WEEKLYCALENDARGRID"
                Dim daysInM As Long, maxDayC As Long
                daysInM = IIf(IsNumeric(paramVal), CLng(paramVal), 31&)
                maxDayC = IIf(IsNumeric(opVal), CLng(opVal), 31&)
                If Len(targetVal) > 0 Then
                    Me.ApplyWeeklyCalendarGrid targetVal, daysInM, maxDayC
                End If
            Case "GROUPGRID", "COLUMNGROUPGRID"
                If Len(targetVal) > 0 And Len(opVal) > 0 Then
                    Dim gSpec(0 To 0) As Variant
                    gSpec(0) = Array(paramVal, targetVal, opVal)
                    Me.ApplyColumnGroupGrid gSpec, -1, _
                        IIf(fontClr >= 0, fontClr, 6903111), _
                        IIf(fillClr >= 0, fillClr, 16381425), True
                End If
            Case "GROUPBANNER", "SUPERHEADER"
                If Len(targetVal) > 0 And Len(opVal) > 0 Then
                    Dim bSpec(0 To 0) As Variant
                    bSpec(0) = Array(paramVal, targetVal, opVal, IIf(fillClr >= 0, fillClr, RGB(30, 41, 59)))
                    Me.ApplyColumnGroupBanner bSpec, -1, 20#, IIf(fontClr >= 0, fontClr, 16777215)
                End If
            Case "TOTALROW", "ADDTOTALROW"
                If Len(targetVal) > 0 Then
                    Me.AddTotalRow targetVal, IIf(Len(paramVal) > 0, paramVal, "Total")
                Else
                    Me.AddTotalRow
                End If
            Case "AUTOFIT"
                Me.AutoFitColumns IIf(IsNumeric(paramVal), CDbl(paramVal), 3#)
        End Select
NextStyleRow:
    Next r

    Set ApplyStyleFromTable = Me
End Function

' ------------------------------------------------------------------------------
' Internal Helpers
' ------------------------------------------------------------------------------

Private Function ResolveStyleConfigDf(ByVal Source As Variant) As cDataFrame
    Dim df As New cDataFrame

    If VarType(Source) = vbString Then
        Dim strName As String
        strName = Trim$(CStr(Source))
        If Len(strName) = 0 Then
            Err.Raise 5, "cTableStyler.ApplyStyleFromTable", "TableName string cannot be blank."
        End If
        If Not mWorksheet Is Nothing Then
            df.LoadFromTableName mWorksheet, strName
        Else
            df.LoadFromTableName Application.ActiveWorkbook, strName
        End If
        Set ResolveStyleConfigDf = df
        Exit Function
    End If

    If Not IsObject(Source) Then
        Err.Raise 5, "cTableStyler.ApplyStyleFromTable", "StyleTableOrDf must be a Table Name string, cDataFrame, ListObject, or Range."
    End If
    If Source Is Nothing Then
        Err.Raise 91, "cTableStyler.ApplyStyleFromTable", "StyleTableOrDf cannot be Nothing."
    End If

    If TypeOf Source Is cDataFrame Then
        Set ResolveStyleConfigDf = Source
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
    Set ResolveStyleConfigDf = df
End Function

Private Function FindStyleColName(ByVal Df As cDataFrame, ByVal Candidates As Variant) As String
    Dim names() As String
    names = Df.ColumnNames()
    Dim i As Long, j As Long
    For j = LBound(Candidates) To UBound(Candidates)
        For i = LBound(names) To UBound(names)
            If StrComp(Trim$(names(i)), CStr(Candidates(j)), vbTextCompare) = 0 Then
                FindStyleColName = names(i)
                Exit Function
            End If
        Next i
    Next j
End Function

Private Function ParseStyleColor(ByVal ColorVal As Variant, ByVal DefaultColor As Long) As Long
    If IsEmpty(ColorVal) Or IsNull(ColorVal) Then
        ParseStyleColor = DefaultColor
        Exit Function
    End If
    Dim s As String
    s = Trim$(CStr(ColorVal))
    If Len(s) = 0 Then
        ParseStyleColor = DefaultColor
        Exit Function
    End If
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    If Len(s) = 6 And Not IsNumeric(s) Then
        On Error Resume Next
        Dim rHex As Long, gHex As Long, bHex As Long
        rHex = CLng("&H" & Left$(s, 2))
        gHex = CLng("&H" & Mid$(s, 3, 2))
        bHex = CLng("&H" & Right$(s, 2))
        If Err.Number = 0 Then
            ParseStyleColor = RGB(rHex, gHex, bHex)
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    If IsNumeric(s) Then
        ParseStyleColor = CLng(s)
    Else
        ParseStyleColor = DefaultColor
    End If
End Function

Private Function ParseAlignCode(ByVal AlignVal As String) As Long
    Select Case UCase$(Trim$(AlignVal))
        Case "LEFT", "L": ParseAlignCode = -4131
        Case "CENTER", "CENTRE", "C", "MID": ParseAlignCode = -4108
        Case "RIGHT", "R": ParseAlignCode = -4152
        Case Else
            If IsNumeric(AlignVal) Then ParseAlignCode = CLng(AlignVal) Else ParseAlignCode = -4131
    End Select
End Function

Private Function ParseStyleBool(ByVal Val As Variant, Optional ByVal DefaultVal As Boolean = True) As Boolean
    If IsEmpty(Val) Or IsNull(Val) Then
        ParseStyleBool = DefaultVal
        Exit Function
    End If
    If VarType(Val) = vbBoolean Then
        ParseStyleBool = CBool(Val)
        Exit Function
    End If
    Dim s As String
    s = UCase$(Trim$(CStr(Val)))
    If Len(s) = 0 Then
        ParseStyleBool = DefaultVal
    ElseIf s = "TRUE" Or s = "YES" Or s = "Y" Or s = "1" Or s = "ON" Then
        ParseStyleBool = True
    ElseIf s = "FALSE" Or s = "NO" Or s = "N" Or s = "0" Or s = "OFF" Then
        ParseStyleBool = False
    Else
        ParseStyleBool = DefaultVal
    End If
End Function

Private Function PureDataRowCount() As Long
    If mTotalRowHighlighted And mRowCount > 0 Then
        PureDataRowCount = mRowCount - 1
    Else
        PureDataRowCount = mRowCount
    End If
End Function

Private Sub ApplyBoxBorders(ByVal Rng As Object, ByVal Weight As Long, ByVal BorderColor As Long)
    Dim b As Variant
    For Each b In Array(7, 8, 9, 10) ' Left, Top, Bottom, Right
        With Rng.Borders(CLng(b))
            .LineStyle = 1
            .Weight = Weight
            .Color = BorderColor
        End With
    Next b
End Sub

Private Function ResolveColIndex(ByVal ColumnKey As Variant) As Long
    If VarType(ColumnKey) = vbString Then
        If mColIndexMap.Exists(ColumnKey) Then
            ResolveColIndex = mColIndexMap.Item(ColumnKey)
        Else
            Err.Raise 5, "cTableStyler", "Column name '" & ColumnKey & "' not found in table."
        End If
    ElseIf IsNumeric(ColumnKey) Then
        ResolveColIndex = CLng(ColumnKey)
    Else
        If mColIndexMap.Exists(CStr(ColumnKey)) Then
            ResolveColIndex = mColIndexMap.Item(CStr(ColumnKey))
        Else
            Err.Raise 5, "cTableStyler", "Column '" & CStr(ColumnKey) & "' not found in table."
        End If
    End If
End Function

Private Function ResolveColIndices(ByVal ColumnKeyOrPattern As Variant) As Long()
    Dim result() As Long
    Dim count As Long
    count = 0

    If IsArray(ColumnKeyOrPattern) Then
        Dim i As Long
        For i = LBound(ColumnKeyOrPattern) To UBound(ColumnKeyOrPattern)
            Dim subIndices() As Long
            subIndices = ResolveColIndices(ColumnKeyOrPattern(i))
            Dim s As Long
            For s = LBound(subIndices) To UBound(subIndices)
                count = count + 1
                ReDim Preserve result(1 To count)
                result(count) = subIndices(s)
            Next s
        Next i
    ElseIf VarType(ColumnKeyOrPattern) = vbString Then
        Dim patternStr As String
        patternStr = CStr(ColumnKeyOrPattern)
        If InStr(patternStr, ",") > 0 Then
            Dim parts() As String
            parts = Split(patternStr, ",")
            Dim p As Long
            For p = LBound(parts) To UBound(parts)
                Dim pIndices() As Long
                pIndices = ResolveColIndices(Trim$(parts(p)))
                For s = LBound(pIndices) To UBound(pIndices)
                    count = count + 1
                    ReDim Preserve result(1 To count)
                    result(count) = pIndices(s)
                Next s
            Next p
        ElseIf InStr(patternStr, "*") > 0 Or InStr(patternStr, "?") > 0 Then
            ' Wildcard pattern matching against all column names in mColIndexMap
            Dim colKey As Variant
            For Each colKey In mColIndexMap.Keys
                If CStr(colKey) Like patternStr Then
                    count = count + 1
                    ReDim Preserve result(1 To count)
                    result(count) = mColIndexMap.Item(colKey)
                End If
            Next colKey
        Else
            ' Exact column name
            count = 1
            ReDim result(1 To 1)
            result(1) = ResolveColIndex(patternStr)
        End If
    ElseIf IsNumeric(ColumnKeyOrPattern) Then
        count = 1
        ReDim result(1 To 1)
        result(1) = CLng(ColumnKeyOrPattern)
    End If

    If count = 0 Then
        ReDim result(0 To -1) ' Empty array
    End If
    ResolveColIndices = result
End Function
