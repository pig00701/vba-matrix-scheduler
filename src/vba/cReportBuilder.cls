VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cReportBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' cReportBuilder: Executive Dashboard & Multi-Table Report Composer
' ==============================================================================
' Provides fluent, automatic flowing layout for building multi-section Excel
' dashboards and executive reports. Coordinates titles, KPI metric cards,
' data tables, total rows, and explanatory footnotes without manual cell coordinate math.
'
' Reliability rules (Phase 0-1 hardening):
'   - Every text written by the builder is sanitized (no formula injection) and
'     stored as literal text (no "00123" -> 123, "34.2%" -> 0.342, "1-2" -> date).
'   - Every Add* call validates arguments and worksheet limits BEFORE touching the
'     sheet, so a rejected call never leaves a half-written section behind.
'   - Unknown theme names raise an error instead of silently falling back.
'   - BeginUpdate/EndUpdate (and Class_Terminate) restore ScreenUpdating,
'     Calculation, EnableEvents and DisplayAlerts even if the caller crashes.
'   - Non-fatal findings (e.g. truncated text, skipped total row) are collected
'     in the Warnings collection instead of being lost.
' ==============================================================================

' Excel constants (late-bound, so no Excel type-library reference is required)
Private Const XL_CENTER As Long = -4108
Private Const XL_LEFT As Long = -4131
Private Const XL_EDGE_LEFT As Long = 7
Private Const XL_EDGE_TOP As Long = 8
Private Const XL_EDGE_BOTTOM As Long = 9
Private Const XL_EDGE_RIGHT As Long = 10
Private Const XL_CONTINUOUS As Long = 1
Private Const XL_THIN As Long = 2
Private Const XL_CALC_MANUAL As Long = -4135
Private Const XL_TEXT_FORMAT As String = "@"

' Worksheet hard limits
Private Const MAX_SHEET_ROWS As Long = 1048576
Private Const MAX_SHEET_COLS As Long = 16384
Private Const MAX_CELL_CHARS As Long = 32767
Private Const MAX_COL_WIDTH As Double = 255#

Private Const SRC As String = "cReportBuilder"

Private mWorksheet As Object
Private mStartRow As Long
Private mStartCol As Long
Private mCurrentRow As Long
Private mMaxCol As Long
Private mInitialTopLeft As Object
Private mWarnings As Collection
Private mLastTableStyler As cTableStyler
Private mLastTableStartRow As Long
Private mLastTableEndRow As Long

' Rollback snapshot state
Private mSnapshotRow As Long
Private mSnapshotMaxCol As Long

' Column grouping state (side-by-side layout)
Private mInColumnGroup As Boolean
Private mGroupStartRow As Long
Private mGroupMaxRow As Long
Private mGroupAnchorCol As Long

' Environment guard state (RAII, restored in EndUpdate / Class_Terminate)
Private mInBatch As Boolean
Private mPrevScreenUpdating As Boolean
Private mPrevCalculation As Long
Private mPrevEnableEvents As Boolean
Private mPrevDisplayAlerts As Boolean

Private Sub Class_Initialize()
    mStartRow = 2
    mStartCol = 2
    mCurrentRow = 2
    mMaxCol = 2
    mSnapshotRow = 2
    mSnapshotMaxCol = 2
    mInColumnGroup = False
    mGroupStartRow = 2
    mGroupMaxRow = 2
    mGroupAnchorCol = 2
    mInBatch = False
    mPrevScreenUpdating = True
    mPrevCalculation = -4105 ' xlCalculationAutomatic
    mPrevEnableEvents = True
    mPrevDisplayAlerts = True
    mLastTableStartRow = 0
    mLastTableEndRow = 0
    Set mLastTableStyler = Nothing
    Set mWarnings = New Collection
End Sub

Private Sub Class_Terminate()
    If mInBatch Then
        EndUpdate
    End If
    Set mLastTableStyler = Nothing
    Set mWorksheet = Nothing
    Set mInitialTopLeft = Nothing
    Set mWarnings = Nothing
End Sub

' ------------------------------------------------------------------------------
' Initialization & Flow Positioning
' ------------------------------------------------------------------------------

Public Function Attach(ByVal TargetWorksheet As Object, _
                       Optional ByVal StartCellAddress As String = "B2") As cReportBuilder
    If TargetWorksheet Is Nothing Then
        Err.Raise 91, SRC & ".Attach", "TargetWorksheet cannot be Nothing."
    End If

    ' Re-attaching: restore the previous application state first
    If mInBatch Then EndUpdate

    If TargetWorksheet.ProtectContents Then
        Err.Raise 1004, SRC & ".Attach", _
            "Cannot build report: Worksheet '" & TargetWorksheet.Name & "' is protected."
    End If

    Dim anchor As Object
    On Error Resume Next
    Set anchor = TargetWorksheet.Range(StartCellAddress).Cells(1, 1)
    On Error GoTo 0
    If anchor Is Nothing Then
        Err.Raise 5, SRC & ".Attach", "Invalid StartCellAddress: '" & StartCellAddress & "'."
    End If

    Set mWorksheet = TargetWorksheet
    Set mInitialTopLeft = anchor
    mStartRow = anchor.Row
    mStartCol = anchor.Column
    mCurrentRow = mStartRow
    mMaxCol = mStartCol
    mSnapshotRow = mStartRow
    mSnapshotMaxCol = mStartCol
    mInColumnGroup = False
    mGroupStartRow = mStartRow
    mGroupMaxRow = mStartRow
    mGroupAnchorCol = mStartCol
    mLastTableStartRow = 0
    mLastTableEndRow = 0
    Set mLastTableStyler = Nothing
    Set mWarnings = New Collection

    Set Attach = Me
End Function

' ------------------------------------------------------------------------------
' Environment Guard & Rollback (crash-safe)
' ------------------------------------------------------------------------------

' Suppresses screen updates, events and automatic calculation while the report is built.
' Also snapshots the current row/col position for atomic Rollback support.
' Class_Terminate automatically calls EndUpdate, so Excel is never left in a frozen state.
Public Function BeginUpdate() As cReportBuilder
    CheckAttached "BeginUpdate"
    If Not mInBatch Then
        mSnapshotRow = mCurrentRow
        mSnapshotMaxCol = mMaxCol

        Dim appObj As Object
        On Error Resume Next
        Set appObj = mWorksheet.Application
        mPrevScreenUpdating = appObj.ScreenUpdating
        mPrevCalculation = appObj.Calculation
        mPrevEnableEvents = appObj.EnableEvents
        mPrevDisplayAlerts = appObj.DisplayAlerts

        appObj.ScreenUpdating = False
        appObj.Calculation = XL_CALC_MANUAL
        appObj.EnableEvents = False
        mInBatch = True
        On Error GoTo 0
    End If
    Set BeginUpdate = Me
End Function

' Restores the application state captured by BeginUpdate.
Public Function EndUpdate() As cReportBuilder
    If mInBatch And Not mWorksheet Is Nothing Then
        Dim appObj As Object
        On Error Resume Next
        Set appObj = mWorksheet.Application
        appObj.Calculation = mPrevCalculation
        appObj.EnableEvents = mPrevEnableEvents
        appObj.DisplayAlerts = mPrevDisplayAlerts
        appObj.ScreenUpdating = mPrevScreenUpdating
        mInBatch = False
        On Error GoTo 0
    End If
    Set EndUpdate = Me
End Function

' Discards all cells written since BeginUpdate (or Attach), clears formats/contents,
' resets CurrentRow back to snapshot, and safely restores application state.
Public Function Rollback() As cReportBuilder
    CheckAttached "Rollback"
    If mSnapshotRow > 0 And Not mWorksheet Is Nothing Then
        Dim lastR As Long, lastC As Long
        lastR = mCurrentRow
        lastC = mMaxCol
        If lastR >= mSnapshotRow And lastC >= mStartCol Then
            Dim dirtyRng As Object
            On Error Resume Next
            Set dirtyRng = mWorksheet.Range(mWorksheet.Cells(mSnapshotRow, mStartCol), mWorksheet.Cells(lastR, lastC))
            If Not dirtyRng Is Nothing Then
                dirtyRng.UnMerge
                dirtyRng.Clear
            End If
            On Error GoTo 0
        End If
        mCurrentRow = mSnapshotRow
        mMaxCol = mSnapshotMaxCol
    End If
    If mInBatch Then EndUpdate
    AddWarning "Rollback: changes discarded and worksheet restored to snapshot."
    Set Rollback = Me
End Function

' ------------------------------------------------------------------------------
' Section Builders
' ------------------------------------------------------------------------------

' Adds an Executive Title Banner spanning WidthColumns
Public Function AddBanner(ByVal Title As String, _
                          Optional ByVal Subtitle As String = "", _
                          Optional ByVal Theme As String = "Navy", _
                          Optional ByVal WidthColumns As Long = 0, _
                          Optional ByVal CenterTitle As Boolean = False, _
                          Optional ByVal UseMerge As Boolean = False) As cReportBuilder
    CheckAttached "AddBanner"
    If WidthColumns <= 0 Then
        If mMaxCol > mStartCol Then
            WidthColumns = mMaxCol - mStartCol + 1
        Else
            WidthColumns = 8
        End If
    End If

    Dim bannerBg As Long, bannerFg As Long, subBg As Long, subFg As Long
    ResolveBannerTheme Theme, bannerBg, bannerFg, subBg, subFg

    Dim hasSubtitle As Boolean
    hasSubtitle = (Len(Trim$(Subtitle)) > 0)

    Dim endCol As Long
    endCol = mStartCol + WidthColumns - 1
    RequireColumns endCol, "AddBanner"
    RequireRows mCurrentRow + IIf(hasSubtitle, 1, 0), "AddBanner"

    If endCol > mMaxCol Then mMaxCol = endCol

    ' Title Row
    Dim titleRng As Object
    Set titleRng = mWorksheet.Range(mWorksheet.Cells(mCurrentRow, mStartCol), mWorksheet.Cells(mCurrentRow, endCol))
    If UseMerge Then
        titleRng.Merge
        titleRng.HorizontalAlignment = IIf(CenterTitle, XL_CENTER, XL_LEFT)
    Else
        titleRng.HorizontalAlignment = IIf(CenterTitle, 7, XL_LEFT) ' 7 = xlCenterAcrossSelection
    End If
    PutValue titleRng.Cells(1, 1), Title
    titleRng.Font.Name = "Segoe UI"
    titleRng.Font.Size = 16
    titleRng.Font.Bold = True
    titleRng.Font.Color = bannerFg
    titleRng.Interior.Color = bannerBg
    titleRng.VerticalAlignment = XL_CENTER
    mWorksheet.Rows(mCurrentRow).RowHeight = 32

    mCurrentRow = mCurrentRow + 1

    ' Optional Subtitle Row
    If hasSubtitle Then
        Dim subRng As Object
        Set subRng = mWorksheet.Range(mWorksheet.Cells(mCurrentRow, mStartCol), mWorksheet.Cells(mCurrentRow, endCol))
        If UseMerge Then
            subRng.Merge
            subRng.HorizontalAlignment = IIf(CenterTitle, XL_CENTER, XL_LEFT)
        Else
            subRng.HorizontalAlignment = IIf(CenterTitle, 7, XL_LEFT) ' 7 = xlCenterAcrossSelection
        End If
        PutValue subRng.Cells(1, 1), Subtitle
        subRng.Font.Name = "Segoe UI"
        subRng.Font.Size = 9.5
        subRng.Font.Italic = True
        subRng.Font.Color = subFg
        subRng.Interior.Color = subBg
        subRng.VerticalAlignment = XL_CENTER
        mWorksheet.Rows(mCurrentRow).RowHeight = 20
        mCurrentRow = mCurrentRow + 1
    End If

    mCurrentRow = mCurrentRow + 1 ' spacing gap
    Set AddBanner = Me
End Function

' Adds a row of horizontal KPI metric summary cards side-by-side.
' Titles / Values / Subtexts accept an array (or a single scalar for one card) and must have equal length.
' Numeric Values are written as numbers; text Values are written as literal text.
' NumberFormats optionally sets formatting per card (e.g. "$#,##0", "0.0%").
' UseMerge:=False uses CenterAcrossSelection to keep sheet cells unmerged and sortable.
' A Subtext starting with "+" or an up-triangle is green; "-" or a down-triangle is red.
Public Function AddKpiRow(ByVal Titles As Variant, _
                          ByVal Values As Variant, _
                          Optional ByVal Subtexts As Variant, _
                          Optional ByVal CardWidthCols As Long = 2, _
                          Optional ByVal NumberFormats As Variant, _
                          Optional ByVal UseMerge As Boolean = False, _
                          Optional ByVal IsFormula As Variant) As cReportBuilder
    CheckAttached "AddKpiRow"
    If CardWidthCols < 1 Then
        Err.Raise 5, SRC & ".AddKpiRow", "CardWidthCols must be >= 1 (got " & CardWidthCols & ")."
    End If

    Dim titleArr As Variant, valueArr As Variant, subArr As Variant, fmtArr As Variant, formArr As Variant
    titleArr = AsArray(Titles)
    valueArr = AsArray(Values)

    Dim numCards As Long
    numCards = ArrayCount(titleArr)
    If numCards < 1 Then
        Err.Raise 5, SRC & ".AddKpiRow", "Titles must contain at least one KPI card."
    End If
    If ArrayCount(valueArr) <> numCards Then
        Err.Raise 5, SRC & ".AddKpiRow", _
            "Values count (" & ArrayCount(valueArr) & ") must match Titles count (" & numCards & ")."
    End If

    Dim hasSub As Boolean
    hasSub = False
    If Not IsMissing(Subtexts) Then
        If Not IsEmpty(Subtexts) Then
            subArr = AsArray(Subtexts)
            If ArrayCount(subArr) <> numCards Then
                Err.Raise 5, SRC & ".AddKpiRow", _
                    "Subtexts count (" & ArrayCount(subArr) & ") must match Titles count (" & numCards & ")."
            End If
            hasSub = True
        End If
    End If

    Dim hasNumFmt As Boolean
    hasNumFmt = False
    If Not IsMissing(NumberFormats) Then
        If Not IsEmpty(NumberFormats) Then
            fmtArr = AsArray(NumberFormats)
            hasNumFmt = True
        End If
    End If

    Dim hasFormArr As Boolean
    hasFormArr = False
    If Not IsMissing(IsFormula) Then
        If Not IsEmpty(IsFormula) Then
            formArr = AsArray(IsFormula)
            hasFormArr = True
        End If
    End If

    Dim cardTop As Long, cardBottom As Long, lastCol As Long
    cardTop = mCurrentRow
    cardBottom = cardTop + IIf(hasSub, 2, 1)
    lastCol = mStartCol + (numCards * CardWidthCols) - 1
    RequireRows cardBottom, "AddKpiRow"
    RequireColumns lastCol, "AddKpiRow"

    If lastCol > mMaxCol Then mMaxCol = lastCol

    Dim curC As Long, cardEndC As Long, k As Long, b As Long
    Dim lbT As Long, lbV As Long, lbS As Long, lbF As Long, lbForm As Long
    lbT = LBound(titleArr)
    lbV = LBound(valueArr)
    If hasSub Then lbS = LBound(subArr)
    If hasNumFmt Then lbF = LBound(fmtArr)
    If hasFormArr Then lbForm = LBound(formArr)

    Dim tRng As Object, vRng As Object, sRng As Object, fullCardRng As Object
    Dim subText As String, fmtStr As String, cardIsForm As Boolean, fStr As String

    curC = mStartCol
    For k = 0 To numCards - 1
        cardEndC = curC + CardWidthCols - 1

        ' Card Title (Row 1)
        Set tRng = mWorksheet.Range(mWorksheet.Cells(cardTop, curC), mWorksheet.Cells(cardTop, cardEndC))
        If UseMerge Then
            tRng.Merge
            tRng.HorizontalAlignment = XL_CENTER
        Else
            If CardWidthCols > 1 Then tRng.HorizontalAlignment = 7 Else tRng.HorizontalAlignment = XL_CENTER
        End If
        PutValue tRng.Cells(1, 1), UCase$(ToPlainText(titleArr(lbT + k)))
        tRng.Font.Name = "Segoe UI"
        tRng.Font.Size = 8.5
        tRng.Font.Bold = True
        tRng.Font.Color = RGB(100, 116, 139) ' Slate 500
        tRng.VerticalAlignment = XL_CENTER

        ' Card Value (Row 2)
        Set vRng = mWorksheet.Range(mWorksheet.Cells(cardTop + 1, curC), mWorksheet.Cells(cardTop + 1, cardEndC))
        If UseMerge Then
            vRng.Merge
            vRng.HorizontalAlignment = XL_CENTER
        Else
            If CardWidthCols > 1 Then vRng.HorizontalAlignment = 7 Else vRng.HorizontalAlignment = XL_CENTER
        End If

        If hasNumFmt Then
            If k <= (UBound(fmtArr) - lbF) Then
                fmtStr = CStr(fmtArr(lbF + k))
                If Len(Trim$(fmtStr)) > 0 Then vRng.NumberFormat = fmtStr
            End If
        End If

        cardIsForm = False
        If hasFormArr Then
            If k <= (UBound(formArr) - lbForm) Then
                cardIsForm = CBool(formArr(lbForm + k))
            End If
        End If

        If cardIsForm Then
            fStr = Trim$(CStr(valueArr(lbV + k)))
            If Left$(fStr, 1) <> "=" Then fStr = "=" & fStr
            vRng.Cells(1, 1).Formula = fStr
        Else
            PutValue vRng.Cells(1, 1), valueArr(lbV + k)
        End If
        vRng.Font.Name = "Segoe UI"
        vRng.Font.Size = 16
        vRng.Font.Bold = True
        vRng.Font.Color = RGB(15, 23, 42)    ' Slate 900
        vRng.VerticalAlignment = XL_CENTER

        ' Card Subtext (Row 3, if present)
        If hasSub Then
            subText = ToPlainText(subArr(lbS + k))
            Set sRng = mWorksheet.Range(mWorksheet.Cells(cardTop + 2, curC), mWorksheet.Cells(cardTop + 2, cardEndC))
            If UseMerge Then
                sRng.Merge
                sRng.HorizontalAlignment = XL_CENTER
            Else
                If CardWidthCols > 1 Then sRng.HorizontalAlignment = 7 Else sRng.HorizontalAlignment = XL_CENTER
            End If
            PutValue sRng.Cells(1, 1), subText
            sRng.Font.Name = "Segoe UI"
            sRng.Font.Size = 8.5
            sRng.Font.Bold = True
            Select Case TrendDirection(subText)
                Case 1
                    sRng.Font.Color = RGB(22, 101, 52)   ' Green
                Case -1
                    sRng.Font.Color = RGB(185, 28, 28)   ' Red
                Case Else
                    sRng.Font.Color = RGB(100, 116, 139) ' Neutral
            End Select
            sRng.VerticalAlignment = XL_CENTER
        End If

        ' Apply Card Box Border and Subtle Fill
        Set fullCardRng = mWorksheet.Range(mWorksheet.Cells(cardTop, curC), mWorksheet.Cells(cardBottom, cardEndC))
        fullCardRng.Interior.Color = RGB(248, 250, 252) ' Light Gray Fill

        For b = XL_EDGE_LEFT To XL_EDGE_RIGHT
            fullCardRng.Borders(b).LineStyle = XL_CONTINUOUS
            fullCardRng.Borders(b).Weight = XL_THIN
            fullCardRng.Borders(b).Color = RGB(203, 213, 225) ' Border color
        Next b

        curC = cardEndC + 1 ' Next card adjacent
    Next k

    mWorksheet.Rows(cardTop).RowHeight = 18
    mWorksheet.Rows(cardTop + 1).RowHeight = 28
    If hasSub Then mWorksheet.Rows(cardTop + 2).RowHeight = 18

    mCurrentRow = cardBottom + 2 ' Gap below cards
    Set AddKpiRow = Me
End Function

' Adds a formatted cDataFrame table with optional section header and auto-total row.
' AddTotalRow sums every purely numeric column (SUBTOTAL(9), filter-aware) and labels the
' first non-numeric column. If there is nothing to total, the row is skipped and a warning is recorded.
Public Function AddTable(ByVal Df As cDataFrame, _
                         Optional ByVal Title As String = "", _
                         Optional ByVal StyleTheme As String = "Navy", _
                         Optional ByVal AddTotalRow As Boolean = False) As cReportBuilder
    CheckAttached "AddTable"
    If Df Is Nothing Then
        Err.Raise 91, SRC & ".AddTable", "Df cannot be Nothing."
    End If
    If Df.ColumnCount < 1 Then
        Err.Raise 5, SRC & ".AddTable", "Df has no columns; nothing to render."
    End If

    Dim stylerTheme As String
    stylerTheme = ResolveTableTheme(StyleTheme)

    ' Decide the total row up-front so the layout can be validated before any write
    Dim numericCols As Variant
    Dim withTotal As Boolean
    withTotal = False
    If AddTotalRow Then
        If Df.RowCount = 0 Then
            AddWarning "AddTable: total row skipped because the table has no data rows."
        Else
            numericCols = NumericColumnIndexes(Df)
            If IsEmpty(numericCols) Then
                AddWarning "AddTable: total row skipped because the table has no numeric column."
            Else
                withTotal = True
            End If
        End If
    End If

    Dim hasTitle As Boolean
    hasTitle = (Len(Trim$(Title)) > 0)

    Dim endCol As Long, headerRow As Long, lastRow As Long
    endCol = mStartCol + Df.ColumnCount - 1
    headerRow = mCurrentRow + IIf(hasTitle, 1, 0)
    lastRow = headerRow + Df.RowCount + IIf(withTotal, 1, 0)
    RequireColumns endCol, "AddTable"
    RequireRows lastRow, "AddTable"

    If endCol > mMaxCol Then mMaxCol = endCol

    ' Section Header Title
    If hasTitle Then
        Dim titleCell As Object
        Set titleCell = mWorksheet.Cells(mCurrentRow, mStartCol)
        PutValue titleCell, Title
        titleCell.Font.Name = "Segoe UI"
        titleCell.Font.Size = 12
        titleCell.Font.Bold = True
        titleCell.Font.Color = RGB(30, 41, 59)
        mWorksheet.Rows(mCurrentRow).RowHeight = 22
        mCurrentRow = mCurrentRow + 1
    End If

    ' Header names are stored as literal text so a column called "=1+1" cannot become a formula
    Dim startCell As Object
    Set startCell = mWorksheet.Cells(mCurrentRow, mStartCol)
    mWorksheet.Range(startCell, mWorksheet.Cells(mCurrentRow, endCol)).NumberFormat = XL_TEXT_FORMAT

    ' Dump DataFrame to Worksheet
    Df.ToRange startCell, IncludeHeaders:=True

    ' Style using cTableStyler
    Dim styler As cTableStyler
    Set styler = New cTableStyler
    styler.Attach startCell, Df, HasHeaders:=True
    styler.ApplyTheme stylerTheme
    styler.ApplyZebraStripes

    If withTotal Then
        Dim labelCol As Variant
        Dim labelIdx As Long
        labelIdx = FirstNonNumericIndex(numericCols, Df.ColumnCount)
        If labelIdx > 0 Then labelCol = labelIdx Else labelCol = Empty
        styler.AddTotalRow SumColumns:=numericCols, LabelColumn:=labelCol, UseSubtotal:=True
    End If

    ' Advance current row (table + one blank gap row)
    Set mLastTableStyler = styler
    mLastTableStartRow = headerRow
    mLastTableEndRow = lastRow
    mCurrentRow = lastRow + 2

    Set AddTable = Me
End Function

' Renders a clean data contract / quality assurance summary badge from a cDataValidator instance.
' Highlights pass (emerald green) or fail (soft red alert) with row/error counts,
' and optionally renders a full error breakdown table if rules were violated.
Public Function AddValidationSummary(ByVal Validator As cDataValidator, _
                                     Optional ByVal WidthColumns As Long = 0, _
                                     Optional ByVal ShowErrorsTable As Boolean = False) As cReportBuilder
    CheckAttached "AddValidationSummary"
    If Validator Is Nothing Then
        Err.Raise 91, SRC & ".AddValidationSummary", "Validator cannot be Nothing."
    End If

    If WidthColumns <= 0 Then
        If mMaxCol > mStartCol Then
            WidthColumns = mMaxCol - mStartCol + 1
        Else
            WidthColumns = 8
        End If
    End If

    Dim endCol As Long
    endCol = mStartCol + WidthColumns - 1
    RequireColumns endCol, "AddValidationSummary"
    RequireRows mCurrentRow + 1, "AddValidationSummary"

    If endCol > mMaxCol Then mMaxCol = endCol

    Dim boxRng As Object
    Set boxRng = mWorksheet.Range(mWorksheet.Cells(mCurrentRow, mStartCol), mWorksheet.Cells(mCurrentRow, endCol))

    Dim summaryMsg As String, bgColor As Long, borderColor As Long, fontColor As Long
    If Validator.IsValid Then
        summaryMsg = "DATA CONTRACT PASS: All " & Validator.TotalRowsChecked & _
                     " rows verified against schema. 0 errors detected."
        bgColor = RGB(240, 253, 244)      ' Emerald 50
        borderColor = RGB(34, 197, 94)    ' Emerald 500
        fontColor = RGB(22, 101, 52)      ' Emerald 800
    Else
        summaryMsg = "DATA CONTRACT ALERT: " & Validator.FailedRowsCount & " of " & _
                     Validator.TotalRowsChecked & " rows failed validation (" & _
                     Validator.ErrorCount & " errors detected)."
        bgColor = RGB(254, 242, 242)      ' Red 50
        borderColor = RGB(239, 68, 68)    ' Red 500
        fontColor = RGB(153, 27, 27)      ' Red 800
        AddWarning summaryMsg
    End If

    boxRng.HorizontalAlignment = 7 ' xlCenterAcrossSelection
    PutValue boxRng.Cells(1, 1), summaryMsg
    boxRng.Font.Name = "Segoe UI"
    boxRng.Font.Size = 9.5
    boxRng.Font.Bold = True
    boxRng.Font.Color = fontColor
    boxRng.Interior.Color = bgColor
    boxRng.VerticalAlignment = XL_CENTER

    Dim b As Long
    For b = XL_EDGE_LEFT To XL_EDGE_RIGHT
        boxRng.Borders(b).LineStyle = XL_CONTINUOUS
        boxRng.Borders(b).Weight = XL_THIN
        boxRng.Borders(b).Color = borderColor
    Next b
    mWorksheet.Rows(mCurrentRow).RowHeight = 24

    mCurrentRow = mCurrentRow + 2 ' Gap

    If ShowErrorsTable And Not Validator.IsValid Then
        If Not Validator.Errors Is Nothing Then
            If Validator.Errors.RowCount > 0 Then
                AddTable Validator.Errors, Title:="Validation Error Details", StyleTheme:="Slate"
            End If
        End If
    End If

    Set AddValidationSummary = Me
End Function

' Renders a full-width dark/colored section bar in the flowing layout, with optional workbook named anchor.
Public Function AddSectionHeader(ByVal Title As String, _
                                 Optional ByVal Theme As String = "Slate", _
                                 Optional ByVal WidthColumns As Long = 0, _
                                 Optional ByVal AnchorName As String = "", _
                                 Optional ByVal CustomBgColor As Variant) As cReportBuilder
    CheckAttached "AddSectionHeader"

    Dim secBg As Long, secFg As Long, dummySubBg As Long, dummySubFg As Long
    ResolveBannerTheme Theme, secBg, secFg, dummySubBg, dummySubFg
    If Not IsMissing(CustomBgColor) And Not IsEmpty(CustomBgColor) Then
        If IsNumeric(CustomBgColor) Then secBg = CLng(CustomBgColor)
    End If

    If WidthColumns <= 0 Then
        If mMaxCol > mStartCol Then
            WidthColumns = mMaxCol - mStartCol + 1
        Else
            WidthColumns = 8
        End If
    End If

    Dim endCol As Long, secRow As Long
    endCol = mStartCol + WidthColumns - 1
    secRow = mCurrentRow
    RequireColumns endCol, "AddSectionHeader"
    RequireRows secRow, "AddSectionHeader"

    If endCol > mMaxCol Then mMaxCol = endCol

    Dim secRng As Object
    Set secRng = mWorksheet.Range(mWorksheet.Cells(secRow, mStartCol), mWorksheet.Cells(secRow, endCol))
    PutValue mWorksheet.Cells(secRow, mStartCol), Title
    secRng.Interior.Color = secBg
    secRng.Font.Name = "Segoe UI"
    secRng.Font.Size = 11
    secRng.Font.Bold = True
    secRng.Font.Color = secFg
    secRng.VerticalAlignment = XL_CENTER
    mWorksheet.Rows(secRow).RowHeight = 25

    If Len(Trim$(AnchorName)) > 0 Then
        On Error Resume Next
        mWorksheet.Parent.Names.Add Name:=Trim$(AnchorName), RefersTo:=mWorksheet.Cells(secRow, mStartCol)
        On Error GoTo 0
    End If

    mCurrentRow = secRow + 1
    Set AddSectionHeader = Me
End Function

' Renders clickable same-sheet or workbook navigation pills at mCurrentRow (or populates a reserved TargetRow).
Public Function AddQuickJumpBar(ByVal LabelText As String, _
                                ByVal PillLabels As Variant, _
                                ByVal SubAddresses As Variant, _
                                Optional ByVal ActiveIndex As Long = 0, _
                                Optional ByVal PillWidthCols As Long = 2, _
                                Optional ByVal TargetRow As Long = 0) As cReportBuilder
    CheckAttached "AddQuickJumpBar"
    If Not IsArray(PillLabels) Or Not IsArray(SubAddresses) Then
        Err.Raise 5, SRC & ".AddQuickJumpBar", "PillLabels and SubAddresses must be arrays."
    End If

    Dim lbL As Long, ubL As Long, lbS As Long, ubS As Long
    lbL = LBound(PillLabels): ubL = UBound(PillLabels)
    lbS = LBound(SubAddresses): ubS = UBound(SubAddresses)
    Dim nPills As Long
    nPills = ubL - lbL + 1
    If (ubS - lbS + 1) <> nPills Then
        Err.Raise 5, SRC & ".AddQuickJumpBar", "PillLabels and SubAddresses must have the same length."
    End If
    If PillWidthCols < 1 Then PillWidthCols = 1

    Dim barRow As Long
    barRow = IIf(TargetRow > 0, TargetRow, mCurrentRow)
    Dim pillStartCol As Long
    pillStartCol = mStartCol + 2
    Dim endCol As Long
    endCol = pillStartCol + (nPills * PillWidthCols) - 1
    RequireColumns endCol, "AddQuickJumpBar"
    RequireRows barRow, "AddQuickJumpBar"
    If endCol > mMaxCol Then mMaxCol = endCol

    mWorksheet.Rows(barRow).RowHeight = 22
    If Len(Trim$(LabelText)) > 0 Then
        PutValue mWorksheet.Cells(barRow, mStartCol), LabelText
        With mWorksheet.Cells(barRow, mStartCol)
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(30, 41, 59)
            .VerticalAlignment = XL_CENTER
        End With
    End If

    Dim k As Long, colIdx As Long, b As Variant
    For k = 0 To nPills - 1
        colIdx = pillStartCol + (k * PillWidthCols)
        Dim pillRng As Object
        Set pillRng = mWorksheet.Range(mWorksheet.Cells(barRow, colIdx), mWorksheet.Cells(barRow, colIdx + PillWidthCols - 1))
        mWorksheet.Cells(barRow, colIdx).NumberFormat = XL_TEXT_FORMAT

        mWorksheet.Hyperlinks.Add Anchor:=mWorksheet.Cells(barRow, colIdx), _
                                  Address:="", _
                                  SubAddress:=CStr(SubAddresses(lbS + k)), _
                                  TextToDisplay:=CStr(PillLabels(lbL + k))

        pillRng.HorizontalAlignment = IIf(PillWidthCols > 1, 7, XL_CENTER)
        pillRng.VerticalAlignment = XL_CENTER
        pillRng.Font.Name = "Segoe UI"
        pillRng.Font.Size = 9.5
        pillRng.Font.Bold = True

        If (k + 1) = ActiveIndex Then
            pillRng.Interior.Color = RGB(217, 119, 6)
            pillRng.Font.Color = RGB(255, 255, 255)
        Else
            pillRng.Interior.Color = RGB(226, 232, 240)
            pillRng.Font.Color = RGB(15, 23, 42)
        End If

        For Each b In Array(XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT)
            With pillRng.Borders(CLng(b))
                .LineStyle = XL_CONTINUOUS
                .Weight = XL_THIN
                .Color = RGB(148, 163, 184)
            End With
        Next b
    Next k

    If TargetRow <= 0 Then
        mCurrentRow = barRow + 2
    End If
    Set AddQuickJumpBar = Me
End Function

' Idempotently creates or updates an xlFreeFloating Form Button macro trigger on the report sheet.
Public Function AddActionButton(ByVal ButtonName As String, _
                                ByVal Caption As String, _
                                ByVal OnActionMacro As String, _
                                Optional ByVal TargetCol As Variant = 30, _
                                Optional ByVal TargetRow As Long = 0, _
                                Optional ByVal WidthPts As Double = 175#, _
                                Optional ByVal HeightPts As Double = 25#) As cReportBuilder
    CheckAttached "AddActionButton"
    Dim useRow As Long
    useRow = IIf(TargetRow > 0, TargetRow, mCurrentRow)
    RequireRows useRow, "AddActionButton"

    Dim btnLeft As Double, btnTop As Double
    btnLeft = mWorksheet.Columns(TargetCol).Left
    btnTop = mWorksheet.Rows(useRow).Top - 2#

    Dim btn As Object
    On Error Resume Next
    Set btn = mWorksheet.Buttons(ButtonName)
    On Error GoTo 0

    If btn Is Nothing Then
        Set btn = mWorksheet.Buttons.Add(btnLeft, btnTop, WidthPts, HeightPts)
        btn.Name = ButtonName
        btn.Placement = 3 ' xlFreeFloating
    Else
        btn.Left = btnLeft
        btn.Top = btnTop
        btn.Width = WidthPts
        btn.Height = HeightPts
    End If
    btn.Text = Caption
    btn.OnAction = OnActionMacro

    Set AddActionButton = Me
End Function

' Adds vertical blank spacing
Public Function AddSpace(Optional ByVal Rows As Long = 2) As cReportBuilder
    CheckAttached "AddSpace"
    If Rows < 0 Then
        Err.Raise 5, SRC & ".AddSpace", "Rows must be >= 0 (got " & Rows & ")."
    End If
    RequireRows mCurrentRow + Rows, "AddSpace"
    mCurrentRow = mCurrentRow + Rows
    Set AddSpace = Me
End Function

' Adds an explanatory footnote or remark
Public Function AddNote(ByVal NoteText As String, _
                        Optional ByVal IsItalic As Boolean = True) As cReportBuilder
    CheckAttached "AddNote"
    RequireRows mCurrentRow, "AddNote"
    Dim noteCell As Object
    Set noteCell = mWorksheet.Cells(mCurrentRow, mStartCol)
    PutValue noteCell, NoteText
    noteCell.Font.Name = "Segoe UI"
    noteCell.Font.Size = 9
    noteCell.Font.Italic = IsItalic
    noteCell.Font.Color = RGB(100, 116, 139)
    mCurrentRow = mCurrentRow + 1
    Set AddNote = Me
End Function

' Appends a standardized audit metadata footer with timestamp, author, source, and checksum.
Public Function AddProvenanceFooter(Optional ByVal Author As String = "", _
                                    Optional ByVal SourceDescription As String = "", _
                                    Optional ByVal Checksum As String = "", _
                                    Optional ByVal WidthColumns As Long = 0) As cReportBuilder
    CheckAttached "AddProvenanceFooter"
    If WidthColumns <= 0 Then
        If mMaxCol > mStartCol Then
            WidthColumns = mMaxCol - mStartCol + 1
        Else
            WidthColumns = 8
        End If
    End If

    Dim endCol As Long
    endCol = mStartCol + WidthColumns - 1
    RequireColumns endCol, "AddProvenanceFooter"
    RequireRows mCurrentRow, "AddProvenanceFooter"

    If endCol > mMaxCol Then mMaxCol = endCol

    Dim metaText As String
    metaText = "Generated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If Len(Trim$(Author)) > 0 Then metaText = metaText & " | Author: " & SanitizeText(Author)
    If Len(Trim$(SourceDescription)) > 0 Then metaText = metaText & " | Source: " & SanitizeText(SourceDescription)
    If Len(Trim$(Checksum)) > 0 Then metaText = metaText & " | Checksum: " & SanitizeText(Checksum)

    Dim footerRng As Object
    Set footerRng = mWorksheet.Range(mWorksheet.Cells(mCurrentRow, mStartCol), mWorksheet.Cells(mCurrentRow, endCol))
    footerRng.HorizontalAlignment = XL_LEFT
    PutValue footerRng.Cells(1, 1), metaText
    footerRng.Font.Name = "Segoe UI"
    footerRng.Font.Size = 8
    footerRng.Font.Italic = True
    footerRng.Font.Color = RGB(148, 163, 184) ' Slate 400
    mWorksheet.Rows(mCurrentRow).RowHeight = 16

    ' Subtle thin top divider
    With footerRng.Borders(XL_EDGE_TOP)
        .LineStyle = XL_CONTINUOUS
        .Weight = XL_THIN
        .Color = RGB(226, 232, 240) ' Slate 200
    End With

    mCurrentRow = mCurrentRow + 1
    Set AddProvenanceFooter = Me
End Function

' Auto-fits all report columns within aesthetic min/max constraints
Public Function AutoFitAll(Optional ByVal MinWidth As Double = 10, _
                           Optional ByVal MaxWidth As Double = 35) As cReportBuilder
    CheckAttached "AutoFitAll"
    If MinWidth <= 0 Or MaxWidth > MAX_COL_WIDTH Or MinWidth > MaxWidth Then
        Err.Raise 5, SRC & ".AutoFitAll", _
            "Width range invalid: need 0 < MinWidth <= MaxWidth <= " & MAX_COL_WIDTH & _
            " (got " & MinWidth & ", " & MaxWidth & ")."
    End If

    Dim c As Long, colObj As Object, curW As Double
    For c = mStartCol To mMaxCol
        Set colObj = mWorksheet.Columns(c)
        colObj.AutoFit
        curW = CDbl(colObj.ColumnWidth)
        If curW < MinWidth Then
            colObj.ColumnWidth = MinWidth
        ElseIf curW > MaxWidth Then
            colObj.ColumnWidth = MaxWidth
        End If
    Next c
    Set AutoFitAll = Me
End Function

' Reconciles two metric values (e.g. KPI value vs Table total) within an optional numerical Tolerance.
' Returns True if values match, or False and records a warning in Warnings collection if they differ.
Public Function Reconcile(ByVal ExpectedValue As Variant, _
                          ByVal ActualValue As Variant, _
                          Optional ByVal Tolerance As Double = 0.0001, _
                          Optional ByVal MetricName As String = "Metric") As Boolean
    Dim expNum As Double, actNum As Double
    On Error Resume Next
    expNum = CDbl(ExpectedValue)
    actNum = CDbl(ActualValue)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        If CStr(ExpectedValue) = CStr(ActualValue) Then
            Reconcile = True
        Else
            AddWarning "Reconciliation Mismatch [" & MetricName & "]: Expected '" & _
                       CStr(ExpectedValue) & "', Got '" & CStr(ActualValue) & "'."
            Reconcile = False
        End If
        Exit Function
    End If
    On Error GoTo 0

    Dim diff As Double
    diff = Abs(expNum - actNum)
    If diff <= Tolerance Then
        Reconcile = True
    Else
        AddWarning "Reconciliation Mismatch [" & MetricName & "]: Expected " & _
                   expNum & ", Got " & actNum & " (Diff: " & diff & " > " & Tolerance & ")."
        Reconcile = False
    End If
End Function

' Clears the area previously written by this builder instance, unmerges cells,
' and resets row/column pointers back to StartCellAddress for clean re-rendering.
Public Function Clear() As cReportBuilder
    CheckAttached "Clear"
    Dim bRng As Object
    On Error Resume Next
    Set bRng = BoundingRange
    If Not bRng Is Nothing Then
        bRng.UnMerge
        bRng.Clear
    End If
    On Error GoTo 0
    mCurrentRow = mStartRow
    mMaxCol = mStartCol
    mInColumnGroup = False
    mGroupStartRow = mStartRow
    mGroupMaxRow = mStartRow
    mGroupAnchorCol = mStartCol
    mLastTableStartRow = 0
    mLastTableEndRow = 0
    Set mLastTableStyler = Nothing
    Set mWarnings = New Collection
    Set Clear = Me
End Function

' High-level single-call convenience API to render a complete executive report
' with banner, styled table, auto-calculated total row, and autofitted columns.
Public Function QuickReport(ByVal TargetWorksheet As Object, _
                            ByVal Df As cDataFrame, _
                            Optional ByVal Title As String = "Executive Report", _
                            Optional ByVal Theme As String = "Navy", _
                            Optional ByVal StartCellAddress As String = "B2") As cReportBuilder
    If TargetWorksheet Is Nothing Then
        Err.Raise 91, SRC & ".QuickReport", "TargetWorksheet cannot be Nothing."
    End If
    If Df Is Nothing Then
        Err.Raise 91, SRC & ".QuickReport", "Df cannot be Nothing."
    End If
    Me.Attach TargetWorksheet, StartCellAddress
    Me.BeginUpdate
    Me.AddBanner Title, Theme:=Theme, WidthColumns:=Df.ColumnCount
    Me.AddSpace 1
    Me.AddTable Df, StyleTheme:=Theme, AddTotalRow:=True
    Me.AddSpace 1
    Me.AutoFitAll
    Me.EndUpdate
    Set QuickReport = Me
End Function

' ------------------------------------------------------------------------------
' Multi-Column Side-by-Side Flow & Coordinate Positioning
' ------------------------------------------------------------------------------

' Begins a multi-column side-by-side section group.
' Sections placed inside this group advance vertically. Call NextColumn to start
' an adjacent column to the right at the exact same vertical start row.
Public Function BeginColumnGroup() As cReportBuilder
    CheckAttached "BeginColumnGroup"
    If mInColumnGroup Then
        Err.Raise 5, SRC & ".BeginColumnGroup", "Already inside a ColumnGroup. Call EndColumnGroup first."
    End If
    mInColumnGroup = True
    mGroupStartRow = mCurrentRow
    mGroupMaxRow = mCurrentRow
    mGroupAnchorCol = mStartCol
    Set BeginColumnGroup = Me
End Function

' Shifts flow to an adjacent column to the right, aligning vertically with the start of the current group.
Public Function NextColumn(Optional ByVal ColGap As Long = 2) As cReportBuilder
    CheckAttached "NextColumn"
    If Not mInColumnGroup Then
        Err.Raise 5, SRC & ".NextColumn", "Not inside a ColumnGroup. Call BeginColumnGroup first."
    End If
    If ColGap < 0 Then ColGap = 1
    If mCurrentRow > mGroupMaxRow Then mGroupMaxRow = mCurrentRow
    mCurrentRow = mGroupStartRow
    mStartCol = mMaxCol + ColGap + 1
    RequireColumns mStartCol, "NextColumn"
    Set NextColumn = Me
End Function

' Concludes a multi-column side-by-side section group and moves CurrentRow below
' the tallest column in the group, restoring horizontal alignment back to the group anchor column.
Public Function EndColumnGroup(Optional ByVal ExtraRowGap As Long = 1) As cReportBuilder
    CheckAttached "EndColumnGroup"
    If Not mInColumnGroup Then
        Err.Raise 5, SRC & ".EndColumnGroup", "Not inside a ColumnGroup."
    End If
    If mCurrentRow > mGroupMaxRow Then mGroupMaxRow = mCurrentRow
    mCurrentRow = mGroupMaxRow + ExtraRowGap
    mStartCol = mGroupAnchorCol
    mInColumnGroup = False
    Set EndColumnGroup = Me
End Function

' Explicitly moves the current write cursor to an absolute worksheet row and column.
Public Function MoveTo(ByVal TargetRow As Long, ByVal TargetCol As Long) As cReportBuilder
    CheckAttached "MoveTo"
    RequireRows TargetRow, "MoveTo"
    RequireColumns TargetCol, "MoveTo"
    mCurrentRow = TargetRow
    mStartCol = TargetCol
    If TargetCol > mMaxCol Then mMaxCol = TargetCol
    Set MoveTo = Me
End Function

' Embeds a native Excel chart onto the worksheet at the current flow cursor.
' ChartType: 51 = xlColumnClustered, 4 = xlLine, 5 = xlPie, 57 = xlBarClustered.
Public Function AddChart(Optional ByVal SourceRange As Object = Nothing, _
                         Optional ByVal Title As String = "Metric Chart", _
                         Optional ByVal ChartType As Long = 51, _
                         Optional ByVal WidthColumns As Long = 6, _
                         Optional ByVal HeightRows As Long = 12) As cReportBuilder
    CheckAttached "AddChart"
    If WidthColumns < 1 Then WidthColumns = 6
    If HeightRows < 1 Then HeightRows = 12

    Dim endCol As Long, endRow As Long
    endCol = mStartCol + WidthColumns - 1
    endRow = mCurrentRow + HeightRows - 1
    RequireColumns endCol, "AddChart"
    RequireRows endRow, "AddChart"

    Dim topCell As Object, bottomCell As Object
    Set topCell = mWorksheet.Cells(mCurrentRow, mStartCol)
    Set bottomCell = mWorksheet.Cells(endRow, endCol)

    Dim chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double
    chartLeft = CDbl(topCell.Left)
    chartTop = CDbl(topCell.Top)
    chartWidth = CDbl(bottomCell.Left + bottomCell.Width - topCell.Left)
    chartHeight = CDbl(bottomCell.Top + bottomCell.Height - topCell.Top)

    Dim co As Object
    Set co = mWorksheet.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    co.Chart.ChartType = ChartType
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = Title

    If Not SourceRange Is Nothing Then
        co.Chart.SetSourceData SourceRange
    End If

    If endCol > mMaxCol Then mMaxCol = endCol
    mCurrentRow = endRow + 2
    Set AddChart = Me
End Function

' Configures print orientation, print area, and fit-to-page settings.
Public Function ConfigurePageSetup(Optional ByVal Orientation As Long = 2, _
                                   Optional ByVal FitToWide As Long = 1, _
                                   Optional ByVal FitToTall As Long = 0) As cReportBuilder
    CheckAttached "ConfigurePageSetup"
    Dim ps As Object
    On Error Resume Next
    Set ps = mWorksheet.PageSetup
    If Not ps Is Nothing Then
        ps.PrintArea = BoundingRange.Address
        ps.Orientation = Orientation ' 1 = xlPortrait, 2 = xlLandscape
        If FitToWide > 0 Then
            ps.Zoom = False
            ps.FitToPagesWide = FitToWide
            If FitToTall > 0 Then
                ps.FitToPagesTall = FitToTall
            Else
                ps.FitToPagesTall = False
            End If
        End If
        ps.CenterHorizontally = True
    End If
    On Error GoTo 0
    Set ConfigurePageSetup = Me
End Function

' Exports the current worksheet report directly to a PDF document.
Public Function ExportPDF(ByVal OutputFilePath As String) As cReportBuilder
    CheckAttached "ExportPDF"
    If Len(Trim$(OutputFilePath)) = 0 Then
        Err.Raise 5, SRC & ".ExportPDF", "OutputFilePath cannot be blank."
    End If
    mWorksheet.ExportAsFixedFormat Type:=0, Filename:=OutputFilePath, Quality:=0, _
                                   IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                                   OpenAfterPublish:=False
    Set ExportPDF = Me
End Function

' Cleans a sheet name to adhere to Excel's 31-character limit and invalid character rules.
Public Function SanitizeSheetName(ByVal Name As String) As String
    Dim s As String, i As Long, ch As String
    s = Trim$(Name)
    If Len(s) = 0 Then s = "Sheet"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case ":", "\", "/", "?", "*", "[", "]"
                Mid$(s, i, 1) = "_"
        End Select
    Next i
    Do While Left$(s, 1) = "'" And Len(s) > 1
        s = Mid$(s, 2)
    Loop
    Do While Right$(s, 1) = "'" And Len(s) > 1
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) > 31 Then s = Left$(s, 31)
    SanitizeSheetName = s
End Function

' Automatically partitions a DataFrame by GroupColumn (via df.SplitBy), generates an executive
' worksheet for each partition, and creates a Table of Contents (TOC) master index sheet with live hyperlinks.
Public Function BuildMultiSheetReport(ByVal TargetWorkbook As Object, _
                                      ByVal Df As cDataFrame, _
                                      ByVal GroupColumn As String, _
                                      Optional ByVal TitlePrefix As String = "Report", _
                                      Optional ByVal Theme As String = "Navy", _
                                      Optional ByVal TocSheetName As String = "TOC") As Collection
    If TargetWorkbook Is Nothing Then
        Err.Raise 91, SRC & ".BuildMultiSheetReport", "TargetWorkbook cannot be Nothing."
    End If
    If Df Is Nothing Then
        Err.Raise 91, SRC & ".BuildMultiSheetReport", "Df cannot be Nothing."
    End If
    If Not Df.HasColumn(GroupColumn) Then
        Err.Raise 5, SRC & ".BuildMultiSheetReport", "GroupColumn '" & GroupColumn & "' not found in DataFrame."
    End If

    Dim splits As Collection
    Set splits = Df.SplitBy(GroupColumn)

    Dim groupColObj As cColumn
    Set groupColObj = Df.Column(GroupColumn)
    Dim uniqueGroups As Variant
    uniqueGroups = groupColObj.Unique()

    Dim createdSheets As New Collection
    Dim wsToc As Object, wsGrp As Object, childRb As cReportBuilder
    Dim i As Long, grpVal As String, safeName As String
    Dim dfSub As cDataFrame

    ' 1. Create individual group sheets
    For i = LBound(uniqueGroups) To UBound(uniqueGroups)
        grpVal = CStr(uniqueGroups(i))
        safeName = SanitizeSheetName(grpVal)
        Set wsGrp = TargetWorkbook.Worksheets.Add
        wsGrp.Name = safeName

        Set dfSub = splits.Item(i - LBound(uniqueGroups) + 1)

        Set childRb = New cReportBuilder
        childRb.QuickReport wsGrp, dfSub, Title:=TitlePrefix & " - " & grpVal, Theme:=Theme
        createdSheets.Add wsGrp
    Next i

    ' 2. Create Master Table of Contents Sheet
    Set wsToc = TargetWorkbook.Worksheets.Add(Before:=TargetWorkbook.Worksheets(1))
    wsToc.Name = SanitizeSheetName(TocSheetName)

    Dim tocRb As New cReportBuilder
    tocRb.Attach wsToc, "B2"
    tocRb.BeginUpdate
    tocRb.AddBanner "Executive Table of Contents", Subtitle:="Multi-Sheet Index (" & TitlePrefix & ")", Theme:=Theme, WidthColumns:=4
    tocRb.AddSpace 1

    Dim dfToc As New cDataFrame
    dfToc.AddColumn "Index"
    dfToc.AddColumn "Group / Sheet"
    dfToc.AddColumn "Records"

    For i = LBound(uniqueGroups) To UBound(uniqueGroups)
        grpVal = CStr(uniqueGroups(i))
        Set dfSub = splits.Item(i - LBound(uniqueGroups) + 1)
        dfToc.AddRow i - LBound(uniqueGroups) + 1, grpVal, dfSub.RowCount
    Next i

    tocRb.AddTable dfToc, Title:="Available Reports", StyleTheme:=Theme, AddTotalRow:=True
    tocRb.AddSpace 1
    tocRb.AddProvenanceFooter Author:="Auto-Report Engine", SourceDescription:="df.SplitBy('" & GroupColumn & "')"
    tocRb.AutoFitAll
    tocRb.EndUpdate

    createdSheets.Add wsToc
    Set BuildMultiSheetReport = createdSheets
End Function

' ------------------------------------------------------------------------------
' Properties
' ------------------------------------------------------------------------------

Public Property Get CurrentRow() As Long
    CurrentRow = mCurrentRow
End Property

Public Property Get StartRow() As Long
    StartRow = mStartRow
End Property

Public Property Get StartCol() As Long
    StartCol = mStartCol
End Property

Public Property Get MaxCol() As Long
    MaxCol = mMaxCol
End Property

Public Property Get LastTableStyler() As cTableStyler
    Set LastTableStyler = mLastTableStyler
End Property

Public Property Get LastTableStartRow() As Long
    LastTableStartRow = mLastTableStartRow
End Property

Public Property Get LastTableEndRow() As Long
    LastTableEndRow = mLastTableEndRow
End Property

' Non-fatal findings recorded while building (truncated text, skipped total row, ...)
Public Property Get Warnings() As Collection
    Set Warnings = mWarnings
End Property

Public Property Get WarningCount() As Long
    WarningCount = mWarnings.Count
End Property

Public Property Get BoundingRange() As Object
    CheckAttached "BoundingRange"
    Dim lastR As Long
    lastR = mCurrentRow - 1
    If lastR < mStartRow Then lastR = mStartRow
    Set BoundingRange = mWorksheet.Range(mWorksheet.Cells(mStartRow, mStartCol), mWorksheet.Cells(lastR, mMaxCol))
End Property

' ------------------------------------------------------------------------------
' Text Safety (public so they can be unit-tested without Excel)
' ------------------------------------------------------------------------------

' Returns text that is safe to store in a cell:
'   - control characters (except Tab/LF/CR) are removed
'   - text longer than the Excel cell limit is truncated (a warning is recorded)
'   - text that Excel could read as a formula or prefix ("=", "+", "-", "@", Tab, CR, "'")
'     gets a leading apostrophe, which Excel consumes and never displays
Public Function SanitizeText(ByVal Value As Variant) As String
    Dim s As String
    s = ToPlainText(Value)
    If Len(s) = 0 Then Exit Function

    Dim i As Long, code As Long, needClean As Boolean
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code >= 0 And code < 32 And code <> 9 And code <> 10 And code <> 13 Then
            needClean = True
            Exit For
        End If
    Next i

    If needClean Then
        Dim cleaned As String
        For i = 1 To Len(s)
            code = AscW(Mid$(s, i, 1))
            If Not (code >= 0 And code < 32 And code <> 9 And code <> 10 And code <> 13) Then
                cleaned = cleaned & Mid$(s, i, 1)
            End If
        Next i
        s = cleaned
        AddWarning "SanitizeText: control characters were removed from a text value."
    End If

    If Len(s) > MAX_CELL_CHARS - 1 Then
        s = Left$(s, MAX_CELL_CHARS - 1)
        AddWarning "SanitizeText: text longer than " & MAX_CELL_CHARS & " characters was truncated."
    End If

    If Len(s) > 0 Then
        Select Case AscW(Left$(s, 1))
            Case 61, 43, 45, 64, 9, 13, 39 ' = + - @ Tab CR '
                s = "'" & s
        End Select
    End If

    SanitizeText = s
End Function

' Direction of a KPI subtext: 1 = up (leading "+" or up-triangle), -1 = down (leading "-",
' minus sign or down-triangle), 0 = neutral. Only the first visible character counts, so
' text such as "Q4-2026" or "n/a" stays neutral.
Public Function TrendDirection(ByVal TextValue As String) As Long
    Dim t As String
    t = LTrim$(TextValue)
    If Len(t) = 0 Then Exit Function

    Select Case AscW(Left$(t, 1))
        Case 43, 9650            ' "+" and U+25B2 up-pointing triangle
            TrendDirection = 1
        Case 45, 8722, 9660      ' "-", U+2212 minus sign and U+25BC down-pointing triangle
            TrendDirection = -1
        Case Else
            TrendDirection = 0
    End Select
End Function

' ------------------------------------------------------------------------------
' Internal Helpers
' ------------------------------------------------------------------------------

' Writes a value into a cell/range: numbers/dates/booleans stay typed, everything else is
' stored as sanitized literal text (NumberFormat "@" stops Excel re-interpreting it).
Private Sub PutValue(ByVal Rng As Object, ByVal v As Variant)
    Dim safe As Variant
    safe = SafeCellValue(v)
    If VarType(safe) = vbString Then
        Rng.NumberFormat = XL_TEXT_FORMAT
    End If
    Rng.Value2 = safe
End Sub

Private Function SafeCellValue(ByVal v As Variant) As Variant
    If IsObject(v) Then
        SafeCellValue = vbNullString
        Exit Function
    End If
    If IsError(v) Then
        SafeCellValue = "#ERR"
        Exit Function
    End If

    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, vbDate, vbBoolean
            SafeCellValue = v
        Case vbEmpty, vbNull
            SafeCellValue = vbNullString
        Case Else
            SafeCellValue = SanitizeText(v)
    End Select
End Function

Private Function ToPlainText(ByVal v As Variant) As String
    If IsObject(v) Then Exit Function
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    ToPlainText = CStr(v)
End Function

Private Function AsArray(ByVal v As Variant) As Variant
    If IsArray(v) Then
        AsArray = v
    Else
        AsArray = Array(v)
    End If
End Function

Private Function ArrayCount(ByVal arr As Variant) As Long
    If Not IsArray(arr) Then Exit Function
    ArrayCount = UBound(arr) - LBound(arr) + 1
End Function

' 1-based indexes of columns whose non-empty cells are all numeric types (and at least one exists).
' Returns Empty when no column qualifies. Indexes (not names) are used so column names containing
' "," or "*" cannot be mistaken for patterns by cTableStyler.
Private Function NumericColumnIndexes(ByVal Df As cDataFrame) As Variant
    Dim names() As String
    names = Df.ColumnNames()

    Dim result() As Variant
    Dim cnt As Long
    Dim i As Long, r As Long
    Dim col As cColumn, v As Variant
    Dim allNumeric As Boolean, sawNumber As Boolean

    For i = LBound(names) To UBound(names)
        Set col = Df.Column(names(i))
        allNumeric = True
        sawNumber = False
        For r = 1 To Df.RowCount
            v = col.Item(r)
            If IsError(v) Then
                allNumeric = False
            ElseIf IsEmpty(v) Or IsNull(v) Then
                ' blank cell: ignored
            Else
                Select Case VarType(v)
                    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
                        sawNumber = True
                    Case Else
                        allNumeric = False
                End Select
            End If
            If Not allNumeric Then Exit For
        Next r

        If allNumeric And sawNumber Then
            cnt = cnt + 1
            ReDim Preserve result(1 To cnt)
            result(cnt) = (i - LBound(names) + 1)
        End If
    Next i

    If cnt = 0 Then
        NumericColumnIndexes = Empty
    Else
        NumericColumnIndexes = result
    End If
End Function

' First column index (1-based) that is not in the numeric set, or 0 when every column is numeric.
Private Function FirstNonNumericIndex(ByVal numericCols As Variant, ByVal totalCols As Long) As Long
    Dim c As Long, j As Long, isNum As Boolean
    For c = 1 To totalCols
        isNum = False
        For j = LBound(numericCols) To UBound(numericCols)
            If CLng(numericCols(j)) = c Then
                isNum = True
                Exit For
            End If
        Next j
        If Not isNum Then
            FirstNonNumericIndex = c
            Exit Function
        End If
    Next c
End Function

Private Sub ResolveBannerTheme(ByVal ThemeName As String, _
                               ByRef BannerBg As Long, ByRef BannerFg As Long, _
                               ByRef SubBg As Long, ByRef SubFg As Long)
    BannerFg = RGB(255, 255, 255)
    Select Case UCase$(Trim$(ThemeName))
        Case "NAVY", "CORPORATE_NAVY", "MODERNBLUE", "MODERN_BLUE", "BLUE"
            BannerBg = RGB(24, 43, 73)      ' Corporate Navy
            SubBg = RGB(240, 244, 250)
            SubFg = RGB(31, 78, 121)
        Case "EMERALD", "GREEN"
            BannerBg = RGB(16, 75, 45)      ' Deep Forest Green
            SubBg = RGB(236, 253, 243)
            SubFg = RGB(22, 101, 52)
        Case "SLATE", "MODERN_CHARCOAL"
            BannerBg = RGB(51, 65, 85)      ' Slate Dark
            SubBg = RGB(241, 245, 249)
            SubFg = RGB(71, 85, 105)
        Case "TEAL"
            BannerBg = RGB(13, 71, 84)      ' Deep Teal
            SubBg = RGB(235, 246, 247)
            SubFg = RGB(15, 118, 110)
        Case "BURGUNDY"
            BannerBg = RGB(128, 0, 32)
            SubBg = RGB(250, 236, 240)
            SubFg = RGB(128, 0, 32)
        Case "INDIGO"
            BannerBg = RGB(63, 81, 181)
            SubBg = RGB(236, 238, 250)
            SubFg = RGB(48, 63, 159)
        Case "DARK_GRAY"
            BannerBg = RGB(64, 64, 64)
            SubBg = RGB(240, 240, 240)
            SubFg = RGB(64, 64, 64)
        Case "MINIMAL", "CLASSIC_GRAY"
            BannerBg = RGB(220, 224, 230)
            BannerFg = RGB(40, 40, 40)
            SubBg = RGB(244, 246, 248)
            SubFg = RGB(70, 70, 70)
        Case Else
            Err.Raise 5, SRC & ".AddBanner", "Unknown theme '" & ThemeName & "'. " & SupportedThemes()
    End Select
End Sub

' Maps a user theme name to the name understood by cTableStyler.ApplyTheme (fail-loud on typos).
Private Function ResolveTableTheme(ByVal ThemeName As String) As String
    Select Case UCase$(Trim$(ThemeName))
        Case "MODERNBLUE", "MODERN_BLUE", "BLUE"
            ResolveTableTheme = "NAVY"
        Case "NAVY", "CORPORATE_NAVY", "SLATE", "MODERN_CHARCOAL", "EMERALD", "GREEN", _
             "MINIMAL", "CLASSIC_GRAY", "TEAL", "BURGUNDY", "INDIGO", "DARK_GRAY"
            ResolveTableTheme = UCase$(Trim$(ThemeName))
        Case Else
            Err.Raise 5, SRC & ".AddTable", "Unknown theme '" & ThemeName & "'. " & SupportedThemes()
    End Select
End Function

Private Function SupportedThemes() As String
    SupportedThemes = "Supported: Navy, Slate, Emerald, Teal, Burgundy, Indigo, Dark_Gray, Minimal " & _
                      "(aliases: ModernBlue, Blue, Green, Corporate_Navy, Modern_Charcoal, Classic_Gray)."
End Function

Private Sub RequireRows(ByVal LastRowNeeded As Long, ByVal MethodName As String)
    If LastRowNeeded > MAX_SHEET_ROWS Then
        Err.Raise 1004, SRC & "." & MethodName, _
            "Report does not fit the worksheet: needs row " & LastRowNeeded & _
            " but the limit is " & MAX_SHEET_ROWS & ". Nothing was written."
    End If
End Sub

Private Sub RequireColumns(ByVal LastColNeeded As Long, ByVal MethodName As String)
    If LastColNeeded > MAX_SHEET_COLS Then
        Err.Raise 1004, SRC & "." & MethodName, _
            "Report does not fit the worksheet: needs column " & LastColNeeded & _
            " but the limit is " & MAX_SHEET_COLS & ". Nothing was written."
    End If
End Sub

Private Sub AddWarning(ByVal Message As String)
    If Not mWarnings Is Nothing Then mWarnings.Add Message
End Sub

Private Sub CheckAttached(ByVal MethodName As String)
    If mWorksheet Is Nothing Then
        Err.Raise 91, SRC & "." & MethodName, "ReportBuilder is not attached to a worksheet. Call Attach() first."
    End If
End Sub
