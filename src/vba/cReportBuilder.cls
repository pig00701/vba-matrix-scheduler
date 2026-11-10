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
' ==============================================================================

Private mWorksheet As Object
Private mStartRow As Long
Private mStartCol As Long
Private mCurrentRow As Long
Private mMaxCol As Long
Private mInitialTopLeft As Object

Private Sub Class_Initialize()
    mStartRow = 2
    mStartCol = 2
    mCurrentRow = 2
    mMaxCol = 2
End Sub

Private Sub Class_Terminate()
    Set mWorksheet = Nothing
    Set mInitialTopLeft = Nothing
End Sub

' ------------------------------------------------------------------------------
' Initialization & Flow Positioning
' ------------------------------------------------------------------------------

Public Function Attach(ByVal TargetWorksheet As Object, _
                       Optional ByVal StartCellAddress As String = "B2") As cReportBuilder
    If TargetWorksheet Is Nothing Then
        Err.Raise 91, "cReportBuilder.Attach", "TargetWorksheet cannot be Nothing."
    End If

    Set mWorksheet = TargetWorksheet
    If mWorksheet.ProtectContents Then
        Err.Raise 1004, "cReportBuilder.Attach", _
            "Cannot build report: Worksheet '" & mWorksheet.Name & "' is protected."
    End If

    Dim anchor As Object
    Set anchor = mWorksheet.Range(StartCellAddress).Cells(1, 1)
    Set mInitialTopLeft = anchor
    mStartRow = anchor.Row
    mStartCol = anchor.Column
    mCurrentRow = mStartRow
    mMaxCol = mStartCol

    Set Attach = Me
End Function

' ------------------------------------------------------------------------------
' Section Builders
' ------------------------------------------------------------------------------

' Adds an Executive Title Banner spanning WidthColumns
Public Function AddBanner(ByVal Title As String, _
                          Optional ByVal Subtitle As String = "", _
                          Optional ByVal Theme As String = "Navy", _
                          Optional ByVal WidthColumns As Long = 8) As cReportBuilder
    CheckAttached "AddBanner"
    If WidthColumns < 1 Then WidthColumns = 8

    Dim endCol As Long
    endCol = mStartCol + WidthColumns - 1
    If endCol > mMaxCol Then mMaxCol = endCol

    Dim bannerBg As Long, bannerFg As Long, subBg As Long, subFg As Long
    Select Case UCase$(Trim$(Theme))
        Case "EMERALD"
            bannerBg = RGB(16, 75, 45)     ' Deep Forest Green
            bannerFg = RGB(255, 255, 255)
            subBg = RGB(236, 253, 243)
            subFg = RGB(22, 101, 52)
        Case "SLATE"
            bannerBg = RGB(51, 65, 85)     ' Slate Dark
            bannerFg = RGB(255, 255, 255)
            subBg = RGB(241, 245, 249)
            subFg = RGB(71, 85, 105)
        Case "TEAL"
            bannerBg = RGB(13, 71, 84)     ' Deep Teal
            bannerFg = RGB(255, 255, 255)
            subBg = RGB(235, 246, 247)
            subFg = RGB(15, 118, 110)
        Case Else ' Navy (Default)
            bannerBg = RGB(24, 43, 73)     ' Corporate Navy
            bannerFg = RGB(255, 255, 255)
            subBg = RGB(240, 244, 250)
            subFg = RGB(31, 78, 120)
    End Select

    ' Title Row
    Dim titleRng As Object
    Set titleRng = mWorksheet.Range(mWorksheet.Cells(mCurrentRow, mStartCol), mWorksheet.Cells(mCurrentRow, endCol))
    titleRng.Merge
    titleRng.Value2 = Title
    titleRng.Font.Name = "Segoe UI"
    titleRng.Font.Size = 16
    titleRng.Font.Bold = True
    titleRng.Font.Color = bannerFg
    titleRng.Interior.Color = bannerBg
    titleRng.HorizontalAlignment = -4131 ' xlLeft
    titleRng.VerticalAlignment = -4108   ' xlCenter
    mWorksheet.Rows(mCurrentRow).RowHeight = 32

    mCurrentRow = mCurrentRow + 1

    ' Optional Subtitle Row
    If Len(Trim$(Subtitle)) > 0 Then
        Dim subRng As Object
        Set subRng = mWorksheet.Range(mWorksheet.Cells(mCurrentRow, mStartCol), mWorksheet.Cells(mCurrentRow, endCol))
        subRng.Merge
        subRng.Value2 = Subtitle
        subRng.Font.Name = "Segoe UI"
        subRng.Font.Size = 9.5
        subRng.Font.Italic = True
        subRng.Font.Color = subFg
        subRng.Interior.Color = subBg
        subRng.HorizontalAlignment = -4131 ' xlLeft
        subRng.VerticalAlignment = -4108   ' xlCenter
        mWorksheet.Rows(mCurrentRow).RowHeight = 20
        mCurrentRow = mCurrentRow + 1
    End If

    mCurrentRow = mCurrentRow + 1 ' spacing gap
    Set AddBanner = Me
End Function

' Adds a row of horizontal KPI metric summary cards side-by-side
Public Function AddKpiRow(ByVal Titles As Variant, _
                          ByVal Values As Variant, _
                          Optional ByVal Subtexts As Variant, _
                          Optional ByVal CardWidthCols As Long = 2) As cReportBuilder
    CheckAttached "AddKpiRow"
    If CardWidthCols < 1 Then CardWidthCols = 2

    Dim lbT As Long, ubT As Long
    lbT = LBound(Titles)
    ubT = UBound(Titles)
    Dim numCards As Long
    numCards = ubT - lbT + 1

    Dim hasSub As Boolean
    hasSub = Not IsMissing(Subtexts) And Not IsEmpty(Subtexts)

    Dim curC As Long
    curC = mStartCol

    Dim k As Long, tStr As String, vStr As Variant, sStr As String
    Dim cardTop As Long, cardBottom As Long
    cardTop = mCurrentRow
    cardBottom = cardTop + IIf(hasSub, 2, 1)

    For k = 0 To numCards - 1
        tStr = CStr(Titles(lbT + k))
        vStr = Values(LBound(Values) + k)
        sStr = ""
        If hasSub Then
            If k <= (UBound(Subtexts) - LBound(Subtexts)) Then
                sStr = CStr(Subtexts(LBound(Subtexts) + k))
            End If
        End If

        Dim cardEndC As Long
        cardEndC = curC + CardWidthCols - 1
        If cardEndC > mMaxCol Then mMaxCol = cardEndC

        ' Card Title (Row 1)
        Dim tRng As Object
        Set tRng = mWorksheet.Range(mWorksheet.Cells(cardTop, curC), mWorksheet.Cells(cardTop, cardEndC))
        tRng.Merge
        tRng.Value2 = UCase$(tStr)
        tRng.Font.Name = "Segoe UI"
        tRng.Font.Size = 8.5
        tRng.Font.Bold = True
        tRng.Font.Color = RGB(100, 116, 139) ' Slate 500
        tRng.HorizontalAlignment = -4108    ' xlCenter
        tRng.VerticalAlignment = -4108

        ' Card Value (Row 2)
        Dim vRng As Object
        Set vRng = mWorksheet.Range(mWorksheet.Cells(cardTop + 1, curC), mWorksheet.Cells(cardTop + 1, cardEndC))
        vRng.Merge
        vRng.Value2 = vStr
        vRng.Font.Name = "Segoe UI"
        vRng.Font.Size = 16
        vRng.Font.Bold = True
        vRng.Font.Color = RGB(15, 23, 42)    ' Slate 900
        vRng.HorizontalAlignment = -4108    ' xlCenter
        vRng.VerticalAlignment = -4108

        ' Card Subtext (Row 3, if present)
        If hasSub Then
            Dim sRng As Object
            Set sRng = mWorksheet.Range(mWorksheet.Cells(cardTop + 2, curC), mWorksheet.Cells(cardTop + 2, cardEndC))
            sRng.Merge
            sRng.Value2 = sStr
            sRng.Font.Name = "Segoe UI"
            sRng.Font.Size = 8.5
            sRng.Font.Bold = True
            If InStr(sStr, "+") > 0 Or InStr(sStr, "?") > 0 Then
                sRng.Font.Color = RGB(22, 101, 52) ' Green
            ElseIf InStr(sStr, "-") > 0 Or InStr(sStr, "?") > 0 Then
                sRng.Font.Color = RGB(185, 28, 28) ' Red
            Else
                sRng.Font.Color = RGB(100, 116, 139)
            End If
            sRng.HorizontalAlignment = -4108 ' xlCenter
            sRng.VerticalAlignment = -4108
        End If

        ' Apply Card Box Border and Subtle Fill
        Dim fullCardRng As Object
        Set fullCardRng = mWorksheet.Range(mWorksheet.Cells(cardTop, curC), mWorksheet.Cells(cardBottom, cardEndC))
        fullCardRng.Interior.Color = RGB(248, 250, 252) ' Light Gray Fill

        Dim b As Long
        For b = 7 To 10 ' xlEdgeLeft to xlEdgeRight
            fullCardRng.Borders(b).LineStyle = 1
            fullCardRng.Borders(b).Weight = 2 ' xlThin
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

' Adds a formatted cDataFrame table with optional section header and auto-total row
Public Function AddTable(ByVal Df As cDataFrame, _
                         Optional ByVal Title As String = "", _
                         Optional ByVal StyleTheme As String = "ModernBlue", _
                         Optional ByVal AddTotalRow As Boolean = False) As cReportBuilder
    CheckAttached "AddTable"
    If Df Is Nothing Then
        Err.Raise 91, "cReportBuilder.AddTable", "Df cannot be Nothing."
    End If

    Dim endCol As Long
    endCol = mStartCol + Df.ColumnCount - 1
    If endCol > mMaxCol Then mMaxCol = endCol

    ' Section Header Title
    If Len(Trim$(Title)) > 0 Then
        Dim titleCell As Object
        Set titleCell = mWorksheet.Cells(mCurrentRow, mStartCol)
        titleCell.Value2 = Title
        titleCell.Font.Name = "Segoe UI"
        titleCell.Font.Size = 12
        titleCell.Font.Bold = True
        titleCell.Font.Color = RGB(30, 41, 59)
        mWorksheet.Rows(mCurrentRow).RowHeight = 22
        mCurrentRow = mCurrentRow + 1
    End If

    ' Dump DataFrame to Worksheet
    Dim startCell As Object
    Set startCell = mWorksheet.Cells(mCurrentRow, mStartCol)
    Df.ToRange startCell, IncludeHeaders:=True

    ' Style using cTableStyler
    Dim styler As New cTableStyler
    styler.Attach startCell, Df, HasHeaders:=True
    styler.ApplyTheme StyleTheme
    styler.ApplyZebraStripes

    If AddTotalRow Then
        styler.AddTotalRow UseSubtotal:=True
    End If

    ' Advance current row
    Dim tableTotalRows As Long
    tableTotalRows = Df.RowCount + 1 + IIf(AddTotalRow, 1, 0)
    mCurrentRow = mCurrentRow + tableTotalRows + 1

    Set AddTable = Me
End Function

' Adds vertical blank spacing
Public Function AddSpace(Optional ByVal Rows As Long = 2) As cReportBuilder
    CheckAttached "AddSpace"
    If Rows < 1 Then Rows = 1
    mCurrentRow = mCurrentRow + Rows
    Set AddSpace = Me
End Function

' Adds an explanatory footnote or remark
Public Function AddNote(ByVal NoteText As String, _
                        Optional ByVal IsItalic As Boolean = True) As cReportBuilder
    CheckAttached "AddNote"
    Dim noteCell As Object
    Set noteCell = mWorksheet.Cells(mCurrentRow, mStartCol)
    noteCell.Value2 = NoteText
    noteCell.Font.Name = "Segoe UI"
    noteCell.Font.Size = 9
    noteCell.Font.Italic = IsItalic
    noteCell.Font.Color = RGB(100, 116, 139)
    mCurrentRow = mCurrentRow + 1
    Set AddNote = Me
End Function

' Auto-fits all report columns within aesthetic min/max constraints
Public Function AutoFitAll(Optional ByVal MinWidth As Double = 10, _
                           Optional ByVal MaxWidth As Double = 35) As cReportBuilder
    CheckAttached "AutoFitAll"
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

Public Property Get BoundingRange() As Object
    CheckAttached "BoundingRange"
    Dim lastR As Long
    lastR = mCurrentRow - 1
    If lastR < mStartRow Then lastR = mStartRow
    Set BoundingRange = mWorksheet.Range(mWorksheet.Cells(mStartRow, mStartCol), mWorksheet.Cells(lastR, mMaxCol))
End Property

Private Sub CheckAttached(ByVal MethodName As String)
    If mWorksheet Is Nothing Then
        Err.Raise 91, "cReportBuilder." & MethodName, "ReportBuilder is not attached to a worksheet. Call Attach() first."
    End If
End Sub
