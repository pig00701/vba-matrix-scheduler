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
' Description: Fluent Excel Cell & Range Styling Engine for cDataFrame.
'              Decouples presentation/formatting concerns from the core in-memory
'              data structure. Provides single-COM-call header theming, border
'              application, vectorized column number formatting, multi-column
'              formatting, batch-union zebra striping, total row highlighting,
'              auto-fitting with custom padding, and screen updating management.
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
Private mInBatchOperation As Boolean   ' Track screen updating suppression
Private mPrevScreenUpdating As Boolean ' Stored ScreenUpdating state

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

' ------------------------------------------------------------------------------
' Screen Updating Optimization Helpers
' ------------------------------------------------------------------------------

' Suppresses Excel screen updates for batch styling operations
Public Function BeginUpdate() As cTableStyler
    If Not mInBatchOperation And Not mWorksheet Is Nothing Then
        On Error Resume Next
        mPrevScreenUpdating = mWorksheet.Application.ScreenUpdating
        mWorksheet.Application.ScreenUpdating = False
        mInBatchOperation = True
        On Error GoTo 0
    End If
    Set BeginUpdate = Me
End Function

' Restores previous Excel screen updating state
Public Function EndUpdate() As cTableStyler
    If mInBatchOperation And Not mWorksheet Is Nothing Then
        On Error Resume Next
        mWorksheet.Application.ScreenUpdating = mPrevScreenUpdating
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

' Sets NumberFormat on data cells for a given column (or 1D array of columns) in a single COM call.
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

    Dim colIdx As Long
    colIdx = ResolveColIndex(ColumnKey)
    If colIdx < 1 Or colIdx > mColumnCount Then
        Err.Raise 9, "cTableStyler.FormatColumn", "Column index out of range: " & CStr(ColumnKey)
    End If

    mDataBodyRange.Columns(colIdx).NumberFormat = NumberFormatString

    Set FormatColumn = Me
End Function

' Sets NumberFormat on data cells for multiple columns (array or comma-delimited string).
Public Function FormatColumns(ByVal ColumnKeys As Variant, ByVal NumberFormatString As String) As cTableStyler
    If mDataBodyRange Is Nothing Or mRowCount = 0 Then
        Set FormatColumns = Me
        Exit Function
    End If

    If IsArray(ColumnKeys) Then
        Dim i As Long
        For i = LBound(ColumnKeys) To UBound(ColumnKeys)
            FormatColumn ColumnKeys(i), NumberFormatString
        Next i
    ElseIf VarType(ColumnKeys) = vbString And InStr(CStr(ColumnKeys), ",") > 0 Then
        Dim parts() As String
        parts = Split(CStr(ColumnKeys), ",")
        Dim p As Long
        For p = LBound(parts) To UBound(parts)
            FormatColumn Trim$(parts(p)), NumberFormatString
        Next p
    Else
        FormatColumn ColumnKeys, NumberFormatString
    End If

    Set FormatColumns = Me
End Function

' Sets HorizontalAlignment on data cells (or entire column including header) for a given column.
Public Function AlignColumn(ByVal ColumnKey As Variant, ByVal Alignment As Long, _
                            Optional ByVal IncludeHeader As Boolean = False) As cTableStyler
    If mTableRange Is Nothing Then
        Set AlignColumn = Me
        Exit Function
    End If

    Dim colIdx As Long
    colIdx = ResolveColIndex(ColumnKey)
    If colIdx < 1 Or colIdx > mColumnCount Then
        Err.Raise 9, "cTableStyler.AlignColumn", "Column index out of range: " & CStr(ColumnKey)
    End If

    If IncludeHeader Then
        mTableRange.Columns(colIdx).HorizontalAlignment = Alignment
    ElseIf Not mDataBodyRange Is Nothing Then
        mDataBodyRange.Columns(colIdx).HorizontalAlignment = Alignment
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

' Auto-fits all table columns and applies extra width padding.
Public Function AutoFitColumns(Optional ByVal Padding As Double = 3#) As cTableStyler
    If mTableRange Is Nothing Then
        Set AutoFitColumns = Me
        Exit Function
    End If

    mTableRange.Columns.AutoFit

    If Padding > 0# Then
        Dim c As Long
        For c = 1 To mColumnCount
            mTableRange.Columns(c).ColumnWidth = mTableRange.Columns(c).ColumnWidth + Padding
        Next c
    End If

    Set AutoFitColumns = Me
End Function

' ------------------------------------------------------------------------------
' Internal Helpers
' ------------------------------------------------------------------------------

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
