VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' Class: cDataFrame
' Description: High-performance Column-Oriented In-Memory Table (VBA DataFrame).
'              Decouples rows and columns using dictionary-indexed cColumn stores,
'              enabling O(1) column additions, drops, and renames, alongside
'              geometric row buffer resizing, bulk Excel Range I/O, fast
'              matrix lookups (SetIndex / Loc / Lookup), in-memory multi-column
'              indirect QuickSort, O(N+M) Hash Joins (Merge), O(N) Hash
'              GroupBy & Pivot table aggregations, O(N) Deduplication
'              & Set Operations (DropDuplicates, Concat), Missing Data
'              Handling (FillNA, DropNA), and Fast CSV & ListObject I/O.
' Author: Korn (AI Backend Developer, Technodevmoo)
' Architecture: Kit (Tech Lead)
' ==============================================================================

Private mCols As Object            ' Scripting.Dictionary of (ColName -> cColumn)
Private mRowCount As Long
Private mRowIndexMap As Object        ' Scripting.Dictionary of (RowKey -> RowIndex)
Private mIndexColumnName As String    ' Active index column name(s)
Private mIndexColumnNames() As String ' 1-based array of active index column names
Private mIndexColCount As Long        ' Number of columns in active index
Private mIndexSkipEmpty As Boolean    ' SkipEmptyKeys flag for index rebuild
Private mIndexDirty As Boolean        ' True if index needs rebuild due to mutation
Private mColList() As cColumn         ' Cached flat array of columns for fast AddRow iteration
Private mColListCount As Long         ' Number of cached columns in mColList
Private mCaseSensitive As Boolean     ' True for vbBinaryCompare, False for vbTextCompare
Private mRenamedHeaders As Object     ' Scripting.Dictionary of renamed duplicate headers

' ------------------------------------------------------------------------------
' Lifecycle
' ------------------------------------------------------------------------------

Private Sub Class_Initialize()
    Set mCols = CreateObject("Scripting.Dictionary")
    mCols.CompareMode = 1          ' vbTextCompare (Case-insensitive column lookups)
    mRowCount = 0
    Set mRowIndexMap = Nothing
    mIndexColumnName = ""
    mIndexColCount = 0
    mIndexSkipEmpty = True
    mIndexDirty = False
    mColListCount = 0
    Erase mColList
    mCaseSensitive = False
    Set mRenamedHeaders = CreateObject("Scripting.Dictionary")
    mRenamedHeaders.CompareMode = 1
End Sub

Private Sub Class_Terminate()
    Set mCols = Nothing
    Set mRowIndexMap = Nothing
    Set mRenamedHeaders = Nothing
    Erase mColList
    mColListCount = 0
End Sub

' ------------------------------------------------------------------------------
' Schema Cache, Case Sensitivity & Header Deduplication
' ------------------------------------------------------------------------------

Private Sub SyncColList()
    Dim numCols As Long, c As Long
    numCols = mCols.Count
    mColListCount = numCols
    If numCols = 0 Then
        Erase mColList
        Exit Sub
    End If
    ReDim mColList(1 To numCols)
    Dim keys As Variant
    keys = mCols.Keys
    For c = 1 To numCols
        Set mColList(c) = mCols.Item(keys(c - 1))
    Next c
End Sub

Public Property Get CaseSensitive() As Boolean
    CaseSensitive = mCaseSensitive
End Property

Public Property Let CaseSensitive(ByVal Value As Boolean)
    If Value = mCaseSensitive Then Exit Property
    mCaseSensitive = Value
    Dim newMode As Long
    newMode = IIf(Value, 0, 1) ' 0 = vbBinaryCompare, 1 = vbTextCompare

    Dim newDict As Object
    Set newDict = CreateObject("Scripting.Dictionary")
    newDict.CompareMode = newMode

    Dim keys As Variant, c As Long
    keys = mCols.Keys
    For c = 0 To mCols.Count - 1
        newDict.Add keys(c), mCols.Item(keys(c))
    Next c
    Set mCols = newDict
    SyncColList

    If Not mRowIndexMap Is Nothing Then
        mIndexDirty = True
    End If
End Property

Public Property Get DuplicateHeadersRenamed() As Variant
    If mRenamedHeaders Is Nothing Then
        DuplicateHeadersRenamed = Array()
        Exit Property
    End If
    DuplicateHeadersRenamed = mRenamedHeaders.Keys
End Property

Public Function ResolveUniqueHeader(ByVal CandidateName As String, _
                                    Optional ByVal AutoRename As Boolean = True) As String
    Dim cleanName As String, testName As String, suffix As Long
    cleanName = Trim$(CandidateName)
    If Len(cleanName) = 0 Then cleanName = "Column"

    If Not mCols.Exists(cleanName) Then
        ResolveUniqueHeader = cleanName
        Exit Function
    End If

    If Not AutoRename Then
        Err.Raise 457, "cDataFrame.ResolveUniqueHeader", "Column already exists: '" & cleanName & "'"
    End If

    suffix = 2
    Do
        testName = cleanName & "_" & suffix
        If Not mCols.Exists(testName) Then
            ResolveUniqueHeader = testName
            If Not mRenamedHeaders.Exists(cleanName) Then
                mRenamedHeaders.Add cleanName, testName
            End If
            Exit Function
        End If
        suffix = suffix + 1
    Loop
End Function

Public Sub ShrinkToFit()
    If mColListCount <> mCols.Count Then
        SyncColList
    End If
    Dim c As Long
    For c = 1 To mColListCount
        mColList(c).ShrinkToFit
    Next c
End Sub

Private Function FormatCompositeKeyPart(ByVal v As Variant) As String
    Dim s As String
    s = SafeKeyString(v)
    FormatCompositeKeyPart = CStr(Len(s)) & ":" & s & "|"
End Function

' ------------------------------------------------------------------------------
' Primary Properties & Default Member
' ------------------------------------------------------------------------------

Public Property Get RowCount() As Long
    RowCount = mRowCount
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = mCols.Count
End Property

' Default member for df(RowIndex, ColumnKey)
Public Property Get Item(ByVal RowIndex As Long, ByVal ColumnKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim col As cColumn
    Set col = mCols.Item(ResolveColName(ColumnKey))

    If IsObject(col.Item(RowIndex)) Then
        Set Item = col.Item(RowIndex)
    Else
        Item = col.Item(RowIndex)
    End If
End Property

Public Property Let Item(ByVal RowIndex As Long, ByVal ColumnKey As Variant, ByVal Value As Variant)
    Dim col As cColumn
    Dim colName As String
    colName = ResolveColName(ColumnKey)
    Set col = mCols.Item(colName)
    col.Item(RowIndex) = Value

    If IsIndexedColumn(colName) Then
        mIndexDirty = True
    End If
End Property

Public Property Set Item(ByVal RowIndex As Long, ByVal ColumnKey As Variant, ByVal Value As Object)
    Dim col As cColumn
    Dim colName As String
    colName = ResolveColName(ColumnKey)
    Set col = mCols.Item(colName)
    Set col.Item(RowIndex) = Value

    If IsIndexedColumn(colName) Then
        mIndexDirty = True
    End If
End Property

' ------------------------------------------------------------------------------
' Dynamic Column API
' ------------------------------------------------------------------------------

' Add a new column (O(1)). Optional DefaultValue or 1D array.
Public Function AddColumn(ByVal Name As String, _
                          Optional ByVal DefaultValue As Variant = Empty, _
                          Optional ByVal Array1D As Variant, _
                          Optional ByVal AutoRename As Boolean = False) As cColumn
    Dim cleanName As String
    cleanName = ResolveUniqueHeader(Name, AutoRename)

    Dim newCol As New cColumn
    newCol.Name = cleanName

    If Not IsMissing(Array1D) And Not IsEmpty(Array1D) Then
        If Not IsArray(Array1D) Then
            Err.Raise 13, "cDataFrame.AddColumn", "Array1D must be an array."
        End If

        Dim lb As Long, ub As Long, arrLen As Long
        lb = LBound(Array1D, 1)
        ub = UBound(Array1D, 1)
        arrLen = ub - lb + 1

        If mCols.Count = 0 And mRowCount = 0 Then
            mRowCount = arrLen
        ElseIf arrLen <> mRowCount Then
            Err.Raise 5, "cDataFrame.AddColumn", _
                "Array length (" & arrLen & ") does not match DataFrame row count (" & mRowCount & ")"
        End If

        newCol.LoadFromArray Array1D
    Else
        newCol.SetCount mRowCount, DefaultValue
    End If

    mCols.Add cleanName, newCol
    SyncColList
    Set AddColumn = newCol
End Function

' Drop column (O(1))
Public Sub DropColumn(ByVal ColumnKey As Variant)
    Dim colName As String
    colName = ResolveColName(ColumnKey)
    mCols.Remove colName

    If IsIndexedColumn(colName) Then
        ClearIndexState
    End If

    If mCols.Count = 0 Then
        mRowCount = 0
    End If
    SyncColList
End Sub

' Rename column (O(1) in-place dictionary key reassignment)
Public Sub RenameColumn(ByVal OldName As String, ByVal NewName As String)
    Dim cleanOld As String, cleanNew As String
    cleanOld = Trim$(OldName)
    cleanNew = Trim$(NewName)

    If Not mCols.Exists(cleanOld) Then
        Err.Raise 5, "cDataFrame.RenameColumn", "Column not found: '" & cleanOld & "'"
    End If

    If Len(cleanNew) = 0 Then
        Err.Raise 5, "cDataFrame.RenameColumn", "New column name cannot be blank."
    End If

    If StrComp(cleanOld, cleanNew, vbTextCompare) <> 0 And mCols.Exists(cleanNew) Then
        Err.Raise 457, "cDataFrame.RenameColumn", "Column already exists: '" & cleanNew & "'"
    End If

    ' Native Scripting.Dictionary O(1) key mutation preserves insertion order!
    mCols.Key(cleanOld) = cleanNew

    Dim col As cColumn
    Set col = mCols.Item(cleanNew)
    col.Name = cleanNew

    If mIndexColCount > 0 Then
        Dim k As Long
        For k = 1 To mIndexColCount
            If StrComp(mIndexColumnNames(k), cleanOld, vbTextCompare) = 0 Then
                mIndexColumnNames(k) = cleanNew
            End If
        Next k
        If mIndexColCount = 1 Then
            mIndexColumnName = mIndexColumnNames(1)
        Else
            mIndexColumnName = Join(mIndexColumnNames, ", ")
        End If
    End If
    SyncColList
End Sub

Public Sub ReorderColumns(ByVal NewOrderArray As Variant)
    If Not IsArray(NewOrderArray) Then
        Err.Raise 13, "cDataFrame.ReorderColumns", "NewOrderArray must be an array."
    End If

    Dim lb As Long, ub As Long, totalNew As Long
    lb = LBound(NewOrderArray)
    ub = UBound(NewOrderArray)
    totalNew = ub - lb + 1

    If totalNew <> mCols.Count Then
        Err.Raise 5, "cDataFrame.ReorderColumns", _
            "Number of columns in NewOrderArray (" & totalNew & ") does not match column count (" & mCols.Count & ")"
    End If

    Dim newDict As Object
    Set newDict = CreateObject("Scripting.Dictionary")
    newDict.CompareMode = 1

    Dim i As Long, cName As String
    For i = lb To ub
        cName = Trim$(CStr(NewOrderArray(i)))
        If Not mCols.Exists(cName) Then
            Err.Raise 5, "cDataFrame.ReorderColumns", "Column not found in existing DataFrame: '" & cName & "'"
        End If
        If newDict.Exists(cName) Then
            Err.Raise 5, "cDataFrame.ReorderColumns", "Duplicate column in NewOrderArray: '" & cName & "'"
        End If
        newDict.Add cName, mCols.Item(cName)
    Next i

    Set mCols = newDict
    SyncColList
End Sub

Public Function HasColumn(ByVal ColumnKey As Variant) As Boolean
    If VarType(ColumnKey) = vbString Then
        HasColumn = mCols.Exists(Trim$(CStr(ColumnKey)))
    ElseIf IsNumeric(ColumnKey) Then
        Dim idx As Long
        idx = CLng(ColumnKey)
        HasColumn = (idx >= 1 And idx <= mCols.Count)
    Else
        HasColumn = False
    End If
End Function

Public Function Column(ByVal ColumnKey As Variant) As cColumn
    Set Column = mCols.Item(ResolveColName(ColumnKey))
End Function

Public Function ColumnNames() As String()
    If mCols.Count = 0 Then
        ColumnNames = Split("")
        Exit Function
    End If

    Dim keys As Variant
    keys = mCols.Keys

    Dim res() As String
    ReDim res(1 To mCols.Count)

    Dim i As Long
    For i = 1 To mCols.Count
        res(i) = CStr(keys(i - 1))
    Next i

    ColumnNames = res
End Function

' ------------------------------------------------------------------------------
' Dynamic Row API
' ------------------------------------------------------------------------------

' Add row with variable values or 1D array. Auto-expands buffer geometrically across columns.
Public Function AddRow(ParamArray Values() As Variant) As Long
    mRowCount = mRowCount + 1

    If mColListCount <> mCols.Count Then
        SyncColList
    End If

    Dim isArrayPassed As Boolean
    Dim passedArr As Variant

    If UBound(Values) >= LBound(Values) Then
        If IsArray(Values(LBound(Values))) Then
            isArrayPassed = True
            passedArr = Values(LBound(Values))
        End If
    End If

    Dim c As Long
    If isArrayPassed Then
        Dim arrLb As Long, arrUb As Long
        arrLb = LBound(passedArr)
        arrUb = UBound(passedArr)
        For c = 1 To mColListCount
            If (arrLb + c - 1) <= arrUb Then
                mColList(c).Append passedArr(arrLb + c - 1)
            Else
                mColList(c).Append Empty
            End If
        Next c
    Else
        Dim numVals As Long
        numVals = UBound(Values) - LBound(Values) + 1
        For c = 1 To mColListCount
            If c <= numVals Then
                mColList(c).Append Values(LBound(Values) + c - 1)
            Else
                mColList(c).Append Empty
            End If
        Next c
    End If

    If Len(mIndexColumnName) > 0 Then
        mIndexDirty = True
    End If

    AddRow = mRowCount
End Function

' High-performance bulk 2D array row append. Pre-allocates column buffers upfront
' and copies column-by-column in tight loops without per-row dictionary lookup overhead.
Public Function AddRows(ByVal Source2DArray As Variant) As Long
    If Not IsArray(Source2DArray) Then
        Err.Raise 13, "cDataFrame.AddRows", "Source2DArray must be a 2D array."
    End If

    If mCols.Count = 0 Then
        Err.Raise 5, "cDataFrame.AddRows", "Cannot AddRows to a DataFrame with 0 columns."
    End If

    Dim rLb As Long, rUb As Long, cLb As Long, cUb As Long
    On Error Resume Next
    rLb = LBound(Source2DArray, 1)
    rUb = UBound(Source2DArray, 1)
    cLb = LBound(Source2DArray, 2)
    cUb = UBound(Source2DArray, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 5, "cDataFrame.AddRows", "Source2DArray must be a 2-dimensional array."
    End If
    On Error GoTo 0

    Dim numNewRows As Long, numSrcCols As Long
    numNewRows = rUb - rLb + 1
    numSrcCols = cUb - cLb + 1

    If numNewRows <= 0 Then
        AddRows = mRowCount
        Exit Function
    End If

    Dim colCount As Long
    colCount = mCols.Count
    Dim keys As Variant
    keys = mCols.Keys

    Dim colObjs() As cColumn
    ReDim colObjs(1 To colCount)

    Dim newTotalRows As Long
    newTotalRows = mRowCount + numNewRows

    Dim c As Long, r As Long
    For c = 1 To colCount
        Set colObjs(c) = mCols.Item(keys(c - 1))
        colObjs(c).EnsureCapacity newTotalRows
        colObjs(c).SetCount newTotalRows
    Next c

    Dim copyCols As Long
    If numSrcCols < colCount Then
        copyCols = numSrcCols
    Else
        copyCols = colCount
    End If

    Dim curCol As cColumn, srcColIdx As Long, baseRow As Long
    baseRow = mRowCount

    For c = 1 To copyCols
        Set curCol = colObjs(c)
        srcColIdx = cLb + c - 1
        For r = 1 To numNewRows
            curCol.Item(baseRow + r) = Source2DArray(rLb + r - 1, srcColIdx)
        Next r
    Next c

    mRowCount = newTotalRows

    If mIndexColCount > 0 Then
        mIndexDirty = True
    End If

    AddRows = mRowCount
End Function

Public Sub InsertRow(ByVal TargetIndex As Long, ParamArray Values() As Variant)
    If TargetIndex < 1 Or TargetIndex > mRowCount + 1 Then
        Err.Raise 9, "cDataFrame.InsertRow", _
            "TargetIndex out of range: " & TargetIndex & " (Valid: 1 To " & (mRowCount + 1) & ")"
    End If

    Dim keys As Variant
    keys = mCols.Keys

    Dim isArrayPassed As Boolean
    Dim passedArr As Variant

    If UBound(Values) >= LBound(Values) Then
        If IsArray(Values(LBound(Values))) Then
            isArrayPassed = True
            passedArr = Values(LBound(Values))
        End If
    End If

    Dim c As Long, col As cColumn
    For c = 0 To mCols.Count - 1
        Set col = mCols.Item(keys(c))
        col.ShiftDown TargetIndex
        If isArrayPassed Then
            If (LBound(passedArr) + c) <= UBound(passedArr) Then
                col.Item(TargetIndex) = passedArr(LBound(passedArr) + c)
            Else
                col.Item(TargetIndex) = Empty
            End If
        Else
            If (LBound(Values) + c) <= UBound(Values) Then
                col.Item(TargetIndex) = Values(LBound(Values) + c)
            Else
                col.Item(TargetIndex) = Empty
            End If
        End If
    Next c

    mRowCount = mRowCount + 1

    If Len(mIndexColumnName) > 0 Then
        mIndexDirty = True
    End If
End Sub

Public Sub DeleteRow(ByVal TargetIndex As Long)
    If TargetIndex < 1 Or TargetIndex > mRowCount Then
        Err.Raise 9, "cDataFrame.DeleteRow", _
            "TargetIndex out of range: " & TargetIndex & " (Valid: 1 To " & mRowCount & ")"
    End If

    Dim keys As Variant
    keys = mCols.Keys

    Dim c As Long, col As cColumn
    For c = 0 To mCols.Count - 1
        Set col = mCols.Item(keys(c))
        col.ShiftUp TargetIndex
    Next c

    mRowCount = mRowCount - 1

    If Len(mIndexColumnName) > 0 Then
        mIndexDirty = True
    End If
End Sub

' ------------------------------------------------------------------------------
' Matrix Lookup & 2-Way Index API
' ------------------------------------------------------------------------------

' Pure VBA linear scan on 1D column array. Returns CVErr(xlErrNA) / CVErr(2042) if not found.
Public Function Lookup(ByVal RowKeyValue As Variant, _
                       ByVal KeyColumn As Variant, _
                       ByVal TargetColumn As Variant) As Variant
    Dim keyCol As cColumn, tgtCol As cColumn
    Set keyCol = mCols.Item(ResolveColName(KeyColumn))
    Set tgtCol = mCols.Item(ResolveColName(TargetColumn))

    Dim r As Long
    r = keyCol.Find(RowKeyValue)
    If r > 0 Then
        If IsObject(tgtCol.Item(r)) Then
            Set Lookup = tgtCol.Item(r)
        Else
            Lookup = tgtCol.Item(r)
        End If
    Else
        Lookup = CVErr(2042)
    End If
End Function

' Builds mRowIndexMap Scripting.Dictionary mapping unique single-column or composite
' multi-column row keys (Array("Col1", "Col2")) to row index 1..N in O(N).
Public Sub SetIndex(ByVal ColumnNames As Variant, Optional ByVal SkipEmptyKeys As Boolean = True)
    Dim resolvedNames() As String
    Dim numIdxCols As Long
    Dim k As Long

    If IsArray(ColumnNames) Then
        Dim lb As Long, ub As Long
        lb = LBound(ColumnNames)
        ub = UBound(ColumnNames)
        numIdxCols = ub - lb + 1
        If numIdxCols <= 0 Then
            Err.Raise 5, "cDataFrame.SetIndex", "ColumnNames array cannot be empty."
        End If
        ReDim resolvedNames(1 To numIdxCols)
        For k = 1 To numIdxCols
            resolvedNames(k) = ResolveColName(ColumnNames(lb + k - 1))
        Next k
    Else
        numIdxCols = 1
        ReDim resolvedNames(1 To 1)
        resolvedNames(1) = ResolveColName(ColumnNames)
    End If

    Dim idxCols() As cColumn
    ReDim idxCols(1 To numIdxCols)
    For k = 1 To numIdxCols
        Set idxCols(k) = mCols.Item(resolvedNames(k))
    Next k

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = IIf(mCaseSensitive, 0, 1)

    Dim r As Long, compKey As String, partKey As String, allEmpty As Boolean
    For r = 1 To mRowCount
        If numIdxCols = 1 Then
            compKey = Trim$(SafeKeyString(idxCols(1).Item(r)))
            allEmpty = (Len(compKey) = 0)
        Else
            compKey = ""
            allEmpty = True
            For k = 1 To numIdxCols
                partKey = SafeKeyString(idxCols(k).Item(r))
                If Len(Trim$(partKey)) > 0 Then allEmpty = False
                compKey = compKey & FormatCompositeKeyPart(partKey)
            Next k
        End If

        If SkipEmptyKeys And allEmpty Then
            ' Skip empty spacer rows or team header separators
        Else
            If dict.Exists(compKey) Then
                Err.Raise 457, "cDataFrame.SetIndex", _
                    "Duplicate key found in index (" & Join(resolvedNames, ", ") & "): '" & compKey & "' at row " & r
            End If
            dict.Add compKey, r
        End If
    Next r

    Set mRowIndexMap = dict
    mIndexColCount = numIdxCols
    ReDim mIndexColumnNames(1 To numIdxCols)
    For k = 1 To numIdxCols
        mIndexColumnNames(k) = resolvedNames(k)
    Next k

    If numIdxCols = 1 Then
        mIndexColumnName = resolvedNames(1)
    Else
        mIndexColumnName = Join(resolvedNames, ", ")
    End If
    mIndexSkipEmpty = SkipEmptyKeys
    mIndexDirty = False
End Sub

Public Function HasIndexKey(ByVal RowKey As Variant) As Boolean
    If mIndexColCount = 0 Or mRowIndexMap Is Nothing Then Exit Function
    If mIndexDirty Then RebuildIndex
    HasIndexKey = mRowIndexMap.Exists(ResolveLocKeyString(RowKey))
End Function

Public Property Get IndexColumnName() As String
    IndexColumnName = mIndexColumnName
End Property

' Fast O(1) 2-way matrix lookup by RowKey (scalar or 1D Array composite key) and ColumnKey
Public Property Get Loc(ByVal RowKey As Variant, ByVal ColumnKey As Variant) As Variant
    If mIndexColCount = 0 Or mRowIndexMap Is Nothing Then
        Err.Raise 5, "cDataFrame.Loc", "Index has not been set. Call SetIndex first."
    End If

    If mIndexDirty Then
        RebuildIndex
    End If

    Dim sKey As String
    sKey = ResolveLocKeyString(RowKey)

    If Not mRowIndexMap.Exists(sKey) Then
        Loc = CVErr(2042)
        Exit Property
    End If

    Dim r As Long
    r = mRowIndexMap.Item(sKey)

    Dim col As cColumn
    Set col = mCols.Item(ResolveColName(ColumnKey))

    If IsObject(col.Item(r)) Then
        Set Loc = col.Item(r)
    Else
        Loc = col.Item(r)
    End If
End Property

' Fast O(1) cell mutation by matrix key (scalar or 1D Array composite key)
Public Property Let Loc(ByVal RowKey As Variant, ByVal ColumnKey As Variant, ByVal Value As Variant)
    If mIndexColCount = 0 Or mRowIndexMap Is Nothing Then
        Err.Raise 5, "cDataFrame.Loc", "Index has not been set. Call SetIndex first."
    End If

    If mIndexDirty Then
        RebuildIndex
    End If

    Dim sKey As String
    sKey = ResolveLocKeyString(RowKey)

    If Not mRowIndexMap.Exists(sKey) Then
        Err.Raise 5, "cDataFrame.Loc", "Row key not found in index: '" & Replace(sKey, Chr$(30), "|") & "'"
    End If

    Dim r As Long
    r = mRowIndexMap.Item(sKey)

    Dim colName As String
    colName = ResolveColName(ColumnKey)

    Dim col As cColumn
    Set col = mCols.Item(colName)
    col.Item(r) = Value

    If IsIndexedColumn(colName) Then
        mIndexDirty = True
    End If
End Property

Public Property Set Loc(ByVal RowKey As Variant, ByVal ColumnKey As Variant, ByVal Value As Object)
    If mIndexColCount = 0 Or mRowIndexMap Is Nothing Then
        Err.Raise 5, "cDataFrame.Loc", "Index has not been set. Call SetIndex first."
    End If

    If mIndexDirty Then
        RebuildIndex
    End If

    Dim sKey As String
    sKey = ResolveLocKeyString(RowKey)

    If Not mRowIndexMap.Exists(sKey) Then
        Err.Raise 5, "cDataFrame.Loc", "Row key not found in index: '" & Replace(sKey, Chr$(30), "|") & "'"
    End If

    Dim r As Long
    r = mRowIndexMap.Item(sKey)

    Dim colName As String
    colName = ResolveColName(ColumnKey)

    Dim col As cColumn
    Set col = mCols.Item(colName)
    Set col.Item(r) = Value

    If IsIndexedColumn(colName) Then
        mIndexDirty = True
    End If
End Property

Private Function ResolveLocKeyString(ByVal RowKey As Variant) As String
    If IsArray(RowKey) Then
        Dim lb As Long, ub As Long, numK As Long, k As Long
        lb = LBound(RowKey)
        ub = UBound(RowKey)
        numK = ub - lb + 1
        If numK <> mIndexColCount Then
            Err.Raise 5, "cDataFrame.Loc", _
                "Composite RowKey length (" & numK & ") does not match index column count (" & mIndexColCount & ")."
        End If
        Dim comp As String
        comp = ""
        For k = 1 To numK
            comp = comp & FormatCompositeKeyPart(RowKey(lb + k - 1))
        Next k
        ResolveLocKeyString = comp
    Else
        If mIndexColCount > 1 Then
            Err.Raise 5, "cDataFrame.Loc", _
                "Index is a " & mIndexColCount & "-column composite index (" & mIndexColumnName & "). Pass Array(...) as RowKey."
        End If
        ResolveLocKeyString = Trim$(SafeKeyString(RowKey))
    End If
End Function

Private Function IsIndexedColumn(ByVal ColName As String) As Boolean
    If mIndexColCount <= 0 Then Exit Function
    Dim k As Long
    For k = 1 To mIndexColCount
        If StrComp(mIndexColumnNames(k), ColName, vbTextCompare) = 0 Then
            IsIndexedColumn = True
            Exit Function
        End If
    Next k
End Function

Private Sub ClearIndexState()
    mIndexColumnName = ""
    mIndexColCount = 0
    Erase mIndexColumnNames
    Set mRowIndexMap = Nothing
    mIndexDirty = False
End Sub

Private Sub RebuildIndex()
    If mIndexColCount <= 0 Then
        ClearIndexState
        Exit Sub
    End If

    Dim k As Long
    For k = 1 To mIndexColCount
        If Not mCols.Exists(mIndexColumnNames(k)) Then
            ClearIndexState
            Exit Sub
        End If
    Next k

    If mIndexColCount = 1 Then
        SetIndex mIndexColumnNames(1), mIndexSkipEmpty
    Else
        SetIndex mIndexColumnNames, mIndexSkipEmpty
    End If
End Sub

' ------------------------------------------------------------------------------
' In-Memory Sort Engine (Pure VBA Indirect QuickSort)
' ------------------------------------------------------------------------------

' Multi-column sorting. SortColumns can be String, Long, or Array.
' Ascending can be single Boolean or Array of Booleans.
Public Sub Sort(ByVal SortColumns As Variant, Optional ByVal Ascending As Variant = True)
    If mRowCount <= 1 Or mCols.Count = 0 Then Exit Sub

    ' 1. Parse SortColumns
    Dim sortColNames() As String
    Dim numKeys As Long
    Dim k As Long

    If IsArray(SortColumns) Then
        Dim lbC As Long, ubC As Long
        lbC = LBound(SortColumns)
        ubC = UBound(SortColumns)
        numKeys = ubC - lbC + 1
        If numKeys <= 0 Then Exit Sub
        ReDim sortColNames(1 To numKeys)
        For k = 1 To numKeys
            sortColNames(k) = ResolveColName(SortColumns(lbC + k - 1))
        Next k
    Else
        numKeys = 1
        ReDim sortColNames(1 To 1)
        sortColNames(1) = ResolveColName(SortColumns)
    End If

    ' 2. Parse Ascending flags
    Dim ascFlags() As Boolean
    ReDim ascFlags(1 To numKeys)

    If IsArray(Ascending) Then
        Dim lbA As Long, ubA As Long
        lbA = LBound(Ascending)
        ubA = UBound(Ascending)
        Dim numA As Long
        numA = ubA - lbA + 1
        If numA <> numKeys Then
            Err.Raise 5, "cDataFrame.Sort", _
                "Length of Ascending array (" & numA & ") does not match SortColumns count (" & numKeys & ")."
        End If
        For k = 1 To numKeys
            ascFlags(k) = CBool(Ascending(lbA + k - 1))
        Next k
    Else
        Dim singleAsc As Boolean
        singleAsc = CBool(Ascending)
        For k = 1 To numKeys
            ascFlags(k) = singleAsc
        Next k
    End If

    ' 3. Resolve Column objects for direct fast access
    Dim sortColObjs() As cColumn
    ReDim sortColObjs(1 To numKeys)
    For k = 1 To numKeys
        Set sortColObjs(k) = mCols.Item(sortColNames(k))
    Next k

    ' 4. Initialize index array 1..mRowCount
    Dim idx() As Long
    ReDim idx(1 To mRowCount)
    Dim i As Long
    For i = 1 To mRowCount
        idx(i) = i
    Next i

    ' 5. Run indirect QuickSort
    QuickSortIndices idx, 1, mRowCount, sortColObjs, ascFlags, numKeys

    ' 6. Permute each column buffer in memory
    Dim colKeys As Variant
    colKeys = mCols.Keys
    Dim c As Long, col As cColumn
    For c = 0 To mCols.Count - 1
        Set col = mCols.Item(colKeys(c))
        col.Permute idx
    Next c

    ' 7. Mark index dirty if index is active
    If Len(mIndexColumnName) > 0 Then
        mIndexDirty = True
    End If
End Sub

Private Sub QuickSortIndices(ByRef idx() As Long, ByVal low As Long, ByVal high As Long, _
                             ByRef sortCols() As cColumn, ByRef ascFlags() As Boolean, ByVal numKeys As Long)
    Do While low < high
        Dim i As Long, j As Long
        Dim pivotRow As Long
        i = low
        j = high
        pivotRow = idx((low + high) \ 2)

        Do While i <= j
            Do While CompareRows(idx(i), pivotRow, sortCols, ascFlags, numKeys) < 0
                i = i + 1
            Loop
            Do While CompareRows(idx(j), pivotRow, sortCols, ascFlags, numKeys) > 0
                j = j - 1
            Loop
            If i <= j Then
                Dim temp As Long
                temp = idx(i)
                idx(i) = idx(j)
                idx(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop

        ' Tail-call optimization: recurse on smaller half, loop on larger half
        ' Guarantees call stack depth <= log2(N) (never exceeds 20 frames for 1M rows)
        If (j - low) < (high - i) Then
            If low < j Then QuickSortIndices idx, low, j, sortCols, ascFlags, numKeys
            low = i
        Else
            If i < high Then QuickSortIndices idx, i, high, sortCols, ascFlags, numKeys
            high = j
        End If
    Loop
End Sub

Private Function CompareRows(ByVal rowA As Long, ByVal rowB As Long, _
                             ByRef sortCols() As cColumn, ByRef ascFlags() As Boolean, ByVal numKeys As Long) As Long
    If rowA = rowB Then
        CompareRows = 0
        Exit Function
    End If

    Dim k As Long, cmp As Long
    For k = 1 To numKeys
        Dim vA As Variant, vB As Variant
        vA = sortCols(k).Item(rowA)
        vB = sortCols(k).Item(rowB)

        cmp = CompareSortValues(vA, vB)
        If cmp <> 0 Then
            If Not ascFlags(k) Then cmp = -cmp
            CompareRows = cmp
            Exit Function
        End If
    Next k

    ' Deterministic tie-breaker by row index
    If rowA < rowB Then
        CompareRows = -1
    Else
        CompareRows = 1
    End If
End Function

Private Function CompareSortValues(ByVal vA As Variant, ByVal vB As Variant) As Long
    On Error GoTo CatchErr

    Dim aErr As Boolean, bErr As Boolean
    aErr = IsError(vA)
    bErr = IsError(vB)
    If aErr Or bErr Then
        If aErr And bErr Then
            If CInt(vA) < CInt(vB) Then
                CompareSortValues = -1
            ElseIf CInt(vA) > CInt(vB) Then
                CompareSortValues = 1
            Else
                CompareSortValues = 0
            End If
        ElseIf aErr Then
            CompareSortValues = 1
        Else
            CompareSortValues = -1
        End If
        Exit Function
    End If

    Dim aEmpty As Boolean, bEmpty As Boolean
    aEmpty = IsEmpty(vA) Or IsNull(vA)
    bEmpty = IsEmpty(vB) Or IsNull(vB)
    If aEmpty And bEmpty Then
        CompareSortValues = 0
        Exit Function
    ElseIf aEmpty Then
        CompareSortValues = -1
        Exit Function
    ElseIf bEmpty Then
        CompareSortValues = 1
        Exit Function
    End If

    If IsNumeric(vA) And IsNumeric(vB) Then
        Dim dblA As Double, dblB As Double
        dblA = CDbl(vA)
        dblB = CDbl(vB)
        If dblA < dblB Then
            CompareSortValues = -1
        ElseIf dblA > dblB Then
            CompareSortValues = 1
        Else
            CompareSortValues = 0
        End If
    ElseIf IsDate(vA) And IsDate(vB) Then
        Dim dtA As Date, dtB As Date
        dtA = CDate(vA)
        dtB = CDate(vB)
        If dtA < dtB Then
            CompareSortValues = -1
        ElseIf dtA > dtB Then
            CompareSortValues = 1
        Else
            CompareSortValues = 0
        End If
    Else
        Dim sc As Long
        sc = StrComp(CStr(vA), CStr(vB), IIf(mCaseSensitive, vbBinaryCompare, vbTextCompare))
        If sc < 0 Then
            CompareSortValues = -1
        ElseIf sc > 0 Then
            CompareSortValues = 1
        Else
            CompareSortValues = 0
        End If
    End If
    Exit Function
CatchErr:
    CompareSortValues = 0
End Function

' ------------------------------------------------------------------------------
' Relational Algebra: Hash Join / Merge
' ------------------------------------------------------------------------------

' Fast O(N + M) table merge using Scripting.Dictionary hash index.
' OnKey accepts a single column key (String/Long) or a 1D Array of column keys.
' Supports "INNER", "LEFT", "RIGHT", "FULL"/"OUTER", "LEFT_ANTI"/"ANTI", and "LEFT_SEMI"/"SEMI".
Public Function Merge(ByVal RightTable As cDataFrame, _
                      ByVal OnKey As Variant, _
                      Optional ByVal How As String = "INNER") As cDataFrame
    If RightTable Is Nothing Then
        Err.Raise 91, "cDataFrame.Merge", "RightTable cannot be Nothing."
    End If

    Dim joinType As String
    joinType = UCase$(Trim$(How))
    If joinType = "OUTER" Then joinType = "FULL"
    If joinType = "ANTI" Then joinType = "LEFT_ANTI"
    If joinType = "SEMI" Then joinType = "LEFT_SEMI"

    If joinType <> "INNER" And joinType <> "LEFT" And joinType <> "RIGHT" And _
       joinType <> "FULL" And joinType <> "LEFT_ANTI" And joinType <> "LEFT_SEMI" Then
        Err.Raise 5, "cDataFrame.Merge", _
            "Unsupported join type: '" & How & "'. Supported: INNER, LEFT, RIGHT, FULL, LEFT_ANTI, LEFT_SEMI"
    End If

    ' 1. Parse and validate join key(s)
    Dim leftKeyNames() As String
    Dim rightKeyNames() As String
    Dim numKeys As Long
    Dim k As Long

    If IsArray(OnKey) Then
        Dim lbK As Long, ubK As Long
        lbK = LBound(OnKey)
        ubK = UBound(OnKey)
        numKeys = ubK - lbK + 1
        If numKeys <= 0 Then
            Err.Raise 5, "cDataFrame.Merge", "OnKey array cannot be empty."
        End If
        ReDim leftKeyNames(1 To numKeys)
        ReDim rightKeyNames(1 To numKeys)
        For k = 1 To numKeys
            leftKeyNames(k) = ResolveColName(OnKey(lbK + k - 1))
            rightKeyNames(k) = RightTable.ResolveColName(OnKey(lbK + k - 1))
        Next k
    Else
        numKeys = 1
        ReDim leftKeyNames(1 To 1)
        ReDim rightKeyNames(1 To 1)
        leftKeyNames(1) = ResolveColName(OnKey)
        rightKeyNames(1) = RightTable.ResolveColName(OnKey)
    End If

    Dim leftKeyCols() As cColumn
    Dim rightKeyCols() As cColumn
    ReDim leftKeyCols(1 To numKeys)
    ReDim rightKeyCols(1 To numKeys)

    Dim leftKeySet As Object, rightKeySet As Object
    Set leftKeySet = CreateObject("Scripting.Dictionary")
    leftKeySet.CompareMode = 1
    Set rightKeySet = CreateObject("Scripting.Dictionary")
    rightKeySet.CompareMode = 1

    For k = 1 To numKeys
        Set leftKeyCols(k) = mCols.Item(leftKeyNames(k))
        Set rightKeyCols(k) = RightTable.Column(rightKeyNames(k))
        leftKeySet.Add leftKeyNames(k), k
        rightKeySet.Add rightKeyNames(k), k
    Next k

    ' 2. Fast path for LEFT_ANTI and LEFT_SEMI (returns only LeftTable schema)
    If joinType = "LEFT_ANTI" Or joinType = "LEFT_SEMI" Then
        Dim semiDf As cDataFrame
        Set semiDf = Me.Clone(True)
        If mRowCount = 0 Then
            Set Merge = semiDf
            Exit Function
        End If

        Dim rightDistinctKeys As Object
        Set rightDistinctKeys = CreateObject("Scripting.Dictionary")
        rightDistinctKeys.CompareMode = 1

        Dim rR As Long, rK As String
        For rR = 1 To RightTable.RowCount
            If numKeys = 1 Then
                rK = SafeKeyString(rightKeyCols(1).Item(rR))
            Else
                rK = ""
                For k = 1 To numKeys
                    rK = rK & FormatCompositeKeyPart(rightKeyCols(k).Item(rR))
                Next k
            End If
            If Not rightDistinctKeys.Exists(rK) Then
                rightDistinctKeys.Add rK, True
            End If
        Next rR

        Dim numLeftCols As Long
        numLeftCols = mCols.Count
        Dim leftColObjs() As cColumn
        ReDim leftColObjs(0 To numLeftCols - 1)
        Dim lKeys As Variant
        lKeys = mCols.Keys
        Dim cIdx As Long
        For cIdx = 0 To numLeftCols - 1
            Set leftColObjs(cIdx) = mCols.Item(lKeys(cIdx))
        Next cIdx

        Dim semiRowVals() As Variant
        ReDim semiRowVals(0 To numLeftCols - 1)

        Dim rL As Long, lK As String, existsInRight As Boolean
        For rL = 1 To mRowCount
            If numKeys = 1 Then
                lK = SafeKeyString(leftKeyCols(1).Item(rL))
            Else
                lK = ""
                For k = 1 To numKeys
                    lK = lK & FormatCompositeKeyPart(leftKeyCols(k).Item(rL))
                Next k
            End If
            existsInRight = rightDistinctKeys.Exists(lK)

            If (joinType = "LEFT_SEMI" And existsInRight) Or _
               (joinType = "LEFT_ANTI" And Not existsInRight) Then
                For cIdx = 0 To numLeftCols - 1
                    semiRowVals(cIdx) = leftColObjs(cIdx).Item(rL)
                Next cIdx
                semiDf.AddRow semiRowVals
            End If
        Next rL

        Set Merge = semiDf
        Exit Function
    End If

    ' 3. Determine output column schema and collision renaming for INNER / LEFT / RIGHT / FULL
    Dim resultDf As New cDataFrame

    Dim leftNames() As String, rightNames() As String
    leftNames = Me.ColumnNames()
    rightNames = RightTable.ColumnNames()

    Dim numL As Long, numR As Long
    numL = Me.ColumnCount
    numR = RightTable.ColumnCount

    Dim outCols As Object
    Set outCols = CreateObject("Scripting.Dictionary")
    outCols.CompareMode = 1

    Dim l As Long, r As Long
    Dim colName As String, resName As String

    ' Add Left columns
    For l = 1 To numL
        colName = leftNames(l)
        If leftKeySet.Exists(colName) Then
            resName = colName
            resultDf.AddColumn resName
            outCols.Add resName, Array(1, colName, CLng(leftKeySet.Item(colName)))
        Else
            If RightTable.HasColumn(colName) And Not rightKeySet.Exists(colName) Then
                resName = colName & "_left"
            Else
                resName = colName
            End If
            resultDf.AddColumn resName
            outCols.Add resName, Array(1, colName, 0&)
        End If
    Next l

    ' Add Right columns (excluding join keys)
    For r = 1 To numR
        colName = rightNames(r)
        If Not rightKeySet.Exists(colName) Then
            If Me.HasColumn(colName) Then
                resName = colName & "_right"
            Else
                resName = colName
            End If

            Dim finalResName As String, dupCount As Long
            finalResName = resName
            dupCount = 1
            Do While resultDf.HasColumn(finalResName)
                dupCount = dupCount + 1
                finalResName = resName & "_" & dupCount
            Loop
            resultDf.AddColumn finalResName
            outCols.Add finalResName, Array(2, colName, 0&)
        End If
    Next r

    ' Pre-resolve column pointers for zero-overhead row assembly
    Dim outCount As Long
    outCount = resultDf.ColumnCount

    Dim outSources() As Long
    Dim outKeyIdx() As Long
    Dim outLeftCols() As cColumn
    Dim outRightCols() As cColumn
    ReDim outSources(0 To outCount - 1)
    ReDim outKeyIdx(0 To outCount - 1)
    ReDim outLeftCols(0 To outCount - 1)
    ReDim outRightCols(0 To outCount - 1)

    Dim outKeys As Variant
    outKeys = outCols.Keys
    Dim i As Long, info As Variant
    For i = 0 To outCount - 1
        info = outCols.Item(outKeys(i))
        outSources(i) = CLng(info(0))
        outKeyIdx(i) = CLng(info(2))
        If outSources(i) = 1 Then
            Set outLeftCols(i) = mCols.Item(CStr(info(1)))
        Else
            Set outRightCols(i) = RightTable.Column(CStr(info(1)))
        End If
    Next i

    If mRowCount = 0 And RightTable.RowCount = 0 Then
        Set Merge = resultDf
        Exit Function
    End If

    If mRowCount = 0 And (joinType = "INNER" Or joinType = "LEFT") Then
        Set Merge = resultDf
        Exit Function
    End If

    ' 4. Build Hash Index on RightTable: composite Key -> Collection of row indices
    Dim rightHash As Object
    Set rightHash = CreateObject("Scripting.Dictionary")
    rightHash.CompareMode = 1

    Dim rRight As Long, kRight As String, rowList As Collection
    For rRight = 1 To RightTable.RowCount
        If numKeys = 1 Then
            kRight = SafeKeyString(rightKeyCols(1).Item(rRight))
        Else
            kRight = ""
            For k = 1 To numKeys
                kRight = kRight & FormatCompositeKeyPart(rightKeyCols(k).Item(rRight))
            Next k
        End If
        If Not rightHash.Exists(kRight) Then
            Set rowList = New Collection
            rightHash.Add kRight, rowList
        Else
            Set rowList = rightHash.Item(kRight)
        End If
        rowList.Add rRight
    Next rRight

    Dim rightMatched() As Boolean
    If RightTable.RowCount > 0 And (joinType = "RIGHT" Or joinType = "FULL") Then
        ReDim rightMatched(1 To RightTable.RowCount)
    End If

    ' 5. Probe Hash Index with LeftTable rows
    Dim rLeft As Long, kLeft As String
    Dim matchRows As Collection, matchedIdx As Variant
    Dim rowVals() As Variant
    ReDim rowVals(0 To outCount - 1)

    For rLeft = 1 To mRowCount
        If numKeys = 1 Then
            kLeft = SafeKeyString(leftKeyCols(1).Item(rLeft))
        Else
            kLeft = ""
            For k = 1 To numKeys
                kLeft = kLeft & FormatCompositeKeyPart(leftKeyCols(k).Item(rLeft))
            Next k
        End If

        If rightHash.Exists(kLeft) Then
            Set matchRows = rightHash.Item(kLeft)
            For Each matchedIdx In matchRows
                rRight = CLng(matchedIdx)
                If joinType = "RIGHT" Or joinType = "FULL" Then
                    rightMatched(rRight) = True
                End If
                For i = 0 To outCount - 1
                    If outSources(i) = 1 Then
                        rowVals(i) = outLeftCols(i).Item(rLeft)
                    Else
                        rowVals(i) = outRightCols(i).Item(rRight)
                    End If
                Next i
                resultDf.AddRow rowVals
            Next matchedIdx
        ElseIf joinType = "LEFT" Or joinType = "FULL" Then
            For i = 0 To outCount - 1
                If outSources(i) = 1 Then
                    rowVals(i) = outLeftCols(i).Item(rLeft)
                Else
                    rowVals(i) = Empty
                End If
            Next i
            resultDf.AddRow rowVals
        End If
    Next rLeft

    ' 6. Emit unmatched RightTable rows for RIGHT and FULL outer joins
    If (joinType = "RIGHT" Or joinType = "FULL") And RightTable.RowCount > 0 Then
        For rRight = 1 To RightTable.RowCount
            If Not rightMatched(rRight) Then
                For i = 0 To outCount - 1
                    If outSources(i) = 2 Then
                        rowVals(i) = outRightCols(i).Item(rRight)
                    ElseIf outKeyIdx(i) > 0 Then
                        rowVals(i) = rightKeyCols(outKeyIdx(i)).Item(rRight)
                    Else
                        rowVals(i) = Empty
                    End If
                Next i
                resultDf.AddRow rowVals
            End If
        Next rRight
    End If

    Set Merge = resultDf
End Function

' ------------------------------------------------------------------------------
' Relational Algebra: GroupBy & Pivot Aggregations (O(N) Hash Grouping)
' ------------------------------------------------------------------------------

' Fast O(N) two-pass hash grouping and aggregation.
' GroupColumns: single column key (String/Long) or 1D Array of column keys.
' AggColumns:   single column key (String/Long) or 1D Array of column keys.
' AggFuncs:     single function name or 1D Array matching AggColumns length.
'               Supported: "SUM", "COUNT", "AVG" / "AVERAGE", "MIN", "MAX".
Public Function GroupBy(ByVal GroupColumns As Variant, _
                        ByVal AggColumns As Variant, _
                        Optional ByVal AggFuncs As Variant = "SUM") As cDataFrame
    ' 1. Parse and validate GroupColumns
    Dim groupColNames() As String
    Dim numGroupCols As Long
    Dim g As Long

    If IsArray(GroupColumns) Then
        Dim lbG As Long, ubG As Long
        lbG = LBound(GroupColumns)
        ubG = UBound(GroupColumns)
        numGroupCols = ubG - lbG + 1
        If numGroupCols <= 0 Then
            Err.Raise 5, "cDataFrame.GroupBy", "GroupColumns array cannot be empty."
        End If
        ReDim groupColNames(1 To numGroupCols)
        For g = 1 To numGroupCols
            groupColNames(g) = ResolveColName(GroupColumns(lbG + g - 1))
        Next g
    Else
        numGroupCols = 1
        ReDim groupColNames(1 To 1)
        groupColNames(1) = ResolveColName(GroupColumns)
    End If

    ' 2. Parse and validate AggColumns
    Dim aggColNames() As String
    Dim numAggs As Long
    Dim a As Long

    If IsArray(AggColumns) Then
        Dim lbA As Long, ubA As Long
        lbA = LBound(AggColumns)
        ubA = UBound(AggColumns)
        numAggs = ubA - lbA + 1
        If numAggs <= 0 Then
            Err.Raise 5, "cDataFrame.GroupBy", "AggColumns array cannot be empty."
        End If
        ReDim aggColNames(1 To numAggs)
        For a = 1 To numAggs
            aggColNames(a) = ResolveColName(AggColumns(lbA + a - 1))
        Next a
    Else
        numAggs = 1
        ReDim aggColNames(1 To 1)
        aggColNames(1) = ResolveColName(AggColumns)
    End If

    ' 3. Parse and normalize AggFuncs
    Dim normFuncs() As String
    ReDim normFuncs(1 To numAggs)

    If IsArray(AggFuncs) Then
        Dim lbF As Long, ubF As Long, numF As Long
        lbF = LBound(AggFuncs)
        ubF = UBound(AggFuncs)
        numF = ubF - lbF + 1
        If numF <> numAggs Then
            Err.Raise 5, "cDataFrame.GroupBy", _
                "Length of AggFuncs array (" & numF & ") does not match AggColumns count (" & numAggs & ")."
        End If
        For a = 1 To numAggs
            normFuncs(a) = NormalizeAggFunc(CStr(AggFuncs(lbF + a - 1)), "cDataFrame.GroupBy")
        Next a
    Else
        Dim singleFunc As String
        singleFunc = NormalizeAggFunc(CStr(AggFuncs), "cDataFrame.GroupBy")
        For a = 1 To numAggs
            normFuncs(a) = singleFunc
        Next a
    End If

    ' 4. Build output DataFrame schema
    Dim resultDf As New cDataFrame
    For g = 1 To numGroupCols
        If resultDf.HasColumn(groupColNames(g)) Then
            Err.Raise 5, "cDataFrame.GroupBy", "Duplicate group column specified: '" & groupColNames(g) & "'"
        End If
        resultDf.AddColumn groupColNames(g)
    Next g

    For a = 1 To numAggs
        Dim baseOutName As String, finalOutName As String, dupIdx As Long
        baseOutName = aggColNames(a) & "_" & normFuncs(a)
        finalOutName = baseOutName
        dupIdx = 1
        Do While resultDf.HasColumn(finalOutName)
            dupIdx = dupIdx + 1
            finalOutName = baseOutName & "_" & dupIdx
        Loop
        resultDf.AddColumn finalOutName
    Next a

    If mRowCount = 0 Then
        Set GroupBy = resultDf
        Exit Function
    End If

    ' 5. Pre-resolve column object references for tight loop access
    Dim groupColObjs() As cColumn
    ReDim groupColObjs(1 To numGroupCols)
    For g = 1 To numGroupCols
        Set groupColObjs(g) = mCols.Item(groupColNames(g))
    Next g

    Dim aggColObjs() As cColumn
    ReDim aggColObjs(1 To numAggs)
    For a = 1 To numAggs
        Set aggColObjs(a) = mCols.Item(aggColNames(a))
    Next a

    ' 6. Pass 1: Discover unique groups (1..numUniqueGroups) in O(N)
    Dim groupDict As Object
    Set groupDict = CreateObject("Scripting.Dictionary")
    groupDict.CompareMode = 1  ' Case-insensitive group keys

    Dim rowGroup() As Long
    Dim groupFirstRow() As Long
    ReDim rowGroup(1 To mRowCount)
    ReDim groupFirstRow(1 To mRowCount)

    Dim numUniqueGroups As Long
    numUniqueGroups = 0

    Dim r As Long, compKey As String, gIdx As Long
    For r = 1 To mRowCount
        If numGroupCols = 1 Then
            compKey = SafeKeyString(groupColObjs(1).Item(r))
        Else
            compKey = ""
            For g = 1 To numGroupCols
                compKey = compKey & FormatCompositeKeyPart(groupColObjs(g).Item(r))
            Next g
        End If

        If groupDict.Exists(compKey) Then
            gIdx = groupDict.Item(compKey)
        Else
            numUniqueGroups = numUniqueGroups + 1
            gIdx = numUniqueGroups
            groupDict.Add compKey, gIdx
            groupFirstRow(gIdx) = r
        End If
        rowGroup(r) = gIdx
    Next r

    ' 7. Pass 2: Allocate exact (1 To numUniqueGroups, 1 To numAggs) accumulators and aggregate
    Dim sums() As Double
    Dim counts() As Long
    Dim numCounts() As Long
    Dim mins() As Variant
    Dim maxs() As Variant
    Dim hasMinMax() As Boolean

    ReDim sums(1 To numUniqueGroups, 1 To numAggs)
    ReDim counts(1 To numUniqueGroups, 1 To numAggs)
    ReDim numCounts(1 To numUniqueGroups, 1 To numAggs)
    ReDim mins(1 To numUniqueGroups, 1 To numAggs)
    ReDim maxs(1 To numUniqueGroups, 1 To numAggs)
    ReDim hasMinMax(1 To numUniqueGroups, 1 To numAggs)

    Dim aggCol As cColumn, fType As String, v As Variant
    For a = 1 To numAggs
        Set aggCol = aggColObjs(a)
        fType = normFuncs(a)

        For r = 1 To mRowCount
            gIdx = rowGroup(r)
            v = aggCol.Item(r)

            If Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
                counts(gIdx, a) = counts(gIdx, a) + 1

                If fType = "SUM" Or fType = "AVG" Then
                    If IsNumeric(v) Then
                        sums(gIdx, a) = sums(gIdx, a) + CDbl(v)
                        numCounts(gIdx, a) = numCounts(gIdx, a) + 1
                    End If
                ElseIf fType = "MIN" Then
                    If Not hasMinMax(gIdx, a) Then
                        mins(gIdx, a) = v
                        hasMinMax(gIdx, a) = True
                    ElseIf CompareSortValues(v, mins(gIdx, a)) < 0 Then
                        mins(gIdx, a) = v
                    End If
                ElseIf fType = "MAX" Then
                    If Not hasMinMax(gIdx, a) Then
                        maxs(gIdx, a) = v
                        hasMinMax(gIdx, a) = True
                    ElseIf CompareSortValues(v, maxs(gIdx, a)) > 0 Then
                        maxs(gIdx, a) = v
                    End If
                End If
            End If
        Next r
    Next a

    ' 8. Emit grouped rows
    Dim totalOutCols As Long
    totalOutCols = numGroupCols + numAggs
    Dim rowVals() As Variant
    ReDim rowVals(0 To totalOutCols - 1)

    Dim firstRow As Long
    For gIdx = 1 To numUniqueGroups
        firstRow = groupFirstRow(gIdx)
        For g = 1 To numGroupCols
            rowVals(g - 1) = groupColObjs(g).Item(firstRow)
        Next g

        For a = 1 To numAggs
            fType = normFuncs(a)
            Select Case fType
                Case "SUM"
                    rowVals(numGroupCols + a - 1) = sums(gIdx, a)
                Case "COUNT"
                    rowVals(numGroupCols + a - 1) = counts(gIdx, a)
                Case "AVG"
                    If numCounts(gIdx, a) > 0 Then
                        rowVals(numGroupCols + a - 1) = sums(gIdx, a) / numCounts(gIdx, a)
                    Else
                        rowVals(numGroupCols + a - 1) = 0#
                    End If
                Case "MIN"
                    If hasMinMax(gIdx, a) Then
                        rowVals(numGroupCols + a - 1) = mins(gIdx, a)
                    Else
                        rowVals(numGroupCols + a - 1) = Empty
                    End If
                Case "MAX"
                    If hasMinMax(gIdx, a) Then
                        rowVals(numGroupCols + a - 1) = maxs(gIdx, a)
                    Else
                        rowVals(numGroupCols + a - 1) = Empty
                    End If
            End Select
        Next a

        resultDf.AddRow rowVals
    Next gIdx

    Set GroupBy = resultDf
End Function

' Splits the DataFrame into a Dictionary of child cDataFrame instances partitioned by unique values of ColumnName.
' Dictionary Keys: Unique group names (String)
' Dictionary Items: cDataFrame containing rows belonging to that group
Public Function SplitBy(ByVal ColumnName As String) As Object
    Dim resolvedCol As String
    resolvedCol = ResolveColName(ColumnName)

    Dim groupDict As Object
    Set groupDict = CreateObject("Scripting.Dictionary")
    groupDict.CompareMode = IIf(mCaseSensitive, 0, 1)

    If mRowCount = 0 Then
        Set SplitBy = groupDict
        Exit Function
    End If

    Dim splitCol As cColumn
    Set splitCol = mCols.Item(resolvedCol)

    Dim numCols As Long
    numCols = mCols.Count
    Dim colKeys As Variant
    colKeys = mCols.Keys
    Dim colObjs() As cColumn
    ReDim colObjs(0 To numCols - 1)
    Dim c As Long
    For c = 0 To numCols - 1
        Set colObjs(c) = mCols.Item(colKeys(c))
    Next c

    Dim r As Long, gKey As String, childDf As cDataFrame
    Dim rowVals() As Variant
    ReDim rowVals(0 To numCols - 1)

    For r = 1 To mRowCount
        gKey = SafeKeyString(splitCol.Item(r))
        If Not groupDict.Exists(gKey) Then
            Set childDf = Me.Clone(True)
            groupDict.Add gKey, childDf
        Else
            Set childDf = groupDict.Item(gKey)
        End If

        For c = 0 To numCols - 1
            rowVals(c) = colObjs(c).Item(r)
        Next c
        childDf.AddRow rowVals
    Next r

    Set SplitBy = groupDict
End Function

' Fast O(N) Cross-Tabulation / Pivot Table generator.
' Unique values of IndexColumn become rows (with active O(1) SetIndex);
' unique values of PivotColumn become dynamic columns;
' ValueColumn is aggregated using AggFunc ("SUM", "COUNT", "AVG"/"AVERAGE", "MIN", "MAX").
Public Function Pivot(ByVal IndexColumn As Variant, _
                      ByVal PivotColumn As Variant, _
                      ByVal ValueColumn As Variant, _
                      Optional ByVal AggFunc As String = "SUM", _
                      Optional ByVal FillEmpty As Variant = 0) As cDataFrame
    Dim idxColName As String, pivColName As String, valColName As String
    idxColName = ResolveColName(IndexColumn)
    pivColName = ResolveColName(PivotColumn)
    valColName = ResolveColName(ValueColumn)

    Dim normFunc As String
    normFunc = NormalizeAggFunc(AggFunc, "cDataFrame.Pivot")

    Dim resultDf As New cDataFrame
    resultDf.AddColumn idxColName

    If mRowCount = 0 Then
        Set Pivot = resultDf
        Exit Function
    End If

    Dim idxCol As cColumn, pivCol As cColumn, valCol As cColumn
    Set idxCol = mCols.Item(idxColName)
    Set pivCol = mCols.Item(pivColName)
    Set valCol = mCols.Item(valColName)

    ' 1. Pass 1: Discover unique Index keys (1..numIdx) and Pivot keys (1..numPiv)
    Dim idxDict As Object, pivDict As Object
    Set idxDict = CreateObject("Scripting.Dictionary")
    idxDict.CompareMode = 1
    Set pivDict = CreateObject("Scripting.Dictionary")
    pivDict.CompareMode = 1

    Dim rowIdxMap() As Long, rowPivMap() As Long
    ReDim rowIdxMap(1 To mRowCount)
    ReDim rowPivMap(1 To mRowCount)

    Dim idxFirstVal() As Variant
    Dim pivHeaders() As String
    ReDim idxFirstVal(1 To mRowCount)
    ReDim pivHeaders(1 To mRowCount)

    Dim numIdx As Long, numPiv As Long
    numIdx = 0
    numPiv = 0

    Dim r As Long, kIdx As String, kPiv As String
    Dim iIdx As Long, iPiv As Long

    For r = 1 To mRowCount
        kIdx = Trim$(SafeKeyString(idxCol.Item(r)))
        If Len(kIdx) = 0 Then kIdx = "(blank)"

        If idxDict.Exists(kIdx) Then
            iIdx = idxDict.Item(kIdx)
        Else
            numIdx = numIdx + 1
            iIdx = numIdx
            idxDict.Add kIdx, iIdx
            If IsEmpty(idxCol.Item(r)) Or IsNull(idxCol.Item(r)) Then
                idxFirstVal(iIdx) = "(blank)"
            Else
                idxFirstVal(iIdx) = idxCol.Item(r)
            End If
        End If
        rowIdxMap(r) = iIdx

        kPiv = Trim$(SafeKeyString(pivCol.Item(r)))
        If Len(kPiv) = 0 Then kPiv = "(blank)"

        If pivDict.Exists(kPiv) Then
            iPiv = pivDict.Item(kPiv)
        Else
            numPiv = numPiv + 1
            iPiv = numPiv
            pivDict.Add kPiv, iPiv
            pivHeaders(iPiv) = kPiv
        End If
        rowPivMap(r) = iPiv
    Next r

    ' 2. Add dynamic pivot columns to resultDf
    Dim p As Long, pColName As String, finalPColName As String, dupIdx As Long
    For p = 1 To numPiv
        pColName = pivHeaders(p)
        finalPColName = pColName
        dupIdx = 1
        Do While resultDf.HasColumn(finalPColName)
            dupIdx = dupIdx + 1
            finalPColName = pColName & "_" & dupIdx
        Loop
        resultDf.AddColumn finalPColName
    Next p

    ' 3. Pass 2: Allocate exact 2D accumulator matrices (1 To numIdx, 1 To numPiv)
    Dim hasCell() As Boolean
    Dim sums() As Double
    Dim counts() As Long
    Dim numCounts() As Long
    Dim mins() As Variant
    Dim maxs() As Variant
    Dim hasMinMax() As Boolean

    ReDim hasCell(1 To numIdx, 1 To numPiv)
    ReDim sums(1 To numIdx, 1 To numPiv)
    ReDim counts(1 To numIdx, 1 To numPiv)
    ReDim numCounts(1 To numIdx, 1 To numPiv)
    ReDim mins(1 To numIdx, 1 To numPiv)
    ReDim maxs(1 To numIdx, 1 To numPiv)
    ReDim hasMinMax(1 To numIdx, 1 To numPiv)

    Dim v As Variant
    For r = 1 To mRowCount
        iIdx = rowIdxMap(r)
        iPiv = rowPivMap(r)
        hasCell(iIdx, iPiv) = True
        v = valCol.Item(r)

        If Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
            counts(iIdx, iPiv) = counts(iIdx, iPiv) + 1

            If normFunc = "SUM" Or normFunc = "AVG" Then
                If IsNumeric(v) Then
                    sums(iIdx, iPiv) = sums(iIdx, iPiv) + CDbl(v)
                    numCounts(iIdx, iPiv) = numCounts(iIdx, iPiv) + 1
                End If
            ElseIf normFunc = "MIN" Then
                If Not hasMinMax(iIdx, iPiv) Then
                    mins(iIdx, iPiv) = v
                    hasMinMax(iIdx, iPiv) = True
                ElseIf CompareSortValues(v, mins(iIdx, iPiv)) < 0 Then
                    mins(iIdx, iPiv) = v
                End If
            ElseIf normFunc = "MAX" Then
                If Not hasMinMax(iIdx, iPiv) Then
                    maxs(iIdx, iPiv) = v
                    hasMinMax(iIdx, iPiv) = True
                ElseIf CompareSortValues(v, maxs(iIdx, iPiv)) > 0 Then
                    maxs(iIdx, iPiv) = v
                End If
            End If
        End If
    Next r

    ' 4. Populate resultDf rows
    Dim rowVals() As Variant
    ReDim rowVals(0 To numPiv)

    For iIdx = 1 To numIdx
        rowVals(0) = idxFirstVal(iIdx)
        For iPiv = 1 To numPiv
            If Not hasCell(iIdx, iPiv) Then
                rowVals(iPiv) = FillEmpty
            Else
                Select Case normFunc
                    Case "SUM"
                        rowVals(iPiv) = sums(iIdx, iPiv)
                    Case "COUNT"
                        rowVals(iPiv) = counts(iIdx, iPiv)
                    Case "AVG"
                        If numCounts(iIdx, iPiv) > 0 Then
                            rowVals(iPiv) = sums(iIdx, iPiv) / numCounts(iIdx, iPiv)
                        Else
                            rowVals(iPiv) = 0#
                        End If
                    Case "MIN"
                        If hasMinMax(iIdx, iPiv) Then
                            rowVals(iPiv) = mins(iIdx, iPiv)
                        Else
                            rowVals(iPiv) = FillEmpty
                        End If
                    Case "MAX"
                        If hasMinMax(iIdx, iPiv) Then
                            rowVals(iPiv) = maxs(iIdx, iPiv)
                        Else
                            rowVals(iPiv) = FillEmpty
                        End If
                End Select
            End If
        Next iPiv
        resultDf.AddRow rowVals
    Next iIdx

    ' 5. Automatically index resultDf on IndexColumn for immediate O(1) Loc(RowKey, PivotColKey)
    If resultDf.RowCount > 0 Then
        resultDf.SetIndex idxColName
    End If

    Set Pivot = resultDf
End Function

Private Function NormalizeAggFunc(ByVal FuncName As String, ByVal Caller As String) As String
    Dim u As String
    u = UCase$(Trim$(FuncName))
    Select Case u
        Case "SUM"
            NormalizeAggFunc = "SUM"
        Case "COUNT"
            NormalizeAggFunc = "COUNT"
        Case "AVG", "AVERAGE"
            NormalizeAggFunc = "AVG"
        Case "MIN"
            NormalizeAggFunc = "MIN"
        Case "MAX"
            NormalizeAggFunc = "MAX"
        Case Else
            Err.Raise 5, Caller, _
                "Unsupported aggregation function: '" & FuncName & "'. Supported: SUM, COUNT, AVG/AVERAGE, MIN, MAX"
    End Select
End Function

Private Function SafeKeyString(ByVal v As Variant) As String
    If IsError(v) Then
        SafeKeyString = "#ERR:" & CStr(CInt(v))
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeKeyString = ""
    Else
        SafeKeyString = CStr(v)
    End If
End Function

Private Function SafeDedupToken(ByVal v As Variant) As String
    If IsError(v) Then
        SafeDedupToken = "#E:" & CStr(CInt(v))
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeDedupToken = "#N"
    Else
        SafeDedupToken = "#V:" & CStr(v)
    End If
End Function

' ------------------------------------------------------------------------------
' Excel Range I/O
' ------------------------------------------------------------------------------

' Loads entire range in a single COM read. Handles 1-cell scalar range cleanly.
Public Function LoadFromRange(ByVal SourceRange As Object, _
                             Optional ByVal HasHeaders As Boolean = True) As cDataFrame
    If SourceRange Is Nothing Then
        Err.Raise 91, "cDataFrame.LoadFromRange", "SourceRange cannot be Nothing."
    End If

    Dim vData As Variant
    On Error Resume Next
    vData = SourceRange.Value2
    If Err.Number <> 0 Then
        Err.Clear
        vData = SourceRange.Value
    End If
    On Error GoTo 0

    ' Guard against single-cell scalar range
    If Not IsArray(vData) Then
        Dim singleCell(1 To 1, 1 To 1) As Variant
        singleCell(1, 1) = vData
        vData = singleCell
    End If

    Dim rLb As Long, rUb As Long, cLb As Long, cUb As Long
    rLb = LBound(vData, 1)
    rUb = UBound(vData, 1)
    cLb = LBound(vData, 2)
    cUb = UBound(vData, 2)

    Dim totalRows As Long, totalCols As Long
    totalRows = rUb - rLb + 1
    totalCols = cUb - cLb + 1

    ' Clear existing data
    mCols.RemoveAll
    mRowCount = 0
    Set mRowIndexMap = Nothing
    mIndexColumnName = ""
    mIndexDirty = False

    Dim dataStartRow As Long, dataRowCount As Long
    If HasHeaders Then
        dataStartRow = rLb + 1
        dataRowCount = totalRows - 1
        If dataRowCount < 0 Then dataRowCount = 0
    Else
        dataStartRow = rLb
        dataRowCount = totalRows
    End If

    mRowCount = dataRowCount

    Dim c As Long, r As Long, headerName As String, col As cColumn
    For c = 1 To totalCols
        If HasHeaders Then
            Dim rawHdr As Variant
            rawHdr = vData(rLb, cLb + c - 1)
            If IsError(rawHdr) Then
                headerName = "Column_Err_" & c
            ElseIf IsNull(rawHdr) Or IsEmpty(rawHdr) Then
                headerName = "Column_" & c
            Else
                headerName = Trim$(CStr(rawHdr))
                If Len(headerName) = 0 Then headerName = "Column_" & c
            End If
        Else
            headerName = "Column_" & c
        End If

        ' Ensure unique column header
        Dim uniqueHeader As String, dupCount As Long
        uniqueHeader = headerName
        dupCount = 1
        Do While mCols.Exists(uniqueHeader)
            dupCount = dupCount + 1
            uniqueHeader = headerName & "_" & dupCount
        Loop

        Set col = New cColumn
        col.Name = uniqueHeader
        col.EnsureCapacity dataRowCount
        col.SetCount dataRowCount

        If dataRowCount > 0 Then
            For r = 1 To dataRowCount
                col.Item(r) = vData(dataStartRow + r - 1, cLb + c - 1)
            Next r
        End If

        mCols.Add uniqueHeader, col
    Next c

    Set LoadFromRange = Me
End Function

' Dumps entire DataFrame to Excel Range in a single COM write.
Public Sub ToRange(ByVal TargetCell As Object, _
                   Optional ByVal IncludeHeaders As Boolean = True)
    If TargetCell Is Nothing Then
        Err.Raise 91, "cDataFrame.ToRange", "TargetCell cannot be Nothing."
    End If

    If mCols.Count = 0 Then Exit Sub

    Dim ws As Object
    Set ws = TargetCell.Worksheet
    If ws.ProtectContents Then
        Err.Raise 1004, "cDataFrame.ToRange", _
            "Cannot write to protected worksheet '" & ws.Name & "'."
    End If

    Dim outputRows As Long
    outputRows = mRowCount + IIf(IncludeHeaders, 1, 0)
    If outputRows = 0 Then Exit Sub

    Dim destRng As Object
    Set destRng = TargetCell.Resize(outputRows, mCols.Count)
    If destRng.MergeCells Then
        destRng.UnMerge
    End If

    Dim outArr() As Variant
    ReDim outArr(1 To outputRows, 1 To mCols.Count)

    Dim keys As Variant
    keys = mCols.Keys

    Dim c As Long, r As Long, col As cColumn, startRow As Long
    startRow = 1

    If IncludeHeaders Then
        For c = 1 To mCols.Count
            outArr(1, c) = keys(c - 1)
        Next c
        startRow = 2
    End If

    If mRowCount > 0 Then
        For c = 1 To mCols.Count
            Set col = mCols.Item(keys(c - 1))
            For r = 1 To mRowCount
                outArr(startRow + r - 1, c) = col.Item(r)
            Next r
        Next c
    End If

    destRng.Value2 = outArr
End Sub

' Loads an Excel Table (ListObject) in at most 1-2 COM reads, safely ignoring TotalsRowRange.
Public Function LoadFromListObject(ByVal TableObj As Object) As cDataFrame
    If TableObj Is Nothing Then
        Err.Raise 91, "cDataFrame.LoadFromListObject", "TableObj cannot be Nothing."
    End If

    Dim totalCols As Long
    totalCols = TableObj.ListColumns.Count

    mCols.RemoveAll
    mRowCount = 0
    Set mRowIndexMap = Nothing
    mIndexColumnName = ""
    mIndexDirty = False

    If totalCols = 0 Then
        Set LoadFromListObject = Me
        Exit Function
    End If

    Dim hasBody As Boolean
    Dim vBody As Variant
    Dim dataRowCount As Long
    Dim rLb As Long, cLb As Long

    If TableObj.DataBodyRange Is Nothing Then
        hasBody = False
        dataRowCount = 0
    Else
        hasBody = True
        On Error Resume Next
        vBody = TableObj.DataBodyRange.Value2
        If Err.Number <> 0 Then
            Err.Clear
            vBody = TableObj.DataBodyRange.Value
        End If
        On Error GoTo 0

        If Not IsArray(vBody) Then
            Dim singleCell(1 To 1, 1 To 1) As Variant
            singleCell(1, 1) = vBody
            vBody = singleCell
        End If

        rLb = LBound(vBody, 1)
        cLb = LBound(vBody, 2)
        dataRowCount = UBound(vBody, 1) - rLb + 1
    End If

    mRowCount = dataRowCount

    Dim c As Long, r As Long, headerName As String, uniqueHeader As String, dupCount As Long
    Dim col As cColumn

    For c = 1 To totalCols
        headerName = Trim$(CStr(TableObj.ListColumns(c).Name))
        If Len(headerName) = 0 Then headerName = "Column_" & c

        uniqueHeader = headerName
        dupCount = 1
        Do While mCols.Exists(uniqueHeader)
            dupCount = dupCount + 1
            uniqueHeader = headerName & "_" & dupCount
        Loop

        Set col = New cColumn
        col.Name = uniqueHeader
        col.EnsureCapacity dataRowCount
        col.SetCount dataRowCount

        If hasBody And dataRowCount > 0 Then
            For r = 1 To dataRowCount
                col.Item(r) = vBody(rLb + r - 1, cLb + c - 1)
            Next r
        End If

        mCols.Add uniqueHeader, col
    Next c

    Set LoadFromListObject = Me
End Function

' ------------------------------------------------------------------------------
' External Workbook I/O (Multi-File Ingestion Engine)
' ------------------------------------------------------------------------------

' Loads data directly from an external Excel workbook (.xlsx, .xlsm, .xlsb, .xls).
' Automatically checks if the workbook is already open in the Excel instance:
' - If already open: reads directly in-memory without disk I/O and leaves it open.
' - If closed: opens headless ReadOnly (ScreenUpdating & EnableEvents suppressed),
'   reads via bulk .Value2, and guarantees workbook is closed in CleanUp.
' Parameters:
'   FilePath: Full or relative path to the Excel file.
'   SheetNameOrIndex: Worksheet name (String) or 1-based index (Long/Integer). Defaults to 1.
'   RangeOrTableName: Range address (e.g. "A1:E100"), ListObject table name, or "" for UsedRange.
'   HasHeaders: True if first row contains column headers (default True).
Public Function LoadFromWorkbook(ByVal FilePath As String, _
                                Optional ByVal SheetNameOrIndex As Variant = 1, _
                                Optional ByVal RangeOrTableName As String = "", _
                                Optional ByVal HasHeaders As Boolean = True) As cDataFrame
    Dim cleanPath As String
    cleanPath = Trim$(FilePath)

    If Len(cleanPath) = 0 Then
        Err.Raise 5, "cDataFrame.LoadFromWorkbook", "FilePath cannot be blank."
    End If

    ' Resolve full path if relative
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FileExists(cleanPath) Then
            cleanPath = fso.GetAbsolutePathName(cleanPath)
        End If
    End If
    On Error GoTo 0

    If Len(Dir(cleanPath)) = 0 Then
        Err.Raise 53, "cDataFrame.LoadFromWorkbook", "File not found: " & cleanPath
    End If

    Dim xlApp As Object
    On Error Resume Next
    Set xlApp = Application
    On Error GoTo 0

    If xlApp Is Nothing Then
        Err.Raise 91, "cDataFrame.LoadFromWorkbook", "Excel Application object not accessible."
    End If

    Dim origScreenUpdating As Boolean
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean
    Dim wb As Object
    Dim ws As Object
    Dim targetRng As Object
    Dim openedByUs As Boolean
    Dim errNum As Long, errDesc As String

    ' Check if workbook is already open
    Dim checkWb As Object
    Dim fileNameOnly As String
    If Not fso Is Nothing Then
        fileNameOnly = fso.GetFileName(cleanPath)
    Else
        fileNameOnly = cleanPath
    End If

    For Each checkWb In xlApp.Workbooks
        If StrComp(checkWb.FullName, cleanPath, vbTextCompare) = 0 Or _
           StrComp(checkWb.Name, fileNameOnly, vbTextCompare) = 0 Then
            Set wb = checkWb
            openedByUs = False
            Exit For
        End If
    Next checkWb

    On Error GoTo ErrorHandler

    ' Save environment states
    origScreenUpdating = xlApp.ScreenUpdating
    origEnableEvents = xlApp.EnableEvents
    origDisplayAlerts = xlApp.DisplayAlerts

    If wb Is Nothing Then
        xlApp.ScreenUpdating = False
        xlApp.EnableEvents = False
        xlApp.DisplayAlerts = False

        Set wb = xlApp.Workbooks.Open(FileName:=cleanPath, ReadOnly:=True, UpdateLinks:=False)
        openedByUs = True
    End If

    ' Resolve worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SheetNameOrIndex)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Err.Raise 9, "cDataFrame.LoadFromWorkbook", "Worksheet '" & CStr(SheetNameOrIndex) & "' not found in " & cleanPath
    End If

    ' Resolve Target Range
    Dim targetName As String
    targetName = Trim$(RangeOrTableName)

    If Len(targetName) = 0 Then
        Set targetRng = ws.UsedRange
    Else
        ' Check if targetName matches a ListObject table name
        Dim lo As Object
        On Error Resume Next
        Set lo = ws.ListObjects(targetName)
        On Error GoTo ErrorHandler

        If Not lo Is Nothing Then
            Set targetRng = lo.Range
        Else
            ' Attempt range address lookup
            On Error Resume Next
            Set targetRng = ws.Range(targetName)
            On Error GoTo ErrorHandler

            If targetRng Is Nothing Then
                Err.Raise 1004, "cDataFrame.LoadFromWorkbook", "Range or Table '" & targetName & "' not found on sheet " & ws.Name
            End If
        End If
    End If

    ' Load into this cDataFrame
    Me.LoadFromRange targetRng, HasHeaders

CleanUp:
    On Error Resume Next
    If openedByUs And Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    If Not xlApp Is Nothing Then
        xlApp.ScreenUpdating = origScreenUpdating
        xlApp.EnableEvents = origEnableEvents
        xlApp.DisplayAlerts = origDisplayAlerts
    End If
    On Error GoTo 0

    If errNum <> 0 Then
        Err.Raise errNum, "cDataFrame.LoadFromWorkbook", errDesc
    End If

    Set LoadFromWorkbook = Me
    Exit Function

ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    Resume CleanUp
End Function

' Dumps DataFrame to worksheet and wraps it in a structured Excel Table (ListObject).
Public Function ToListObject(ByVal TargetCellOrSheet As Object, _
                             Optional ByVal TableName As String = "") As Object
    If TargetCellOrSheet Is Nothing Then
        Err.Raise 91, "cDataFrame.ToListObject", "TargetCellOrSheet cannot be Nothing."
    End If

    If mCols.Count = 0 Then
        Err.Raise 5, "cDataFrame.ToListObject", "Cannot create ListObject from a DataFrame with 0 columns."
    End If

    Dim ws As Object, topCell As Object
    If TypeName(TargetCellOrSheet) = "Worksheet" Then
        Set ws = TargetCellOrSheet
        Set topCell = ws.Range("A1")
    Else
        Set topCell = TargetCellOrSheet.Cells(1, 1)
        Set ws = topCell.Worksheet
    End If

    Me.ToRange topCell, True

    Dim outputRows As Long
    outputRows = mRowCount + 1

    Dim lo As Object
    ' xlSrcRange = 1, xlYes = 1
    Set lo = ws.ListObjects.Add(1, topCell.Resize(outputRows, mCols.Count), , 1)

    If Len(Trim$(TableName)) > 0 Then
        lo.Name = Trim$(TableName)
    End If

    Set ToListObject = lo
End Function

' ------------------------------------------------------------------------------
' Fast CSV File I/O (RFC 4180 Compliant, Worksheet-Free)
' ------------------------------------------------------------------------------

' Loads a CSV file directly into memory without touching any worksheet.
' Supports RFC 4180 quoted fields, escaped quotes (""), UTF-8 BOM stripping,
' and automatic numeric inference on unquoted tokens (preserving leading zeros on quoted strings).
Public Function LoadFromCSV(ByVal FilePath As String, _
                            Optional ByVal Delimiter As String = ",", _
                            Optional ByVal HasHeaders As Boolean = True, _
                            Optional ByVal AutoInferNumeric As Boolean = True) As cDataFrame
    Dim cleanPath As String
    cleanPath = Trim$(FilePath)
    If Len(cleanPath) = 0 Then
        Err.Raise 5, "cDataFrame.LoadFromCSV", "FilePath cannot be blank."
    End If
    If Len(Delimiter) = 0 Then
        Err.Raise 5, "cDataFrame.LoadFromCSV", "Delimiter cannot be empty."
    End If
    If Len(Dir$(cleanPath)) = 0 Then
        Err.Raise 53, "cDataFrame.LoadFromCSV", "CSV file not found: '" & cleanPath & "'"
    End If

    mCols.RemoveAll
    mRowCount = 0
    Set mRowIndexMap = Nothing
    mIndexColumnName = ""
    mIndexDirty = False

    Dim fNum As Integer
    fNum = FreeFile
    Open cleanPath For Input As #fNum

    On Error GoTo CleanFail

    Dim rawLine As String
    Dim isFirstLine As Boolean
    isFirstLine = True

    Dim tokens() As String
    Dim wasQuoted() As Boolean
    Dim fieldCount As Long
    Dim numCols As Long
    Dim c As Long
    Dim rowVals() As Variant

    Do While Not EOF(fNum)
        Line Input #fNum, rawLine

        If isFirstLine Then
            ' Strip UTF-8 BOM (EF BB BF or FEFF) if present
            If Left$(rawLine, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then
                rawLine = Mid$(rawLine, 4)
            ElseIf Left$(rawLine, 1) = ChrW$(&HFEFF) Then
                rawLine = Mid$(rawLine, 2)
            End If
        End If

        If Len(Trim$(rawLine)) > 0 Then
            fieldCount = ParseCSVLine(rawLine, Delimiter, tokens, wasQuoted)

            If isFirstLine Then
                isFirstLine = False
                numCols = fieldCount
                Dim headerName As String, uniqueHeader As String, dupCount As Long

                For c = 1 To numCols
                    If HasHeaders Then
                        headerName = Trim$(tokens(c))
                        If Len(headerName) = 0 Then headerName = "Column_" & c
                    Else
                        headerName = "Column_" & c
                    End If

                    uniqueHeader = headerName
                    dupCount = 1
                    Do While mCols.Exists(uniqueHeader)
                        dupCount = dupCount + 1
                        uniqueHeader = headerName & "_" & dupCount
                    Loop
                    Me.AddColumn uniqueHeader
                Next c

                ReDim rowVals(0 To numCols - 1)

                If Not HasHeaders Then
                    For c = 1 To numCols
                        rowVals(c - 1) = ConvertCSVToken(tokens(c), wasQuoted(c), AutoInferNumeric)
                    Next c
                    Me.AddRow rowVals
                End If
            Else
                For c = 1 To numCols
                    If c <= fieldCount Then
                        rowVals(c - 1) = ConvertCSVToken(tokens(c), wasQuoted(c), AutoInferNumeric)
                    Else
                        rowVals(c - 1) = Empty
                    End If
                Next c
                Me.AddRow rowVals
            End If
        End If
    Loop

    Close #fNum
    Set LoadFromCSV = Me
    Exit Function

CleanFail:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    Err.Raise errNum, "cDataFrame.LoadFromCSV", errDesc
End Function

' Writes the DataFrame directly to a CSV file with RFC 4180 quoting, escaping,
' and optional CSV formula injection protection (SanitizeFormulas:=True).
Public Sub ToCSV(ByVal FilePath As String, _
                 Optional ByVal Delimiter As String = ",", _
                 Optional ByVal IncludeHeaders As Boolean = True, _
                 Optional ByVal SanitizeFormulas As Boolean = False)
    Dim cleanPath As String
    cleanPath = Trim$(FilePath)
    If Len(cleanPath) = 0 Then
        Err.Raise 5, "cDataFrame.ToCSV", "FilePath cannot be blank."
    End If
    If Len(Delimiter) = 0 Then
        Err.Raise 5, "cDataFrame.ToCSV", "Delimiter cannot be empty."
    End If

    Dim fNum As Integer
    fNum = FreeFile
    Open cleanPath For Output As #fNum

    On Error GoTo CleanWriteFail

    If mCols.Count = 0 Then
        Close #fNum
        Exit Sub
    End If

    Dim colCount As Long
    colCount = mCols.Count
    Dim keys As Variant
    keys = mCols.Keys

    Dim colObjs() As cColumn
    ReDim colObjs(0 To colCount - 1)
    Dim c As Long
    For c = 0 To colCount - 1
        Set colObjs(c) = mCols.Item(keys(c))
    Next c

    Dim lineParts() As String
    ReDim lineParts(0 To colCount - 1)

    If IncludeHeaders Then
        For c = 0 To colCount - 1
            lineParts(c) = FormatCSVField(CStr(keys(c)), Delimiter, SanitizeFormulas)
        Next c
        Print #fNum, Join(lineParts, Delimiter)
    End If

    Dim r As Long
    For r = 1 To mRowCount
        For c = 0 To colCount - 1
            lineParts(c) = FormatCSVField(colObjs(c).Item(r), Delimiter, SanitizeFormulas)
        Next c
        Print #fNum, Join(lineParts, Delimiter)
    Next r

    Close #fNum
    Exit Sub

CleanWriteFail:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    Err.Raise errNum, "cDataFrame.ToCSV", errDesc
End Sub

Private Function ParseCSVLine(ByVal LineText As String, _
                              ByVal Delimiter As String, _
                              ByRef OutTokens() As String, _
                              ByRef OutWasQuoted() As Boolean) As Long
    Dim delimChar As String
    delimChar = Left$(Delimiter, 1)

    Dim cap As Long
    cap = 16
    ReDim OutTokens(1 To cap)
    ReDim OutWasQuoted(1 To cap)

    Dim count As Long
    count = 0

    Dim pos As Long, lineLen As Long
    lineLen = Len(LineText)
    pos = 1

    Dim inQuotes As Boolean, fieldQuoted As Boolean
    Dim curToken As String, ch As String

    inQuotes = False
    fieldQuoted = False
    curToken = ""

    Do While pos <= lineLen
        ch = Mid$(LineText, pos, 1)

        If inQuotes Then
            If ch = """" Then
                If pos < lineLen And Mid$(LineText, pos + 1, 1) = """" Then
                    curToken = curToken & """"
                    pos = pos + 1
                Else
                    inQuotes = False
                End If
            Else
                curToken = curToken & ch
            End If
        Else
            If ch = """" And Len(curToken) = 0 Then
                inQuotes = True
                fieldQuoted = True
            ElseIf ch = delimChar Then
                count = count + 1
                If count > cap Then
                    cap = cap * 2
                    ReDim Preserve OutTokens(1 To cap)
                    ReDim Preserve OutWasQuoted(1 To cap)
                End If
                OutTokens(count) = curToken
                OutWasQuoted(count) = fieldQuoted
                curToken = ""
                fieldQuoted = False
            Else
                curToken = curToken & ch
            End If
        End If

        pos = pos + 1
    Loop

    count = count + 1
    If count > cap Then
        cap = count
        ReDim Preserve OutTokens(1 To cap)
        ReDim Preserve OutWasQuoted(1 To cap)
    End If
    OutTokens(count) = curToken
    OutWasQuoted(count) = fieldQuoted

    ParseCSVLine = count
End Function

Private Function ConvertCSVToken(ByVal Token As String, _
                                 ByVal WasQuoted As Boolean, _
                                 ByVal AutoInferNumeric As Boolean) As Variant
    If WasQuoted Then
        ConvertCSVToken = Token
    Else
        Dim trimmed As String
        trimmed = Trim$(Token)
        If Len(trimmed) = 0 Then
            ConvertCSVToken = Empty
        ElseIf AutoInferNumeric And IsNumeric(trimmed) Then
            ConvertCSVToken = CDbl(trimmed)
        Else
            ConvertCSVToken = Token
        End If
    End If
End Function

Private Function FormatCSVField(ByVal v As Variant, _
                                ByVal Delimiter As String, _
                                Optional ByVal SanitizeFormulas As Boolean = False) As String
    If IsError(v) Then
        FormatCSVField = "#N/A"
        Exit Function
    End If
    If IsEmpty(v) Or IsNull(v) Then
        FormatCSVField = ""
        Exit Function
    End If

    Dim s As String
    s = CStr(v)

    If SanitizeFormulas And VarType(v) = vbString Then
        Dim trimmedS As String
        trimmedS = LTrim$(s)
        If Len(trimmedS) > 0 Then
            Dim firstCh As String
            firstCh = Left$(trimmedS, 1)
            If firstCh = "=" Or firstCh = "+" Or firstCh = "-" Or _
               firstCh = "@" Or firstCh = vbTab Or firstCh = vbCr Then
                s = "'" & s
            End If
        End If
    End If

    If InStr(1, s, Delimiter, vbBinaryCompare) > 0 Or _
       InStr(1, s, """", vbBinaryCompare) > 0 Or _
       InStr(1, s, vbCr, vbBinaryCompare) > 0 Or _
       InStr(1, s, vbLf, vbBinaryCompare) > 0 Then
        FormatCSVField = """" & Replace(s, """", """""") & """"
    Else
        FormatCSVField = s
    End If
End Function

' ------------------------------------------------------------------------------
' Transformations & Queries
' ------------------------------------------------------------------------------

' Query rows matching condition and return new cDataFrame
Public Function Filter(ByVal ColumnName As String, _
                       ByVal Operator As String, _
                       ByVal CompareValue As Variant) As cDataFrame
    Dim targetCol As cColumn
    Set targetCol = mCols.Item(ResolveColName(ColumnName))

    Dim resultDf As New cDataFrame
    Dim keys As Variant
    keys = mCols.Keys

    Dim c As Long
    For c = 0 To mCols.Count - 1
        resultDf.AddColumn CStr(keys(c))
    Next c

    If mRowCount = 0 Then
        Set Filter = resultDf
        Exit Function
    End If

    Dim r As Long
    For r = 1 To mRowCount
        If EvaluateCondition(targetCol.Item(r), UCase$(Trim$(Operator)), CompareValue) Then
            Dim rowVals() As Variant
            ReDim rowVals(0 To mCols.Count - 1)
            For c = 0 To mCols.Count - 1
                Dim srcCol As cColumn
                Set srcCol = mCols.Item(keys(c))
                rowVals(c) = srcCol.Item(r)
            Next c

            resultDf.AddRow rowVals
        End If
    Next r

    Set Filter = resultDf
End Function

' Multi-condition compound filter with short-circuit AND / OR evaluation.
' Conditions is a 1D array of Array(ColumnKey, Operator, CompareValue).
Public Function FilterCompound(ByVal Conditions As Variant, _
                               Optional ByVal Logic As String = "AND") As cDataFrame
    Dim normLogic As String
    normLogic = UCase$(Trim$(Logic))
    If normLogic <> "AND" And normLogic <> "OR" Then
        Err.Raise 5, "cDataFrame.FilterCompound", _
            "Unsupported Logic mode: '" & Logic & "'. Supported: AND, OR"
    End If

    If Not IsArray(Conditions) Then
        Err.Raise 5, "cDataFrame.FilterCompound", _
            "Conditions must be a 1D array of Array(ColumnKey, Operator, CompareValue)."
    End If

    Dim lbC As Long, ubC As Long, numConds As Long
    lbC = LBound(Conditions)
    ubC = UBound(Conditions)
    numConds = ubC - lbC + 1
    If numConds <= 0 Then
        Err.Raise 5, "cDataFrame.FilterCompound", "Conditions array cannot be empty."
    End If

    Dim condCols() As cColumn
    Dim condOps() As String
    Dim condVals() As Variant
    ReDim condCols(1 To numConds)
    ReDim condOps(1 To numConds)
    ReDim condVals(1 To numConds)

    Dim k As Long, cItem As Variant, lbI As Long, ubI As Long
    For k = 1 To numConds
        cItem = Conditions(lbC + k - 1)
        If Not IsArray(cItem) Then
            Err.Raise 5, "cDataFrame.FilterCompound", _
                "Condition #" & k & " must be a 3-element Array(ColumnKey, Operator, CompareValue)."
        End If
        lbI = LBound(cItem)
        ubI = UBound(cItem)
        If (ubI - lbI + 1) < 3 Then
            Err.Raise 5, "cDataFrame.FilterCompound", _
                "Condition #" & k & " must have 3 elements: Array(ColumnKey, Operator, CompareValue)."
        End If

        Set condCols(k) = mCols.Item(ResolveColName(cItem(lbI)))
        condOps(k) = UCase$(Trim$(CStr(cItem(lbI + 1))))
        condVals(k) = cItem(lbI + 2)
    Next k

    Dim resultDf As cDataFrame
    Set resultDf = Me.Clone(True)

    If mRowCount = 0 Or mCols.Count = 0 Then
        Set FilterCompound = resultDf
        Exit Function
    End If

    Dim allColCount As Long
    allColCount = mCols.Count
    Dim keys As Variant
    keys = mCols.Keys
    Dim allCols() As cColumn
    ReDim allCols(0 To allColCount - 1)
    Dim c As Long
    For c = 0 To allColCount - 1
        Set allCols(c) = mCols.Item(keys(c))
    Next c

    Dim rowVals() As Variant
    ReDim rowVals(0 To allColCount - 1)

    Dim r As Long, rowMatch As Boolean, singleMatch As Boolean
    For r = 1 To mRowCount
        If normLogic = "AND" Then
            rowMatch = True
            For k = 1 To numConds
                singleMatch = EvaluateCondition(condCols(k).Item(r), condOps(k), condVals(k))
                If Not singleMatch Then
                    rowMatch = False
                    Exit For
                End If
            Next k
        Else
            rowMatch = False
            For k = 1 To numConds
                singleMatch = EvaluateCondition(condCols(k).Item(r), condOps(k), condVals(k))
                If singleMatch Then
                    rowMatch = True
                    Exit For
                End If
            Next k
        End If

        If rowMatch Then
            For c = 0 To allColCount - 1
                rowVals(c) = allCols(c).Item(r)
            Next c
            resultDf.AddRow rowVals
        End If
    Next r

    Set FilterCompound = resultDf
End Function

' Reshapes (unpivots) wide columns into tall (IdVars..., VarName, ValName) format in O(R * V)
Public Function Melt(ByVal IdVars As Variant, _
                     ByVal ValueVars As Variant, _
                     Optional ByVal VarName As String = "Variable", _
                     Optional ByVal ValName As String = "Value") As cDataFrame
    Dim cleanVarName As String, cleanValName As String
    cleanVarName = Trim$(VarName)
    cleanValName = Trim$(ValName)
    If Len(cleanVarName) = 0 Or Len(cleanValName) = 0 Then
        Err.Raise 5, "cDataFrame.Melt", "VarName and ValName cannot be blank."
    End If

    Dim idNames() As String
    Dim numIds As Long
    Dim i As Long

    If IsArray(IdVars) Then
        Dim lbI As Long, ubI As Long
        lbI = LBound(IdVars)
        ubI = UBound(IdVars)
        numIds = ubI - lbI + 1
        If numIds <= 0 Then
            Err.Raise 5, "cDataFrame.Melt", "IdVars array cannot be empty."
        End If
        ReDim idNames(1 To numIds)
        For i = 1 To numIds
            idNames(i) = ResolveColName(IdVars(lbI + i - 1))
        Next i
    Else
        numIds = 1
        ReDim idNames(1 To 1)
        idNames(1) = ResolveColName(IdVars)
    End If

    Dim valNames() As String
    Dim numVals As Long
    Dim v As Long

    If IsArray(ValueVars) Then
        Dim lbV As Long, ubV As Long
        lbV = LBound(ValueVars)
        ubV = UBound(ValueVars)
        numVals = ubV - lbV + 1
        If numVals <= 0 Then
            Err.Raise 5, "cDataFrame.Melt", "ValueVars array cannot be empty."
        End If
        ReDim valNames(1 To numVals)
        For v = 1 To numVals
            valNames(v) = ResolveColName(ValueVars(lbV + v - 1))
        Next v
    Else
        numVals = 1
        ReDim valNames(1 To 1)
        valNames(1) = ResolveColName(ValueVars)
    End If

    Dim resultDf As New cDataFrame
    Dim totalOutRows As Long
    totalOutRows = mRowCount * numVals

    If totalOutRows = 0 Then
        For i = 1 To numIds
            resultDf.AddColumn idNames(i)
        Next i
        resultDf.AddColumn cleanVarName
        resultDf.AddColumn cleanValName
        Set Melt = resultDf
        Exit Function
    End If

    Dim idCols() As cColumn
    ReDim idCols(1 To numIds)
    For i = 1 To numIds
        Set idCols(i) = mCols.Item(idNames(i))
    Next i

    Dim valCols() As cColumn
    ReDim valCols(1 To numVals)
    For v = 1 To numVals
        Set valCols(v) = mCols.Item(valNames(v))
    Next v

    Dim r As Long, outIdx As Long
    Dim colBuf() As Variant

    ' Populate IdVars columns
    For i = 1 To numIds
        ReDim colBuf(1 To totalOutRows)
        outIdx = 0
        For r = 1 To mRowCount
            For v = 1 To numVals
                outIdx = outIdx + 1
                colBuf(outIdx) = idCols(i).Item(r)
            Next v
        Next r
        resultDf.AddColumn idNames(i), , colBuf
    Next i

    ' Populate VarName column
    ReDim colBuf(1 To totalOutRows)
    outIdx = 0
    For r = 1 To mRowCount
        For v = 1 To numVals
            outIdx = outIdx + 1
            colBuf(outIdx) = valNames(v)
        Next v
    Next r
    resultDf.AddColumn cleanVarName, , colBuf

    ' Populate ValName column
    ReDim colBuf(1 To totalOutRows)
    outIdx = 0
    For r = 1 To mRowCount
        For v = 1 To numVals
            outIdx = outIdx + 1
            colBuf(outIdx) = valCols(v).Item(r)
        Next v
    Next r
    resultDf.AddColumn cleanValName, , colBuf

    Set Melt = resultDf
End Function

' ------------------------------------------------------------------------------
' Fast JSON Serialization & Deserialization (ToJSON & LoadFromJSON)
' ------------------------------------------------------------------------------

' Serializes DataFrame to JSON string.
' Orient: "RECORDS" -> [{"ColA":1,"ColB":"x"},...]
'         "SPLIT"   -> {"columns":["ColA","ColB"],"data":[[1,"x"],...]}
Public Function ToJSON(Optional ByVal Orient As String = "RECORDS") As String
    Dim normOrient As String
    normOrient = UCase$(Trim$(Orient))
    If normOrient <> "RECORDS" And normOrient <> "SPLIT" Then
        Err.Raise 5, "cDataFrame.ToJSON", _
            "Unsupported JSON Orient: '" & Orient & "'. Supported: RECORDS, SPLIT"
    End If

    Dim colCount As Long
    colCount = mCols.Count

    If colCount = 0 Then
        If normOrient = "RECORDS" Then
            ToJSON = "[]"
        Else
            ToJSON = "{""columns"":[],""data"":[]}"
        End If
        Exit Function
    End If

    Dim keys As Variant
    keys = mCols.Keys

    Dim colObjs() As cColumn
    Dim escapedColNames() As String
    ReDim colObjs(0 To colCount - 1)
    ReDim escapedColNames(0 To colCount - 1)

    Dim c As Long
    For c = 0 To colCount - 1
        Set colObjs(c) = mCols.Item(keys(c))
        escapedColNames(c) = """" & EscapeJSONString(CStr(keys(c))) & """"
    Next c

    Dim r As Long
    Dim rowParts() As String
    Dim cellParts() As String
    ReDim cellParts(0 To colCount - 1)

    If normOrient = "RECORDS" Then
        If mRowCount = 0 Then
            ToJSON = "[]"
            Exit Function
        End If

        ReDim rowParts(1 To mRowCount)
        For r = 1 To mRowCount
            For c = 0 To colCount - 1
                cellParts(c) = escapedColNames(c) & ":" & FormatJSONValue(colObjs(c).Item(r))
            Next c
            rowParts(r) = "{" & Join(cellParts, ",") & "}"
        Next r
        ToJSON = "[" & Join(rowParts, ",") & "]"
    Else
        Dim colHeaderJson As String
        colHeaderJson = "[" & Join(escapedColNames, ",") & "]"

        If mRowCount = 0 Then
            ToJSON = "{""columns"":" & colHeaderJson & ",""data"":[]}"
            Exit Function
        End If

        ReDim rowParts(1 To mRowCount)
        For r = 1 To mRowCount
            For c = 0 To colCount - 1
                cellParts(c) = FormatJSONValue(colObjs(c).Item(r))
            Next c
            rowParts(r) = "[" & Join(cellParts, ",") & "]"
        Next r
        ToJSON = "{""columns"":" & colHeaderJson & ",""data"":[" & Join(rowParts, ",") & "]}"
    End If
End Function

' Parses JSON string in either RECORDS ([{...},...]) or SPLIT ({"columns":[...],"data":[[...]]}) format.
Public Function LoadFromJSON(ByVal JsonString As String) As cDataFrame
    Dim s As String
    s = Trim$(JsonString)
    If Len(s) = 0 Then
        Err.Raise 5, "cDataFrame.LoadFromJSON", "JsonString cannot be blank."
    End If

    mCols.RemoveAll
    mRowCount = 0
    Set mRowIndexMap = Nothing
    mIndexColumnName = ""
    mIndexDirty = False

    Dim firstCh As String
    firstCh = Left$(s, 1)

    If firstCh = "[" Then
        ParseJSONRecordsIntoMe s
    ElseIf firstCh = "{" Then
        ParseJSONSplitIntoMe s
    Else
        Err.Raise 5, "cDataFrame.LoadFromJSON", _
            "Invalid JSON string: must start with '[' (RECORDS) or '{' (SPLIT)."
    End If

    Set LoadFromJSON = Me
End Function

Private Sub ParseJSONRecordsIntoMe(ByRef s As String)
    Dim pos As Long, maxLen As Long
    maxLen = Len(s)
    pos = 2 ' Skip opening '['

    SkipJSONWhitespace s, pos, maxLen
    If pos <= maxLen And Mid$(s, pos, 1) = "]" Then Exit Sub

    Dim rowDicts As New Collection
    Dim curObj As Object
    Dim keyName As String
    Dim valItem As Variant

    Do While pos <= maxLen
        SkipJSONWhitespace s, pos, maxLen
        If pos > maxLen Then Exit Do

        Dim ch As String
        ch = Mid$(s, pos, 1)
        If ch = "]" Then Exit Do
        If ch = "," Then
            pos = pos + 1
            SkipJSONWhitespace s, pos, maxLen
            If pos > maxLen Then Exit Do
            ch = Mid$(s, pos, 1)
        End If

        If ch <> "{" Then
            Err.Raise 5, "cDataFrame.LoadFromJSON", "Expected '{' in JSON RECORDS array at position " & pos
        End If
        pos = pos + 1 ' Skip '{'

        Set curObj = CreateObject("Scripting.Dictionary")
        curObj.CompareMode = 1

        Do While pos <= maxLen
            SkipJSONWhitespace s, pos, maxLen
            If pos > maxLen Then Exit Do
            ch = Mid$(s, pos, 1)
            If ch = "}" Then
                pos = pos + 1
                Exit Do
            End If
            If ch = "," Then
                pos = pos + 1
                SkipJSONWhitespace s, pos, maxLen
            End If

            keyName = ParseJSONStringToken(s, pos, maxLen)
            SkipJSONWhitespace s, pos, maxLen
            If pos > maxLen Or Mid$(s, pos, 1) <> ":" Then
                Err.Raise 5, "cDataFrame.LoadFromJSON", "Expected ':' after key '" & keyName & "'"
            End If
            pos = pos + 1 ' Skip ':'

            valItem = ParseJSONScalarValue(s, pos, maxLen)
            curObj.Item(keyName) = valItem

            If Not mCols.Exists(keyName) Then
                Me.AddColumn keyName
            End If
        Loop

        rowDicts.Add curObj
    Loop

    Dim numCols As Long
    numCols = mCols.Count
    If numCols = 0 Or rowDicts.Count = 0 Then Exit Sub

    Dim colKeys As Variant
    colKeys = mCols.Keys
    Dim rowVals() As Variant
    ReDim rowVals(0 To numCols - 1)

    Dim rIdx As Long, c As Long, kStr As String
    For rIdx = 1 To rowDicts.Count
        Set curObj = rowDicts(rIdx)
        For c = 0 To numCols - 1
            kStr = CStr(colKeys(c))
            If curObj.Exists(kStr) Then
                rowVals(c) = curObj.Item(kStr)
            Else
                rowVals(c) = Empty
            End If
        Next c
        Me.AddRow rowVals
    Next rIdx
End Sub

Private Sub ParseJSONSplitIntoMe(ByRef s As String)
    Dim pos As Long, maxLen As Long
    maxLen = Len(s)
    pos = 1

    ' Locate "columns" array
    Dim colKeyPos As Long
    colKeyPos = InStr(1, s, """columns""", vbTextCompare)
    If colKeyPos = 0 Then
        Err.Raise 5, "cDataFrame.LoadFromJSON", "Missing ""columns"" key in SPLIT JSON."
    End If

    pos = InStr(colKeyPos + 9, s, "[")
    If pos = 0 Then
        Err.Raise 5, "cDataFrame.LoadFromJSON", "Missing '[' for ""columns"" in SPLIT JSON."
    End If
    pos = pos + 1 ' Skip '['

    Dim ch As String, colName As String
    Do While pos <= maxLen
        SkipJSONWhitespace s, pos, maxLen
        If pos > maxLen Then Exit Do
        ch = Mid$(s, pos, 1)
        If ch = "]" Then
            pos = pos + 1
            Exit Do
        End If
        If ch = "," Then
            pos = pos + 1
            SkipJSONWhitespace s, pos, maxLen
        End If
        colName = ParseJSONStringToken(s, pos, maxLen)
        Me.AddColumn colName
    Loop

    Dim numCols As Long
    numCols = mCols.Count
    If numCols = 0 Then Exit Sub

    Dim dataKeyPos As Long
    dataKeyPos = InStr(pos, s, """data""", vbTextCompare)
    If dataKeyPos = 0 Then Exit Sub

    pos = InStr(dataKeyPos + 6, s, "[")
    If pos = 0 Then Exit Sub
    pos = pos + 1 ' Skip outer '[' of data

    Dim rowVals() As Variant
    ReDim rowVals(0 To numCols - 1)
    Dim c As Long

    Do While pos <= maxLen
        SkipJSONWhitespace s, pos, maxLen
        If pos > maxLen Then Exit Do
        ch = Mid$(s, pos, 1)
        If ch = "]" Then Exit Do
        If ch = "," Then
            pos = pos + 1
            SkipJSONWhitespace s, pos, maxLen
            If pos > maxLen Then Exit Do
            ch = Mid$(s, pos, 1)
        End If

        If ch <> "[" Then
            Err.Raise 5, "cDataFrame.LoadFromJSON", "Expected '[' for data row at position " & pos
        End If
        pos = pos + 1 ' Skip row '['

        For c = 0 To numCols - 1
             rowVals(c) = Empty
        Next c

        c = 0
        Do While pos <= maxLen
            SkipJSONWhitespace s, pos, maxLen
            If pos > maxLen Then Exit Do
            ch = Mid$(s, pos, 1)
            If ch = "]" Then
                pos = pos + 1
                Exit Do
            End If
            If ch = "," Then
                pos = pos + 1
                SkipJSONWhitespace s, pos, maxLen
            End If
            If c < numCols Then
                rowVals(c) = ParseJSONScalarValue(s, pos, maxLen)
            Else
                Call ParseJSONScalarValue(s, pos, maxLen)
            End If
            c = c + 1
        Loop

        Me.AddRow rowVals
    Loop
End Sub

Private Sub SkipJSONWhitespace(ByRef s As String, ByRef pos As Long, ByVal maxLen As Long)
    Dim ch As String
    Do While pos <= maxLen
        ch = Mid$(s, pos, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function ParseJSONStringToken(ByRef s As String, ByRef pos As Long, ByVal maxLen As Long) As String
    SkipJSONWhitespace s, pos, maxLen
    If pos > maxLen Or Mid$(s, pos, 1) <> """" Then
        Err.Raise 5, "cDataFrame.LoadFromJSON", "Expected '""' at position " & pos
    End If
    pos = pos + 1 ' Skip opening quote

    Dim res As String, ch As String, esc As String
    res = ""

    Do While pos <= maxLen
        ch = Mid$(s, pos, 1)
        If ch = "\" Then
            pos = pos + 1
            If pos > maxLen Then Exit Do
            esc = Mid$(s, pos, 1)
            Select Case esc
                Case """"
                    res = res & """"
                Case "\"
                    res = res & "\"
                Case "/"
                    res = res & "/"
                Case "r"
                    res = res & vbCr
                Case "n"
                    res = res & vbLf
                Case "t"
                    res = res & vbTab
                Case Else
                    res = res & esc
            End Select
            pos = pos + 1
        ElseIf ch = """" Then
            pos = pos + 1 ' Skip closing quote
            ParseJSONStringToken = res
            Exit Function
        Else
            res = res & ch
            pos = pos + 1
        End If
    Loop

    ParseJSONStringToken = res
End Function

Private Function ParseJSONScalarValue(ByRef s As String, ByRef pos As Long, ByVal maxLen As Long) As Variant
    SkipJSONWhitespace s, pos, maxLen
    If pos > maxLen Then
        ParseJSONScalarValue = Empty
        Exit Function
    End If

    If Mid$(s, pos, 1) = """" Then
        ParseJSONScalarValue = ParseJSONStringToken(s, pos, maxLen)
        Exit Function
    End If

    Dim startPos As Long, ch As String
    startPos = pos
    Do While pos <= maxLen
        ch = Mid$(s, pos, 1)
        If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            Exit Do
        End If
        pos = pos + 1
    Loop

    Dim rawTok As String, lowTok As String
    rawTok = Trim$(Mid$(s, startPos, pos - startPos))
    lowTok = LCase$(rawTok)

    If lowTok = "null" Or Len(rawTok) = 0 Then
        ParseJSONScalarValue = Empty
    ElseIf lowTok = "true" Then
        ParseJSONScalarValue = True
    ElseIf lowTok = "false" Then
        ParseJSONScalarValue = False
    Else
        ParseJSONScalarValue = CDbl(Val(rawTok))
    End If
End Function

Private Function FormatJSONValue(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        FormatJSONValue = "null"
        Exit Function
    End If

    Dim vt As VbVarType
    vt = VarType(v)

    If vt = vbBoolean Then
        If CBool(v) Then
            FormatJSONValue = "true"
        Else
            FormatJSONValue = "false"
        End If
    ElseIf vt = vbInteger Or vt = vbLong Or vt = vbSingle Or vt = vbDouble Or _
           vt = vbCurrency Or vt = vbDecimal Or vt = vbByte Then
        Dim numStr As String
        numStr = Trim$(Str$(CDbl(v)))
        If Left$(numStr, 1) = "." Then numStr = "0" & numStr
        If Left$(numStr, 2) = "-." Then numStr = "-0." & Mid$(numStr, 3)
        FormatJSONValue = numStr
    Else
        FormatJSONValue = """" & EscapeJSONString(CStr(v)) & """"
    End If
End Function

Private Function EscapeJSONString(ByVal s As String) As String
    Dim res As String
    res = Replace(s, "\", "\\")
    res = Replace(res, """", "\""")
    res = Replace(res, vbCr, "\r")
    res = Replace(res, vbLf, "\n")
    res = Replace(res, vbTab, "\t")
    EscapeJSONString = res
End Function

' Returns the top N rows as a new cDataFrame
Public Function Head(Optional ByVal N As Long = 5) As cDataFrame
    If N <= 0 Or mRowCount = 0 Then
        Set Head = Me.Clone(True)
        Exit Function
    End If

    Dim endR As Long
    endR = N
    If endR > mRowCount Then endR = mRowCount
    Set Head = Me.Slice(1, endR)
End Function

' Returns the bottom N rows as a new cDataFrame
Public Function Tail(Optional ByVal N As Long = 5) As cDataFrame
    If N <= 0 Or mRowCount = 0 Then
        Set Tail = Me.Clone(True)
        Exit Function
    End If

    Dim startR As Long
    startR = mRowCount - N + 1
    If startR < 1 Then startR = 1
    Set Tail = Me.Slice(startR, mRowCount)
End Function

' Returns a contiguous slice of rows [FromRow .. ToRow] (1-based, clamped to valid range)
Public Function Slice(ByVal FromRow As Long, ByVal ToRow As Long) As cDataFrame
    If mCols.Count = 0 Or mRowCount = 0 Or FromRow > ToRow Or FromRow > mRowCount Or ToRow < 1 Then
        Set Slice = Me.Clone(True)
        Exit Function
    End If

    Dim startR As Long, endR As Long, sliceLen As Long
    startR = FromRow
    If startR < 1 Then startR = 1
    endR = ToRow
    If endR > mRowCount Then endR = mRowCount

    If startR > endR Then
        Set Slice = Me.Clone(True)
        Exit Function
    End If

    sliceLen = endR - startR + 1

    Dim resDf As New cDataFrame
    Dim keys As Variant
    keys = mCols.Keys

    Dim c As Long, r As Long, srcCol As cColumn
    Dim colBuf() As Variant

    For c = 0 To mCols.Count - 1
        Set srcCol = mCols.Item(keys(c))
        ReDim colBuf(1 To sliceLen)
        For r = 1 To sliceLen
            If IsObject(srcCol.Item(startR + r - 1)) Then
                Set colBuf(r) = srcCol.Item(startR + r - 1)
            Else
                colBuf(r) = srcCol.Item(startR + r - 1)
            End If
        Next r
        resDf.AddColumn srcCol.Name, , colBuf
    Next c

    Set Slice = resDf
End Function

' Projects a subset of columns in the specified order into a new cDataFrame
Public Function SelectColumns(ByVal ColumnKeys As Variant) As cDataFrame
    Dim selNames() As String
    Dim numSel As Long
    Dim s As Long

    If IsArray(ColumnKeys) Then
        Dim lbS As Long, ubS As Long
        lbS = LBound(ColumnKeys)
        ubS = UBound(ColumnKeys)
        numSel = ubS - lbS + 1
        If numSel <= 0 Then
            Err.Raise 5, "cDataFrame.SelectColumns", "ColumnKeys array cannot be empty."
        End If
        ReDim selNames(1 To numSel)
        For s = 1 To numSel
            selNames(s) = ResolveColName(ColumnKeys(lbS + s - 1))
        Next s
    Else
        numSel = 1
        ReDim selNames(1 To 1)
        selNames(1) = ResolveColName(ColumnKeys)
    End If

    Dim resDf As New cDataFrame
    Dim col As cColumn

    For s = 1 To numSel
        If resDf.HasColumn(selNames(s)) Then
            Err.Raise 5, "cDataFrame.SelectColumns", "Duplicate column specified in SelectColumns: '" & selNames(s) & "'"
        End If
        Set col = mCols.Item(selNames(s))
        If mRowCount = 0 Then
            resDf.AddColumn col.Name
        Else
            resDf.AddColumn col.Name, , col.ToArray()
        End If
    Next s

    If mRowCount > 0 And Len(mIndexColumnName) > 0 And resDf.HasColumn(mIndexColumnName) Then
        resDf.SetIndex mIndexColumnName
    End If

    Set SelectColumns = resDf
End Function

' Generates descriptive statistics summary table for all numeric columns (similar to pandas.DataFrame.describe())
' Returns a new cDataFrame containing: count, mean, std, min, 25%, 50%, 75%, max
Public Function Describe() As cDataFrame
    Dim dfDesc As New cDataFrame
    dfDesc.AddColumn "Metric"

    Dim metrics As Variant
    metrics = Array("count", "mean", "std", "min", "25%", "50%", "75%", "max")

    Dim m As Long
    For m = LBound(metrics) To UBound(metrics)
        dfDesc.AddRow CStr(metrics(m))
    Next m

    If mCols.Count = 0 Or mRowCount = 0 Then
        dfDesc.SetIndex "Metric"
        Set Describe = dfDesc
        Exit Function
    End If

    Dim colKey As Variant
    Dim colObj As cColumn
    Dim i As Long, numCount As Long
    Dim v As Variant

    For Each colKey In mCols.Keys
        Set colObj = mCols(colKey)
        numCount = 0

        For i = 1 To colObj.Count
            v = colObj.Item(i)
            If IsNumeric(v) And Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
                numCount = numCount + 1
            End If
        Next i

        If numCount > 0 Then
            Dim cntVal As Double, meanVal As Double, stdVal As Double
            Dim minVal As Variant, q25Val As Double, q50Val As Double, q75Val As Double, maxVal As Variant

            cntVal = CDbl(numCount)
            meanVal = colObj.Average()
            stdVal = colObj.StdDev(True)
            minVal = colObj.Min()
            q25Val = colObj.Quantile(0.25)
            q50Val = colObj.Quantile(0.5)
            q75Val = colObj.Quantile(0.75)
            maxVal = colObj.Max()

            dfDesc.AddColumn CStr(colKey), , Array(cntVal, meanVal, stdVal, CDbl(minVal), q25Val, q50Val, q75Val, CDbl(maxVal))
        End If
    Next colKey

    dfDesc.SetIndex "Metric"
    Set Describe = dfDesc
End Function

' Computes pairwise Pearson correlation matrix for all numeric columns (similar to pandas.DataFrame.corr())
' Returns an M x M square cDataFrame indexed by column names.
Public Function Corr() As cDataFrame
    Dim dfCorr As New cDataFrame
    dfCorr.AddColumn "Column"

    If mCols.Count = 0 Or mRowCount = 0 Then
        dfCorr.SetIndex "Column"
        Set Corr = dfCorr
        Exit Function
    End If

    ' Find all numeric columns
    Dim numCols As Object
    Set numCols = CreateObject("Scripting.Dictionary")

    Dim colKey As Variant
    Dim colObj As cColumn
    Dim i As Long, numCount As Long

    For Each colKey In mCols.Keys
        Set colObj = mCols(colKey)
        numCount = 0
        For i = 1 To colObj.Count
            If IsNumeric(colObj.Item(i)) And Not IsEmpty(colObj.Item(i)) And Not IsNull(colObj.Item(i)) And Not IsError(colObj.Item(i)) Then
                numCount = numCount + 1
            End If
        Next i
        If numCount > 0 Then
            numCols.Add CStr(colKey), colObj
        End If
    Next colKey

    Dim colNames() As String
    Dim numM As Long
    numM = numCols.Count

    If numM = 0 Then
        dfCorr.SetIndex "Column"
        Set Corr = dfCorr
        Exit Function
    End If

    ReDim colNames(0 To numM - 1)
    Dim k As Long
    k = 0
    For Each colKey In numCols.Keys
        colNames(k) = CStr(colKey)
        dfCorr.AddRow CStr(colKey)
        k = k + 1
    Next colKey

    ' Calculate pairwise correlations
    Dim col1 As cColumn, col2 As cColumn
    Dim rIdx As Long, cIdx As Long
    Dim rVal As Double
    Dim colVals() As Variant

    For cIdx = 0 To numM - 1
        Set col2 = numCols(colNames(cIdx))
        ReDim colVals(0 To numM - 1)

        For rIdx = 0 To numM - 1
            If rIdx = cIdx Then
                colVals(rIdx) = 1#
            Else
                Set col1 = numCols(colNames(rIdx))
                rVal = col1.Correlation(col2)
                colVals(rIdx) = Round(rVal, 6)
            End If
        Next rIdx

        dfCorr.AddColumn colNames(cIdx), , colVals
    Next cIdx

    dfCorr.SetIndex "Column"
    Set Corr = dfCorr
End Function

' Exports DataFrame as a GitHub Flavored Markdown table string
Public Function ToMarkdown(Optional ByVal MaxRows As Long = 0) As String
    If mCols.Count = 0 Then
        ToMarkdown = "*(Empty DataFrame)*"
        Exit Function
    End If

    Dim sb As String
    Dim colKey As Variant, c As Long, r As Long
    Dim colNames() As String
    Dim numCols As Long
    numCols = mCols.Count
    ReDim colNames(1 To numCols)

    ' Header row
    sb = "| "
    c = 1
    For Each colKey In mCols.Keys
        colNames(c) = CStr(colKey)
        sb = sb & Replace(colNames(c), "|", "\|") & " | "
        c = c + 1
    Next colKey
    sb = sb & vbCrLf

    ' Separator row
    sb = sb & "| "
    For c = 1 To numCols
        sb = sb & ":--- | "
    Next c
    sb = sb & vbCrLf

    If mRowCount = 0 Then
        ToMarkdown = sb
        Exit Function
    End If

    Dim renderRows As Long
    renderRows = mRowCount
    If MaxRows > 0 And MaxRows < renderRows Then
        renderRows = MaxRows
    End If

    Dim v As Variant, strVal As String
    For r = 1 To renderRows
        sb = sb & "| "
        For c = 1 To numCols
            v = mCols(colNames(c)).Item(r)
            If IsError(v) Then
                strVal = "#ERR:" & CStr(CLng(v))
            ElseIf IsEmpty(v) Or IsNull(v) Then
                strVal = ""
            Else
                strVal = CStr(v)
            End If
            sb = sb & Replace(strVal, "|", "\|") & " | "
        Next c
        sb = sb & vbCrLf
    Next r

    If MaxRows > 0 And mRowCount > MaxRows Then
        sb = sb & "| *... (" & CStr(mRowCount - MaxRows) & " more rows) ...* |" & vbCrLf
    End If

    ToMarkdown = sb
End Function

' Exports DataFrame as an HTML table string with optional modern responsive CSS styling
Public Function ToHTML(Optional ByVal TableClass As String = "dataframe", _
                       Optional ByVal IncludeCSS As Boolean = True) As String
    If mCols.Count = 0 Then
        ToHTML = "<table class=""" & TableClass & """><tbody><tr><td>(Empty DataFrame)</td></tr></tbody></table>"
        Exit Function
    End If

    Dim sb As String
    If IncludeCSS Then
        sb = "<style>" & vbCrLf & _
             "  table." & TableClass & " { border-collapse: collapse; width: 100%; font-family: Segoe UI, Arial, sans-serif; font-size: 13px; }" & vbCrLf & _
             "  table." & TableClass & " th { background-color: #1F4E79; color: white; font-weight: 600; text-align: left; padding: 8px 12px; border: 1px solid #D9D9D9; }" & vbCrLf & _
             "  table." & TableClass & " td { padding: 6px 12px; border: 1px solid #E0E0E0; }" & vbCrLf & _
             "  table." & TableClass & " tr:nth-child(even) { background-color: #F8F9FA; }" & vbCrLf & _
             "  table." & TableClass & " tr:hover { background-color: #F0F4F8; }" & vbCrLf & _
             "</style>" & vbCrLf
    End If

    sb = sb & "<table class=""" & TableClass & """>" & vbCrLf
    sb = sb & "  <thead>" & vbCrLf & "    <tr>" & vbCrLf

    Dim colKey As Variant, c As Long, r As Long
    Dim colNames() As String
    Dim numCols As Long
    numCols = mCols.Count
    ReDim colNames(1 To numCols)

    c = 1
    For Each colKey In mCols.Keys
        colNames(c) = CStr(colKey)
        sb = sb & "      <th>" & colNames(c) & "</th>" & vbCrLf
        c = c + 1
    Next colKey

    sb = sb & "    </tr>" & vbCrLf & "  </thead>" & vbCrLf
    sb = sb & "  <tbody>" & vbCrLf

    Dim v As Variant, strVal As String
    For r = 1 To mRowCount
        sb = sb & "    <tr>" & vbCrLf
        For c = 1 To numCols
            v = mCols(colNames(c)).Item(r)
            If IsError(v) Then
                strVal = "#ERR:" & CStr(CLng(v))
            ElseIf IsEmpty(v) Or IsNull(v) Then
                strVal = ""
            Else
                strVal = Replace(Replace(CStr(v), "<", "&lt;"), ">", "&gt;")
            End If
            sb = sb & "      <td>" & strVal & "</td>" & vbCrLf
        Next c
        sb = sb & "    </tr>" & vbCrLf
    Next r

    sb = sb & "  </tbody>" & vbCrLf & "</table>"
    ToHTML = sb
End Function

' Evaluates an arithmetic expression containing column names, numbers, operators (+, -, *, /) and parentheses
' Returns a newly computed cColumn evaluated entirely in RAM via Vectorized RPN Execution
Public Function Eval(ByVal Expression As String, Optional ByVal ResultColName As String = "") As cColumn
    Dim cleanExpr As String
    cleanExpr = Trim$(Expression)
    If Len(cleanExpr) = 0 Then
        Err.Raise 5, "cDataFrame.Eval", "Expression cannot be blank."
    End If

    ' 1. Tokenize expression into tokens (operators, parenthesis, identifiers, numbers)
    Dim tokens As New Collection
    Dim p As Long, L As Long
    Dim ch As String, curToken As String
    L = Len(cleanExpr)
    p = 1

    Do While p <= L
        ch = Mid$(cleanExpr, p, 1)

        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            p = p + 1
        ElseIf ch = "(" Or ch = ")" Or ch = "+" Or ch = "-" Or ch = "*" Or ch = "/" Then
            tokens.Add ch
            p = p + 1
        Else
            curToken = ""
            Do While p <= L
                ch = Mid$(cleanExpr, p, 1)
                If ch = " " Or ch = vbTab Or ch = "(" Or ch = ")" Or ch = "+" Or ch = "-" Or ch = "*" Or ch = "/" Then
                    Exit Do
                End If
                curToken = curToken & ch
                p = p + 1
            Loop
            tokens.Add curToken
        End If
    Loop

    ' 2. Shunting-Yard Algorithm: Infix -> RPN (Reverse Polish Notation)
    Dim outputQueue As New Collection
    Dim opStack As New Collection
    Dim tVal As String, topOp As String
    Dim precT As Integer, precTop As Integer

    For p = 1 To tokens.Count
        tVal = CStr(tokens(p))

        If tVal = "+" Or tVal = "-" Or tVal = "*" Or tVal = "/" Then
            precT = IIf(tVal = "*" Or tVal = "/", 2, 1)
            Do While opStack.Count > 0
                topOp = CStr(opStack(opStack.Count))
                If topOp = "(" Then Exit Do
                precTop = IIf(topOp = "*" Or topOp = "/", 2, 1)
                If precTop >= precT Then
                    outputQueue.Add topOp
                    opStack.Remove opStack.Count
                Else
                    Exit Do
                End If
            Loop
            opStack.Add tVal
        ElseIf tVal = "(" Then
            opStack.Add tVal
        ElseIf tVal = ")" Then
            Do While opStack.Count > 0
                topOp = CStr(opStack(opStack.Count))
                opStack.Remove opStack.Count
                If topOp = "(" Then Exit Do
                outputQueue.Add topOp
            Loop
        Else
            outputQueue.Add tVal
        End If
    Next p

    Do While opStack.Count > 0
        outputQueue.Add CStr(opStack(opStack.Count))
        opStack.Remove opStack.Count
    Loop

    ' 3. Evaluate RPN Stack on Vectorized Columns / Numeric Scalars (Typed Stack)
    Dim stackObj(0 To 64) As cColumn
    Dim stackNum(0 To 64) As Double
    Dim stackIsObj(0 To 64) As Boolean
    Dim stackTop As Long
    stackTop = 0

    Dim tok As Variant, tStr As String
    Dim isObjA As Boolean, isObjB As Boolean
    Dim colA As cColumn, colB As cColumn, colRes As cColumn
    Dim numA As Double, numB As Double

    For p = 1 To outputQueue.Count
        tok = outputQueue(p)
        tStr = CStr(tok)

        If tStr = "+" Or tStr = "-" Or tStr = "*" Or tStr = "/" Then
            If stackTop < 2 Then
                Err.Raise 5, "cDataFrame.Eval", "Malformed expression: insufficient operands for '" & tStr & "'"
            End If

            ' Pop B
            isObjB = stackIsObj(stackTop)
            If isObjB Then
                Set colB = stackObj(stackTop)
                Set stackObj(stackTop) = Nothing
            Else
                numB = stackNum(stackTop)
            End If
            stackTop = stackTop - 1

            ' Pop A
            isObjA = stackIsObj(stackTop)
            If isObjA Then
                Set colA = stackObj(stackTop)
                Set stackObj(stackTop) = Nothing
            Else
                numA = stackNum(stackTop)
            End If
            stackTop = stackTop - 1

            ' Vectorized arithmetic
            If isObjA And isObjB Then
                Set colRes = colA.VectorOp(tStr, colB)
                stackTop = stackTop + 1
                Set stackObj(stackTop) = colRes
                stackIsObj(stackTop) = True
            ElseIf isObjA And Not isObjB Then
                Set colRes = colA.VectorOp(tStr, numB)
                stackTop = stackTop + 1
                Set stackObj(stackTop) = colRes
                stackIsObj(stackTop) = True
            ElseIf Not isObjA And isObjB Then
                Dim colScalar As New cColumn
                colScalar.EnsureCapacity mRowCount
                colScalar.SetCount mRowCount, numA
                Set colRes = colScalar.VectorOp(tStr, colB)
                stackTop = stackTop + 1
                Set stackObj(stackTop) = colRes
                stackIsObj(stackTop) = True
            Else
                ' Pure scalar
                Dim numRes As Double
                Select Case tStr
                    Case "+": numRes = numA + numB
                    Case "-": numRes = numA - numB
                    Case "*": numRes = numA * numB
                    Case "/":
                        If numB <> 0# Then
                            numRes = numA / numB
                        Else
                            numRes = 0#
                        End If
                End Select
                stackTop = stackTop + 1
                stackNum(stackTop) = numRes
                stackIsObj(stackTop) = False
            End If
        Else
            ' Operand: column or number
            If Me.HasColumn(tStr) Then
                stackTop = stackTop + 1
                Set stackObj(stackTop) = Me.Column(tStr)
                stackIsObj(stackTop) = True
            ElseIf IsNumeric(tStr) Then
                stackTop = stackTop + 1
                stackNum(stackTop) = CDbl(tStr)
                stackIsObj(stackTop) = False
            Else
                Err.Raise 5, "cDataFrame.Eval", "Unrecognized token or missing column: '" & tStr & "'"
            End If
        End If
    Next p

    If stackTop <> 1 Then
        Err.Raise 5, "cDataFrame.Eval", "Malformed expression: leftover operands on stack."
    End If

    If stackIsObj(1) Then
        Set colRes = stackObj(1)
    Else
        Set colRes = New cColumn
        colRes.EnsureCapacity mRowCount
        colRes.SetCount mRowCount, stackNum(1)
    End If

    If Len(Trim$(ResultColName)) > 0 Then
        colRes.Name = Trim$(ResultColName)
    Else
        colRes.Name = cleanExpr
    End If

    Set Eval = colRes
End Function

' Vectorized conditional ternary evaluation (IIf) over all rows returning a new cColumn.
' TrueValue and FalseValue can each be a scalar Variant or a cColumn of matching length.
Public Function Where(ByVal ColumnName As String, _
                      ByVal Operator As String, _
                      ByVal CompareValue As Variant, _
                      ByVal TrueValue As Variant, _
                      ByVal FalseValue As Variant, _
                      Optional ByVal ResultColName As String = "") As cColumn
    Dim targetCol As cColumn
    Set targetCol = mCols.Item(ResolveColName(ColumnName))

    Dim res As New cColumn
    If Len(Trim$(ResultColName)) > 0 Then
        res.Name = Trim$(ResultColName)
    Else
        res.Name = targetCol.Name & "_Where"
    End If

    If mRowCount = 0 Then
        Set Where = res
        Exit Function
    End If

    Dim trueIsCol As Boolean, falseIsCol As Boolean
    Dim trueCol As cColumn, falseCol As cColumn

    If IsObject(TrueValue) Then
        If TypeName(TrueValue) = "cColumn" Then
            trueIsCol = True
            Set trueCol = TrueValue
            If trueCol.Count <> mRowCount Then
                Err.Raise 5, "cDataFrame.Where", "TrueValue cColumn count does not match DataFrame RowCount."
            End If
        End If
    End If

    If IsObject(FalseValue) Then
        If TypeName(FalseValue) = "cColumn" Then
            falseIsCol = True
            Set falseCol = FalseValue
            If falseCol.Count <> mRowCount Then
                Err.Raise 5, "cDataFrame.Where", "FalseValue cColumn count does not match DataFrame RowCount."
            End If
        End If
    End If

    res.EnsureCapacity mRowCount
    res.SetCount mRowCount

    Dim normOp As String
    normOp = UCase$(Trim$(Operator))

    Dim r As Long
    For r = 1 To mRowCount
        If EvaluateCondition(targetCol.Item(r), normOp, CompareValue) Then
            If trueIsCol Then
                res.Item(r) = trueCol.Item(r)
            Else
                res.Item(r) = TrueValue
            End If
        Else
            If falseIsCol Then
                res.Item(r) = falseCol.Item(r)
            Else
                res.Item(r) = FalseValue
            End If
        End If
    Next r

    Set Where = res
End Function

Public Function Clone(Optional ByVal StructureOnly As Boolean = False) As cDataFrame
    Dim newDf As New cDataFrame
    Dim keys As Variant
    keys = mCols.Keys

    Dim c As Long, col As cColumn
    For c = 0 To mCols.Count - 1
        Set col = mCols.Item(keys(c))
        If StructureOnly Then
            newDf.AddColumn col.Name
        Else
            newDf.AddColumn col.Name, , col.ToArray()
        End If
    Next c

    If Not StructureOnly And mIndexColCount > 0 Then
        If mIndexColCount = 1 Then
            newDf.SetIndex mIndexColumnNames(1), mIndexSkipEmpty
        Else
            newDf.SetIndex mIndexColumnNames, mIndexSkipEmpty
        End If
    End If

    Set Clone = newDf
End Function

' Fast O(N) row deduplication using Scripting.Dictionary hash keys.
' SubsetColumns: optional single column key or 1D Array of column keys (defaults to all columns).
' Keep: "FIRST" (default), "LAST", or "NONE"/"FALSE" (drops all keys that appear more than once).
Public Function DropDuplicates(Optional ByVal SubsetColumns As Variant, _
                               Optional ByVal Keep As String = "FIRST") As cDataFrame
    Dim keepMode As String
    keepMode = UCase$(Trim$(Keep))
    If keepMode = "FALSE" Then keepMode = "NONE"

    If keepMode <> "FIRST" And keepMode <> "LAST" And keepMode <> "NONE" Then
        Err.Raise 5, "cDataFrame.DropDuplicates", _
            "Unsupported Keep mode: '" & Keep & "'. Supported: FIRST, LAST, NONE"
    End If

    If mCols.Count = 0 Then
        Dim emptySchemaDf As New cDataFrame
        Set DropDuplicates = emptySchemaDf
        Exit Function
    End If

    ' 1. Resolve subset columns (all columns if omitted/empty)
    Dim subColNames() As String
    Dim numSubCols As Long
    Dim s As Long

    If IsMissing(SubsetColumns) Or IsEmpty(SubsetColumns) Then
        numSubCols = mCols.Count
        subColNames = Me.ColumnNames()
    ElseIf IsArray(SubsetColumns) Then
        Dim lbS As Long, ubS As Long
        lbS = LBound(SubsetColumns)
        ubS = UBound(SubsetColumns)
        numSubCols = ubS - lbS + 1
        If numSubCols <= 0 Then
            Err.Raise 5, "cDataFrame.DropDuplicates", "SubsetColumns array cannot be empty."
        End If
        ReDim subColNames(1 To numSubCols)
        For s = 1 To numSubCols
            subColNames(s) = ResolveColName(SubsetColumns(lbS + s - 1))
        Next s
    Else
        numSubCols = 1
        ReDim subColNames(1 To 1)
        subColNames(1) = ResolveColName(SubsetColumns)
    End If

    Dim resultDf As cDataFrame
    Set resultDf = Me.Clone(True)

    If mRowCount = 0 Then
        Set DropDuplicates = resultDf
        Exit Function
    End If

    ' 2. Pre-resolve subset column objects and all column objects
    Dim subColObjs() As cColumn
    ReDim subColObjs(1 To numSubCols)
    For s = 1 To numSubCols
        Set subColObjs(s) = mCols.Item(subColNames(s))
    Next s

    Dim allColCount As Long
    allColCount = mCols.Count
    Dim allKeys As Variant
    allKeys = mCols.Keys
    Dim allColObjs() As cColumn
    ReDim allColObjs(0 To allColCount - 1)
    Dim c As Long
    For c = 0 To allColCount - 1
        Set allColObjs(c) = mCols.Item(allKeys(c))
    Next c

    ' 3. Pass 1: Track firstRow, lastRow, and count per unique composite key in O(N)
    Dim groupDict As Object
    Set groupDict = CreateObject("Scripting.Dictionary")
    groupDict.CompareMode = 1

    Dim firstRow() As Long
    Dim lastRow() As Long
    Dim grpCount() As Long
    ReDim firstRow(1 To mRowCount)
    ReDim lastRow(1 To mRowCount)
    ReDim grpCount(1 To mRowCount)

    Dim numGroups As Long
    numGroups = 0

    Dim r As Long, compKey As String, gIdx As Long
    For r = 1 To mRowCount
        If numSubCols = 1 Then
            compKey = SafeDedupToken(subColObjs(1).Item(r))
        Else
            compKey = ""
            For s = 1 To numSubCols
                Dim dedupTok As String
                dedupTok = SafeDedupToken(subColObjs(s).Item(r))
                compKey = compKey & CStr(Len(dedupTok)) & ":" & dedupTok & "|"
            Next s
        End If

        If groupDict.Exists(compKey) Then
            gIdx = groupDict.Item(compKey)
            lastRow(gIdx) = r
            grpCount(gIdx) = grpCount(gIdx) + 1
        Else
            numGroups = numGroups + 1
            gIdx = numGroups
            groupDict.Add compKey, gIdx
            firstRow(gIdx) = r
            lastRow(gIdx) = r
            grpCount(gIdx) = 1
        End If
    Next r

    ' 4. Mark winning rows in boolean mask to preserve original row ordering
    Dim keepRow() As Boolean
    ReDim keepRow(1 To mRowCount)

    For gIdx = 1 To numGroups
        If keepMode = "FIRST" Then
            keepRow(firstRow(gIdx)) = True
        ElseIf keepMode = "LAST" Then
            keepRow(lastRow(gIdx)) = True
        ElseIf keepMode = "NONE" Then
            If grpCount(gIdx) = 1 Then
                keepRow(firstRow(gIdx)) = True
            End If
        End If
    Next gIdx

    ' 5. Emit surviving rows in original order
    Dim rowVals() As Variant
    ReDim rowVals(0 To allColCount - 1)

    For r = 1 To mRowCount
        If keepRow(r) Then
            For c = 0 To allColCount - 1
                rowVals(c) = allColObjs(c).Item(r)
            Next c
            resultDf.AddRow rowVals
        End If
    Next r

    Set DropDuplicates = resultDf
End Function

' Vertical concatenation (UNION ALL) appending OtherTable rows below Me.
' When AllowMissingColumns = False (default), requires identical column sets (in any order).
' When AllowMissingColumns = True, builds the superset of columns and fills missing slots with Empty.
Public Function Concat(ByVal OtherTable As cDataFrame, _
                       Optional ByVal AllowMissingColumns As Boolean = False) As cDataFrame
    If OtherTable Is Nothing Then
        Err.Raise 91, "cDataFrame.Concat", "OtherTable cannot be Nothing."
    End If

    Dim myNames() As String, otherNames() As String
    myNames = Me.ColumnNames()
    otherNames = OtherTable.ColumnNames()

    Dim numMyCols As Long, numOtherCols As Long
    numMyCols = Me.ColumnCount
    numOtherCols = OtherTable.ColumnCount

    Dim outColDict As Object
    Set outColDict = CreateObject("Scripting.Dictionary")
    outColDict.CompareMode = 1

    Dim c As Long, colName As String

    If Not AllowMissingColumns Then
        If numMyCols <> numOtherCols Then
            Err.Raise 5, "cDataFrame.Concat", _
                "Column count mismatch (" & numMyCols & " vs " & numOtherCols & "). Set AllowMissingColumns:=True to union disparate schemas."
        End If
        For c = 1 To numMyCols
            colName = myNames(c)
            If Not OtherTable.HasColumn(colName) Then
                Err.Raise 5, "cDataFrame.Concat", _
                    "Column '" & colName & "' is missing in OtherTable. Set AllowMissingColumns:=True to fill with Empty."
            End If
            outColDict.Add colName, colName
        Next c
    Else
        For c = 1 To numMyCols
            colName = myNames(c)
            outColDict.Add colName, colName
        Next c
        For c = 1 To numOtherCols
            colName = otherNames(c)
            If Not outColDict.Exists(colName) Then
                outColDict.Add colName, colName
            End If
        Next c
    End If

    Dim resultDf As New cDataFrame
    Dim totalOutCols As Long
    totalOutCols = outColDict.Count
    If totalOutCols = 0 Then
        Set Concat = resultDf
        Exit Function
    End If

    Dim myRows As Long, otherRows As Long, totalRows As Long
    myRows = mRowCount
    otherRows = OtherTable.RowCount
    totalRows = myRows + otherRows

    Dim outKeys As Variant
    outKeys = outColDict.Keys

    If totalRows = 0 Then
        For c = 0 To totalOutCols - 1
            resultDf.AddColumn CStr(outKeys(c))
        Next c
        Set Concat = resultDf
        Exit Function
    End If

    ' Bulk build each 1D column vector (1 To totalRows) and attach via AddColumn
    Dim colBuf() As Variant
    Dim r As Long
    Dim srcCol As cColumn

    For c = 0 To totalOutCols - 1
        colName = CStr(outKeys(c))
        ReDim colBuf(1 To totalRows)

        If Me.HasColumn(colName) And myRows > 0 Then
            Set srcCol = mCols.Item(colName)
            For r = 1 To myRows
                If IsObject(srcCol.Item(r)) Then
                    Set colBuf(r) = srcCol.Item(r)
                Else
                    colBuf(r) = srcCol.Item(r)
                End If
            Next r
        End If

        If OtherTable.HasColumn(colName) And otherRows > 0 Then
            Set srcCol = OtherTable.Column(colName)
            For r = 1 To otherRows
                If IsObject(srcCol.Item(r)) Then
                    Set colBuf(myRows + r) = srcCol.Item(r)
                Else
                    colBuf(myRows + r) = srcCol.Item(r)
                End If
            Next r
        End If

        resultDf.AddColumn colName, , colBuf
    Next c

    Set Concat = resultDf
End Function

' Replaces missing values (Empty, Null, Error, and optionally blank strings) in-place
' across all columns or a specified subset of columns.
Public Sub FillNA(ByVal FillValue As Variant, _
                  Optional ByVal SubsetColumns As Variant, _
                  Optional ByVal TreatBlankStringAsNA As Boolean = False)
    If mCols.Count = 0 Or mRowCount = 0 Then Exit Sub

    Dim subColNames() As String
    Dim numSubCols As Long
    Dim s As Long

    If IsMissing(SubsetColumns) Or IsEmpty(SubsetColumns) Then
        numSubCols = mCols.Count
        subColNames = Me.ColumnNames()
    ElseIf IsArray(SubsetColumns) Then
        Dim lbS As Long, ubS As Long
        lbS = LBound(SubsetColumns)
        ubS = UBound(SubsetColumns)
        numSubCols = ubS - lbS + 1
        If numSubCols <= 0 Then
            Err.Raise 5, "cDataFrame.FillNA", "SubsetColumns array cannot be empty."
        End If
        ReDim subColNames(1 To numSubCols)
        For s = 1 To numSubCols
            subColNames(s) = ResolveColName(SubsetColumns(lbS + s - 1))
        Next s
    Else
        numSubCols = 1
        ReDim subColNames(1 To 1)
        subColNames(1) = ResolveColName(SubsetColumns)
    End If

    Dim col As cColumn, colName As String
    For s = 1 To numSubCols
        colName = subColNames(s)
        Set col = mCols.Item(colName)
        col.FillNA FillValue, TreatBlankStringAsNA
        If IsIndexedColumn(colName) Then
            mIndexDirty = True
        End If
    Next s
End Sub

' Returns a new cDataFrame with rows containing missing values removed in O(N).
' How: "ANY" (default, drops row if any subset column is missing) or
'      "ALL" (drops row only if all subset columns are missing).
Public Function DropNA(Optional ByVal SubsetColumns As Variant, _
                       Optional ByVal How As String = "ANY", _
                       Optional ByVal TreatBlankStringAsNA As Boolean = False) As cDataFrame
    Dim howMode As String
    howMode = UCase$(Trim$(How))
    If howMode <> "ANY" And howMode <> "ALL" Then
        Err.Raise 5, "cDataFrame.DropNA", _
            "Unsupported How mode: '" & How & "'. Supported: ANY, ALL"
    End If

    If mCols.Count = 0 Then
        Dim emptySchemaDf As New cDataFrame
        Set DropNA = emptySchemaDf
        Exit Function
    End If

    Dim subColNames() As String
    Dim numSubCols As Long
    Dim s As Long

    If IsMissing(SubsetColumns) Or IsEmpty(SubsetColumns) Then
        numSubCols = mCols.Count
        subColNames = Me.ColumnNames()
    ElseIf IsArray(SubsetColumns) Then
        Dim lbS As Long, ubS As Long
        lbS = LBound(SubsetColumns)
        ubS = UBound(SubsetColumns)
        numSubCols = ubS - lbS + 1
        If numSubCols <= 0 Then
            Err.Raise 5, "cDataFrame.DropNA", "SubsetColumns array cannot be empty."
        End If
        ReDim subColNames(1 To numSubCols)
        For s = 1 To numSubCols
            subColNames(s) = ResolveColName(SubsetColumns(lbS + s - 1))
        Next s
    Else
        numSubCols = 1
        ReDim subColNames(1 To 1)
        subColNames(1) = ResolveColName(SubsetColumns)
    End If

    Dim resultDf As cDataFrame
    Set resultDf = Me.Clone(True)

    If mRowCount = 0 Then
        Set DropNA = resultDf
        Exit Function
    End If

    Dim subColObjs() As cColumn
    ReDim subColObjs(1 To numSubCols)
    For s = 1 To numSubCols
        Set subColObjs(s) = mCols.Item(subColNames(s))
    Next s

    Dim allColCount As Long
    allColCount = mCols.Count
    Dim allKeys As Variant
    allKeys = mCols.Keys
    Dim allColObjs() As cColumn
    ReDim allColObjs(0 To allColCount - 1)
    Dim c As Long
    For c = 0 To allColCount - 1
        Set allColObjs(c) = mCols.Item(allKeys(c))
    Next c

    Dim rowVals() As Variant
    ReDim rowVals(0 To allColCount - 1)

    Dim r As Long, missingCount As Long, keepRow As Boolean
    For r = 1 To mRowCount
        missingCount = 0
        For s = 1 To numSubCols
            If IsMissingCellValue(subColObjs(s).Item(r), TreatBlankStringAsNA) Then
                missingCount = missingCount + 1
                If howMode = "ANY" Then Exit For
            End If
        Next s

        If howMode = "ANY" Then
            keepRow = (missingCount = 0)
        Else
            keepRow = (missingCount < numSubCols)
        End If

        If keepRow Then
            For c = 0 To allColCount - 1
                rowVals(c) = allColObjs(c).Item(r)
            Next c
            resultDf.AddRow rowVals
        End If
    Next r

    Set DropNA = resultDf
End Function

Private Function IsMissingCellValue(ByVal v As Variant, ByVal TreatBlankStringAsNA As Boolean) As Boolean
    If IsError(v) Then
        IsMissingCellValue = True
    ElseIf IsEmpty(v) Or IsNull(v) Then
        IsMissingCellValue = True
    ElseIf TreatBlankStringAsNA And VarType(v) = vbString Then
        IsMissingCellValue = (Len(Trim$(CStr(v))) = 0)
    Else
        IsMissingCellValue = False
    End If
End Function

' ------------------------------------------------------------------------------
' Internal Helpers
' ------------------------------------------------------------------------------

Public Function ResolveColName(ByVal ColumnKey As Variant) As String
    If VarType(ColumnKey) = vbString Then
        Dim sKey As String
        sKey = Trim$(CStr(ColumnKey))
        If Not mCols.Exists(sKey) Then
            Err.Raise 5, "cDataFrame.ResolveColName", "Column not found: '" & sKey & "'"
        End If
        ResolveColName = sKey
    ElseIf IsNumeric(ColumnKey) Then
        Dim idx As Long
        idx = CLng(ColumnKey)
        If idx < 1 Or idx > mCols.Count Then
            Err.Raise 9, "cDataFrame.ResolveColName", _
                "Column index out of range: " & idx & " (Valid: 1 To " & mCols.Count & ")"
        End If
        Dim keys As Variant
        keys = mCols.Keys
        ResolveColName = CStr(keys(idx - 1))
    Else
        Err.Raise 13, "cDataFrame.ResolveColName", _
            "Invalid column key type: " & TypeName(ColumnKey)
    End If
End Function

Private Function EvaluateCondition(ByVal ActualVal As Variant, _
                                  ByVal Op As String, _
                                  ByVal CompVal As Variant) As Boolean
    On Error GoTo MatchError

    If IsNull(ActualVal) Or IsEmpty(ActualVal) Then
        If Op = "=" Or Op = "==" Then
            EvaluateCondition = (IsNull(CompVal) Or IsEmpty(CompVal))
        ElseIf Op = "<>" Then
            EvaluateCondition = Not (IsNull(CompVal) Or IsEmpty(CompVal))
        Else
            EvaluateCondition = False
        End If
        Exit Function
    End If

    Select Case Op
        Case "=", "=="
            If IsNumeric(ActualVal) And IsNumeric(CompVal) Then
                EvaluateCondition = (CDbl(ActualVal) = CDbl(CompVal))
            Else
                EvaluateCondition = (CStr(ActualVal) = CStr(CompVal))
            End If

        Case "<>"
            If IsNumeric(ActualVal) And IsNumeric(CompVal) Then
                EvaluateCondition = (CDbl(ActualVal) <> CDbl(CompVal))
            Else
                EvaluateCondition = (CStr(ActualVal) <> CStr(CompVal))
            End If

        Case ">"
            If IsNumeric(ActualVal) And IsNumeric(CompVal) Then
                EvaluateCondition = (CDbl(ActualVal) > CDbl(CompVal))
            Else
                EvaluateCondition = (CStr(ActualVal) > CStr(CompVal))
            End If

        Case ">="
            If IsNumeric(ActualVal) And IsNumeric(CompVal) Then
                EvaluateCondition = (CDbl(ActualVal) >= CDbl(CompVal))
            Else
                EvaluateCondition = (CStr(ActualVal) >= CStr(CompVal))
            End If

        Case "<"
            If IsNumeric(ActualVal) And IsNumeric(CompVal) Then
                EvaluateCondition = (CDbl(ActualVal) < CDbl(CompVal))
            Else
                EvaluateCondition = (CStr(ActualVal) < CStr(CompVal))
            End If

        Case "<="
            If IsNumeric(ActualVal) And IsNumeric(CompVal) Then
                EvaluateCondition = (CDbl(ActualVal) <= CDbl(CompVal))
            Else
                EvaluateCondition = (CStr(ActualVal) <= CStr(CompVal))
            End If

        Case "LIKE"
            EvaluateCondition = (CStr(ActualVal) Like CStr(CompVal))

        Case "CONTAINS"
            EvaluateCondition = (InStr(1, CStr(ActualVal), CStr(CompVal), vbTextCompare) > 0)

        Case Else
            Err.Raise 5, "cDataFrame.Filter", "Unsupported filter operator: '" & Op & "'"
    End Select
    Exit Function

MatchError:
    EvaluateCondition = False
End Function

