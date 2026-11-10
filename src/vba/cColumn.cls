VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' Class: cColumn
' Description: High-performance 1D columnar store buffer for VBA DataFrame.
'              Implements geometric buffer expansion (amortized O(1) appends),
'              native 1D array operations, and vectorized math helpers.
' Author: Korn (AI Backend Developer, Technodevmoo)
' Architecture: Kit (Tech Lead)
' ==============================================================================

Private Const DEFAULT_CAPACITY As Long = 16
Private Const GROWTH_FACTOR As Double = 2#

Private mName As String
Private mData() As Variant
Private mCount As Long
Private mCapacity As Long

' ------------------------------------------------------------------------------
' Lifecycle
' ------------------------------------------------------------------------------

Private Sub Class_Initialize()
    mCapacity = DEFAULT_CAPACITY
    mCount = 0
    ReDim mData(1 To mCapacity)
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    On Error Resume Next
    For i = 1 To mCount
        If IsObject(mData(i)) Then
            Set mData(i) = Nothing
        End If
    Next i
    On Error GoTo 0
    Erase mData
End Sub

' ------------------------------------------------------------------------------
' Properties
' ------------------------------------------------------------------------------

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal Value As String)
    mName = Trim$(Value)
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

' Default member for col(RowIndex)
Public Property Get Item(ByVal RowIndex As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If RowIndex < 1 Or RowIndex > mCount Then
        Err.Raise 9, "cColumn.Item", "Subscript out of range: " & RowIndex & " (Valid range: 1 To " & mCount & ")"
    End If

    If IsObject(mData(RowIndex)) Then
        Set Item = mData(RowIndex)
    Else
        Item = mData(RowIndex)
    End If
End Property

Public Property Let Item(ByVal RowIndex As Long, ByVal Value As Variant)
    If RowIndex < 1 Or RowIndex > mCount Then
        Err.Raise 9, "cColumn.Item", "Subscript out of range: " & RowIndex & " (Valid range: 1 To " & mCount & ")"
    End If
    mData(RowIndex) = Value
End Property

Public Property Set Item(ByVal RowIndex As Long, ByVal Value As Object)
    If RowIndex < 1 Or RowIndex > mCount Then
        Err.Raise 9, "cColumn.Item", "Subscript out of range: " & RowIndex & " (Valid range: 1 To " & mCount & ")"
    End If
    Set mData(RowIndex) = Value
End Property

' ------------------------------------------------------------------------------
' Buffer Management & Dynamic Row Operations
' ------------------------------------------------------------------------------

Public Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    If RequiredCapacity <= mCapacity Then Exit Sub

    Dim newCap As Long
    newCap = mCapacity
    If newCap < DEFAULT_CAPACITY Then newCap = DEFAULT_CAPACITY

    Do While newCap < RequiredCapacity
        If newCap < 10000 Then
            newCap = CLng(newCap * 2#)
        ElseIf newCap < 100000 Then
            newCap = CLng(newCap * 1.5)
        Else
            newCap = CLng(newCap * 1.25)
        End If
    Loop

    ReDim Preserve mData(1 To newCap)
    mCapacity = newCap
End Sub

Public Sub ShrinkToFit()
    If mCapacity > mCount Then
        mCapacity = mCount
        If mCapacity < DEFAULT_CAPACITY Then mCapacity = DEFAULT_CAPACITY
        ReDim Preserve mData(1 To mCapacity)
    End If
End Sub

Public Sub SetCount(ByVal NewCount As Long, Optional ByVal DefaultValue As Variant = Empty)
    If NewCount < 0 Then
        Err.Raise 5, "cColumn.SetCount", "NewCount cannot be negative: " & NewCount
    End If

    If NewCount > mCapacity Then
        EnsureCapacity NewCount
    End If

    Dim i As Long
    If NewCount > mCount Then
        If Not IsEmpty(DefaultValue) Then
            For i = mCount + 1 To NewCount
                mData(i) = DefaultValue
            Next i
        End If
    ElseIf NewCount < mCount Then
        ' Release objects / clean memory for slots beyond NewCount
        For i = NewCount + 1 To mCount
            mData(i) = Empty
        Next i
    End If

    mCount = NewCount
End Sub

Public Sub Append(ByVal Value As Variant)
    EnsureCapacity mCount + 1
    mCount = mCount + 1

    If IsObject(Value) Then
        Set mData(mCount) = Value
    Else
        mData(mCount) = Value
    End If
End Sub

' Shifts elements [FromIndex .. mCount] down to [FromIndex+1 .. mCount+1] to open a hole
Public Sub ShiftDown(ByVal FromIndex As Long)
    If FromIndex < 1 Or FromIndex > mCount + 1 Then
        Err.Raise 9, "cColumn.ShiftDown", "Index out of range for ShiftDown: " & FromIndex & " (Valid range: 1 To " & (mCount + 1) & ")"
    End If

    EnsureCapacity mCount + 1

    Dim i As Long
    For i = mCount To FromIndex Step -1
        If IsObject(mData(i)) Then
            Set mData(i + 1) = mData(i)
        Else
            mData(i + 1) = mData(i)
        End If
    Next i

    mData(FromIndex) = Empty
    mCount = mCount + 1
End Sub

' Shifts elements [FromIndex+1 .. mCount] up to [FromIndex .. mCount-1] to remove element
Public Sub ShiftUp(ByVal FromIndex As Long)
    If FromIndex < 1 Or FromIndex > mCount Then
        Err.Raise 9, "cColumn.ShiftUp", "Index out of range for ShiftUp: " & FromIndex & " (Valid range: 1 To " & mCount & ")"
    End If

    Dim i As Long
    For i = FromIndex To mCount - 1
        If IsObject(mData(i + 1)) Then
            Set mData(i) = mData(i + 1)
        Else
            mData(i) = mData(i + 1)
        End If
    Next i

    mData(mCount) = Empty
    mCount = mCount - 1
End Sub

Public Sub Fill(ByVal Value As Variant, Optional ByVal FromRow As Long = 1, Optional ByVal ToRow As Long = 0)
    If mCount = 0 Then Exit Sub
    If ToRow <= 0 Or ToRow > mCount Then ToRow = mCount
    If FromRow < 1 Then FromRow = 1
    If FromRow > ToRow Then Exit Sub

    Dim i As Long
    For i = FromRow To ToRow
        If IsObject(Value) Then
            Set mData(i) = Value
        Else
            mData(i) = Value
        End If
    Next i
End Sub

' ------------------------------------------------------------------------------
' Buffer Permutation & Search
' ------------------------------------------------------------------------------

Public Sub Permute(ByRef indices() As Long)
    If mCount <= 1 Then Exit Sub

    Dim temp() As Variant
    ReDim temp(1 To mCount)

    Dim i As Long
    For i = 1 To mCount
        If IsObject(mData(indices(i))) Then
            Set temp(i) = mData(indices(i))
        Else
            temp(i) = mData(indices(i))
        End If
    Next i

    For i = 1 To mCount
        If IsObject(temp(i)) Then
            Set mData(i) = temp(i)
        Else
            mData(i) = temp(i)
        End If
    Next i

    Erase temp
End Sub

Public Function Find(ByVal Value As Variant, Optional ByVal StartIndex As Long = 1) As Long
    If StartIndex < 1 Then StartIndex = 1
    If StartIndex > mCount Then
        Find = 0
        Exit Function
    End If

    Dim i As Long
    For i = StartIndex To mCount
        If IsError(mData(i)) Or IsError(Value) Then
            If IsError(mData(i)) And IsError(Value) Then
                If CInt(mData(i)) = CInt(Value) Then
                    Find = i
                    Exit Function
                End If
            End If
        ElseIf IsEmpty(Value) Or IsNull(Value) Then
            If IsEmpty(mData(i)) Or IsNull(mData(i)) Then
                Find = i
                Exit Function
            End If
        ElseIf IsNumeric(Value) And IsNumeric(mData(i)) Then
            If CDbl(mData(i)) = CDbl(Value) Then
                Find = i
                Exit Function
            End If
        Else
            If StrComp(CStr(mData(i)), CStr(Value), vbTextCompare) = 0 Then
                Find = i
                Exit Function
            End If
        End If
    Next i

    Find = 0
End Function

' ------------------------------------------------------------------------------
' Array Conversions
' ------------------------------------------------------------------------------

Public Function ToArray() As Variant
    If mCount = 0 Then
        ToArray = Array()
        Exit Function
    End If

    Dim res() As Variant
    ReDim res(1 To mCount)

    Dim i As Long
    For i = 1 To mCount
        If IsObject(mData(i)) Then
            Set res(i) = mData(i)
        Else
            res(i) = mData(i)
        End If
    Next i

    ToArray = res
End Function

Public Function ToDoubleArray() As Double()
    Dim res() As Double
    If mCount = 0 Then
        ToDoubleArray = res
        Exit Function
    End If
    ReDim res(1 To mCount)
    Dim i As Long
    For i = 1 To mCount
        If IsNumeric(mData(i)) And Not IsEmpty(mData(i)) Then
            res(i) = CDbl(mData(i))
        Else
            res(i) = 0#
        End If
    Next i
    ToDoubleArray = res
End Function

Public Function ToLongArray() As Long()
    Dim res() As Long
    If mCount = 0 Then
        ToLongArray = res
        Exit Function
    End If
    ReDim res(1 To mCount)
    Dim i As Long
    For i = 1 To mCount
        If IsNumeric(mData(i)) And Not IsEmpty(mData(i)) Then
            res(i) = CLng(mData(i))
        Else
            res(i) = 0&
        End If
    Next i
    ToLongArray = res
End Function

Public Function ToStringArray() As String()
    Dim res() As String
    If mCount = 0 Then
        ToStringArray = res
        Exit Function
    End If
    ReDim res(1 To mCount)
    Dim i As Long
    For i = 1 To mCount
        res(i) = CStr(mData(i))
    Next i
    ToStringArray = res
End Function

Public Sub LoadFromArray(ByVal SourceArray As Variant)
    If Not IsArray(SourceArray) Then
        Err.Raise 13, "cColumn.LoadFromArray", "SourceArray must be an array"
    End If

    Dim lb As Long, ub As Long
    lb = LBound(SourceArray, 1)
    ub = UBound(SourceArray, 1)

    Dim numItems As Long
    numItems = ub - lb + 1

    EnsureCapacity numItems
    mCount = numItems

    Dim i As Long
    For i = 1 To numItems
        If IsObject(SourceArray(lb + i - 1)) Then
            Set mData(i) = SourceArray(lb + i - 1)
        Else
            mData(i) = SourceArray(lb + i - 1)
        End If
    Next i
End Sub

' ------------------------------------------------------------------------------
' Vector Math & Analytics Helpers
' ------------------------------------------------------------------------------

Public Function Sum(Optional ByVal UseDecimal As Boolean = False) As Variant
    Dim i As Long
    If UseDecimal Then
        Dim decTotal As Variant
        decTotal = CDec(0)
        For i = 1 To mCount
            If IsNumeric(mData(i)) And Not IsEmpty(mData(i)) And Not IsNull(mData(i)) Then
                decTotal = decTotal + CDec(mData(i))
            End If
        Next i
        Sum = decTotal
    Else
        Dim total As Double
        For i = 1 To mCount
            If IsNumeric(mData(i)) And Not IsEmpty(mData(i)) And Not IsNull(mData(i)) Then
                total = total + CDbl(mData(i))
            End If
        Next i
        Sum = total
    End If
End Function

Public Function Average() As Double
    Dim total As Double
    Dim numNumeric As Long
    Dim i As Long

    For i = 1 To mCount
        If IsNumeric(mData(i)) And Not IsEmpty(mData(i)) And Not IsNull(mData(i)) Then
            total = total + CDbl(mData(i))
            numNumeric = numNumeric + 1
        End If
    Next i

    If numNumeric > 0 Then
        Average = total / numNumeric
    Else
        Average = 0#
    End If
End Function

Public Function Min() As Variant
    If mCount = 0 Then
        Min = Empty
        Exit Function
    End If

    Dim currentMin As Variant
    Dim foundFirst As Boolean
    Dim i As Long

    For i = 1 To mCount
        If Not IsEmpty(mData(i)) And Not IsNull(mData(i)) And Not IsError(mData(i)) Then
            If Not foundFirst Then
                currentMin = mData(i)
                foundFirst = True
            ElseIf mData(i) < currentMin Then
                currentMin = mData(i)
            End If
        End If
    Next i

    If foundFirst Then
        Min = currentMin
    Else
        Min = Empty
    End If
End Function

Public Function Max() As Variant
    If mCount = 0 Then
        Max = Empty
        Exit Function
    End If

    Dim currentMax As Variant
    Dim foundFirst As Boolean
    Dim i As Long

    For i = 1 To mCount
        If Not IsEmpty(mData(i)) And Not IsNull(mData(i)) And Not IsError(mData(i)) Then
            If Not foundFirst Then
                currentMax = mData(i)
                foundFirst = True
            ElseIf mData(i) > currentMax Then
                currentMax = mData(i)
            End If
        End If
    Next i

    If foundFirst Then
        Max = currentMax
    Else
        Max = Empty
    End If
End Function

' Helper to extract clean numeric array for statistical calculations
Private Function GetNumericArray(ByRef outCount As Long) As Double()
    Dim cnt As Long
    cnt = 0

    Dim i As Long
    For i = 1 To mCount
        If IsNumeric(mData(i)) And Not IsEmpty(mData(i)) And Not IsNull(mData(i)) And Not IsError(mData(i)) Then
            cnt = cnt + 1
        End If
    Next i

    Dim nums() As Double
    If cnt > 0 Then
        ReDim nums(0 To cnt - 1)
        Dim idx As Long
        idx = 0
        For i = 1 To mCount
            If IsNumeric(mData(i)) And Not IsEmpty(mData(i)) And Not IsNull(mData(i)) And Not IsError(mData(i)) Then
                nums(idx) = CDbl(mData(i))
                idx = idx + 1
            End If
        Next i
    End If

    outCount = cnt
    GetNumericArray = nums
End Function

' Fast in-place QuickSort for 1D Double array
Private Sub QuickSortDouble(ByRef arr() As Double, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long
    Dim pivot As Double, temp As Double

    i = low
    j = high
    pivot = arr((low + high) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If low < j Then QuickSortDouble arr, low, j
    If i < high Then QuickSortDouble arr, i, high
End Sub

' Calculates the variance of numeric values in the column (Sample or Population)
Public Function Variance(Optional ByVal Sample As Boolean = True) As Double
    Dim cnt As Long
    Dim nums() As Double
    nums = GetNumericArray(cnt)

    If cnt = 0 Then
        Variance = 0#
        Exit Function
    End If
    If Sample And cnt = 1 Then
        Variance = 0#
        Exit Function
    End If

    Dim meanVal As Double
    Dim total As Double
    Dim i As Long

    For i = 0 To cnt - 1
        total = total + nums(i)
    Next i
    meanVal = total / cnt

    Dim sumSq As Double
    For i = 0 To cnt - 1
        sumSq = sumSq + (nums(i) - meanVal) * (nums(i) - meanVal)
    Next i

    If Sample Then
        Variance = sumSq / (cnt - 1)
    Else
        Variance = sumSq / cnt
    End If
End Function

' Calculates the standard deviation of numeric values in the column
Public Function StdDev(Optional ByVal Sample As Boolean = True) As Double
    Dim varVal As Double
    varVal = Me.Variance(Sample)
    If varVal > 0# Then
        StdDev = Sqr(varVal)
    Else
        StdDev = 0#
    End If
End Function

' Calculates the median (50th percentile) of numeric values
Public Function Median() As Double
    Median = Me.Quantile(0.5)
End Function

' Calculates the q-th quantile (0.0 <= q <= 1.0) with linear interpolation
Public Function Quantile(ByVal q As Double) As Double
    If q < 0# Or q > 1# Then
        Err.Raise 5, "cColumn.Quantile", "q must be between 0.0 and 1.0 (got " & CStr(q) & ")"
    End If

    Dim cnt As Long
    Dim nums() As Double
    nums = GetNumericArray(cnt)

    If cnt = 0 Then
        Quantile = 0#
        Exit Function
    End If

    If cnt = 1 Then
        Quantile = nums(0)
        Exit Function
    End If

    QuickSortDouble nums, 0, cnt - 1

    Dim idx As Double
    idx = (cnt - 1) * q

    Dim k As Long
    k = Fix(idx)
    Dim d As Double
    d = idx - k

    If k >= cnt - 1 Then
        Quantile = nums(cnt - 1)
    Else
        Quantile = (1# - d) * nums(k) + d * nums(k + 1)
    End If
End Function

' Returns the mode (most frequently occurring value)
Public Function Mode() As Variant
    If mCount = 0 Then
        Mode = Empty
        Exit Function
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, v As Variant, keyStr As String
    Dim maxFreq As Long, bestVal As Variant
    maxFreq = 0

    For i = 1 To mCount
        v = mData(i)
        If Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
            keyStr = CStr(v)
            If dict.Exists(keyStr) Then
                dict(keyStr) = dict(keyStr) + 1
            Else
                dict(keyStr) = 1
            End If

            If dict(keyStr) > maxFreq Then
                maxFreq = dict(keyStr)
                bestVal = v
            End If
        End If
    Next i

    If maxFreq > 0 Then
        Mode = bestVal
    Else
        Mode = Empty
    End If
End Function

' Calculates sample or population covariance between this column and another column
Public Function Covariance(ByVal OtherCol As cColumn, Optional ByVal Sample As Boolean = True) As Double
    If OtherCol Is Nothing Then
        Err.Raise 91, "cColumn.Covariance", "OtherCol cannot be Nothing."
    End If

    Dim n As Long
    n = mCount
    If OtherCol.Count < n Then n = OtherCol.Count

    If n = 0 Then
        Covariance = 0#
        Exit Function
    End If
    If Sample And n = 1 Then
        Covariance = 0#
        Exit Function
    End If

    Dim validCount As Long
    Dim sumX As Double, sumY As Double
    Dim i As Long
    Dim vX As Variant, vY As Variant

    For i = 1 To n
        vX = mData(i)
        vY = OtherCol.Item(i)
        If IsNumeric(vX) And IsNumeric(vY) And _
           Not IsEmpty(vX) And Not IsEmpty(vY) And _
           Not IsNull(vX) And Not IsNull(vY) And _
           Not IsError(vX) And Not IsError(vY) Then
            validCount = validCount + 1
            sumX = sumX + CDbl(vX)
            sumY = sumY + CDbl(vY)
        End If
    Next i

    If validCount = 0 Or (Sample And validCount = 1) Then
        Covariance = 0#
        Exit Function
    End If

    Dim meanX As Double, meanY As Double
    meanX = sumX / validCount
    meanY = sumY / validCount

    Dim sumProdDiff As Double
    For i = 1 To n
        vX = mData(i)
        vY = OtherCol.Item(i)
        If IsNumeric(vX) And IsNumeric(vY) And _
           Not IsEmpty(vX) And Not IsEmpty(vY) And _
           Not IsNull(vX) And Not IsNull(vY) And _
           Not IsError(vX) And Not IsError(vY) Then
            sumProdDiff = sumProdDiff + (CDbl(vX) - meanX) * (CDbl(vY) - meanY)
        End If
    Next i

    If Sample Then
        Covariance = sumProdDiff / (validCount - 1)
    Else
        Covariance = sumProdDiff / validCount
    End If
End Function

' Calculates Pearson correlation coefficient (r) between this column and another column (-1.0 to 1.0)
Public Function Correlation(ByVal OtherCol As cColumn) As Double
    If OtherCol Is Nothing Then
        Err.Raise 91, "cColumn.Correlation", "OtherCol cannot be Nothing."
    End If

    Dim n As Long
    n = mCount
    If OtherCol.Count < n Then n = OtherCol.Count
    If n <= 1 Then
        Correlation = 0#
        Exit Function
    End If

    Dim validCount As Long
    Dim sumX As Double, sumY As Double
    Dim i As Long
    Dim vX As Variant, vY As Variant

    For i = 1 To n
        vX = mData(i)
        vY = OtherCol.Item(i)
        If IsNumeric(vX) And IsNumeric(vY) And _
           Not IsEmpty(vX) And Not IsEmpty(vY) And _
           Not IsNull(vX) And Not IsNull(vY) And _
           Not IsError(vX) And Not IsError(vY) Then
            validCount = validCount + 1
            sumX = sumX + CDbl(vX)
            sumY = sumY + CDbl(vY)
        End If
    Next i

    If validCount <= 1 Then
        Correlation = 0#
        Exit Function
    End If

    Dim meanX As Double, meanY As Double
    meanX = sumX / validCount
    meanY = sumY / validCount

    Dim sumDiffX2 As Double, sumDiffY2 As Double, sumDiffXY As Double
    Dim diffX As Double, diffY As Double

    For i = 1 To n
        vX = mData(i)
        vY = OtherCol.Item(i)
        If IsNumeric(vX) And IsNumeric(vY) And _
           Not IsEmpty(vX) And Not IsEmpty(vY) And _
           Not IsNull(vX) And Not IsNull(vY) And _
           Not IsError(vX) And Not IsError(vY) Then
            diffX = CDbl(vX) - meanX
            diffY = CDbl(vY) - meanY
            sumDiffXY = sumDiffXY + diffX * diffY
            sumDiffX2 = sumDiffX2 + diffX * diffX
            sumDiffY2 = sumDiffY2 + diffY * diffY
        End If
    Next i

    Dim denom As Double
    denom = Sqr(sumDiffX2 * sumDiffY2)

    If denom > 0# Then
        Dim r As Double
        r = sumDiffXY / denom
        If r > 1# Then r = 1#
        If r < -1# Then r = -1#
        Correlation = r
    Else
        Correlation = 0#
    End If
End Function

' Calculates rolling moving sum over a sliding window of size WindowSize
Public Function RollingSum(ByVal WindowSize As Long, Optional ByVal ResultName As String = "") As cColumn
    If WindowSize <= 0 Then
        Err.Raise 5, "cColumn.RollingSum", "WindowSize must be greater than 0."
    End If

    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName & "_RollingSum" & WindowSize
    End If

    If mCount = 0 Then
        Set RollingSum = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    Dim i As Long, startIdx As Long, j As Long
    Dim currentSum As Double
    Dim v As Variant, hasValid As Boolean

    For i = 1 To mCount
        If i < WindowSize Then
            res.Append CVErr(2042)
        Else
            currentSum = 0#
            hasValid = False
            startIdx = i - WindowSize + 1
            For j = startIdx To i
                v = mData(j)
                If IsNumeric(v) And Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
                    currentSum = currentSum + CDbl(v)
                    hasValid = True
                End If
            Next j
            If hasValid Then
                res.Append currentSum
            Else
                res.Append CVErr(2042)
            End If
        End If
    Next i

    Set RollingSum = res
End Function

' Calculates rolling moving average over a sliding window of size WindowSize
Public Function RollingAvg(ByVal WindowSize As Long, Optional ByVal ResultName As String = "") As cColumn
    If WindowSize <= 0 Then
        Err.Raise 5, "cColumn.RollingAvg", "WindowSize must be greater than 0."
    End If

    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName & "_RollingAvg" & WindowSize
    End If

    If mCount = 0 Then
        Set RollingAvg = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    Dim i As Long, startIdx As Long, j As Long
    Dim currentSum As Double, numValid As Long
    Dim v As Variant

    For i = 1 To mCount
        If i < WindowSize Then
            res.Append CVErr(2042)
        Else
            currentSum = 0#
            numValid = 0
            startIdx = i - WindowSize + 1
            For j = startIdx To i
                v = mData(j)
                If IsNumeric(v) And Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
                    currentSum = currentSum + CDbl(v)
                    numValid = numValid + 1
                End If
            Next j
            If numValid > 0 Then
                res.Append currentSum / numValid
            Else
                res.Append CVErr(2042)
            End If
        End If
    Next i

    Set RollingAvg = res
End Function

' Calculates rolling moving minimum over a sliding window of size WindowSize
Public Function RollingMin(ByVal WindowSize As Long, Optional ByVal ResultName As String = "") As cColumn
    If WindowSize <= 0 Then
        Err.Raise 5, "cColumn.RollingMin", "WindowSize must be greater than 0."
    End If

    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName & "_RollingMin" & WindowSize
    End If

    If mCount = 0 Then
        Set RollingMin = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    Dim i As Long, startIdx As Long, j As Long
    Dim minVal As Variant, v As Variant, found As Boolean

    For i = 1 To mCount
        If i < WindowSize Then
            res.Append CVErr(2042)
        Else
            found = False
            startIdx = i - WindowSize + 1
            For j = startIdx To i
                v = mData(j)
                If Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
                    If Not found Then
                        minVal = v
                        found = True
                    ElseIf v < minVal Then
                        minVal = v
                    End If
                End If
            Next j
            If found Then
                res.Append minVal
            Else
                res.Append CVErr(2042)
            End If
        End If
    Next i

    Set RollingMin = res
End Function

' Calculates rolling moving maximum over a sliding window of size WindowSize
Public Function RollingMax(ByVal WindowSize As Long, Optional ByVal ResultName As String = "") As cColumn
    If WindowSize <= 0 Then
        Err.Raise 5, "cColumn.RollingMax", "WindowSize must be greater than 0."
    End If

    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName & "_RollingMax" & WindowSize
    End If

    If mCount = 0 Then
        Set RollingMax = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    Dim i As Long, startIdx As Long, j As Long
    Dim maxVal As Variant, v As Variant, found As Boolean

    For i = 1 To mCount
        If i < WindowSize Then
            res.Append CVErr(2042)
        Else
            found = False
            startIdx = i - WindowSize + 1
            For j = startIdx To i
                v = mData(j)
                If Not IsEmpty(v) And Not IsNull(v) And Not IsError(v) Then
                    If Not found Then
                        maxVal = v
                        found = True
                    ElseIf v > maxVal Then
                        maxVal = v
                    End If
                End If
            Next j
            If found Then
                res.Append maxVal
            Else
                res.Append CVErr(2042)
            End If
        End If
    Next i

    Set RollingMax = res
End Function

' Replaces Empty, Null, Error (and optionally blank strings) in-place with FillValue in O(N)
Public Sub FillNA(ByVal FillValue As Variant, Optional ByVal TreatBlankStringAsNA As Boolean = False)
    If mCount = 0 Then Exit Sub

    Dim i As Long, v As Variant, isNA As Boolean
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            isNA = True
        ElseIf IsEmpty(v) Or IsNull(v) Then
            isNA = True
        ElseIf TreatBlankStringAsNA And VarType(v) = vbString Then
            isNA = (Len(Trim$(CStr(v))) = 0)
        Else
            isNA = False
        End If

        If isNA Then
            If IsObject(FillValue) Then
                Set mData(i) = FillValue
            Else
                mData(i) = FillValue
            End If
        End If
    Next i
End Sub

' Element-wise vectorized arithmetic (+, -, *, /) against a scalar, cColumn, or 1D Array
Public Function VectorOp(ByVal Operator As String, _
                         ByVal Operand As Variant, _
                         Optional ByVal ResultName As String = "") As cColumn
    Dim normOp As String
    Select Case UCase$(Trim$(Operator))
        Case "+", "ADD"
            normOp = "+"
        Case "-", "SUB", "SUBTRACT"
            normOp = "-"
        Case "*", "MUL", "MULTIPLY"
            normOp = "*"
        Case "/", "DIV", "DIVIDE"
            normOp = "/"
        Case Else
            Err.Raise 5, "cColumn.VectorOp", _
                "Unsupported vector operator: '" & Operator & "'. Supported: +, -, *, /"
    End Select

    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName
    End If

    If mCount = 0 Then
        Set VectorOp = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long
    If IsObject(Operand) Then
        If Operand Is Nothing Then
            Err.Raise 91, "cColumn.VectorOp", "Operand object cannot be Nothing."
        End If
        If TypeName(Operand) <> "cColumn" Then
            Err.Raise 13, "cColumn.VectorOp", "Object operand must be a cColumn instance."
        End If

        Dim rhsCol As cColumn
        Set rhsCol = Operand
        If rhsCol.Count <> mCount Then
            Err.Raise 5, "cColumn.VectorOp", _
                "Operand cColumn count (" & rhsCol.Count & ") does not match column count (" & mCount & ")."
        End If

        For i = 1 To mCount
            res.Item(i) = EvalPairArithmetic(mData(i), rhsCol.Item(i), normOp)
        Next i

    ElseIf IsArray(Operand) Then
        Dim lb As Long, ub As Long, arrLen As Long
        lb = LBound(Operand, 1)
        ub = UBound(Operand, 1)
        arrLen = ub - lb + 1
        If arrLen <> mCount Then
            Err.Raise 5, "cColumn.VectorOp", _
                "Operand array length (" & arrLen & ") does not match column count (" & mCount & ")."
        End If

        For i = 1 To mCount
            res.Item(i) = EvalPairArithmetic(mData(i), Operand(lb + i - 1), normOp)
        Next i

    Else
        For i = 1 To mCount
            res.Item(i) = EvalPairArithmetic(mData(i), Operand, normOp)
        Next i
    End If

    Set VectorOp = res
End Function

Public Function Add(ByVal Operand As Variant, Optional ByVal ResultName As String = "") As cColumn
    Set Add = VectorOp("+", Operand, ResultName)
End Function

Public Function Subtract(ByVal Operand As Variant, Optional ByVal ResultName As String = "") As cColumn
    Set Subtract = VectorOp("-", Operand, ResultName)
End Function

Public Function Multiply(ByVal Operand As Variant, Optional ByVal ResultName As String = "") As cColumn
    Set Multiply = VectorOp("*", Operand, ResultName)
End Function

Public Function Divide(ByVal Operand As Variant, Optional ByVal ResultName As String = "") As cColumn
    Set Divide = VectorOp("/", Operand, ResultName)
End Function

Private Function EvalPairArithmetic(ByVal vA As Variant, ByVal vB As Variant, ByVal normOp As String) As Variant
    If IsError(vA) Then
        EvalPairArithmetic = vA
        Exit Function
    End If
    If IsError(vB) Then
        EvalPairArithmetic = vB
        Exit Function
    End If
    If IsEmpty(vA) Or IsNull(vA) Or IsEmpty(vB) Or IsNull(vB) Then
        EvalPairArithmetic = Empty
        Exit Function
    End If
    If Not IsNumeric(vA) Or Not IsNumeric(vB) Then
        EvalPairArithmetic = CVErr(2015) ' xlErrValue (#VALUE!)
        Exit Function
    End If

    Dim dA As Double, dB As Double
    dA = CDbl(vA)
    dB = CDbl(vB)

    Select Case normOp
        Case "+"
            EvalPairArithmetic = dA + dB
        Case "-"
            EvalPairArithmetic = dA - dB
        Case "*"
            EvalPairArithmetic = dA * dB
        Case "/"
            If dB = 0# Then
                EvalPairArithmetic = CVErr(2007) ' xlErrDiv0 (#DIV/0!)
            Else
                EvalPairArithmetic = dA / dB
            End If
    End Select
End Function

' Returns a new cColumn with distinct values in order of first appearance (O(N) via Dictionary)
Public Function Unique() As cColumn
    Dim res As New cColumn
    res.Name = mName
    If mCount = 0 Then
        Set Unique = res
        Exit Function
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1  ' Case-insensitive string keys

    Dim seenEmpty As Boolean
    seenEmpty = False

    Dim i As Long, v As Variant, k As String
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            k = "#E:" & CStr(CInt(v))
            If Not dict.Exists(k) Then
                dict.Add k, True
                res.Append v
            End If
        ElseIf IsEmpty(v) Or IsNull(v) Then
            If Not seenEmpty Then
                seenEmpty = True
                res.Append Empty
            End If
        Else
            k = "#V:" & CStr(v)
            If Not dict.Exists(k) Then
                dict.Add k, True
                res.Append v
            End If
        End If
    Next i

    Set Unique = res
End Function

' ------------------------------------------------------------------------------
' Window & Cumulative Functions (CumSum, Shift, PctChange, Rank)
' ------------------------------------------------------------------------------

' Running cumulative sum over 1 To mCount (skips Empty/Null/non-numeric while carrying running total)
Public Function CumSum(Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName
    End If

    If mCount = 0 Then
        Set CumSum = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim runningTotal As Double
    runningTotal = 0#

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            res.Item(i) = v
        Else
            If Not IsEmpty(v) And Not IsNull(v) Then
                If IsNumeric(v) Then
                    runningTotal = runningTotal + CDbl(v)
                End If
            End If
            res.Item(i) = runningTotal
        End If
    Next i

    Set CumSum = res
End Function

' Shifts (lags when Periods > 0, leads when Periods < 0) column elements by Periods
Public Function Shift(Optional ByVal Periods As Long = 1, _
                      Optional ByVal FillValue As Variant = Empty, _
                      Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName
    End If

    If mCount = 0 Then
        Set Shift = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, srcIdx As Long
    For i = 1 To mCount
        srcIdx = i - Periods
        If srcIdx >= 1 And srcIdx <= mCount Then
            If IsObject(mData(srcIdx)) Then
                Set res.Item(i) = mData(srcIdx)
            Else
                res.Item(i) = mData(srcIdx)
            End If
        Else
            If IsObject(FillValue) Then
                Set res.Item(i) = FillValue
            Else
                res.Item(i) = FillValue
            End If
        End If
    Next i

    Set Shift = res
End Function

' Fluent alias for Shift(Periods, FillValue): Retrieves value from Periods rows prior
Public Function Lag(Optional ByVal Periods As Long = 1, _
                    Optional ByVal FillValue As Variant = Empty, _
                    Optional ByVal ResultName As String = "") As cColumn
    Set Lag = Shift(Periods, FillValue, ResultName)
End Function

' Fluent alias for Shift(-Periods, FillValue): Retrieves value from Periods rows ahead
Public Function Lead(Optional ByVal Periods As Long = 1, _
                     Optional ByVal FillValue As Variant = Empty, _
                     Optional ByVal ResultName As String = "") As cColumn
    Set Lead = Shift(-Periods, FillValue, ResultName)
End Function

' Computes the difference between current element and previous element: v(i) - v(i - Periods)
Public Function Diff(Optional ByVal Periods As Long = 1, _
                     Optional ByVal ResultName As String = "") As cColumn
    If Periods <= 0 Then
        Err.Raise 5, "cColumn.Diff", "Periods must be greater than 0."
    End If

    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName & "_Diff"
    End If

    If mCount = 0 Then
        Set Diff = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, srcIdx As Long, vCur As Variant, vPrev As Variant
    For i = 1 To mCount
        srcIdx = i - Periods
        If srcIdx >= 1 Then
            vCur = mData(i)
            vPrev = mData(srcIdx)
            If IsNumeric(vCur) And IsNumeric(vPrev) And Not IsEmpty(vCur) And Not IsEmpty(vPrev) Then
                res.Item(i) = CDbl(vCur) - CDbl(vPrev)
            Else
                res.Item(i) = Empty
            End If
        Else
            res.Item(i) = Empty
        End If
    Next i

    Set Diff = res
End Function

' Calculates fractional percentage change (v_cur - v_prev) / v_prev over Periods
Public Function PctChange(Optional ByVal Periods As Long = 1, _
                          Optional ByVal FillValue As Variant = Empty, _
                          Optional ByVal ResultName As String = "") As cColumn
    If Periods = 0 Then
        Err.Raise 5, "cColumn.PctChange", "Periods cannot be 0."
    End If

    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName
    End If

    If mCount = 0 Then
        Set PctChange = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, prevIdx As Long
    Dim vCur As Variant, vPrev As Variant
    Dim dCur As Double, dPrev As Double

    For i = 1 To mCount
        prevIdx = i - Periods
        If prevIdx < 1 Or prevIdx > mCount Then
            res.Item(i) = FillValue
        Else
            vCur = mData(i)
            vPrev = mData(prevIdx)

            If IsError(vCur) Then
                res.Item(i) = vCur
            ElseIf IsError(vPrev) Then
                res.Item(i) = vPrev
            ElseIf IsEmpty(vCur) Or IsNull(vCur) Or IsEmpty(vPrev) Or IsNull(vPrev) Then
                res.Item(i) = FillValue
            ElseIf Not IsNumeric(vCur) Or Not IsNumeric(vPrev) Then
                res.Item(i) = FillValue
            Else
                dCur = CDbl(vCur)
                dPrev = CDbl(vPrev)
                If dPrev = 0# Then
                    res.Item(i) = CVErr(2007) ' #DIV/0!
                Else
                    res.Item(i) = (dCur - dPrev) / dPrev
                End If
            End If
        End If
    Next i

    Set PctChange = res
End Function

' Computes rank of values in O(N log N) via indirect QuickSort.
' Method: "AVERAGE" (default), "MIN", or "DENSE". Empty/Null/Error slots remain Empty.
Public Function Rank(Optional ByVal Ascending As Boolean = True, _
                     Optional ByVal Method As String = "AVERAGE", _
                     Optional ByVal ResultName As String = "") As cColumn
    Dim rankMethod As String
    rankMethod = UCase$(Trim$(Method))
    If rankMethod = "AVG" Then rankMethod = "AVERAGE"

    If rankMethod <> "AVERAGE" And rankMethod <> "MIN" And rankMethod <> "DENSE" Then
        Err.Raise 5, "cColumn.Rank", _
            "Unsupported Rank method: '" & Method & "'. Supported: AVERAGE, MIN, DENSE"
    End If

    Dim res As New cColumn
    If Len(Trim$(ResultName)) > 0 Then
        res.Name = Trim$(ResultName)
    Else
        res.Name = mName
    End If

    If mCount = 0 Then
        Set Rank = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim validIdx() As Long
    ReDim validIdx(1 To mCount)
    Dim numValid As Long
    numValid = 0

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = Empty
        Else
            numValid = numValid + 1
            validIdx(numValid) = i
        End If
    Next i

    If numValid = 0 Then
        Set Rank = res
        Exit Function
    End If

    If numValid > 1 Then
        QuickSortRankIndices validIdx, 1, numValid, Ascending
    End If

    Dim tieStart As Long, tieEnd As Long, t As Long
    Dim denseGroup As Long
    Dim rankVal As Double

    tieStart = 1
    denseGroup = 0

    Do While tieStart <= numValid
        tieEnd = tieStart
        Do While tieEnd < numValid
            If CompareRankValues(mData(validIdx(tieStart)), mData(validIdx(tieEnd + 1)), Ascending) = 0 Then
                tieEnd = tieEnd + 1
            Else
                Exit Do
            End If
        Loop

        denseGroup = denseGroup + 1

        If rankMethod = "AVERAGE" Then
            rankVal = (CDbl(tieStart) + CDbl(tieEnd)) / 2#
        ElseIf rankMethod = "MIN" Then
            rankVal = CDbl(tieStart)
        Else
            rankVal = CDbl(denseGroup)
        End If

        For t = tieStart To tieEnd
            res.Item(validIdx(t)) = rankVal
        Next t

        tieStart = tieEnd + 1
    Loop

    Set Rank = res
End Function

Private Sub QuickSortRankIndices(ByRef idx() As Long, ByVal low As Long, ByVal high As Long, ByVal Ascending As Boolean)
    Dim i As Long, j As Long
    Dim pivotVal As Variant
    i = low
    j = high
    pivotVal = mData(idx((low + high) \ 2))

    Do While i <= j
        Do While CompareRankValues(mData(idx(i)), pivotVal, Ascending) < 0
            i = i + 1
        Loop
        Do While CompareRankValues(mData(idx(j)), pivotVal, Ascending) > 0
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

    If low < j Then QuickSortRankIndices idx, low, j, Ascending
    If i < high Then QuickSortRankIndices idx, i, high, Ascending
End Sub

Private Function CompareRankValues(ByVal vA As Variant, ByVal vB As Variant, ByVal Ascending As Boolean) As Long
    Dim cmp As Long
    If IsNumeric(vA) And IsNumeric(vB) Then
        Dim dA As Double, dB As Double
        dA = CDbl(vA)
        dB = CDbl(vB)
        If dA < dB Then
            cmp = -1
        ElseIf dA > dB Then
            cmp = 1
        Else
            cmp = 0
        End If
    Else
        Dim sc As Long
        sc = StrComp(CStr(vA), CStr(vB), vbTextCompare)
        If sc < 0 Then
            cmp = -1
        ElseIf sc > 0 Then
            cmp = 1
        Else
            cmp = 0
        End If
    End If

    If Not Ascending Then cmp = -cmp
    CompareRankValues = cmp
End Function

' ------------------------------------------------------------------------------
' Vectorized String & Numeric Clamping Helpers (StrTrim, StrUpper, StrLower, StrReplace, Clip)
' ------------------------------------------------------------------------------

Public Function StrTrim(Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set StrTrim = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = v
        Else
            res.Item(i) = Trim$(CStr(v))
        End If
    Next i

    Set StrTrim = res
End Function

Public Function StrUpper(Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set StrUpper = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = v
        Else
            res.Item(i) = UCase$(CStr(v))
        End If
    Next i

    Set StrUpper = res
End Function

Public Function StrLower(Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set StrLower = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = v
        Else
            res.Item(i) = LCase$(CStr(v))
        End If
    Next i

    Set StrLower = res
End Function

Public Function StrReplace(ByVal FindStr As String, _
                           ByVal ReplaceWith As String, _
                           Optional ByVal ResultName As String = "", _
                           Optional ByVal IgnoreCase As Boolean = False) As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set StrReplace = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim cmpMode As VbCompareMethod
    If IgnoreCase Then
        cmpMode = vbTextCompare
    Else
        cmpMode = vbBinaryCompare
    End If

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = v
        Else
            res.Item(i) = Replace(CStr(v), FindStr, ReplaceWith, 1, -1, cmpMode)
        End If
    Next i

    Set StrReplace = res
End Function

Public Function Clip(ByVal Lower As Variant, _
                     ByVal Upper As Variant, _
                     Optional ByVal ResultName As String = "") As cColumn
    Dim hasLower As Boolean, hasUpper As Boolean
    Dim dLow As Double, dHigh As Double

    hasLower = Not (IsEmpty(Lower) Or IsNull(Lower))
    hasUpper = Not (IsEmpty(Upper) Or IsNull(Upper))

    If hasLower Then
        If Not IsNumeric(Lower) Then
            Err.Raise 5, "cColumn.Clip", "Lower bound must be numeric."
        End If
        dLow = CDbl(Lower)
    End If

    If hasUpper Then
        If Not IsNumeric(Upper) Then
            Err.Raise 5, "cColumn.Clip", "Upper bound must be numeric."
        End If
        dHigh = CDbl(Upper)
    End If

    If hasLower And hasUpper Then
        If dLow > dHigh Then
            Err.Raise 5, "cColumn.Clip", "Lower bound (" & dLow & ") cannot exceed Upper bound (" & dHigh & ")."
        End If
    End If

    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set Clip = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant, dVal As Double
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = v
        ElseIf IsNumeric(v) Then
            dVal = CDbl(v)
            If hasLower And dVal < dLow Then dVal = dLow
            If hasUpper And dVal > dHigh Then dVal = dHigh
            res.Item(i) = dVal
        Else
            res.Item(i) = v
        End If
    Next i

    Set Clip = res
End Function

' ------------------------------------------------------------------------------
' Vectorized Date & Time Operations (Year, Month, Day, Quarter, DayOfWeek, DateAdd, DateDiff, FormatDate)
' ------------------------------------------------------------------------------

Public Function Year(Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set Year = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            res.Item(i) = v
        ElseIf IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = Empty
        ElseIf IsDate(v) Then
            res.Item(i) = CLng(VBA.Year(CDate(v)))
        Else
            res.Item(i) = Empty
        End If
    Next i

    Set Year = res
End Function

Public Function Month(Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set Month = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            res.Item(i) = v
        ElseIf IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = Empty
        ElseIf IsDate(v) Then
            res.Item(i) = CLng(VBA.Month(CDate(v)))
        Else
            res.Item(i) = Empty
        End If
    Next i

    Set Month = res
End Function

Public Function Day(Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set Day = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            res.Item(i) = v
        ElseIf IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = Empty
        ElseIf IsDate(v) Then
            res.Item(i) = CLng(VBA.Day(CDate(v)))
        Else
            res.Item(i) = Empty
        End If
    Next i

    Set Day = res
End Function

Public Function Quarter(Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set Quarter = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            res.Item(i) = v
        ElseIf IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = Empty
        ElseIf IsDate(v) Then
            res.Item(i) = CLng((VBA.Month(CDate(v)) - 1) \ 3 + 1)
        Else
            res.Item(i) = Empty
        End If
    Next i

    Set Quarter = res
End Function

Public Function DayOfWeek(Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
                          Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set DayOfWeek = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            res.Item(i) = v
        ElseIf IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = Empty
        ElseIf IsDate(v) Then
            res.Item(i) = CLng(VBA.Weekday(CDate(v), FirstDayOfWeek))
        Else
            res.Item(i) = Empty
        End If
    Next i

    Set DayOfWeek = res
End Function

Public Function DateAdd(ByVal Interval As String, _
                        ByVal Number As Double, _
                        Optional ByVal ResultName As String = "") As cColumn
    Dim cleanInt As String
    cleanInt = Trim$(Interval)
    If Len(cleanInt) = 0 Then
        Err.Raise 5, "cColumn.DateAdd", "Interval cannot be blank."
    End If

    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set DateAdd = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            res.Item(i) = v
        ElseIf IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = Empty
        ElseIf IsDate(v) Then
            res.Item(i) = VBA.DateAdd(cleanInt, Number, CDate(v))
        Else
            res.Item(i) = Empty
        End If
    Next i

    Set DateAdd = res
End Function

Public Function DateDiff(ByVal Interval As String, _
                         ByVal OtherDateOrCol As Variant, _
                         Optional ByVal ResultName As String = "") As cColumn
    Dim cleanInt As String
    cleanInt = Trim$(Interval)
    If Len(cleanInt) = 0 Then
        Err.Raise 5, "cColumn.DateDiff", "Interval cannot be blank."
    End If

    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set DateDiff = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, vA As Variant, vB As Variant

    If IsObject(OtherDateOrCol) Then
        If OtherDateOrCol Is Nothing Then
            Err.Raise 91, "cColumn.DateDiff", "OtherDateOrCol object cannot be Nothing."
        End If
        If TypeName(OtherDateOrCol) <> "cColumn" Then
            Err.Raise 13, "cColumn.DateDiff", "Object operand must be a cColumn instance."
        End If

        Dim rhsCol As cColumn
        Set rhsCol = OtherDateOrCol
        If rhsCol.Count <> mCount Then
            Err.Raise 5, "cColumn.DateDiff", _
                "OtherDateOrCol cColumn count (" & rhsCol.Count & ") does not match column count (" & mCount & ")."
        End If

        For i = 1 To mCount
            vA = mData(i)
            vB = rhsCol.Item(i)
            If IsError(vA) Then
                res.Item(i) = vA
            ElseIf IsError(vB) Then
                res.Item(i) = vB
            ElseIf IsEmpty(vA) Or IsNull(vA) Or IsEmpty(vB) Or IsNull(vB) Then
                res.Item(i) = Empty
            ElseIf IsDate(vA) And IsDate(vB) Then
                res.Item(i) = CLng(VBA.DateDiff(cleanInt, CDate(vA), CDate(vB)))
            Else
                res.Item(i) = Empty
            End If
        Next i
    Else
        If Not IsDate(OtherDateOrCol) Then
            Err.Raise 5, "cColumn.DateDiff", "Scalar OtherDateOrCol must be a valid date."
        End If
        Dim dtTarget As Date
        dtTarget = CDate(OtherDateOrCol)

        For i = 1 To mCount
            vA = mData(i)
            If IsError(vA) Then
                res.Item(i) = vA
            ElseIf IsEmpty(vA) Or IsNull(vA) Then
                res.Item(i) = Empty
            ElseIf IsDate(vA) Then
                res.Item(i) = CLng(VBA.DateDiff(cleanInt, CDate(vA), dtTarget))
            Else
                res.Item(i) = Empty
            End If
        Next i
    End If

    Set DateDiff = res
End Function

Public Function FormatDate(ByVal FormatString As String, _
                           Optional ByVal ResultName As String = "") As cColumn
    Dim res As New cColumn
    res.Name = IIf(Len(Trim$(ResultName)) > 0, Trim$(ResultName), mName)
    If mCount = 0 Then
        Set FormatDate = res
        Exit Function
    End If

    res.EnsureCapacity mCount
    res.SetCount mCount

    Dim i As Long, v As Variant
    For i = 1 To mCount
        v = mData(i)
        If IsError(v) Then
            res.Item(i) = v
        ElseIf IsEmpty(v) Or IsNull(v) Then
            res.Item(i) = Empty
        ElseIf IsDate(v) Then
            res.Item(i) = Format$(CDate(v), FormatString)
        Else
            res.Item(i) = Empty
        End If
    Next i

    Set FormatDate = res
End Function

' ------------------------------------------------------------------------------
' Cloning
' ------------------------------------------------------------------------------

Public Function Clone() As cColumn
    Dim copyCol As New cColumn
    copyCol.Name = mName
    copyCol.EnsureCapacity mCapacity
    copyCol.SetCount mCount

    Dim i As Long
    For i = 1 To mCount
        If IsObject(mData(i)) Then
            Set copyCol.Item(i) = mData(i)
        Else
            copyCol.Item(i) = mData(i)
        End If
    Next i

    Set Clone = copyCol
End Function

