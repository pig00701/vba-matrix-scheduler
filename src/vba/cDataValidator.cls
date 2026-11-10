VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDataValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' cDataValidator: Schema Contract & Data Quality Assurance Engine
' ==============================================================================
' Provides fluent schema validation and data hygiene checks for cDataFrame.
' Detects schema mismatches, null values, type incompatibilities, range
' violations, uniqueness breaches, and custom list membership in O(N).
' Can generate an error audit report or extract clean / invalid rows.
' ==============================================================================

Private mRules As Object            ' Dictionary: ColumnName -> Rule Object (Dictionary)
Private mErrorsDf As cDataFrame     ' Stores validation error records
Private mTotalRowsChecked As Long
Private mFailedRowsCount As Long
Private mIsValid As Boolean

Private Sub Class_Initialize()
    Set mRules = CreateObject("Scripting.Dictionary")
    mRules.CompareMode = 1 ' Case-insensitive column matching
    ResetState
End Sub

Private Sub Class_Terminate()
    Set mRules = Nothing
    Set mErrorsDf = Nothing
End Sub

Private Sub ResetState()
    mTotalRowsChecked = 0
    mFailedRowsCount = 0
    mIsValid = True
    Set mErrorsDf = New cDataFrame
    mErrorsDf.AddColumn "RowIndex"
    mErrorsDf.AddColumn "ColumnName"
    mErrorsDf.AddColumn "Value"
    mErrorsDf.AddColumn "RuleViolated"
    mErrorsDf.AddColumn "ErrorMessage"
End Sub

' ------------------------------------------------------------------------------
' Fluent Rule Definition API
' ------------------------------------------------------------------------------

' Generic rule definition
Public Function AddRule(ByVal ColumnName As String, _
                        Optional ByVal ExpectedType As VbVarType = vbEmpty, _
                        Optional ByVal AllowNull As Boolean = True, _
                        Optional ByVal IsUnique As Boolean = False, _
                        Optional ByVal MinVal As Variant, _
                        Optional ByVal MaxVal As Variant, _
                        Optional ByVal AllowedValues As Variant, _
                        Optional ByVal LikePattern As String = "") As cDataValidator
    Dim colKey As String
    colKey = Trim$(ColumnName)
    If Len(colKey) = 0 Then
        Err.Raise 5, "cDataValidator.AddRule", "ColumnName cannot be blank."
    End If

    Dim rule As Object
    Set rule = CreateObject("Scripting.Dictionary")
    rule.CompareMode = 1

    rule.Item("ExpectedType") = CLng(ExpectedType)
    rule.Item("AllowNull") = AllowNull
    rule.Item("IsUnique") = IsUnique
    rule.Item("LikePattern") = LikePattern

    If Not IsMissing(MinVal) And Not IsEmpty(MinVal) Then
        rule.Item("HasMin") = True
        rule.Item("MinVal") = MinVal
    Else
        rule.Item("HasMin") = False
    End If

    If Not IsMissing(MaxVal) And Not IsEmpty(MaxVal) Then
        rule.Item("HasMax") = True
        rule.Item("MaxVal") = MaxVal
    Else
        rule.Item("HasMax") = False
    End If

    If Not IsMissing(AllowedValues) And Not IsEmpty(AllowedValues) Then
        Dim allowedDict As Object
        Set allowedDict = CreateObject("Scripting.Dictionary")
        allowedDict.CompareMode = 1
        Dim item As Variant
        If IsArray(AllowedValues) Then
            For Each item In AllowedValues
                allowedDict.Item(CStr(item)) = True
            Next item
        Else
            allowedDict.Item(CStr(AllowedValues)) = True
        End If
        Set rule.Item("AllowedDict") = allowedDict
    End If

    Set mRules.Item(colKey) = rule
    Set AddRule = Me
End Function

' Convenience method: Require column existence and non-null values
Public Function Require(ByVal ColumnName As String, _
                        Optional ByVal ExpectedType As VbVarType = vbEmpty, _
                        Optional ByVal IsUnique As Boolean = False) As cDataValidator
    Set Require = AddRule(ColumnName, ExpectedType:=ExpectedType, AllowNull:=False, IsUnique:=IsUnique)
End Function

' Convenience method: Check numeric range [MinVal, MaxVal]
Public Function AddNumericRange(ByVal ColumnName As String, _
                                ByVal MinVal As Double, _
                                ByVal MaxVal As Double) As cDataValidator
    Set AddNumericRange = AddRule(ColumnName, ExpectedType:=vbDouble, MinVal:=MinVal, MaxVal:=MaxVal)
End Function

' Convenience method: Check against allowed whitelist of values
Public Function AddAllowedList(ByVal ColumnName As String, _
                               ByVal AllowedValues As Variant) As cDataValidator
    Set AddAllowedList = AddRule(ColumnName, AllowedValues:=AllowedValues)
End Function

' ------------------------------------------------------------------------------
' Validation Execution Engine
' ------------------------------------------------------------------------------

Public Function Validate(ByVal TargetDf As cDataFrame) As Boolean
    If TargetDf Is Nothing Then
        Err.Raise 91, "cDataValidator.Validate", "TargetDf cannot be Nothing."
    End If

    ResetState
    mTotalRowsChecked = TargetDf.RowCount

    If mRules.Count = 0 Then
        mIsValid = True
        Validate = True
        Exit Function
    End If

    Dim failedRowTracker As Object
    Set failedRowTracker = CreateObject("Scripting.Dictionary")

    Dim ruleColNames As Variant
    ruleColNames = mRules.Keys

    Dim i As Long, colName As String, rule As Object
    Dim colObj As cColumn
    Dim expectedType As Long, allowNull As Boolean, isUnique As Boolean
    Dim hasMin As Boolean, minV As Variant, hasMax As Boolean, maxV As Variant
    Dim likePattern As String, hasAllowed As Boolean, allowedDict As Object

    For i = 0 To mRules.Count - 1
        colName = CStr(ruleColNames(i))
        Set rule = mRules.Item(colName)

        ' Check 1: Column existence
        If Not TargetDf.HasColumn(colName) Then
            mIsValid = False
            AddError 0, colName, Empty, "MISSING_COLUMN", _
                     "Required column '" & colName & "' does not exist in DataFrame."
        Else
            Set colObj = TargetDf.Column(colName)
            expectedType = CLng(rule.Item("ExpectedType"))
            allowNull = CBool(rule.Item("AllowNull"))
            isUnique = CBool(rule.Item("IsUnique"))
            hasMin = CBool(rule.Item("HasMin"))
            If hasMin Then minV = rule.Item("MinVal")
            hasMax = CBool(rule.Item("HasMax"))
            If hasMax Then maxV = rule.Item("MaxVal")
            likePattern = CStr(rule.Item("LikePattern"))

            hasAllowed = rule.Exists("AllowedDict")
            If hasAllowed Then Set allowedDict = rule.Item("AllowedDict")

            Dim seenValues As Object
            If isUnique Then
                Set seenValues = CreateObject("Scripting.Dictionary")
                seenValues.CompareMode = IIf(TargetDf.CaseSensitive, 0, 1)
            End If

            Dim r As Long, v As Variant, isCellNull As Boolean
            For r = 1 To mTotalRowsChecked
                v = colObj.Item(r)
                isCellNull = IsNull(v) Or IsEmpty(v)
                If Not isCellNull And VarType(v) = vbString Then
                    If Len(Trim$(CStr(v))) = 0 Then isCellNull = True
                End If

                ' Check 2: Nullability
                If isCellNull Then
                    If Not allowNull Then
                        mIsValid = False
                        failedRowTracker.Item(r) = True
                        AddError r, colName, v, "NOT_NULL", "Value is empty/null but column is non-nullable."
                    End If
                Else
                    ' Check 3: Data Type
                    If expectedType <> vbEmpty Then
                        If Not CheckTypeMatch(v, expectedType) Then
                            mIsValid = False
                            failedRowTracker.Item(r) = True
                            AddError r, colName, v, "TYPE_MISMATCH", _
                                     "Expected type " & TypeNameFromVbType(expectedType) & " but got " & TypeName(v) & "."
                        End If
                    End If

                    ' Check 4: Numeric / Date Range
                    If hasMin Then
                        If IsNumeric(v) And IsNumeric(minV) Then
                            If CDbl(v) < CDbl(minV) Then
                                mIsValid = False
                                failedRowTracker.Item(r) = True
                                AddError r, colName, v, "OUT_OF_RANGE", "Value " & CStr(v) & " is less than minimum " & CStr(minV) & "."
                            End If
                        ElseIf IsDate(v) And IsDate(minV) Then
                            If CDate(v) < CDate(minV) Then
                                mIsValid = False
                                failedRowTracker.Item(r) = True
                                AddError r, colName, v, "OUT_OF_RANGE", "Date " & CStr(v) & " is earlier than minimum " & CStr(minV) & "."
                            End If
                        End If
                    End If

                    If hasMax Then
                        If IsNumeric(v) And IsNumeric(maxV) Then
                            If CDbl(v) > CDbl(maxV) Then
                                mIsValid = False
                                failedRowTracker.Item(r) = True
                                AddError r, colName, v, "OUT_OF_RANGE", "Value " & CStr(v) & " is greater than maximum " & CStr(maxV) & "."
                            End If
                        ElseIf IsDate(v) And IsDate(maxV) Then
                            If CDate(v) > CDate(maxV) Then
                                mIsValid = False
                                failedRowTracker.Item(r) = True
                                AddError r, colName, v, "OUT_OF_RANGE", "Date " & CStr(v) & " is later than maximum " & CStr(maxV) & "."
                            End If
                        End If
                    End If

                    ' Check 5: Whitelist
                    If hasAllowed Then
                        If Not allowedDict.Exists(CStr(v)) Then
                            mIsValid = False
                            failedRowTracker.Item(r) = True
                            AddError r, colName, v, "NOT_IN_LIST", "Value '" & CStr(v) & "' is not in allowed value list."
                        End If
                    End If

                    ' Check 6: Like Pattern
                    If Len(likePattern) > 0 Then
                        If Not (CStr(v) Like likePattern) Then
                            mIsValid = False
                            failedRowTracker.Item(r) = True
                            AddError r, colName, v, "PATTERN_MISMATCH", "Value '" & CStr(v) & "' does not match pattern '" & likePattern & "'."
                        End If
                    End If

                    ' Check 7: Uniqueness
                    If isUnique Then
                        Dim sKey As String
                        sKey = CStr(v)
                        If seenValues.Exists(sKey) Then
                            mIsValid = False
                            failedRowTracker.Item(r) = True
                            AddError r, colName, v, "DUPLICATE", _
                                     "Duplicate value '" & sKey & "' found (first seen at row " & seenValues.Item(sKey) & ")."
                        Else
                            seenValues.Item(sKey) = r
                        End If
                    End If
                End If
            Next r
        End If
    Next i

    mFailedRowsCount = failedRowTracker.Count
    Validate = mIsValid
End Function

Private Sub AddError(ByVal RowIdx As Long, ByVal ColName As String, ByVal Val As Variant, _
                     ByVal RuleName As String, ByVal Msg As String)
    mErrorsDf.AddRow RowIdx, ColName, Val, RuleName, Msg
End Sub

Private Function CheckTypeMatch(ByVal v As Variant, ByVal Expected As Long) As Boolean
    Select Case Expected
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency, vbDecimal
            CheckTypeMatch = IsNumeric(v)
        Case vbDate
            CheckTypeMatch = IsDate(v)
        Case vbString
            CheckTypeMatch = (VarType(v) = vbString)
        Case vbBoolean
            CheckTypeMatch = (VarType(v) = vbBoolean)
        Case Else
            CheckTypeMatch = True
    End Select
End Function

Private Function TypeNameFromVbType(ByVal vt As Long) As String
    Select Case vt
        Case vbInteger: TypeNameFromVbType = "Integer"
        Case vbLong: TypeNameFromVbType = "Long"
        Case vbDouble: TypeNameFromVbType = "Double"
        Case vbString: TypeNameFromVbType = "String"
        Case vbDate: TypeNameFromVbType = "Date"
        Case vbBoolean: TypeNameFromVbType = "Boolean"
        Case vbCurrency: TypeNameFromVbType = "Currency"
        Case Else: TypeNameFromVbType = "Variant"
    End Select
End Function

' ------------------------------------------------------------------------------
' Results & Reporting Properties
' ------------------------------------------------------------------------------

Public Property Get IsValid() As Boolean
    IsValid = mIsValid
End Property

Public Property Get ErrorCount() As Long
    ErrorCount = mErrorsDf.RowCount
End Property

Public Property Get TotalRowsChecked() As Long
    TotalRowsChecked = mTotalRowsChecked
End Property

Public Property Get FailedRowsCount() As Long
    FailedRowsCount = mFailedRowsCount
End Property

Public Property Get Errors() As cDataFrame
    Set Errors = mErrorsDf
End Property

' Extracts a new cDataFrame containing only passing rows
Public Function GetValidRows(ByVal SourceDf As cDataFrame) As cDataFrame
    If SourceDf Is Nothing Then
        Err.Raise 91, "cDataValidator.GetValidRows", "SourceDf cannot be Nothing."
    End If

    Dim cleanDf As cDataFrame
    Set cleanDf = SourceDf.Clone(True)

    If SourceDf.RowCount = 0 Then
        Set GetValidRows = cleanDf
        Exit Function
    End If

    Dim badRows As Object
    Set badRows = CreateObject("Scripting.Dictionary")

    Dim errRowCol As cColumn
    Set errRowCol = mErrorsDf.Column("RowIndex")
    Dim e As Long, badR As Long
    For e = 1 To mErrorsDf.RowCount
        badR = CLng(errRowCol.Item(e))
        If badR > 0 Then badRows.Item(badR) = True
    Next e

    Dim colNames() As String
    colNames = SourceDf.ColumnNames()
    Dim numCols As Long
    numCols = UBound(colNames) - LBound(colNames) + 1

    Dim colObjs() As cColumn
    ReDim colObjs(1 To numCols)
    Dim c As Long
    For c = 1 To numCols
        Set colObjs(c) = SourceDf.Column(colNames(LBound(colNames) + c - 1))
    Next c

    Dim rowVals() As Variant
    ReDim rowVals(0 To numCols - 1)

    Dim r As Long
    For r = 1 To SourceDf.RowCount
        If Not badRows.Exists(r) Then
            For c = 1 To numCols
                rowVals(c - 1) = colObjs(c).Item(r)
            Next c
            cleanDf.AddRow rowVals
        End If
    Next r

    Set GetValidRows = cleanDf
End Function

' Extracts a new cDataFrame containing only rows with errors
Public Function GetInvalidRows(ByVal SourceDf As cDataFrame) As cDataFrame
    If SourceDf Is Nothing Then
        Err.Raise 91, "cDataValidator.GetInvalidRows", "SourceDf cannot be Nothing."
    End If

    Dim dirtyDf As cDataFrame
    Set dirtyDf = SourceDf.Clone(True)

    If SourceDf.RowCount = 0 Then
        Set GetInvalidRows = dirtyDf
        Exit Function
    End If

    Dim badRows As Object
    Set badRows = CreateObject("Scripting.Dictionary")

    Dim errRowCol As cColumn
    Set errRowCol = mErrorsDf.Column("RowIndex")
    Dim e As Long, badR As Long
    For e = 1 To mErrorsDf.RowCount
        badR = CLng(errRowCol.Item(e))
        If badR > 0 Then badRows.Item(badR) = True
    Next e

    Dim colNames() As String
    colNames = SourceDf.ColumnNames()
    Dim numCols As Long
    numCols = UBound(colNames) - LBound(colNames) + 1

    Dim colObjs() As cColumn
    ReDim colObjs(1 To numCols)
    Dim c As Long
    For c = 1 To numCols
        Set colObjs(c) = SourceDf.Column(colNames(LBound(colNames) + c - 1))
    Next c

    Dim rowVals() As Variant
    ReDim rowVals(0 To numCols - 1)

    Dim r As Long
    For r = 1 To SourceDf.RowCount
        If badRows.Exists(r) Then
            For c = 1 To numCols
                rowVals(c - 1) = colObjs(c).Item(r)
            Next c
            dirtyDf.AddRow rowVals
        End If
    Next r

    Set GetInvalidRows = dirtyDf
End Function
