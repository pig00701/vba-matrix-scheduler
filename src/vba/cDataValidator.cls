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

' Loads validation rules from an Excel Table (ListObject), Range, or cDataFrame.
' Supported columns (case-insensitive): ColumnName/Column, ExpectedType/Type,
' AllowNull/Required, IsUnique/Unique, MinVal/Min, MaxVal/Max, AllowedValues, LikePattern/Pattern, Enabled.
Public Function AddRulesFromTable(ByVal RuleTableOrDf As Variant) As cDataValidator
    Dim ruleDf As cDataFrame
    Set ruleDf = ResolveRuleDf(RuleTableOrDf)
    If ruleDf.RowCount = 0 Then
        Set AddRulesFromTable = Me
        Exit Function
    End If

    Dim colNameCol As String, typeCol As String, allowNullCol As String, reqCol As String
    Dim uniqCol As String, minCol As String, maxCol As String
    Dim allowedCol As String, patternCol As String, enabledCol As String

    colNameCol = FindRuleColName(ruleDf, Array("ColumnName", "Column", "Field"))
    If Len(colNameCol) = 0 Then
        Err.Raise 5, "cDataValidator.AddRulesFromTable", "Rule table requires a 'ColumnName' (or 'Column') column."
    End If

    typeCol = FindRuleColName(ruleDf, Array("ExpectedType", "Type", "DataType"))
    allowNullCol = FindRuleColName(ruleDf, Array("AllowNull", "Nullable"))
    reqCol = FindRuleColName(ruleDf, Array("Required", "IsRequired"))
    uniqCol = FindRuleColName(ruleDf, Array("IsUnique", "Unique"))
    minCol = FindRuleColName(ruleDf, Array("MinVal", "Min", "MinValue"))
    maxCol = FindRuleColName(ruleDf, Array("MaxVal", "Max", "MaxValue"))
    allowedCol = FindRuleColName(ruleDf, Array("AllowedValues", "Allowed", "Whitelist"))
    patternCol = FindRuleColName(ruleDf, Array("LikePattern", "Pattern"))
    enabledCol = FindRuleColName(ruleDf, Array("Enabled", "Active", "IsEnabled"))

    Dim r As Long
    For r = 1 To ruleDf.RowCount
        If Len(enabledCol) > 0 Then
            If Not ParseRuleBool(ruleDf.Item(r, enabledCol), True) Then GoTo NextRuleRow
        End If

        Dim targetCol As String
        targetCol = Trim$(CStr(ruleDf.Item(r, colNameCol)))
        If Len(targetCol) = 0 Then GoTo NextRuleRow

        Dim expType As VbVarType
        expType = vbEmpty
        If Len(typeCol) > 0 Then expType = ParseExpectedVarType(CStr(ruleDf.Item(r, typeCol)))

        Dim allowNull As Boolean
        allowNull = True
        If Len(allowNullCol) > 0 Then
            allowNull = ParseRuleBool(ruleDf.Item(r, allowNullCol), True)
        ElseIf Len(reqCol) > 0 Then
            allowNull = Not ParseRuleBool(ruleDf.Item(r, reqCol), False)
        End If

        Dim isUniq As Boolean
        isUniq = False
        If Len(uniqCol) > 0 Then isUniq = ParseRuleBool(ruleDf.Item(r, uniqCol), False)

        Dim minV As Variant, maxV As Variant
        minV = Empty
        maxV = Empty
        If Len(minCol) > 0 Then
            Dim rawMin As Variant
            rawMin = ruleDf.Item(r, minCol)
            If Not IsEmpty(rawMin) And Not IsNull(rawMin) Then
                If Len(Trim$(CStr(rawMin))) > 0 Then minV = CDbl(rawMin)
            End If
        End If
        If Len(maxCol) > 0 Then
            Dim rawMax As Variant
            rawMax = ruleDf.Item(r, maxCol)
            If Not IsEmpty(rawMax) And Not IsNull(rawMax) Then
                If Len(Trim$(CStr(rawMax))) > 0 Then maxV = CDbl(rawMax)
            End If
        End If

        Dim allowedV As Variant
        allowedV = Empty
        If Len(allowedCol) > 0 Then
            Dim rawAllowed As String
            rawAllowed = Trim$(CStr(ruleDf.Item(r, allowedCol)))
            If Len(rawAllowed) > 0 Then
                Dim parts() As String, idx As Long
                If InStr(rawAllowed, "|") > 0 Then
                    parts = Split(rawAllowed, "|")
                Else
                    parts = Split(rawAllowed, ",")
                End If
                For idx = LBound(parts) To UBound(parts)
                    parts(idx) = Trim$(parts(idx))
                Next idx
                allowedV = parts
            End If
        End If

        Dim patStr As String
        patStr = ""
        If Len(patternCol) > 0 Then patStr = Trim$(CStr(ruleDf.Item(r, patternCol)))

        Me.AddRule targetCol, ExpectedType:=expType, AllowNull:=allowNull, _
                   IsUnique:=isUniq, MinVal:=minV, MaxVal:=maxV, _
                   AllowedValues:=allowedV, LikePattern:=patStr
NextRuleRow:
    Next r

    Set AddRulesFromTable = Me
End Function

Private Function ResolveRuleDf(ByVal Source As Variant) As cDataFrame
    Dim df As New cDataFrame

    If VarType(Source) = vbString Then
        Dim strName As String
        strName = Trim$(CStr(Source))
        If Len(strName) = 0 Then
            Err.Raise 5, "cDataValidator.AddRulesFromTable", "TableName string cannot be blank."
        End If
        df.LoadFromTableName Application.ActiveWorkbook, strName
        Set ResolveRuleDf = df
        Exit Function
    End If

    If Not IsObject(Source) Then
        Err.Raise 5, "cDataValidator.AddRulesFromTable", "RuleTableOrDf must be a Table Name string, cDataFrame, ListObject, or Range."
    End If
    If Source Is Nothing Then
        Err.Raise 91, "cDataValidator.AddRulesFromTable", "RuleTableOrDf cannot be Nothing."
    End If
    If TypeOf Source Is cDataFrame Then
        Set ResolveRuleDf = Source
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
    Set ResolveRuleDf = df
End Function

Private Function FindRuleColName(ByVal Df As cDataFrame, ByVal Candidates As Variant) As String
    Dim names() As String
    names = Df.ColumnNames()
    Dim i As Long, j As Long
    For j = LBound(Candidates) To UBound(Candidates)
        For i = LBound(names) To UBound(names)
            If StrComp(Trim$(names(i)), CStr(Candidates(j)), vbTextCompare) = 0 Then
                FindRuleColName = names(i)
                Exit Function
            End If
        Next i
    Next j
End Function

Private Function ParseExpectedVarType(ByVal TypeStr As String) As VbVarType
    Select Case UCase$(Trim$(TypeStr))
        Case "STRING", "TEXT", "STR": ParseExpectedVarType = vbString
        Case "LONG", "INT", "INTEGER": ParseExpectedVarType = vbLong
        Case "DOUBLE", "NUMBER", "NUM", "FLOAT", "DECIMAL": ParseExpectedVarType = vbDouble
        Case "DATE", "DATETIME": ParseExpectedVarType = vbDate
        Case "BOOL", "BOOLEAN": ParseExpectedVarType = vbBoolean
        Case Else: ParseExpectedVarType = vbEmpty
    End Select
End Function

Private Function ParseRuleBool(ByVal Val As Variant, Optional ByVal DefaultVal As Boolean = False) As Boolean
    If IsEmpty(Val) Or IsNull(Val) Then
        ParseRuleBool = DefaultVal
        Exit Function
    End If
    If VarType(Val) = vbBoolean Then
        ParseRuleBool = CBool(Val)
        Exit Function
    End If
    Dim s As String
    s = UCase$(Trim$(CStr(Val)))
    If Len(s) = 0 Then
        ParseRuleBool = DefaultVal
    ElseIf s = "TRUE" Or s = "YES" Or s = "Y" Or s = "1" Then
        ParseRuleBool = True
    ElseIf s = "FALSE" Or s = "NO" Or s = "N" Or s = "0" Then
        ParseRuleBool = False
    Else
        ParseRuleBool = DefaultVal
    End If
End Function
