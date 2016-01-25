Attribute VB_Name = "MVBALogic"
' ==========================================================================
' Module      : MVBALogic
' Type        : Module
' Description : Procedures for logical comparisons.
' --------------------------------------------------------------------------
' Procedures  : AnyEqual            Boolean
'               IsAlpha             Boolean
'               IsBetween           Boolean
'               IsBoolean           Boolean
'               IsBracket           Boolean
'               IsIn                Boolean
'               IsOperator          Boolean
'               IsOutside           Boolean
'               ObjectExists        Boolean
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE As String = "MVBALogic"

Public Function AnyEqual(ByVal Value As Variant, _
                         ParamArray CompareValues() As Variant) As Boolean
' ==========================================================================
' Description : Do a comparison to see if any values match a given value
'
' Parameters  : Value         The value to look for
'               CompareValues The values to compare to Value
'
' Returns     : Boolean
'
' Comments    : This function does a direct, case-sensitive
'               comparison, which is slightly differently than IsIn.
' ==========================================================================

    Dim bRtn        As Boolean
    Dim vElement    As Variant
    Dim vVal        As Variant

    ' ----------------------------------------------------------------------
    ' Compare each value passed to the function
    ' -----------------------------------------

    For Each vVal In CompareValues

        If IsArray(vVal) Then
            ' If the value passed is an array
            ' parse and compare each value
            ' -------------------------------

            For Each vElement In vVal

                If (vElement = Value) Then
                    bRtn = True
                    Exit For
                End If

            Next vElement

        Else
            ' Compare singleton value
            ' -----------------------

            If vVal = Value Then
                bRtn = True
                Exit For
            End If
        End If

        If bRtn Then
            Exit For
        End If

    Next vVal

    ' ----------------------------------------------------------------------

    AnyEqual = bRtn

End Function

Public Function IsAlpha(ByVal Value As String) As Boolean
' ==========================================================================
' Description : Determines if a value is alphabetic
'
' Parameters  : Value       The value to test
'
' Returns     : Boolean
'
' Comments    : This test is only valid for the basic
'               26-character western latin alphabet, and does not support
'               glyphs, ligatures, accents, or other diacritical markings.
' ==========================================================================

    Const sPROC As String = "IsAlpha"

    Dim bRtn    As Boolean

    Dim lAsc    As Long
    Dim lIdx    As Long
    Dim lLen    As Long

    ' ----------------------------------------------------------------------

    lLen = Len(Value)

    For lIdx = 1 To lLen
        lAsc = Asc(Mid$(Value, lIdx, 1))

        Select Case lAsc
        Case 65 To 90   ' Upper case
            bRtn = True

        Case 97 To 122  ' Lower case
            bRtn = True

        Case Else       ' Non-alpha
            bRtn = False
            Exit For
        End Select
    Next lIdx

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsAlpha = bRtn

End Function

Public Function IsBetween(ByVal Value As Variant, _
                          ByVal RangeStart As Variant, _
                          ByVal RangeEnd As Variant, _
                 Optional ByVal Inclusive As Boolean = True) As Boolean
' ==========================================================================
' Purpose   : Determines if a value is between two other values.
'
' Arguments : Val           The value to test
'
'             RangeStart    The start of the range of values to compare to
'
'             RangeEnd      The end of the range of values to compare to
'
'             Inclusive     Indicates if the comparison should consider
'                           the start and end values as between
'
' Returns   : Boolean
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    If Inclusive Then
        bRtn = ((Value >= RangeStart) And (Value <= RangeEnd))
    Else
        bRtn = ((Value > RangeStart) And (Value < RangeEnd))
    End If

    ' ----------------------------------------------------------------------

    IsBetween = bRtn

End Function

Public Function IsBoolean(ByVal Exp As String) As Boolean
' ==========================================================================
' Description : Determine if a string expression is a boolean
'
' Parameters  : Exp   The expression to analyze
'
' Returns     : Boolean
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------
    ' First look for simple booleans
    ' ------------------------------

    Select Case UCase(Exp)
    Case "Y", "YES", "N", "NO", "T", "TRUE", "F", "FALSE", "X"
        bRtn = True
        GoTo PROC_EXIT
    End Select

    ' Look for a boolean expression
    ' -----------------------------
    If (InStr(1, Exp, "=", vbTextCompare) > 0) Then
        bRtn = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsBoolean = bRtn

End Function

Public Function IsBracket(ByVal Char As String) As Boolean
' ==========================================================================
' Description : Evaluate an character to determine
'               if it is a bracketing character.
'
' Parameters  : Char     The character to evaluate
'
' Returns     : Boolean
' ==========================================================================

    Const sBRACKETS As String = "(){}[]<>"

    Dim bRtn        As Boolean

    ' ----------------------------------------------------------------------

    If (Len(Trim(Char)) > 1) Then
        GoTo PROC_EXIT
    ElseIf (Trim(Char) = vbNullString) Then
        GoTo PROC_EXIT
    End If

    If (InStr(1, sBRACKETS, Trim(Char), vbBinaryCompare) > 0) Then
        bRtn = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsBracket = bRtn

End Function

Public Function IsIn(ByVal Value As Variant, _
                     ParamArray Params() As Variant) As Boolean
' ==========================================================================
' Description : Test if a value is in a list
'
' Parameters  : Value   The value to search for
'               Params  The list of items to search in
'
' Returns     : Boolean
'
' Comments    : This function operates slightly differently than AnyEqual
'               by including string conversion, trimming, and ignoring case.
' ==========================================================================

    Dim bRtn        As Boolean

    Dim vParam      As Variant
    Dim vElement    As Variant

    ' ----------------------------------------------------------------------

    For Each vParam In Params

        If IsArray(vParam) Then

            For Each vElement In vParam
                If (Value = vElement) Then
                    bRtn = True
                    GoTo PROC_EXIT
                ElseIf (StrComp(Trim$(CStr(Value)), _
                                Trim$(CStr(vElement)), _
                                vbTextCompare) = 0) Then
                    bRtn = True
                    GoTo PROC_EXIT
                End If
            Next vElement

        Else

            If (Value = vParam) Then
                bRtn = True
                GoTo PROC_EXIT
            ElseIf (StrComp(Trim$(CStr(Value)), _
                            Trim$(CStr(vParam)), _
                            vbTextCompare) = 0) Then
                bRtn = True
                GoTo PROC_EXIT
            End If

        End If
    Next vParam

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsIn = bRtn

End Function

Public Function IsOperator(ByVal Char As String, _
                  Optional ByVal IncludeAssignment As Boolean = True) _
       As Boolean
' ==========================================================================
' Description : Evaluate an character to determine
'               if it is a mathematical operator.
'
' Parameters  : Char     The character to evaluate
'
' Returns     : Boolean
' ==========================================================================

    Dim sOperators  As String

    Dim bRtn        As Boolean

    ' ----------------------------------------------------------------------
    ' Determine if assignment ("=") should
    ' be considered a mathematical operator
    ' -------------------------------------

    If IncludeAssignment Then
        sOperators = "+-*/\^=,"
    Else
        sOperators = "+-*/\^"
    End If

    ' ----------------------------------------------------------------------

    If (Len(Trim(Char)) > 1) Then
        GoTo PROC_EXIT
    ElseIf (Trim(Char) = vbNullString) Then
        GoTo PROC_EXIT
    End If

    If (InStr(1, sOperators, Trim(Char), vbBinaryCompare) > 0) Then
        bRtn = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsOperator = bRtn

End Function

Public Function IsOutside(ByVal Value As Variant, _
                          ByVal RangeStart As Variant, _
                          ByVal RangeEnd As Variant) As Boolean
' ==========================================================================
' Purpose   : Determines if a value is outside a range of values.
'
' Arguments : Value         The value to test
'
'             RangeStart    The start of the range of values to compare to
'
'             RangeEnd      The end of the range of values to compare to
'
' Returns   : Boolean
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    bRtn = ((Value < RangeStart) Or (Value > RangeEnd))

    ' ----------------------------------------------------------------------

    IsOutside = bRtn

End Function

Public Function ObjectExists(ByRef obj As Object) As Boolean
' ==========================================================================
' Description : Determines if a given object is instantiated
'               or references a valid object
'
' Parameters  : Obj     The object to test
'
' Returns     : Boolean
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    bRtn = (Not (obj Is Nothing))

    ' ----------------------------------------------------------------------

    ObjectExists = bRtn

End Function
