Attribute VB_Name = "MVBAMath"
' ==========================================================================
' Module      : MVBAMath
' Type        : Module
' Description : Procedures for math-related operations.
' --------------------------------------------------------------------------
' Procedures  : AtLeast             Variant
'               AtMost              Variant
'               MaxVal              Variant
'               MinVal              Variant
'               Within              Variant
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

Private Const msMODULE As String = "MVBAMath"

Public Function AtLeast(ByRef Value As Variant, _
                        ByVal MinVal As Variant) As Variant
' ==========================================================================
' Description : Return the greater of two values.
'
' Parameters  : Value       The value to check
'               MinVal      The minimum allowed value
'
' Returns     : Variant     If Value is less than MinVal, return MinVal.
' ==========================================================================

    Dim vRtn    As Variant

    ' ----------------------------------------------------------------------

    If (Value < MinVal) Then
        vRtn = MinVal
    Else
        vRtn = Value
    End If

    ' ----------------------------------------------------------------------

    AtLeast = vRtn

End Function

Public Function AtMost(ByRef Value As Variant, _
                       ByVal MaxVal As Variant) As Variant
' ==========================================================================
' Description : Return the lesser of two values.
'
' Parameters  : Value       The value to check
'               MaxVal      The maximum allowed value
'
' Returns     : Variant     If Value is greater than MaxVal, return MaxVal.
' ==========================================================================

    Dim vRtn    As Variant

    ' ----------------------------------------------------------------------

    If (Value > MaxVal) Then
        vRtn = MaxVal
    Else
        vRtn = Value
    End If

    ' ----------------------------------------------------------------------

    AtMost = vRtn

End Function

Private Function MaxOfTwo(ByVal Val1 As Variant, _
                          ByVal Val2 As Variant) As Variant
' ==========================================================================
' Purpose   : Compare two values
'
' Arguments : Val1    The first value
'             Val2    The second value
'
' Returns   : Variant
'
' NOTES     : This is a helper function to the MinMaxProcess.
' ==========================================================================

    Dim vRtn    As Variant

    ' ----------------------------------------------------------------------
    ' Both are null - return null
    ' ---------------------------

    If (IsNull(Val1) And IsNull(Val2)) _
    Or ((VarType(Val1) = vbEmpty) And (VarType(Val2) = vbEmpty)) Then
        vRtn = Val1

    ElseIf (IsNull(Val1) Or (VarType(Val1) = vbEmpty)) Then
        vRtn = Val2

    ElseIf (IsNull(Val2) Or (VarType(Val2) = vbEmpty)) Then
        vRtn = Val1

    ElseIf (Val1 > Val2) Then
        vRtn = Val1

    ElseIf (Val1 < Val2) Then
        vRtn = Val2

    ElseIf (Val1 = Val2) Then
        vRtn = Val1

    End If

    ' ----------------------------------------------------------------------

    MaxOfTwo = vRtn

End Function

Public Function MaxVal(ParamArray Values() As Variant) As Variant
' ==========================================================================
' Purpose   : Determine the maximum (greatest) value
'
' Arguments : ParamArray    Two or more values to be compared
'
' Returns   : Variant
' ==========================================================================

    Dim vRtn    As Variant

    ' ----------------------------------------------------------------------

    vRtn = MinMaxProcess(Values, True)

    ' ----------------------------------------------------------------------

    MaxVal = vRtn

End Function

Private Function MinMaxProcess(ByVal Values As Variant, _
                      Optional ByVal UseMax As Boolean) As Variant
' ==========================================================================
' Purpose   : Find the extent value (min or max) in a group
'
' Arguments : Values    The values to inspect
'             UseMax    Return the maximum value
'
' Returns   : Variant
'
' NOTES     : This is a helper function to the MinVal and MaxVal functions.
' ==========================================================================

    Dim sTAB    As String: sTAB = Chr(9)
    Dim sCOMMA  As String: sCOMMA = Chr(44)
    Dim sSPACE  As String: sSPACE = Chr(32)

    Dim lLB     As Long
    Dim lUB     As Long
    Dim lIdx    As Long
    Dim vRtn    As Variant
    Dim vVal    As Variant

    ' ----------------------------------------------------------------------
    ' Determine the size of the array
    ' -------------------------------
    lLB = LBound(Values)
    lUB = UBound(Values)

    ' Test each value
    ' ---------------
    For lIdx = lLB To lUB

        vVal = Values(lIdx)

        If IsArray(vVal) Then
            vRtn = MinMaxProcess(vVal, UseMax)

        ElseIf (VarType(vVal) = vbString) Then
            ' Test if it is a delimited string
            ' --------------------------------
            If (InStr(1, vVal, sTAB, vbTextCompare) > 0) Then
                vRtn = MinMaxProcess(Split(vVal, _
                                           sTAB, _
                                           -1, _
                                           vbTextCompare), UseMax)

            ElseIf (InStr(1, vVal, sCOMMA, vbTextCompare) > 0) Then
                vRtn = MinMaxProcess(Split(vVal, _
                                           sCOMMA, _
                                           -1, _
                                           vbTextCompare), UseMax)

            ElseIf (InStr(1, vVal, sSPACE, vbTextCompare) > 0) Then
                vRtn = MinMaxProcess(Split(vVal, _
                                           sSPACE, _
                                           -1, _
                                           vbTextCompare), UseMax)

                ' Otherwise handle as normal
                ' --------------------------
            Else
                If UseMax Then
                    vRtn = MaxOfTwo(vVal, vRtn)
                Else
                    vRtn = MinOfTwo(vVal, vRtn)
                End If
            End If

        Else
            If UseMax Then
                vRtn = MaxOfTwo(vVal, vRtn)
            Else
                vRtn = MinOfTwo(vVal, vRtn)
            End If
        End If

    Next

    ' ----------------------------------------------------------------------

    MinMaxProcess = vRtn

End Function

Private Function MinOfTwo(ByVal Val1 As Variant, _
                          ByVal Val2 As Variant) As Variant
' ==========================================================================
' Purpose   : Compare two values
'
' Arguments : Val1    The first value
'             Val2    The second value
'
' Returns   : Variant
'
' NOTES     : This is a helper function to the MinMaxProcess.
' ==========================================================================

    Dim vRtn    As Variant

    ' ----------------------------------------------------------------------
    ' Both are null - return null
    ' ---------------------------
    If (IsNull(Val1) And IsNull(Val2)) _
    Or ((VarType(Val1) = vbEmpty) And (VarType(Val2) = vbEmpty)) Then
        vRtn = Val1

    ElseIf IsNull(Val1) Or (VarType(Val1) = vbEmpty) Then
        vRtn = Val2

    ElseIf IsNull(Val2) Or (VarType(Val2) = vbEmpty) Then
        vRtn = Val1

    ElseIf Val1 < Val2 Then
        vRtn = Val1

    ElseIf Val1 > Val2 Then
        vRtn = Val2

    ElseIf Val1 = Val2 Then
        vRtn = Val1

    End If

    ' ----------------------------------------------------------------------

    MinOfTwo = vRtn

End Function

Public Function MinVal(ParamArray Values() As Variant) As Variant
' ==========================================================================
' Purpose   : Return the minimum value from an array of values
'
' Arguments : ParamArray    Two or more values to be compared
'
' Returns   : Variant
' ==========================================================================

    Dim vRtn    As Variant

    ' ----------------------------------------------------------------------

    vRtn = MinMaxProcess(Values)

    ' ----------------------------------------------------------------------

    MinVal = vRtn

End Function

Public Function Within(ByRef Value As Variant, _
                       ByVal MinVal As Variant, _
                       ByVal MaxVal As Variant) As Variant
' ==========================================================================
' Description : Return a value within the given range.
'
' Parameters  : Value       The value to check
'               MinVal      The minimum allowed value
'               MaxVal      The maximum allowed value
'
' Returns     : Variant     If Value is less than MinVal, return MinVal.
'                           If Value is more than MaxVal, return MaxVal.
' ==========================================================================

    Dim vRtn    As Variant

    ' ----------------------------------------------------------------------

    If (Value < MinVal) Then
        vRtn = MinVal
    ElseIf (Value > MaxVal) Then
        vRtn = MaxVal
    Else
        vRtn = Value
    End If

    ' ----------------------------------------------------------------------

    Within = vRtn

End Function
