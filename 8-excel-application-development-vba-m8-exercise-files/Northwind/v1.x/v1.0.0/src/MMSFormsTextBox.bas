Attribute VB_Name = "MMSFormsTextBox"
' ==========================================================================
' Module      : MMSFormsTextBox
' Type        : Module
' Description : Procedures for use with TextBox controls
' --------------------------------------------------------------------------
' Procedures  : KeyCapitalize       Boolean
'               KeyFilter
'               KeySpaceFill        Boolean
' --------------------------------------------------------------------------
' Dependencies: MVBAMath
' --------------------------------------------------------------------------
' References  : Microsoft Forms 2.0 Object Library
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

Private Const msMODULE As String = "MMSFormsTextBox"

Public Function KeyCapitalize(ByRef KeyAscii As MSForms.ReturnInteger) _
       As Boolean
' ==========================================================================
' Description : Capitalize a KeyPress keystroke
'
' Parameters  : KeyAscii    The character to capitalize
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC         As String = "KeyCapitalize"
    Const lLCASE_OFFSET As Long = 32

    Dim bRtn            As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If IsBetween(CDbl(KeyAscii), _
                 vbKeyA + lLCASE_OFFSET, _
                 vbKeyZ + lLCASE_OFFSET) Then
        KeyAscii = KeyAscii - lLCASE_OFFSET
        bRtn = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    KeyCapitalize = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
      Stop
      Resume
    Else
      Resume PROC_EXIT
    End If

End Function

Public Sub KeyFilter(ByRef KeyAscii As MSForms.ReturnInteger, _
                     ByVal AllowAlpha As Boolean, _
                     ByVal AllowNumeric As Boolean, _
                     ParamArray AlsoAllow() As Variant)
' ==========================================================================
' Description : Limit text entry to specific characters or types
'
' Parameters  : KeyAscii        The ascii code of the key that was pressed
'               AllowNumeric    Set to True to allow numeric input
'               AllowAlpha      Set to True to allow alpha input
'               AlsoAllow       Any additional characters that are allowed
'
' Notes       : This helper function is intended for use with TextBox and
'               ComboBox controls and their associated KeyPress events
' ==========================================================================

    Const sPROC         As String = "KeyFilter"
    Const lLCASE_OFFSET As Long = 32

    Dim bOK             As Boolean
    Dim vOtherKey       As Variant


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Some keys are always allowed
    ' ----------------------------
    If AnyEqual(KeyAscii, _
                vbKeyBack, _
                vbKeyUp, _
                vbKeyDown, _
                vbKeyLeft, _
                vbKeyRight) Then
        bOK = True

        ' If numerics allowed then check if KeyAscii is a number
        ' ------------------------------------------------------
    ElseIf (AllowNumeric _
            And IsBetween(CDbl(KeyAscii), vbKey0, vbKey9, True)) Then
        bOK = True

        ' If alpha allowed then check if KeyAscii is an alpha
        ' ---------------------------------------------------
    ElseIf (AllowAlpha _
       And (IsBetween(CDbl(KeyAscii), vbKeyA, vbKeyZ, True) _
         Or IsBetween(CDbl(KeyAscii), _
                      vbKeyA + lLCASE_OFFSET, _
                      vbKeyZ + lLCASE_OFFSET, _
                      True))) Then
        bOK = True
    Else

        ' If not a number or alpha then
        ' check for other allowed key values
        ' ----------------------------------
        For Each vOtherKey In AlsoAllow

            ' Convert strings to Ascii value
            ' assume numerics are key constants
            ' ---------------------------------
            If KeyAscii = IIf(TypeName(vOtherKey) = gsTYPENAME_STRING, _
                              Asc(vOtherKey), _
                              vOtherKey) Then
                bOK = True
                Exit For
            End If
        Next vOtherKey

    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    If (Not bOK) Then
        KeyAscii = 0
    End If

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub

Public Function KeySpaceFill(ByRef KeyAscii As MSForms.ReturnInteger, _
                    Optional ByVal SpaceFill As String = "_") As Boolean
' ==========================================================================
' Description : Replace spaces with a specified fill character
'
' Parameters  : KeyAscii        The ascii code of the key that was pressed
'               SpaceFill       The character to replace spaces with
'
' Notes       : This helper function is intended for use with TextBox and
'               ComboBox controls and their associated KeyPress events
' ==========================================================================

    Const sPROC As String = "KeySpaceFill"

    Dim bRtn     As Boolean


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Test for space
    ' --------------
    If (KeyAscii = 32) Then
        KeyAscii = Asc(SpaceFill)
        bRtn = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    KeySpaceFill = bRtn

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function
