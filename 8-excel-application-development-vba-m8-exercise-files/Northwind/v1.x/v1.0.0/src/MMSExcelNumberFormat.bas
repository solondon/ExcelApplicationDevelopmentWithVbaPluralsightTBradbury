Attribute VB_Name = "MMSExcelNumberFormat"
' ==========================================================================
' Module      : MMSExcelNumberFormat
' Type        : Module
' Description : Support for working with number
'               formatting using the Format function
' --------------------------------------------------------------------------
' Procedures  : NumberFormatToValue
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

' Number formats
' --------------
Public Const gsEXCEL_NUMFMT_GENERAL         As String = "General"

Public Const gsEXCEL_NUMFMT_GENERALDATE     As String = "General Date"
Public Const gsEXCEL_NUMFMT_LONGDATE        As String = "Long Date"
Public Const gsEXCEL_NUMFMT_LONGTIME        As String = "Long Time"
Public Const gsEXCEL_NUMFMT_MEDIUMDATE      As String = "Medium Date"
Public Const gsEXCEL_NUMFMT_MEDIUMTIME      As String = "Medium Time"
Public Const gsEXCEL_NUMFMT_SHORTDATE       As String = "Short Date"
Public Const gsEXCEL_NUMFMT_SHORTTIME       As String = "Short Time"

Public Const gsEXCEL_NUMFMT_GENERALNUMBER   As String = "General Number"
Public Const gsEXCEL_NUMFMT_CURRENCY        As String = "Currency"
Public Const gsEXCEL_NUMFMT_FIXED           As String = "Fixed"
Public Const gsEXCEL_NUMFMT_STANDARD        As String = "Standard"
Public Const gsEXCEL_NUMFMT_PERCENT         As String = "Percent"
Public Const gsEXCEL_NUMFMT_SCIENTIFIC      As String = "Scientific"
Public Const gsEXCEL_NUMFMT_YESNO           As String = "Yes/No"
Public Const gsEXCEL_NUMFMT_TRUEFALSE       As String = "True/False"
Public Const gsEXCEL_NUMFMT_ONOFF           As String = "On/Off"

Public Const gsEXCEL_NUMFMT_TEXT            As String = "@"

' ----------------
' Module Level
' ----------------

Private Const msMODULE                      As String = "MMSExcelFormat"

Public Sub NumberFormatToValue(ByRef Target As Excel.Range)
' ==========================================================================
' Description : Move the NumberFormat string to text
'
' Parameters  : Target    The range of cells to modify
' ==========================================================================

    Const sPROC As String = "NumberFormatToValue"

    Dim sNumFmt As String
    Dim Cell    As Excel.Range


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    For Each Cell In Target.Cells

        ' Reset
        ' -----
        sNumFmt = vbNullString

        ' Get the format
        ' --------------
        If IsNumeric(Left$(Cell.NumberFormat, 1)) Then
            sNumFmt = "'" & Cell.NumberFormat
        Else
            sNumFmt = Cell.NumberFormat
        End If

        ' Clear the format
        ' ----------------
        Cell.NumberFormat = gsEXCEL_NUMFMT_GENERAL

        ' Copy the converted format
        ' -------------------------
        If (sNumFmt <> gsEXCEL_NUMFMT_GENERAL) Then
            Cell = sNumFmt
        End If

    Next Cell

    Call ClearErrors(Target)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set Cell = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, Target.Address)
    On Error GoTo 0

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
