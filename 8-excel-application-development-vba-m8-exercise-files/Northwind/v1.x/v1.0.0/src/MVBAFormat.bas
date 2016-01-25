Attribute VB_Name = "MVBAFormat"
' ==========================================================================
' Module      : MVBAFormat
' Type        : Module
' Description : Support for working with the Format function.
' --------------------------------------------------------------------------
' Procedures  : XXX
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

' Date/Time formats
' -----------------
Public Const gsVBA_FMTDTM_GENERALDATE   As String = "General Date"
Public Const gsVBA_FMTDTM_LONGDATE      As String = "Long Date"
Public Const gsVBA_FMTDTM_LONGTIME      As String = "Long Time"
Public Const gsVBA_FMTDTM_MEDIUMDATE    As String = "Medium Date"
Public Const gsVBA_FMTDTM_MEDIUMTIME    As String = "Medium Time"
Public Const gsVBA_FMTDTM_SHORTDATE     As String = "Short Date"
Public Const gsVBA_FMTDTM_SHORTTIME     As String = "Short Time"

' Number formats
' --------------
Public Const gsVBA_FMTNUM_GENERALNUMBER As String = "General Number"
Public Const gsVBA_FMTNUM_CURRENCY      As String = "Currency"
Public Const gsVBA_FMTNUM_FIXED         As String = "Fixed"
Public Const gsVBA_FMTNUM_STANDARD      As String = "Standard"
Public Const gsVBA_FMTNUM_PERCENT       As String = "Percent"
Public Const gsVBA_FMTNUM_SCIENTIFIC    As String = "Scientific"
Public Const gsVBA_FMTNUM_YESNO         As String = "Yes/No"
Public Const gsVBA_FMTNUM_TRUEFALSE     As String = "True/False"
Public Const gsVBA_FMTNUM_ONOFF         As String = "On/Off"
