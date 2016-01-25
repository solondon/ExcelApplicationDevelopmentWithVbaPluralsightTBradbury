Attribute VB_Name = "MRibbonSep"
' ==========================================================================
' Module      : MRibbonSep
' Type        : Module
' Description : Support for the IRibbonControl separator
' --------------------------------------------------------------------------
' Procedures  : sep_getVisible
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

Private Const msMODULE As String = "MRibbonSep"

Public Sub sep_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
Attribute sep_getVisible.VB_Description = "Get the visibility for an IRibbonControl"
' ==========================================================================
' Description : Get the visibility for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Visible     Returns the visible state for the control
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        bRtn = True
    End Select

    ' ----------------------------------------------------------------------

    Visible = bRtn

End Sub
