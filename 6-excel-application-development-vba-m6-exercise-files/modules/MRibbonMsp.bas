Attribute VB_Name = "MRibbonMsp"
' ==========================================================================
' Module      : MRibbonMsp
' Type        : Module
' Description : Support for the IRibbonControl menuSeparator
' --------------------------------------------------------------------------
' Procedures  : msp_getTitle
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

Private Const msMODULE As String = "MRibbonMsp"

Public Sub msp_getTitle(ByRef Control As IRibbonControl, _
                        ByRef Title As Variant)
' ==========================================================================
' Description : Get the title for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Title       Returns the title for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "menuSeparator (" & Control.Id & ")"
    End Select

    ' ----------------------------------------------------------------------

    Title = sRtn

End Sub
