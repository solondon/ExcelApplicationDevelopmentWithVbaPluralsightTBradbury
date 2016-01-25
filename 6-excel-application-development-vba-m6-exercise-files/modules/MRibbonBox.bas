Attribute VB_Name = "MRibbonBox"
' ==========================================================================
' Module      : MRibbonBox
' Type        : Module
' Description : Support for the IRibbonControl box
' --------------------------------------------------------------------------
' Procedures  : box_getVisible
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

Private Const msMODULE As String = "MRibbonBox"

Public Const gsRXCTL_BOX_H1 As String = "rxboxH1"
Public Const gsRXCTL_BOX_H2 As String = "rxboxH2"
Public Const gsRXCTL_BOX_H3 As String = "rxboxH3"

Public Const gsRXCTL_BOX_V1 As String = "rxboxV1"
Public Const gsRXCTL_BOX_V2 As String = "rxboxV2"
Public Const gsRXCTL_BOX_V3 As String = "rxboxV3"

Public Sub box_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
Attribute box_getVisible.VB_Description = "Get the visibility for an IRibbonControl"
' ==========================================================================
' Description : Get the visibility for an IRibbonControl
'
' Parameters  : Control   The control initiating the callback
'               Visible   Returns the visible state for the control
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case gsRXCTL_BOX_V1
        bRtn = False
    Case Else
        bRtn = True
    End Select

    ' ----------------------------------------------------------------------

    Visible = bRtn

End Sub
