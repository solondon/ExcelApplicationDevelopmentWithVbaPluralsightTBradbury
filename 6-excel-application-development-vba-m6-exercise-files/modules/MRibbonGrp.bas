Attribute VB_Name = "MRibbonGrp"
' ==========================================================================
' Module      : MRibbonGrp
' Type        : Module
' Description : Support for the IRibbonControl groupBox
' --------------------------------------------------------------------------
' Procedures  : grp_getImage
'               grp_getKeytip
'               grp_getLabel
'               grp_getScreentip
'               grp_getSupertip
'               grp_getVisible
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

Private Const msMODULE As String = "MRibbonGrp"

Public Sub grp_getImage(ByRef Control As IRibbonControl, _
                        ByRef Image As Variant)
' ==========================================================================
' Description : Get the image for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Image       Returns the image for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = gsRXCTL_DFLT_IMAGE
    End Select

    ' ----------------------------------------------------------------------

    Image = sRtn

End Sub

Public Sub grp_getKeytip(ByRef Control As IRibbonControl, _
                         ByRef Keytip As Variant)
' ==========================================================================
' Description : Get the keytip for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Keytip      Returns the keytip for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = Keytip
    End Select

    ' ----------------------------------------------------------------------

    Keytip = sRtn

End Sub

Public Sub grp_getLabel(ByRef Control As IRibbonControl, _
                        ByRef Label As Variant)
' ==========================================================================
' Description : Get the label for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Label       Returns the label for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = Control.Id
    End Select

    ' ----------------------------------------------------------------------

    Label = sRtn

End Sub

Public Sub grp_getScreentip(ByRef Control As IRibbonControl, _
                            ByRef Screentip As Variant)
' ==========================================================================
' Description : Get the screentip for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Screentip   Returns the screentip for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Screentip for " & Control.Id
    End Select

    ' ----------------------------------------------------------------------

    Screentip = sRtn

End Sub

Public Sub grp_getSupertip(ByRef Control As IRibbonControl, _
                           ByRef Supertip As Variant)
' ==========================================================================
' Description : Get the supertip for an IRibbonControl label
'
' Parameters  : Control     The control initiating the callback
'               Supertip    Returns the supertip for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Supertip for " & Control.Id
    End Select

    ' ----------------------------------------------------------------------

    Supertip = sRtn

End Sub

Public Sub grp_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
Attribute grp_getVisible.VB_Description = "Get the visibility for an IRibbonControl"
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
