Attribute VB_Name = "MRibbonChk"
' ==========================================================================
' Module      : MRibbonChk
' Type        : Module
' Description : Support for the IRibbonControl checkBox
' --------------------------------------------------------------------------
' Callbacks   : chk_getDescription
'               chk_getEnabled
'               chk_getKeytip
'               chk_getLabel
'               chk_getPressed
'               chk_getScreentip
'               chk_getSupertip
'               chk_getVisible
'               chk_onAction
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
Private Const msMODULE  As String = "MRibbonChk"

Public Sub chk_getDescription(ByRef Control As IRibbonControl, _
                              ByRef Description As Variant)
Attribute chk_getDescription.VB_Description = "Get the description for an IRibbonControl"
' ==========================================================================
' Description : Get the description for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Description Returns the description for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Description of " & Control.Id
    End Select

    ' ----------------------------------------------------------------------

    Description = sRtn

End Sub

Public Sub chk_getEnabled(ByRef Control As IRibbonControl, _
                          ByRef Enabled As Variant)
' ==========================================================================
' Description : Get the enabled state for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Enabled     Returns the enabled state for the control
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        bRtn = True
    End Select

    ' ----------------------------------------------------------------------

    Enabled = bRtn

End Sub

Public Sub chk_getKeytip(ByRef Control As IRibbonControl, _
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

Public Sub chk_getLabel(ByRef Control As IRibbonControl, _
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

Public Sub chk_getPressed(ByRef Control As IRibbonControl, _
                          ByRef Pressed As Variant)
' ==========================================================================
' Description : Get the pressed state for an IRibbonControl
'
' Parameters  : Control   The control initiating the callback
'               Pressed   Returns the pressed state for the control
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        bRtn = True
    End Select

    ' ----------------------------------------------------------------------

    Pressed = bRtn

End Sub

Public Sub chk_getScreentip(ByRef Control As IRibbonControl, _
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

Public Sub chk_getSupertip(ByRef Control As IRibbonControl, _
                           ByRef Supertip As Variant)
' ==========================================================================
' Description : Get the supertip for an IRibbonControl
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

Public Sub chk_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
Attribute chk_getVisible.VB_Description = "Get the visibility for an IRibbonControl"
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

Public Sub chk_onAction(ByRef Control As IRibbonControl, _
                        ByRef Pressed As Boolean)
' ==========================================================================
' Description : Respond to an IRibbonControl action
'
' Parameters  : Control     The control initiating the callback
'               Pressed     The pressed state for the control
' ==========================================================================

    Dim sTitle      As String: sTitle = Control.Id
    Dim sPrompt     As String
    Dim eButtons    As VbMsgBoxStyle: eButtons = vbInformation Or vbOKOnly
    Dim eMBR        As VbMsgBoxResult

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        If CBool(Pressed) Then
            sPrompt = "Pressed."
        Else
            sPrompt = "Not pressed."
        End If
        eMBR = MsgBox(sPrompt, eButtons, sTitle)
    End Select

End Sub
