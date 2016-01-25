Attribute VB_Name = "MRibbonEdt"
' ==========================================================================
' Module      : MRibbonEdt
' Type        : Module
' Description : Support for the IRibbonControl editBox
' --------------------------------------------------------------------------
' Callbacks   : edt_getEnabled
'               edt_getImage
'               edt_getKeytip
'               edt_getLabel
'               edt_getScreentip
'               edt_getShowImage
'               edt_getShowLabel
'               edt_getSupertip
'               edt_getText
'               edt_getVisible
'               edt_onChange
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

Private Const msMODULE As String = "MRibbonEdt"

Public Sub edt_getEnabled(ByRef Control As IRibbonControl, _
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

Public Sub edt_getImage(ByRef Control As IRibbonControl, _
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

Public Sub edt_getKeytip(ByRef Control As IRibbonControl, _
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

Public Sub edt_getLabel(ByRef Control As IRibbonControl, _
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

Public Sub edt_getScreentip(ByRef Control As IRibbonControl, _
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

Public Sub edt_getShowImage(ByRef Control As IRibbonControl, _
                            ByRef ShowImage As Variant)
' ==========================================================================
' Description : Get the visibility for an IRibbonControl image
'
' Parameters  : Control     The control initiating the callback
'               ShowImage   Returns the visibility of the image
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        bRtn = True
    End Select

    ' ----------------------------------------------------------------------

    ShowImage = bRtn

End Sub

Public Sub edt_getShowLabel(ByRef Control As IRibbonControl, _
                            ByRef ShowLabel As Variant)
' ==========================================================================
' Description : Get the visibility for an IRibbonControl label
'
' Parameters  : Control     The control initiating the callback
'               ShowLabel   Returns the visibility of the label
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        bRtn = True
    End Select

    ' ----------------------------------------------------------------------

    ShowLabel = bRtn

End Sub

Public Sub edt_getSupertip(ByRef Control As IRibbonControl, _
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

Public Sub edt_getText(ByRef Control As IRibbonControl, _
                       ByRef Text As Variant)
' ==========================================================================
' Description : Get the text for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Text        Returns the text for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = vbNullString
    End Select

    ' ----------------------------------------------------------------------

    Text = sRtn

End Sub

Public Sub edt_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
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

Public Sub edt_onChange(ByRef Control As IRibbonControl, _
                        ByRef Text As String)
' ==========================================================================
' Description : Respond to an IRibbonControl action
'
' Parameters  : Control     The control initiating the callback
'               Text        The text for the control
' ==========================================================================

    Dim sTitle      As String: sTitle = Control.Id
    Dim sPrompt     As String: sPrompt = Text
    Dim eButtons    As VbMsgBoxStyle: eButtons = vbInformation Or vbOKOnly
    Dim eMBR        As VbMsgBoxResult

    Dim lPos        As Long

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        eMBR = MsgBox(sPrompt, eButtons, sTitle)
    End Select

End Sub

