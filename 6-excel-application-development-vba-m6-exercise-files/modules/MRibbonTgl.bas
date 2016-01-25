Attribute VB_Name = "MRibbonTgl"
' ==========================================================================
' Module      : MRibbonTgl
' Type        : Module
' Description : Support for the IRibbonControl toggleButton
' --------------------------------------------------------------------------
' Procedures  : tgl_getDescription
'               tgl_getEnabled
'               tgl_getImage
'               tgl_getLabel
'               tgl_getPressed
'               tgl_getScreentip
'               tgl_getShowImage
'               tgl_getShowLabel
'               tgl_getSize
'               tgl_getSupertip
'               tgl_getVisible
'               tgl_onAction
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

Public Const gsRXCTL_TGL_ALIGNLEFT      As String = "rxtglAlignLeft"
Public Const gsRXCTL_TGL_ALIGNCENTER    As String = "rxtglAlignCenter"
Public Const gsRXCTL_TGL_ALIGNRIGHT     As String = "rxtglAlignRight"

' ----------------
' Module Level
' ----------------

Private Const msMODULE                  As String = "MRibbonTgl"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Module Level
' ----------------

Private Enum enuTextAlign
    taUnknown = 0
    taLeft
    taCenter
    taRight
End Enum

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private meAlign                         As enuTextAlign

Public Sub tgl_getDescription(ByRef Control As IRibbonControl, _
                              ByRef Description As Variant)
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

Public Sub tgl_getEnabled(ByRef Control As IRibbonControl, _
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

Public Sub tgl_getImage(ByRef Control As IRibbonControl, _
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

Public Sub tgl_getKeytip(ByRef Control As IRibbonControl, _
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

Public Sub tgl_getLabel(ByRef Control As IRibbonControl, _
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

Public Sub tgl_getPressed(ByRef Control As IRibbonControl, _
                          ByRef Pressed As Variant)
' ==========================================================================
' Description : Get the Pressed state for a ribbon control
'
' Parameters  : Control   The control initiating the callback
'               Pressed   Returns the Pressed state
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        bRtn = False
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Pressed = bRtn

End Sub

Public Sub tgl_getScreentip(ByRef Control As IRibbonControl, _
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

Public Sub tgl_getShowImage(ByRef Control As IRibbonControl, _
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

Public Sub tgl_getShowLabel(ByRef Control As IRibbonControl, _
                            ByRef ShowLabel As Variant)
' ==========================================================================
' Description : Get the visibility for an IRibbonControl label
'
' Parameters  : Control     The control initiating the callback
'               ShowLabel   Returns the visible state of the label
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

Public Sub tgl_getSize(ByRef Control As IRibbonControl, _
                       ByRef Size As Variant)
' ==========================================================================
' Description : Get the size for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Size        Returns the size for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = gsRXCTL_DFLT_SIZE
    End Select

    ' ----------------------------------------------------------------------

    Size = sRtn

End Sub

Public Sub tgl_getSupertip(ByRef Control As IRibbonControl, _
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

Public Sub tgl_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
Attribute tgl_getVisible.VB_Description = "Get the visibility for an IRibbonControl"
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

Public Sub tgl_onAction(ByRef Control As IRibbonControl, _
                        ByRef Pressed As Boolean)
' ==========================================================================
' Description : Respond to an IRibbonControl action
'
' Parameters  : Control   The control initiating the callback
'               Pressed   The pressed state for the control
' ==========================================================================

    Dim sTitle      As String: sTitle = Control.Id
    Dim eButtons    As VbMsgBoxStyle: eButtons = vbInformation Or vbOKOnly
    Dim eMBR        As VbMsgBoxResult
    Dim sPrompt     As String

    Select Case Control.Id
    Case gsRXCTL_TGL_ALIGNLEFT
        If (meAlign = taLeft) Then
            meAlign = taUnknown
        Else
            meAlign = taLeft
        End If
        goRibbon.InvalidateControl gsRXCTL_TGL_ALIGNCENTER
        goRibbon.InvalidateControl gsRXCTL_TGL_ALIGNRIGHT

    Case gsRXCTL_TGL_ALIGNCENTER
        If (meAlign = taCenter) Then
            meAlign = taUnknown
        Else
            meAlign = taCenter
        End If
        goRibbon.InvalidateControl gsRXCTL_TGL_ALIGNLEFT
        goRibbon.InvalidateControl gsRXCTL_TGL_ALIGNRIGHT

    Case gsRXCTL_TGL_ALIGNRIGHT
        If (meAlign = taRight) Then
            meAlign = taUnknown
        Else
            meAlign = taRight
        End If
        goRibbon.InvalidateControl gsRXCTL_TGL_ALIGNLEFT
        goRibbon.InvalidateControl gsRXCTL_TGL_ALIGNCENTER

    Case Else
        If CBool(Pressed) Then
            sPrompt = "Pressed."
        Else
            sPrompt = "Not pressed."
        End If

        eMBR = MsgBox(sPrompt, eButtons, sTitle)
    End Select

End Sub
