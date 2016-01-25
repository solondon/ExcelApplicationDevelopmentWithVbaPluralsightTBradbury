Attribute VB_Name = "MRibbonBtn"
' ==========================================================================
' Module      : MRibbonBtn
' Type        : Module
' Description : Support for the IRibbonControl button
' --------------------------------------------------------------------------
' Callbacks   : btn_getDescription
'               btn_getEnabled
'               btn_getImage
'               btn_getKeytip
'               btn_getLabel
'               btn_getScreentip
'               btn_getShowImage
'               btn_getShowLabel
'               btn_getSize
'               btn_getSupertip
'               btn_getVisible
'               btn_onAction
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

Private Const msMODULE              As String = "MRibbonBtn"

' Ribbon controls
' ---------------
Public Const gsRXCTL_BTN_BUTTON    As String = "rxbtnButton"
Public Const gsRXCTL_BTN_BUTTON1   As String = "rxbtnButton1"
Public Const gsRXCTL_BTN_BUTTON2   As String = "rxbtnButton2"
Public Const gsRXCTL_BTN_BUTTON3   As String = "rxbtnButton3"
Public Const gsRXCTL_BTN_BUTTON4   As String = "rxbtnButton4"

Public Sub btn_getDescription(ByRef Control As IRibbonControl, _
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

Public Sub btn_getEnabled(ByRef Control As IRibbonControl, _
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

Public Sub btn_getImage(ByRef Control As IRibbonControl, _
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
    Case gsRXCTL_BTN_BUTTON
        sRtn = "MicrosoftExcel"
    Case Else
        sRtn = gsRXCTL_DFLT_IMAGE
    End Select

    ' ----------------------------------------------------------------------

    Image = sRtn

End Sub

Public Sub btn_getKeytip(ByRef Control As IRibbonControl, _
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

Public Sub btn_getLabel(ByRef Control As IRibbonControl, _
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
'    Case gsRXCTL_BTN_BUTTON4
'        sRtn = "Happy"
    Case Else
        If (Right$(Control.Id, 2) = "2B") Then
            sRtn = "Long Label"
        Else
            sRtn = Mid$(Control.Id, 6)
        End If
    End Select

    ' ----------------------------------------------------------------------

    Label = sRtn

End Sub

Public Sub btn_getScreentip(ByRef Control As IRibbonControl, _
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
    Case gsRXCTL_BTN_BUTTON
        sRtn = "My Screentip"
    Case Else
        sRtn = "Screentip for " & Control.Id
    End Select

    ' ----------------------------------------------------------------------

    Screentip = sRtn

End Sub

Public Sub btn_getShowImage(ByRef Control As IRibbonControl, _
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

Public Sub btn_getShowLabel(ByRef Control As IRibbonControl, _
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
'    Case gsRXCTL_BTN_BUTTON1
'        bRtn = False
'    Case gsRXCTL_BTN_BUTTON2
'        bRtn = False
'    Case gsRXCTL_BTN_BUTTON3
'        bRtn = False
    Case Else
        bRtn = True
    End Select

    ' ----------------------------------------------------------------------

    ShowLabel = bRtn

End Sub

Public Sub btn_getSize(ByRef Control As IRibbonControl, _
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
    Case gsRXCTL_BTN_BUTTON4
        sRtn = gsRXCTL_SIZE_LARGE
    Case Else
        sRtn = gsRXCTL_DFLT_SIZE
    End Select

    ' ----------------------------------------------------------------------

    Size = sRtn

End Sub

Public Sub btn_getSupertip(ByRef Control As IRibbonControl, _
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
    Case gsRXCTL_BTN_BUTTON
        sRtn = "My Supertip"
    Case Else
        sRtn = "Supertip for " & Control.Id
    End Select

    ' ----------------------------------------------------------------------

    Supertip = sRtn

End Sub

Public Sub btn_getVisible(ByRef Control As IRibbonControl, _
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

Public Sub btn_onAction(ByRef Control As IRibbonControl)
' ==========================================================================
' Description : Respond to an IRibbonControl action
'
' Parameters  : Control     The control initiating the callback
' ==========================================================================

    Select Case Control.Id
    Case Else
        If (Len(Control.Tag) > 0) Then
            Call MsgBox(Control.Id & vbNewLine & "Tag = " & Control.Tag, _
                        vbInformation Or vbOKOnly)
        Else
            Call MsgBox(Control.Id, vbInformation Or vbOKOnly)
        End If
    End Select

End Sub
