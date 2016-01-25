Attribute VB_Name = "MRibbonDmn"
' ==========================================================================
' Module      : MRibbonDmn
' Type        : Module
' Description : Support for the IRibbonControl dynamicMenu
' --------------------------------------------------------------------------
' Procedures  : dmn_getContent
'               dmn_getDescription
'               dmn_getEnabled
'               dmn_getImage
'               dmn_getKeytip
'               dmn_getLabel
'               dmn_getScreentip
'               dmn_getShowImage
'               dmn_getShowLabel
'               dmn_getSize
'               dmn_getSupertip
'               dmn_getVisible
' --------------------------------------------------------------------------
' Comments    : The default settings will populate
'               a dynamicMenu with a unique instance of a button.
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

Private Const msMODULE As String = "MRibbonDmn"

Public Sub dmn_getContent(ByRef Control As IRibbonControl, _
                          ByRef Content As Variant)
' ==========================================================================
' Description : Get the XML content for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Content     Returns the content for the control
' ==========================================================================

    Static lsCnt    As Long     ' Ensure this instance is unique

    Dim sXML        As String

    ' ----------------------------------------------------------------------
    ' Build the standard wrapper
    ' --------------------------

    sXML = "<menu xmlns=" & Chr(34)

    If gbXMLNS_CUSTOMUI_ONLY Then
        sXML = sXML & gsXMLNS_CUSTOMUI

    ElseIf (Application.Version = OfficeVersion2007) Then
        sXML = sXML & gsXMLNS_CUSTOMUI

    Else
        sXML = sXML & gsXMLNS_CUSTOMUI14
    End If

    sXML = sXML & Chr(34) & ">" & vbNewLine
    
    ' ----------------------------------------------------------------------
    
    Select Case Control.Id
    Case Else
        lsCnt = lsCnt + 1

        sXML = sXML _
             & "  <button" & vbNewLine _
             & "    id=" & Chr(34) & "rxbtnDmnButton" & lsCnt & Chr(34) & vbNewLine _
             & "    getDescription=" & Chr(34) & "btn_getDescription" & Chr(34) & vbNewLine _
             & "    getEnabled=" & Chr(34) & "btn_getEnabled" & Chr(34) & vbNewLine _
             & "    getImage=" & Chr(34) & "btn_getImage" & Chr(34) & vbNewLine _
             & "    getKeytip=" & Chr(34) & "btn_getKeytip" & Chr(34) & vbNewLine _
             & "    getLabel=" & Chr(34) & "btn_getLabel" & Chr(34) & vbNewLine _
             & "    getScreentip=" & Chr(34) & "btn_getScreentip" & Chr(34) & vbNewLine _
             & "    getShowImage=" & Chr(34) & "btn_getShowImage" & Chr(34) & vbNewLine _
             & "    getShowLabel=" & Chr(34) & "btn_getShowLabel" & Chr(34) & vbNewLine _
             & "    getSupertip=" & Chr(34) & "btn_getSupertip" & Chr(34) & vbNewLine _
             & "    getVisible=" & Chr(34) & "btn_getVisible" & Chr(34) & vbNewLine _
             & "    onAction=" & Chr(34) & "btn_onAction" & Chr(34) & " />" & vbNewLine

    End Select

    ' ----------------------------------------------------------------------
    
    sXML = sXML & "</menu>"
    Debug.Print sXML

    ' ----------------------------------------------------------------------

    Content = sXML

End Sub

Public Sub dmn_getDescription(ByRef Control As IRibbonControl, _
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

Public Sub dmn_getEnabled(ByRef Control As IRibbonControl, _
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

Public Sub dmn_getImage(ByRef Control As IRibbonControl, _
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

Public Sub dmn_getKeytip(ByRef Control As IRibbonControl, _
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

Public Sub dmn_getLabel(ByRef Control As IRibbonControl, _
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

Public Sub dmn_getScreentip(ByRef Control As IRibbonControl, _
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

Public Sub dmn_getShowImage(ByRef Control As IRibbonControl, _
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

Public Sub dmn_getShowLabel(ByRef Control As IRibbonControl, _
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

Public Sub dmn_getSize(ByRef Control As IRibbonControl, _
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
        sRtn = gsRXCTL_SIZE_LARGE
    End Select

    ' ----------------------------------------------------------------------

    Size = sRtn

End Sub

Public Sub dmn_getSupertip(ByRef Control As IRibbonControl, _
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

Public Sub dmn_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
Attribute dmn_getVisible.VB_Description = "Get the visibility for an IRibbonControl"
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
