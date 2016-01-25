Attribute VB_Name = "MRibbonCbo"
' ==========================================================================
' Module      : MRibbonCbo
' Type        : Module
' Description : Support for the IRibbonControl comboBox
' --------------------------------------------------------------------------
' Callbacks   : cbo_getEnabled
'               cbo_getImage
'               cbo_getItemCount
'               cbo_getItemId
'               cbo_getItemImage
'               cbo_getItemLabel
'               cbo_getItemScreentip
'               cbo_getItemSupertip
'               cbo_getKeytip
'               cbo_getLabel
'               cbo_getScreentip
'               cbo_getShowImage
'               cbo_getShowLabel
'               cbo_getSupertip
'               cbo_getText
'               cbo_getVisible
'               cbo_onChange
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

Private Const msMODULE As String = "MRibbonCbo"

Public Sub cbo_getEnabled(ByRef Control As IRibbonControl, _
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

Public Sub cbo_getImage(ByRef Control As IRibbonControl, _
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

Public Sub cbo_getItemCount(ByRef Control As IRibbonControl, _
                            ByRef ItemCount As Variant)
' ==========================================================================
' Description : Get the item count for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               ItemCount   Returns the item count for the control
' ==========================================================================

    Dim iRtn    As Integer  ' Cannot be > 1000

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        iRtn = 5
    End Select

    ' ----------------------------------------------------------------------

    ItemCount = iRtn

End Sub

Public Sub cbo_getItemId(ByRef Control As IRibbonControl, _
                         ByRef Index As Integer, _
                         ByRef ItemId As Variant)
' ==========================================================================
' Description : Get the id for an IRibbonControl item
'
' Parameters  : Control       The control initiating the callback
'               Index         The index position of the item
'               ItemId        Returns the id to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Item" & Right$("00" & CStr(Index), 3)
    End Select

    ' ----------------------------------------------------------------------

    ItemId = sRtn

End Sub

Public Sub cbo_getItemImage(ByRef Control As IRibbonControl, _
                            ByRef Index As Integer, _
                            ByRef ItemImage As Variant)
' ==========================================================================
' Description : Get the image for an IRibbonControl item
'
' Parameters  : Control       The ribbon control to be affected
'               Index         The index position of the item
'               ItemImage     Returns the image to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = gsRXCTL_DFLT_IMAGE
    End Select

    ' ----------------------------------------------------------------------

    ItemImage = sRtn

End Sub

Public Sub cbo_getItemLabel(ByRef Control As IRibbonControl, _
                            ByRef Index As Integer, _
                            ByRef ItemLabel As Variant)
' ==========================================================================
' Description : Get the label for an IRibbonControl item
'
' Parameters  : Control       The ribbon control to be affected
'               Index         The index position of the item
'               ItemLabel     Returns the label to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Item" & Right$("00" & CStr(Index), 3)
    End Select

    ' ----------------------------------------------------------------------

    ItemLabel = sRtn

End Sub

Public Sub cbo_getItemScreentip(ByRef Control As IRibbonControl, _
                                ByRef Index As Integer, _
                                ByRef ItemScreentip As Variant)
' ==========================================================================
' Description : Get the screentip for an IRibbonControl item
'
' Parameters  : Control       The ribbon control to be affected
'               Index         The index position of the item
'               Screentip     Returns the screentip to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Screentip for Item " & CStr(Index)
    End Select

    ' ----------------------------------------------------------------------

    ItemScreentip = sRtn

End Sub

Public Sub cbo_getItemSupertip(ByRef Control As IRibbonControl, _
                               ByRef Index As Integer, _
                               ByRef ItemSupertip As Variant)
' ==========================================================================
' Description : Get the supertip for an IRibbonControl item
'
' Parameters  : Control       The ribbon control to be affected
'               Index         The index position of the item
'               ItemSupertip  Returns the supertip to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Supertip for Item " & CStr(Index)
    End Select

    ' ----------------------------------------------------------------------

    ItemSupertip = sRtn

End Sub

Public Sub cbo_getKeytip(ByRef Control As IRibbonControl, _
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

Public Sub cbo_getLabel(ByRef Control As IRibbonControl, _
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

Public Sub cbo_getScreentip(ByRef Control As IRibbonControl, _
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
        sRtn = Control.Id
    End Select

    ' ----------------------------------------------------------------------

    Screentip = sRtn

End Sub

Public Sub cbo_getShowImage(ByRef Control As IRibbonControl, _
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

Public Sub cbo_getShowLabel(ByRef Control As IRibbonControl, _
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

Public Sub cbo_getSupertip(ByRef Control As IRibbonControl, _
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

Sub cbo_getText(ByRef Control As IRibbonControl, _
                ByRef Text As Variant)
' ==========================================================================
' Description : Get the text for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Text        Returns the text for the control
' ==========================================================================

    Dim sRtn    As String

    Select Case Control.Id
    Case Else
        sRtn = "Item000"
    End Select

    Text = sRtn

End Sub

Public Sub cbo_getVisible(ByRef Control As IRibbonControl, _
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

Public Sub cbo_onChange(ByRef Control As IRibbonControl, _
                        ByRef Text As String)
' ==========================================================================
' Description : Respond to an IRibbonControl action
'
' Parameters  : Control     The control initiating the callback
'               Text        The new value for the control
' ==========================================================================

    Select Case Control.Id
    Case Else
        MsgBox Text, vbInformation Or vbOKOnly, Control.Id
    End Select

End Sub
