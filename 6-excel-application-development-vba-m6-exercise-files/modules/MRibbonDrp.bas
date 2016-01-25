Attribute VB_Name = "MRibbonDrp"
' ==========================================================================
' Module      : MRibbonDrp
' Type        : Module
' Description : Support for the IRibbonControl dropDown
' --------------------------------------------------------------------------
' Callbacks   : drp_getEnabled
'               drp_getImage
'               drp_getItemCount
'               drp_getItemID
'               drp_getItemImage
'               drp_getItemLabel
'               drp_getItemScreentip
'               drp_getItemSupertip
'               drp_getKeytip
'               drp_getLabel
'               drp_getScreentip
'               drp_getSelectedItemId
'               drp_getSelectedItemIndex
'               drp_getShowImage
'               drp_getShowLabel
'               drp_getSupertip
'               drp_getVisible
'               drp_onAction
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

Private Const msMODULE As String = "MRibbonDrp"

Public Sub drp_getEnabled(ByRef Control As IRibbonControl, _
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

Public Sub drp_getImage(ByRef Control As IRibbonControl, _
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

Public Sub drp_getItemCount(ByRef Control As IRibbonControl, _
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

Public Sub drp_getItemId(ByRef Control As IRibbonControl, _
                         ByRef Index As Integer, _
                         ByRef ItemId As Variant)
' ==========================================================================
' Description : Get the id for an IRibbonControl item
'
' Parameters  : Control       The ribbon control to be affected
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

Public Sub drp_getItemImage(ByRef Control As IRibbonControl, _
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

Public Sub drp_getItemLabel(ByRef Control As IRibbonControl, _
                            ByRef Index As Integer, _
                            ByRef ItemLabel As Variant)
' ==========================================================================
' Description : Get the label for an IRibbonControl item
'
' Parameters  : Control     The ribbon control to be affected
'               Index       The index position of the item
'               ItemLabel   Returns the label to assign to the item
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

Public Sub drp_getItemScreentip(ByRef Control As IRibbonControl, _
                                ByRef Index As Integer, _
                                ByRef ItemScreentip As Variant)
' ==========================================================================
' Description : Get the screentip for an IRibbonControl item
'
' Parameters  : Control         The ribbon control to be affected
'               Index           The index position of the item
'               ItemScreentip   Returns the screentip to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Screentip for item " & CStr(Index)
    End Select

    ' ----------------------------------------------------------------------

    ItemScreentip = sRtn

End Sub

Public Sub drp_getItemSupertip(ByRef Control As IRibbonControl, _
                               ByRef Index As Integer, _
                               ByRef ItemSupertip As Variant)
' ==========================================================================
' Description : Get the supertip for an IRibbonControl item
'
' Parameters  : Control         The ribbon control to be affected
'               Index           The index position of the item
'               ItemSupertip    Returns the supertip to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Supertip for index " & CStr(Index)
    End Select

    ' ----------------------------------------------------------------------

    ItemSupertip = sRtn

End Sub

Public Sub drp_getKeytip(ByRef Control As IRibbonControl, _
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

Public Sub drp_getLabel(ByRef Control As IRibbonControl, _
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

Public Sub drp_getScreentip(ByRef Control As IRibbonControl, _
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

Public Sub drp_getSelectedItemId(ByRef Control As IRibbonControl, _
                                 ByRef SelectedItemId As Variant)
' ==========================================================================
' Description : Get the id for the selected IRibbonControl item
'
' Parameters  : Control         The control initiating the callback
'               SelectedItemId  Returns the id for the selected item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Item000"
    End Select

    ' ----------------------------------------------------------------------

    SelectedItemId = sRtn

End Sub

Public Sub drp_getSelectedItemIndex(ByRef Control As IRibbonControl, _
                                    ByRef SelectedItemIndex As Variant)
' ==========================================================================
' Description : Respond to a ribbon control action
'
' Parameters  : Control             The control initiating the callback
'               SelectedItemIndex   Returns the index for the selected item
' ==========================================================================

    Dim lRtn    As Long

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        lRtn = 0
    End Select

    ' ----------------------------------------------------------------------

    SelectedItemIndex = lRtn

End Sub

Public Sub drp_getShowImage(ByRef Control As IRibbonControl, _
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

Public Sub drp_getShowLabel(ByRef Control As IRibbonControl, _
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

Public Sub drp_getSupertip(ByRef Control As IRibbonControl, _
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

Public Sub drp_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
Attribute drp_getVisible.VB_Description = "Get the visibility for an IRibbonControl"
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

Public Sub drp_onAction(ByRef Control As IRibbonControl, _
                        ByRef SelectedId As String, _
                        ByRef SelectedIndex As Integer)
' ==========================================================================
' Description : Respond to an IRibbonControl action
'
' Parameters  : Control         The control initiating the callback
'               SelectedId      The id of the selected item
'               SelectedIndex   The index of the selected item
' ==========================================================================

    Dim sTitle      As String: sTitle = Control.Id
    Dim sPrompt     As String: sPrompt = "Index = " & CStr(SelectedIndex)
    Dim eButtons    As VbMsgBoxStyle: eButtons = vbInformation Or vbOKOnly
    Dim eMBR        As VbMsgBoxResult


    Select Case Control.Id
    Case Else
        If (SelectedId <> Control.Id) Then
            sPrompt = sPrompt & vbNewLine _
                    & "ID = " & SelectedId
        End If
        eMBR = MsgBox(sPrompt, eButtons, sTitle)
    End Select

End Sub
