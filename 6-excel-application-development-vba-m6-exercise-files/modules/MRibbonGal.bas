Attribute VB_Name = "MRibbonGal"
' ==========================================================================
' Module      : MRibbonGal
' Type        : Module
' Description : Support for the IRibbonControl gallery
' --------------------------------------------------------------------------
' Procedures  : gal_getDescription
'               gal_getEnabled
'               gal_getImage
'               gal_getItemCount
'               gal_getItemHeight
'               gal_getItemId
'               gal_getItemImage
'               gal_getItemLabel
'               gal_getItemScreentip
'               gal_getItemSupertip
'               gal_getItemWidth
'               gal_getKeytip
'               gal_getLabel
'               gal_getScreentip
'               gal_getSelectedItemIndex
'               gal_getShowImage
'               gal_getShowLabel
'               gal_getSize
'               gal_getSupertip
'               gal_getVisible
'               gal_onAction
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

Private Const msMODULE As String = "MRibbonGal"

Public Sub gal_getDescription(ByRef Control As IRibbonControl, _
                              ByRef Description As Variant)
Attribute gal_getDescription.VB_Description = "Get the description for an IRibbonControl"
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

Public Sub gal_getEnabled(ByRef Control As IRibbonControl, _
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

Public Sub gal_getImage(ByRef Control As IRibbonControl, _
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

Public Sub gal_getItemCount(ByRef Control As IRibbonControl, _
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
        iRtn = 12
    End Select

    ' ----------------------------------------------------------------------

    ItemCount = iRtn

End Sub

Public Sub gal_getItemHeight(ByRef Control As IRibbonControl, _
                             ByRef ItemHeight As Variant)
' ==========================================================================
' Description : Returns the height for an IRibbonControl item
'
' Parameters  : Control     The control initiating the callback
'               Height      Returns the height for the item
' ==========================================================================

    Dim lRtn    As Long

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        lRtn = glRXITEM_DFLT_HEIGHT
    End Select

    ' ----------------------------------------------------------------------

    ItemHeight = lRtn

End Sub

Public Sub gal_getItemId(ByRef Control As IRibbonControl, _
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
        sRtn = "Item" & Right$("0" & CStr(Index), 2)
    End Select

    ' ----------------------------------------------------------------------

    ItemId = sRtn

End Sub

Public Sub gal_getItemImage(ByRef Control As IRibbonControl, _
                            ByRef Index As Integer, _
                            ByRef ItemImage As Variant)
' ==========================================================================
' Description : Determine the image for a ribbon item
'
' Parameters  : Control       The control initiating the callback
'               Index         The index position of the item
'               ItemImage     The image to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case "rxgalGallery"
        Select Case Index
            Case 0 To 2
                sRtn = "MicrosoftAccess"
            Case 3 To 5
                sRtn = "MicrosoftExcel"
            Case 6 To 8
                sRtn = "MicrosoftPowerPoint"
            Case 9 To 11
                sRtn = "MicrosoftOutlook"
        End Select
    Case Else
        sRtn = gsRXCTL_DFLT_IMAGE
    End Select

    ' ----------------------------------------------------------------------

    ItemImage = sRtn

End Sub

Public Sub gal_getItemLabel(ByRef Control As IRibbonControl, _
                            ByRef Index As Integer, _
                            ByRef Label As Variant)
' ==========================================================================
' Description : Determine the label for a ribbon item
'
' Parameters  : Control       The control initiating the callback
'               Index         The index position of the item
'               Label         The label to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = MonthName(Index + 1)
    End Select

    ' ----------------------------------------------------------------------

    Label = sRtn

End Sub

Public Sub gal_getItemScreentip(ByRef Control As IRibbonControl, _
                                ByRef Index As Integer, _
                                ByRef Screentip As Variant)
' ==========================================================================
' Description : Determine the screentip for a ribbon item
'
' Parameters  : Control       The control initiating the callback
'               Index         The index position of the item
'               Screentip     The screentip to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Screentip for item " & CStr(Index)
    End Select

    ' ----------------------------------------------------------------------

    Screentip = sRtn

End Sub

Public Sub gal_getItemSupertip(ByRef Control As IRibbonControl, _
                               ByRef Index As Integer, _
                               ByRef Supertip As Variant)
' ==========================================================================
' Description : Determine the screentip for a ribbon item
'
' Parameters  : Control       The ribbon control to be affected
'               Index         The index position of the item
'               Supertip      The supertip to assign to the item
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        sRtn = "Supertip for item " & CStr(Index)
    End Select

    ' ----------------------------------------------------------------------

    Supertip = sRtn

End Sub

Public Sub gal_getItemWidth(ByRef Control As IRibbonControl, _
                            ByRef Width As Variant)
' ==========================================================================
' Description : Provide the width for an item in the control
'
' Parameters  : Control     The control initiating the callback
'               Width       Returns the width for the items
' ==========================================================================

    Dim lRtn    As Long

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case Else
        lRtn = glRXITEM_DFLT_WIDTH
    End Select

    ' ----------------------------------------------------------------------

    Width = lRtn

End Sub

Public Sub gal_getKeytip(ByRef Control As IRibbonControl, _
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

Public Sub gal_getLabel(ByRef Control As IRibbonControl, _
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

Public Sub gal_getScreentip(ByRef Control As IRibbonControl, _
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

Sub gal_getSelectedItemIndex(ByRef Control As IRibbonControl, _
                             ByRef SelectedItemIndex As Variant)
' ==========================================================================
' Description : Respond to a ribbon control action
'
' Parameters  : Control     The control initiating the callback
'               ItemIndex   The new value for the control
' ==========================================================================

    Dim lRtn As Long

    Select Case Control.Id
    Case Else
        lRtn = Month(Now) - 1
    End Select

    SelectedItemIndex = lRtn

End Sub

Public Sub gal_getShowImage(ByRef Control As IRibbonControl, _
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

Public Sub gal_getShowLabel(ByRef Control As IRibbonControl, _
                            ByRef ShowLabel As Variant)
' ==========================================================================
' Description : Dynamically set the visibility for an IRibbonControl label
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

Public Sub gal_getSize(ByRef Control As IRibbonControl, _
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

Public Sub gal_getSupertip(ByRef Control As IRibbonControl, _
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

Public Sub gal_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
Attribute gal_getVisible.VB_Description = "Get the visibility for an IRibbonControl"
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

Public Sub gal_onAction(ByRef Control As IRibbonControl, _
                        ByRef Id As String, _
                        ByRef Index As Integer)
' ==========================================================================
' Description : Respond to a ribbon control action
'
' Parameters  : Control     The control initiating the callback
'               id          The caller control id
'               index       The gallery index
' ==========================================================================

    Select Case Control.Id
    Case Else
        Call MsgBox("Id = " & Id & vbNewLine & "Index = " & CStr(Index), _
                    vbInformation Or vbOKOnly, _
                    Control.Id)
    End Select

End Sub
