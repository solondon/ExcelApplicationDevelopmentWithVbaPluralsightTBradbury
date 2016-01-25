Attribute VB_Name = "MNWRibbonBtn"
' ==========================================================================
' Module      : M_RibbonBtn
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

Private Const msMODULE                  As String = "M_RibbonBtn"

Private Const msRXID_BTNCUSTOMERS       As String = "rxbtnCustomers"
Private Const msRXID_BTNEMPLOYEES       As String = "rxbtnEmployees"
Private Const msRXID_BTNORDERS          As String = "rxbtnOrders"
Private Const msRXID_BTNORDERDETAILS    As String = "rxbtnOrderDetails"
Private Const msRXID_BTNPRODUCTS        As String = "rxbtnProducts"
Private Const msRXID_BTNCATEGORIES      As String = "rxbtnCategories"
Private Const msRXID_BTNSUPPLIERS       As String = "rxbtnSuppliers"
Private Const msRXID_BTNSHIPPERS        As String = "rxbtnShippers"
Private Const msRXID_BTNCOUNTRIES       As String = "rxbtnCountries"
Private Const msRXID_BTNREGIONS         As String = "rxbtnRegions"

Private Const msRXID_BTNABOUT           As String = "rxbtnAbout"

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
    Case Else
        sRtn = Control.Id
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
    Case msRXID_BTNCUSTOMERS
        Call ChangeSubject(gsNW_WKSCN_CUSTOMERS)
        Call NW_CustomerEdit
    Case msRXID_BTNEMPLOYEES
        Call ChangeSubject(gsNW_WKSCN_EMPLOYEES)
    Case msRXID_BTNORDERS
        Call ChangeSubject(gsNW_WKSCN_ORDERS)
    Case msRXID_BTNORDERDETAILS
        Call ChangeSubject(gsNW_WKSCN_ORDERDETAILS)
    Case msRXID_BTNPRODUCTS
        Call ChangeSubject(gsNW_WKSCN_PRODUCTS)
    Case msRXID_BTNCATEGORIES
        Call ChangeSubject(gsNW_WKSCN_CATEGORIES)
    Case msRXID_BTNSUPPLIERS
        Call ChangeSubject(gsNW_WKSCN_SUPPLIERS)
    Case msRXID_BTNSHIPPERS
        Call ChangeSubject(gsNW_WKSCN_SHIPPERS)
    Case msRXID_BTNCOUNTRIES
        Call ChangeSubject(gsNW_WKSCN_COUNTRIES)
    Case msRXID_BTNREGIONS
        Call ChangeSubject(gsNW_WKSCN_REGIONS)

    Case msRXID_BTNABOUT
        Call NW_About
    Case Else
        Call MsgBox(Control.Id, vbInformation Or vbOKOnly)
    End Select

End Sub
