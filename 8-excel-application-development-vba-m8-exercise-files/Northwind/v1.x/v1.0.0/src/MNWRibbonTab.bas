Attribute VB_Name = "MNWRibbonTab"
' ==========================================================================
' Module      : MNWRibbonTab
' Type        : Module
' Description : Support for the IRibbonControl tab
' --------------------------------------------------------------------------
' Procedures  : tab_getKeytip
'               tab_getLabel
'               tab_getVisible
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

Private Const msMODULE              As String = "MNWRibbonTab"

Private Const msRXID_TABNORTHWIND   As String = "rxtabNorthwind"

Public Sub tab_getKeytip(ByRef Control As IRibbonControl, _
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

Public Sub tab_getLabel(ByRef Control As IRibbonControl, _
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
    Case msRXID_TABNORTHWIND
        sRtn = "Northwind"
    Case Else
        sRtn = Control.Id
    End Select

    ' ----------------------------------------------------------------------
    ' Tabs for 2013 are capitalized
    ' -----------------------------

    If (Application.Version = OfficeVersion2013) Then
        sRtn = UCase(sRtn)
    End If

    Label = sRtn

End Sub

Public Sub tab_getVisible(ByRef Control As IRibbonControl, _
                          ByRef Visible As Variant)
Attribute tab_getVisible.VB_Description = "Get the visibility for an IRibbonControl"
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
