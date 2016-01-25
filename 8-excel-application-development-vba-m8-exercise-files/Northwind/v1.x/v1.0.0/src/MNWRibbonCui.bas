Attribute VB_Name = "MNWRibbonCui"
' ==========================================================================
' Module      : MNWRibbonCui
' Type        : Module
' Description : Support for the IRibbonUI interface
' --------------------------------------------------------------------------
' Procedures  : cui_onLoad
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

Private Const msMODULE As String = "MNWRibbonCui"

Public Sub cui_onLoad(ByRef Ribbon As IRibbonUI)
' ==========================================================================
' Description : Store a reference to the ribbon
'
' Parameters  : Ribbon    The Ribbon object
' ==========================================================================

    ' Create the App object
    ' ---------------------
    If (goApp Is Nothing) Then
        Set goApp = New CNWApp
    End If

    Set goApp.Ribbon = Ribbon
    Call ActivateRibbonTab("rxtabNorthwind")

End Sub
