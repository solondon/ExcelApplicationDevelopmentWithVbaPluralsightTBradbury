Attribute VB_Name = "MRibbonCui"
' ==========================================================================
' Module      : MRibbonCui
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

Private Const msMODULE As String = "M_RibbonRbn"

Public Sub cui_onLoad(ByRef Ribbon As IRibbonUI)
' ==========================================================================
' Description : Store a reference to the ribbon
'
' Parameters  : Ribbon    The Ribbon object
' ==========================================================================

    Set goRibbon = Ribbon

End Sub
