VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNWApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' Module      : CNWApp
' Type        : Class Module
' Description : App object for the application
' --------------------------------------------------------------------------
' Properties  : Processing          (Get)   Boolean
'               Processing          (Let)   Boolean
'               Ribbon              (Get)   IRibbonUI
'               Ribbon              (Set)   IRibbonUI
'               Version             (Get)   CVersion
' --------------------------------------------------------------------------
' Events      : Class_Initialize
'               Class_Terminate
' --------------------------------------------------------------------------
' Dependencies: CNWWorkbookEvents
'               CVersion
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE      As String = "CNWApp"

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private mwkbEvents          As CNWWorkbookEvents

Private mb_Processing       As Boolean
Private mo_Ribbon           As IRibbonUI
Private mo_Version          As CVersion

Public Property Get Processing() As Boolean
' ==========================================================================

    Processing = mb_Processing

End Property

Public Property Let Processing(ByVal bProcessing As Boolean)
' ==========================================================================

    mb_Processing = bProcessing

End Property

Public Property Get Ribbon() As IRibbonUI
' ==========================================================================

    Set Ribbon = mo_Ribbon

End Property

Public Property Set Ribbon(Ribbon As IRibbonUI)
' ==========================================================================

    Set mo_Ribbon = Ribbon

End Property

Public Property Get Version() As CVersion
' ==========================================================================

    Set Version = mo_Version

End Property

Private Sub Class_Initialize()
' ==========================================================================

    Const sPROC As String = "Class_Initialize"

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    Set mo_Version = New CVersion
    Set mwkbEvents = New CNWWorkbookEvents

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)

End Sub

Private Sub Class_Terminate()
' ==========================================================================

    Const sPROC As String = "Class_Terminate"

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    Set mo_Ribbon = Nothing
    Set mo_Version = Nothing
    Set mwkbEvents = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)

End Sub
