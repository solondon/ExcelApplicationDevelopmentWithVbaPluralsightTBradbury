VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNWWorkbookEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' Module      : CNWWorkbookEvents
' Type        : Class Module
' Description : Event handler for the workbook
' --------------------------------------------------------------------------
' Events      : Class_Initialize
'               Class_Terminate
'               Workbook_SheetActivate
'               Workbook_SheetDeactivate
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

Private Const msMODULE As String = "CNWWorkbookEvents"

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private WithEvents Workbook As Excel.Workbook
Attribute Workbook.VB_VarHelpID = -1

Private Sub Class_Initialize()
' ==========================================================================

    Set Workbook = ThisWorkbook

End Sub

Private Sub Class_Terminate()
' ==========================================================================

    Set Workbook = Nothing

End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
' ==========================================================================
' Description : Process a sheet activation
'
' Parameters  : Sh    The newly activated sheet
' ==========================================================================

    Const sPROC As String = "Workbook_SheetActivate"


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Sh.CodeName)

    ' ----------------------------------------------------------------------

    If (TypeName(Sh) = gsTYPENAME_EXCEL_WKS) Then
'        Call OnEntryStart(Sh)
'        Call OnKeyStart(Sh)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
' ==========================================================================
' Description : Remove sheet-specific processing
'
' Parameters  : Sh    The worksheet that is losing focus
' ==========================================================================

    Const sPROC As String = "Workbook_SheetDeactivate"


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Sh.CodeName)

    ' ----------------------------------------------------------------------

    If (TypeName(Sh) = gsTYPENAME_EXCEL_WKS) Then
'        Call OnEntryStop(Sh)
'        Call OnKeyStop(Sh)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub
