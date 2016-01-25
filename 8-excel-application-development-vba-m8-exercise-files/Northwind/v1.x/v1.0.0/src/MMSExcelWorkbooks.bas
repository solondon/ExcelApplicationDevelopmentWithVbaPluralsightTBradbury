Attribute VB_Name = "MMSExcelWorkbooks"
' ==========================================================================
' Module      : MMSExcelWorkbooks
' Type        : Module
' Description : Support for working with Excel Workbooks
' --------------------------------------------------------------------------
' Procedures  : CloseOtherWorkbooks
'               GetWorkbook                      Excel.Workbook
'               WorkbookIsOpen                   Boolean
'               WorkbookSavedHasChanged          Boolean
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

Private Const msMODULE As String = "MMSExcelWorkbooks"

Public Sub CloseOtherWorkbooks(Optional ByVal SaveChanges As Boolean)
' ==========================================================================
' Description : Close all other workbooks
' ==========================================================================

    Const sPROC As String = "CloseOtherWorkbooks"

    Dim wkb     As Excel.Workbook


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (Application.Workbooks.Count = 1) Then
        GoTo PROC_EXIT
    End If

    For Each wkb In Application.Workbooks
        DoEvents

        If (wkb.FullName <> ThisWorkbook.FullName) Then
            If ((Not wkb.Saved) And SaveChanges) Then
                wkb.Save
            End If

            Call wkb.Close(False)
        End If

    Next wkb

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set wkb = Nothing

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

Public Function GetWorkbook(ByVal FullName As String, _
                   Optional ByVal UseName As Boolean) As Excel.Workbook
' ==========================================================================
' Description : Locate a workbook using the full name and path
'
' Parameters  : FullName    The full name and path of the workbook
'               UseName     Use the base name only
'
' Returns     : Excel.Workbook
' ==========================================================================

    Const sPROC As String = "GetWorkbook"

    Dim lRtn    As Long

    Dim wkb     As Excel.Workbook


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If UseName Then
        For Each wkb In Workbooks
            lRtn = StrComp(wkb.Name, FullName, vbTextCompare)
            If (lRtn = 0) Then
                Set GetWorkbook = wkb
                Exit For
            End If
        Next wkb

    Else
        For Each wkb In Workbooks
            lRtn = StrComp(wkb.FullName, FullName, vbTextCompare)
            If (lRtn = 0) Then
                Set GetWorkbook = wkb
                Exit For
            End If
        Next wkb
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set wkb = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function WorkbookIsOpen(ByVal FullName As String) As Boolean
' ==========================================================================
' Description : Determine if a workbook is already loaded
'
' Parameters  : FullName    The full path and name of the workbook
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "WorkbookIsOpen"

    Dim bRtn    As Boolean
    Dim lRtn    As Long

    Dim wkb     As Excel.Workbook


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    For Each wkb In Workbooks
        lRtn = StrComp(wkb.FullName, FullName, vbTextCompare)
        If (lRtn = 0) Then
            bRtn = True
            Exit For
        End If
    Next wkb

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WorkbookIsOpen = bRtn

    Set wkb = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function WorkbookOpen(ByVal FileName As String, _
                    Optional ByVal AllowMacros As Boolean, _
                    Optional ByVal ReadOnly As Boolean = True) _
       As Excel.Workbook
' ==========================================================================
' Description : Open a workbook from code.
'
' Parameters  : FileName    The name of the workbook to open
'               AllowMacros Whether to allow the workbook macros to run
'               ReadOnly    Opens the workbook in read only mode
'
' Returns     : Excel.Workbook
' ==========================================================================

    Const sPROC     As String = "WorkbookOpen"

    Dim sTitle      As String: sTitle = gsAPP_NAME
    Dim sPrompt     As String
    Dim eButtons    As VbMsgBoxStyle: eButtons = vbInformation Or vbOKOnly
    Dim eMBR        As VbMsgBoxResult

    Dim eAutoSec    As MsoAutomationSecurity

    Dim wkb         As Excel.Workbook
    Dim wkbAct      As Excel.Workbook

    Dim udtAppProps As TApplicationProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtAppProps)
    Application.ScreenUpdating = gbDEBUG_MODE
    Set wkbAct = ActiveWorkbook

    ' Validate the file
    ' -----------------
    If (Not FileExists(FileName)) Then
        sPrompt = "The specified file" & vbNewLine _
                & Chr(34) & FileName & Chr(34) & vbNewLine _
                & "does not exist."
        eMBR = MsgBox(sPrompt, eButtons, sTitle)
        GoTo PROC_EXIT
    End If

    ' Check if it is already loaded
    ' -----------------------------
    For Each wkb In Workbooks
        If (StrComp(wkb.FullName, FileName, vbTextCompare) = 0) Then
            GoTo PROC_EXIT
        End If
    Next wkb

    ' ----------------------------------------------------------------------
    ' Open the file
    ' -------------

    eAutoSec = Application.AutomationSecurity
    If (Not AllowMacros) Then
        Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Else
        Application.AutomationSecurity = msoAutomationSecurityByUI
    End If
    Set wkb = Workbooks.Open(FileName:=FileName, _
                             ReadOnly:=ReadOnly)
    Application.AutomationSecurity = eAutoSec

    wkbAct.Activate

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set WorkbookOpen = wkb

    Set wkb = Nothing
    Set wkbAct = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtAppProps)

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function WorkbookSavedHasChanged(ByVal OriginalState As Boolean, _
                               Optional ByVal Reset As Boolean = True) _
       As Boolean
' ==========================================================================
' Description : This is a utility for use during startup routines.
'               Use this to check if the Saved state of the Workbook
'               has changed during a procedure, and optionally reset.
'
' Parameters  : OriginalState   A boolean value, usually stored at the start
'                               of a procedure, containing the Saved
'                               attribute of the Workbook
'               Reset           Indicates if the Saved flag should be reset.
'                               This is optional, as it is almost always
'                               set to True.
'
' Returns     : Boolean         Returns True if the OriginalState flag
'                               (the value of ThisWorkbook.Saved) has
'                               changed before this routine was called.
' ==========================================================================

    Const sPROC As String = "WorkbookSavedHasChanged"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------

    bRtn = (OriginalState <> ThisWorkbook.Saved)

    If (bRtn And Reset) Then
        ThisWorkbook.Saved = OriginalState
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WorkbookSavedHasChanged = bRtn

    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function
