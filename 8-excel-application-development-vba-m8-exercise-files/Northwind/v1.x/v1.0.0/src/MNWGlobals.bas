Attribute VB_Name = "MNWGlobals"
' ==========================================================================
' Module      : MNWGlobals
' Type        : Module
' Description : Global functions, constants and variables
' --------------------------------------------------------------------------
' Procedures  : AppStart
'               AppStop
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

Public Const gsAPP_CODE             As String = "NWDEMO"
Public Const gsAPP_NAME             As String = "Northwind Demo"
Public Const gsVBA_PROJ             As String = "NWDemo"

Public Const gsAPP_EMAIL            As String = "yourusername@example.com"

Public Const gsAPP_WKBNM_DEFAULT    As String = "Northwind.xlsm"

Public Const gsAPP_COMPANY          As String = "Pluralsight"

' ----------------
' Module Level
' ----------------

Private Const msMODULE              As String = "MNWGlobals"

' -----------------------------------
' Variable declarations
' -----------------------------------
' Global Level
' ----------------

Public goApp                        As CNWApp

Public Sub AppStart(Optional ByVal Initializing As Boolean)
' ==========================================================================
' Description : Manually start the application
' ==========================================================================

    Const sPROC As String = "AppStart"

    Dim bSaved  As Boolean: bSaved = ThisWorkbook.Saved


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Initializing)

    ' ----------------------------------------------------------------------
    ' Start the global app object
    ' ---------------------------

    If (goApp Is Nothing) Then
        Call Trace(tlVerbose, msMODULE, sPROC, "Creating app object")
        Set goApp = New CNWApp
        DoEvents
    End If

    If Initializing Then
        goApp.Processing = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Call WorkbookSavedHasChanged(bSaved)

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

Public Sub AppStop()
' ==========================================================================
' Description : Manually stop the application
' ==========================================================================

    Const sPROC As String = "AppStop"

    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Stop the global app object
    ' --------------------------
    Set goApp = Nothing

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
