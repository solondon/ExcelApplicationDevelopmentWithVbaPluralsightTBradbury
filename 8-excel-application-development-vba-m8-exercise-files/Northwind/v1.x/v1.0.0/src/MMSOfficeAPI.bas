Attribute VB_Name = "MMSOfficeAPI"
' ==========================================================================
' Module      : MMSOfficeAPI
' Type        : Module
' Description : Support for working with Office APIs
' --------------------------------------------------------------------------
' Procedures  : WindowHandle        LongPtr
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

Private Const msMODULE As String = "MMSOfficeAPI"

' -----------------------------------
' External Function declarations
' -----------------------------------

' The FindWindowEx function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633500(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function FindWindowEx _
            Lib "User32" _
            Alias "FindWindowExA" (ByVal hWndParent As LongPtr, _
                                   ByVal hWndChildAfter As LongPtr, _
                                   ByVal lpszClass As String, _
                                   ByVal lpszWindow As String) _
            As LongPtr
#Else
    Private Declare _
            Function FindWindowEx _
            Lib "User32" _
            Alias "FindWindowExA" (ByVal hWndParent As Long, _
                                   ByVal hWndChildAfter As Long, _
                                   ByVal lpszClass As String, _
                                   ByVal lpszWindow As String) _
            As Long
#End If

#If VBA7 Then
Public Function WindowHandle(ByRef Win As Window) As LongPtr
#Else
Public Function WindowHandle(ByRef Win As Window) As Long
#End If
' ==========================================================================
' Description : Returns the handle of the specified window.
'
' Parameters  : Win         The window to query
'
' Returns     : Long        The hWnd of the window
'
' Comments    : WindowHandle (for Office windows) works differently than
'               GetWindowHandle (for standard Windows).
' ==========================================================================

    Const sPROC         As String = "WindowHandle"

    #If VBA7 Then
        Dim hWnd        As LongPtr
        Dim hWndApp     As LongPtr
        Dim hWndDesktop As LongPtr
    #Else
        Dim hWnd        As Long
        Dim hWndApp     As Long
        Dim hWndDesktop As Long
    #End If

    Dim sCaption        As String


    On Error GoTo PROC_ERR
    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the handle of the application
    ' ---------------------------------
    Select Case Application.Name
    Case gsOFFICE_APPNAME_EXCEL
        hWndApp = Application.hWnd
        
        ' Get the handle of the desktop
        ' -----------------------------
        hWndDesktop = FindWindowEx(hWndApp, _
                                   0&, _
                                   gsCLASSNAME_EXCEL_DESKTOP, _
                                   vbNullString)

        If (hWndDesktop > 0) Then
            'sCaption = GetWindowCaption(Win)
            sCaption = Win.Caption
            hWnd = FindWindowEx(hWndDesktop, _
                                0&, _
                                gsCLASSNAME_EXCEL_WINDOW, _
                                sCaption)
        End If

    Case Else
        GoTo PROC_EXIT
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WindowHandle = hWnd

    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_EXIT)
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
