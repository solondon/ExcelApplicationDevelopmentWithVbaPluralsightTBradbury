Attribute VB_Name = "MWinAPIUser32Window"
' ==========================================================================
' Module      : MWinAPIUser32Window
' Type        : Module
' Description : Support for window operations
' --------------------------------------------------------------------------
' Procedures  : CloseWindowByHandle     Boolean
'               EnableWindow            Boolean
'               GetClassName            String
'               GetFocusHandle          LongPtr
'               GetWindowHandle         LongPtr
'               GetWindowText           String
'               IsValidWindowHandle     Boolean
'               SetTopmostWindow        Boolean
'               SetTopWindow            Boolean
'               ShowWindowByHandle      Boolean
'               WindowIsEnabled         Boolean
'               WindowIsMaximized       Boolean
'               WindowIsMinimized       Boolean
' --------------------------------------------------------------------------
' Comments    : Most procedures in this module depend on the window handle.
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

Private Const msMODULE      As String = "MWinAPIUser32Window"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

' ----------------
' Module Level
' ----------------

Private Enum enuhWndInsertAfterFlags
    HWND_BOTTOM = 1
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
End Enum

Private Enum enuSetWindowPosFlags
    SWP_ASYNCWINDOWPOS = &H4000
    SWP_DEFERERASE = &H2000
    SWP_DRAWFRAME = &H20
    SWP_FRAMECHANGED = &H20
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = &H200
    SWP_NOSENDCHANGING = &H400
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The EnableWindow function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms646291(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function EnablehWindow _
            Lib "User32" _
            Alias "EnableWindow" (ByVal hWnd As LongPtr, _
                                  ByVal bEnable As Boolean) _
            As Boolean
#Else
    Private Declare _
            Function EnablehWindow _
            Lib "User32" _
            Alias "EnableWindow" (ByVal hWnd As Long, _
                                  ByVal bEnable As Boolean) _
            As Boolean
#End If

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

' The GetClassNameStr function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633582(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetClassNameStr _
            Lib "User32" _
            Alias "GetClassNameA" (ByVal hWnd As LongPtr, _
                                   ByVal lpClassName As String, _
                                   ByVal nMaxcount As Long) _
            As LongPtr
#Else
    Private Declare _
            Function GetClassNameStr _
            Lib "User32" _
            Alias "GetClassNameA" (ByVal hWnd As Long, _
                                   ByVal lpClassName As String, _
                                   ByVal nMaxcount As Long) _
            As Long
#End If

' The GetCurrentProcessId function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms683180(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetCurrentProcessId _
            Lib "Kernel32" () _
            As LongPtr
#Else
    Private Declare _
            Function GetCurrentProcessId _
            Lib "Kernel32" () _
            As Long
#End If

' The GetDesktopWindow function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633504(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetDesktopWindow _
            Lib "User32" () _
            As LongPtr
#Else
    Private Declare _
            Function GetDesktopWindow _
            Lib "User32" () _
            As Long
#End If

' The GetFocusHWnd function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms646294(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetFocusHWnd _
            Lib "User32" _
            Alias "GetFocus" () _
            As LongPtr
#Else
    Private Declare _
            Function GetFocusHWnd _
            Lib "User32" _
            Alias "GetFocus" () _
            As Long
#End If

' The GetWindowText function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633520(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetWindowTxt _
            Lib "User32" _
            Alias "GetWindowTextA" (ByVal hWnd As LongPtr, _
                                    ByVal lpString As String, _
                                    ByVal nMaxcount As Long) _
            As LongPtr
#Else
    Private Declare _
            Function GetWindowTxt _
            Lib "User32" _
            Alias "GetWindowTextA" (ByVal hWnd As Long, _
                                    ByVal lpString As String, _
                                    ByVal nMaxcount As Long) _
            As Long
#End If

' The GetWindowThreadProcessId function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633522(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetWindowThreadProcessId _
            Lib "User32" (ByVal hWnd As LongPtr, _
                          ByRef lpdwProcessId As LongPtr) _
            As LongPtr
#Else
    Private Declare _
            Function GetWindowThreadProcessId _
            Lib "User32" (ByVal hWnd As Long, _
                          ByRef lpdwProcessId As Long) _
            As Long
#End If

' The IsIconic function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633527(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare _
            Function IsIconic _
            Lib "User32" (ByVal hWnd As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function IsIconic _
            Lib "User32" (ByVal hWnd As Long) _
            As Boolean
#End If

' The IsWindow function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633528(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function IsWindow _
            Lib "User32" (ByVal hWnd As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function IsWindow _
            Lib "User32" (ByVal hWnd As Long) _
            As Boolean
#End If

' The IsWindowEnabled function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms646303(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function IsWindowEnabled _
            Lib "User32" (ByVal hWnd As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function IsWindowEnabled _
            Lib "User32" (ByVal hWnd As Long) _
            As Boolean
#End If

' The IsZoomed function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633531(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare _
            Function IsZoomed _
            Lib "User32" (ByVal hWnd As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function IsZoomed _
            Lib "User32" (ByVal hWnd As Long) _
            As Boolean
#End If

' The PostMessage function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms644944(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function PostMessage _
            Lib "User32" _
            Alias "PostMessageA" (ByVal hWnd As LongPtr, _
                                  ByVal wMsg As enuWindowMessage, _
                                  ByVal wParam As Long, _
                                  ByVal lParam As Long) _
            As Long
#Else
    Private Declare _
            Function PostMessage _
            Lib "User32" _
            Alias "PostMessageA" (ByVal hWnd As Long, _
                                  ByVal wMsg As enuWindowMessage, _
                                  ByVal wParam As Long, _
                                  ByVal lParam As Long) _
            As Long
#End If

' The SetWindowPos function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633545(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SetWindowPos _
            Lib "User32" (ByVal hWnd As LongPtr, _
                          ByVal hWndInsertAfter As LongPtr, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal cx As Long, _
                          ByVal cy As Long, _
                          ByVal uFlags As enuSetWindowPosFlags) _
            As Boolean
#Else
    Private Declare _
            Function SetWindowPos _
            Lib "User32" (ByVal hWnd As Long, _
                          ByVal hWndInsertAfter As Long, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal cx As Long, _
                          ByVal cy As Long, _
                          ByVal uFlags As enuSetWindowPosFlags) _
            As Boolean
#End If

' The ShowWindow function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633548(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function ShowWindow _
            Lib "User32" (ByVal hWnd As LongPtr, _
                          ByVal nCmdShow As enuShowWindowCommand) _
            As Boolean
#Else
    Private Declare _
            Function ShowWindow _
            Lib "User32" (ByVal hWnd As Long, _
                          ByVal nCmdShow As enuShowWindowCommand) _
            As Boolean
#End If

#If VBA7 Then
Public Function CloseWindowByHandle(ByVal hWnd As LongPtr) As Boolean
#Else
Public Function CloseWindowByHandle(ByVal hWnd As Long) As Boolean
#End If
' ==========================================================================
' Description : Send a close message to a window
'
' Parameters  : hWnd        The handle of the window
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "CloseWindowByHandle"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If IsWindow(hWnd) Then
        bRtn = PostMessage(hWnd, WM_CLOSE, 0&, 0&)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    CloseWindowByHandle = bRtn

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

#If VBA7 Then
Public Function EnableWindow(ByVal hWnd As LongPtr, _
                    Optional ByVal Enable As Boolean = True) As Boolean
#Else
Public Function EnableWindow(ByVal hWnd As Long, _
                    Optional ByVal Enable As Boolean = True) As Boolean
#End If
' ==========================================================================
' Description : Enables or disables mouse and keyboard
'               input to the specified window.
'
' Parameters  : hWnd        Handle to the window to be enabled or disabled.
'               Enable      Enable or disable the window.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "EnableWindow"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = EnablehWindow(hWnd, Enable)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    EnableWindow = bRtn

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

#If VBA7 Then
Public Function GetClassName(ByVal hWnd As LongPtr) As String
#Else
Public Function GetClassName(ByVal hWnd As Long) As String
#End If
' ==========================================================================
' Description : Retrieves the name of the window class.
'
' Parameters  : hWnd        The handle of the window
'
' Returns     : String
' ==========================================================================

    Const sPROC         As String = "GetClassName"
    Const lBUFFER_SIZE  As Long = 255

    Dim lRtn            As Long: lRtn = lBUFFER_SIZE
    Dim sRtn            As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ------------------------------------------------------------------------

    sRtn = String$(lRtn, vbNullChar)
    lRtn = GetClassNameStr(hWnd, sRtn, lRtn)
    sRtn = Left$(sRtn, lRtn)

    ' ------------------------------------------------------------------------

PROC_EXIT:

    GetClassName = sRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

' ------------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

#If VBA7 Then
Public Function GetFocusHandle() As LongPtr
#Else
Public Function GetFocusHandle() As Long
#End If
' ==========================================================================
' Description : Returns the handle of the currently focused control.
'
' Returns     : Long
' ==========================================================================

    Const sPROC     As String = "GetFocusHandle"

    #If VBA7 Then
        Dim lRtn    As LongPtr
    #Else
        Dim lRtn    As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ------------------------------------------------------------------------

    lRtn = GetFocusHWnd()

    ' ------------------------------------------------------------------------

PROC_EXIT:

    GetFocusHandle = lRtn

    Call Trace(tlMaximum, msMODULE, sPROC, lRtn)
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

#If VBA7 Then
Public Function _
       GetWindowHandle(Optional ByVal Class As String, _
                       Optional ByVal Caption As String) As LongPtr
#Else
Public Function _
       GetWindowHandle(Optional ByVal Class As String, _
                       Optional ByVal Caption As String) As Long
#End If
' ==========================================================================
' Description : Get the handle of a Window or Form
'
' Parameters  : Class       The ClassName of the Window
'               Caption     The title of the window
'
' Returns     : LongPtr     The handle of the window
' ==========================================================================

    Const sPROC     As String = "GetWindowHandle"

    #If VBA7 Then
        Dim hWnd    As LongPtr
        Dim hWndDesktop As LongPtr
        Dim hProcApp As LongPtr
        Dim hProcWin As LongPtr
    #Else
        Dim hWnd    As Long
        Dim hWndDesktop As Long
        Dim hProcApp As Long
        Dim hProcWin As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_ENTER)

    ' ------------------------------------------------------------------------
    ' All top-level windows are children of the desktop
    ' -------------------------------------------------
    hWndDesktop = GetDesktopWindow()

    ' Get the proc ID of this instance
    ' --------------------------------
    hProcApp = GetCurrentProcessId()


    ' Find the child window of the desktop that
    ' matches the window class and/or caption.
    ' hWnd will be 0 the first time, so get the first match.
    ' Each loop will pass the handle of the window from the
    ' previous pass, and wil eventually find the window handle
    ' --------------------------------------------------------
    Do
        hWnd = FindWindowEx(hWndDesktop, hWnd, Class, Caption)

        ' Get the process ID of the window owner
        ' --------------------------------------
        Call GetWindowThreadProcessId(hWnd, hProcWin)

        ' Loop until the window process matches the app process
        ' -----------------------------------------------------
    Loop Until ((hProcWin = hProcApp) Or (hWnd = 0))

    ' ------------------------------------------------------------------------

PROC_EXIT:

    GetWindowHandle = hWnd

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

#If VBA7 Then
Public Function GetWindowText(ByVal hWnd As LongPtr) As String
#Else
Public Function GetWindowText(ByVal hWnd As Long) As String
#End If
' ==========================================================================
' Description : Retrieves the caption for the window.
'
' Parameters  : hWnd        The handle of the window
'
' Returns     : String
' ==========================================================================

    Const sPROC         As String = "GetWindowText"
    Const lBUFFER_SIZE  As Long = 255

    Dim lRtn            As Long: lRtn = lBUFFER_SIZE
    Dim sRtn            As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    sRtn = String$(lRtn, vbNullChar)
    lRtn = GetWindowTxt(hWnd, sRtn, lRtn)
    sRtn = Left$(sRtn, lRtn)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetWindowText = sRtn

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

#If VBA7 Then
Public Function IsValidWindowHandle(ByVal hWnd As LongPtr) As Boolean
#Else
Public Function IsValidWindowHandle(ByVal hWnd As Long) As Boolean
#End If
' ==========================================================================
' Description : Determines whether the specified window handle
'               identifies an existing window.
'
' Parameters  : hWnd    The window handle to use
'
' Returns     : Boolean Returns True if the handle is a valid window
' ==========================================================================

    Const sPROC As String = "IsValidWindowHandle"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = IsWindow(hWnd)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsValidWindowHandle = bRtn

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

#If VBA7 Then
Public Function SetTopmostWindow(ByVal hWnd As LongPtr, _
                                 ByVal Topmost As Boolean) As Boolean
#Else
Public Function SetTopmostWindow(ByVal hWnd As Long, _
                                 ByVal Topmost As Boolean) As Boolean
#End If
' ==========================================================================
' Description : Set or clear a window as topmost (front of z-order)
'
' Parameters  : hWnd        The handle of the window
'               Topmost     If True, set as the topmost window
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "SetTopmostWindow"

    Dim bRtn    As Boolean

    
    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If Topmost Then
        bRtn = SetWindowPos(hWnd, _
                            HWND_TOPMOST, _
                            0, 0, 0, 0, _
                            SWP_NOMOVE Or SWP_NOSIZE)
    Else
        bRtn = SetWindowPos(hWnd, _
                            HWND_NOTOPMOST, _
                            0, 0, 0, 0, _
                            SWP_NOMOVE Or SWP_NOSIZE)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    SetTopmostWindow = bRtn
    
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

#If VBA7 Then
Public Function SetTopWindow(ByVal hWnd As LongPtr, _
                             ByVal Top As Boolean) As Boolean
#Else
Public Function SetTopWindow(ByVal hWnd As Long, _
                             ByVal Top As Boolean) As Boolean
#End If
' ==========================================================================
' Description : Set or clear a window as top (front of z-order)
'
' Parameters  : hWnd        The handle of the window
'               Top         If True, set as a top-level window
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "SetTopWindow"

    Dim bRtn    As Boolean

    
    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If Top Then
        bRtn = SetWindowPos(hWnd, _
                            HWND_TOP, _
                            0, 0, 0, 0, _
                            SWP_NOMOVE Or SWP_NOSIZE)
    Else
        bRtn = SetWindowPos(hWnd, _
                            HWND_NOTOPMOST, _
                            0, 0, 0, 0, _
                            SWP_NOMOVE Or SWP_NOSIZE)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    SetTopWindow = bRtn
    
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

#If VBA7 Then
Public Function ShowWindowByHandle(ByVal hWnd As LongPtr, _
                                   ByVal Cmd As enuShowWindowCommand) _
       As Boolean
#Else
Public Function ShowWindowByHandle(ByVal hWnd As Long, _
                                   ByVal Cmd As enuShowWindowCommand) _
       As Boolean
#End If
' ==========================================================================
' Description : Processes a ShowWindow command using the supplied handle
'
' Parameters  : hWnd    The window handle to use
'               Cmd     The command to execute
'
' Returns     : Boolean If the window was previously hidden,
'                       the return value is False
' ==========================================================================

    Const sPROC As String = "ShowWindowByHandle"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = ShowWindow(hWnd, Cmd)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ShowWindowByHandle = bRtn

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

#If VBA7 Then
    Public Function WindowIsEnabled(ByVal hWnd As LongPtr) As Boolean
#Else
    Public Function WindowIsEnabled(ByVal hWnd As Long) As Boolean
#End If
' ==========================================================================
' Description : Determines if a window is enabled
'
' Parameters  : hWnd        The handle of the window
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "WindowIsEnabled"

    Dim bRtn        As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (hWnd = 0) Then
        GoTo PROC_EXIT
    End If
    
    bRtn = IsWindowEnabled(hWnd)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WindowIsEnabled = bRtn

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

#If VBA7 Then
    Public Function WindowIsMaximized(ByVal hWnd As LongPtr) As Boolean
#Else
    Public Function WindowIsMaximized(ByVal hWnd As Long) As Boolean
#End If
' ==========================================================================
' Description : Determines if a window is maximized
'
' Parameters  : hWnd        The handle of the window
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "WindowIsMaximized"

    Dim bRtn        As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (hWnd = 0) Then
        GoTo PROC_EXIT
    End If
    
    bRtn = IsZoomed(hWnd)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WindowIsMaximized = bRtn

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

#If VBA7 Then
    Public Function WindowIsMinimized(ByVal hWnd As LongPtr) As Boolean
#Else
    Public Function WindowIsMinimized(ByVal hWnd As Long) As Boolean
#End If
' ==========================================================================
' Description : Determines if a window is minimized
'
' Parameters  : hWnd    The handle of the window
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "WindowIsMinimized"

    Dim bRtn        As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (hWnd = 0) Then
        GoTo PROC_EXIT
    End If
    
    bRtn = IsIconic(hWnd)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WindowIsMinimized = bRtn

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
