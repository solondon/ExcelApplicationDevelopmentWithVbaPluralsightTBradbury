Attribute VB_Name = "MWinAPIUser32Style"
' ==========================================================================
' Module      : MWinAPIUser32Style
' Type        : Module
' Description : Support for window style operations
' --------------------------------------------------------------------------
' Procedures  : GetWindowOpacity        Byte
'               SetWindowOpacity        Boolean
'               SetWindowStyle          Boolean
'               WindowStyleIsSet        Boolean
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

Private Const msMODULE As String = "MWinAPIUser32Style"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

' GetWindowLong indexes are described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633584(VS.85).aspx
' -----------------------------------
Public Enum enuGetWindowLongIndex
    GWL_WNDPROC = (-4)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_STYLE = (-16)
    GWL_EXSTYLE = (-20)
    GWL_USERDATA = (-21)
    GWL_ID = (-12)
End Enum

' LayeredWindowAttributes are described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633540(VS.85).aspx
' -----------------------------------
Public Enum enuLayeredWindowAttribute
    LWA_ALPHA = &O2
    LWA_COLORKEY = &O1
End Enum

' Window styles are described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms632600(VS.85).aspx
' -----------------------------------
Public Enum enuWindowStyle
    WS_OVERLAPPED = &H0&
    WS_POPUP = &H80000000
    WS_CHILD = &H40000000
    WS_MINIMIZE = &H20000000
    WS_VISIBLE = &H10000000
    WS_DISABLED = &H8000000
    WS_CLIPSIBLINGS = &H4000000
    WS_CLIPCHILDREN = &H2000000
    WS_MAXIMIZE = &H1000000
    WS_CAPTION = &HC00000               ' WS_BORDER Or WS_DLGFRAME
    WS_BORDER = &H800000
    WS_DLGFRAME = &H400000
    WS_VSCROLL = &H200000
    WS_HSCROLL = &H100000
    WS_SYSMENU = &H80000
    WS_THICKFRAME = &H40000
    WS_GROUP = &H20000
    WS_TABSTOP = &H10000

    WS_MINIMIZEBOX = &H20000
    WS_MAXIMIZEBOX = &H10000

    WS_TILED = &H0&                     ' WS_OVERLAPPED
    WS_ICONIC = &H20000000              ' WS_MINIMIZE
    WS_SIZEBOX = &H40000                ' WS_THICKFRAME

    WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
    WS_CHILDWINDOW = &H40000000         ' WS_CHILD

    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
End Enum

' Extended Window Styles are described on MSDN at
' http://msdn.microsoft.com/en-us/library/ff700543(v=VS.85).aspx
' -----------------------------------
Public Enum enuWindowStyleEx
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_TRANSPARENT = &H20&

    WS_EX_MDICHILD = &H40&
    WS_EX_TOOLWINDOW = &H80&
    WS_EX_WINDOWEDGE = &H100&
    WS_EX_CLIENTEDGE = &H200&
    WS_EX_CONTEXTHELP = &H400&

    WS_EX_RIGHT = &H1000&
    WS_EX_LEFT = &H0&
    WS_EX_RTLREADING = &H2000&
    WS_EX_LTRREADING = &H0&
    WS_EX_LEFTSCROLLBAR = &H4000&
    WS_EX_RIGHTSCROLLBAR = &H0&

    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000

    WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
    WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST

    WS_EX_LAYERED = &H80000

    WS_EX_NOINHERITLAYOUT = &H100000    ' Disable inheritence of mirroring by children
    WS_EX_LAYOUTRTL = &H400000          ' Right to left mirroring

    WS_EX_COMPOSITED = &H2000000

    WS_EX_NOACTIVATE = &H8000000
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------

' The GetLayeredWindowAttributes function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633508(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetLayeredWindowAttributes _
            Lib "user32.dll" (ByVal hWnd As LongPtr, _
                          ByRef pcrKey As Long, _
                          ByRef pbAlpha As Byte, _
                          ByRef pdwFlags As enuLayeredWindowAttribute) _
            As Boolean
#Else
    Private Declare _
            Function GetLayeredWindowAttributes _
            Lib "user32.dll" (ByVal hWnd As Long, _
                          ByRef pcrKey As Byte, _
                          ByRef pbAlpha As Byte, _
                          ByRef pdwFlags As enuLayeredWindowAttribute) _
            As Boolean
#End If

' The GetWindowLong function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633584(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetWindowLong _
            Lib "user32.dll" _
            Alias "GetWindowLongA" (ByVal hWnd As LongPtr, _
                                    ByVal nIndex As enuGetWindowLongIndex) _
            As Long
#Else
    Private Declare _
            Function GetWindowLong _
            Lib "user32.dll" _
            Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                    ByVal nIndex As enuGetWindowLongIndex) _
            As Long
#End If

' The SetLayeredWindowAttributes function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633540(VS.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SetLayeredWindowAttributes _
            Lib "user32.dll" (ByVal hWnd As LongPtr, _
                          ByVal crKey As Byte, _
                          ByVal bAlpha As Byte, _
                          ByVal dwFlags As enuLayeredWindowAttribute) _
            As Boolean
#Else
    Private Declare _
            Function SetLayeredWindowAttributes _
            Lib "user32.dll" (ByVal hWnd As Long, _
                          ByVal crKey As Byte, _
                          ByVal bAlpha As Byte, _
                          ByVal dwFlags As enuLayeredWindowAttribute) _
            As Boolean
#End If

' The SetWindowLong function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633591(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SetWindowLong _
            Lib "user32.dll" _
            Alias "SetWindowLongA" (ByVal hWnd As LongPtr, _
                                    ByVal nIndex As enuGetWindowLongIndex, _
                                    ByVal dwNewLong As Long) _
            As Long
#Else
    Private Declare _
            Function SetWindowLong _
            Lib "user32.dll" _
            Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                    ByVal nIndex As enuGetWindowLongIndex, _
                                    ByVal dwNewLong As Long) _
            As Long
#End If

#If VBA7 Then
Public Function GetWindowOpacity(ByVal hWnd As LongPtr) As Byte
#Else
Public Function GetWindowOpacity(ByVal hWnd As Long) As Byte
#End If
' ==========================================================================
' Description : Get the opacity (transparency) of a window.
'
' Parameters  : hWnd    The handle of the window
'
' Returns     : Byte    The transparency level of the form (255 = opaque)
' ==========================================================================

    Const sPROC     As String = "GetWindowOpacity"

    Dim bRtn        As Boolean
    Dim bytRtn      As Byte


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the opacity level
    ' ---------------------

    bRtn = GetLayeredWindowAttributes(hWnd, 0, bytRtn, LWA_ALPHA)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetWindowOpacity = bytRtn

    Call Trace(tlMaximum, msMODULE, sPROC, bytRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    bRtn = False

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

#If VBA7 Then
Public Function SetWindowOpacity(ByVal hWnd As LongPtr, _
                                 ByVal Opacity As Byte) As Boolean
#Else
Public Function SetWindowOpacity(ByVal hWnd As Long, _
                                 ByVal Opacity As Byte) As Boolean
#End If
' ==========================================================================
' Description : Set the opacity (transparency) of a window.
'
' Parameters  : hWnd        The handle of the window
'               Opacity     Specifies the level of opacity from
'                           0 (transparent) to 255 (opaque).
'
' Returns     : Boolean     Returns True if successful.
' ==========================================================================

    Const sPROC As String = "SetWindowOpacity"

    Dim bRtn    As Boolean

    Dim lRtn    As Long
    Dim lStyle  As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Quit if the handle is not available
    ' -----------------------------------
    If (hWnd = 0) Then
        GoTo PROC_EXIT
    End If

    ' Get the style bits
    ' ------------------
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (lStyle = 0) Then
        GoTo PROC_EXIT
    End If

    ' Prepare the window for opacity info
    ' Only a layered window can have opacity
    ' --------------------------------------
    lRtn = SetWindowLong(hWnd, GWL_EXSTYLE, lStyle Or WS_EX_LAYERED)
    If (lRtn = 0) Then
        Exit Function
    End If

    ' Set the opacity level
    ' ---------------------
    bRtn = SetLayeredWindowAttributes(hWnd, 0, Opacity, LWA_ALPHA)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    SetWindowOpacity = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, Opacity)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    bRtn = False

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

#If VBA7 Then
Public Function SetWindowStyle(ByVal hWnd As LongPtr, _
                               ByVal Style As enuWindowStyle, _
                      Optional ByVal SetStyle As Boolean = True) As Boolean
#Else
Public Function SetWindowStyle(ByVal hWnd As Long, _
                               ByVal Style As enuWindowStyle, _
                      Optional ByVal SetStyle As Boolean = True) As Boolean
#End If
' ==========================================================================
' Description : Sets the style bits in a windows long
'
' Parameters  : hWnd        The handle of the window
'               Style       The style bits to set
'               SetStyle    If True, then set the style bits
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "SetWindowStyle"

    Dim bRtn        As Boolean

    Dim lRtn        As Long
    Dim lStyle      As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Quit if the handle is not available
    ' -----------------------------------
    If (hWnd = 0) Then
        GoTo PROC_EXIT
    End If

    ' Get the window style
    ' --------------------
    lStyle = GetWindowLong(hWnd, GWL_STYLE)

    ' Set the style bits
    ' ------------------
    If SetStyle Then
        lStyle = (lStyle Or Style)
    Else
        lStyle = lStyle And (Not Style)
    End If

    ' Update the window style
    ' -----------------------
    lRtn = SetWindowLong(hWnd, GWL_STYLE, lStyle)

    bRtn = CBool(lRtn)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    SetWindowStyle = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, bRtn)
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
Public Function WindowStyleIsSet(ByVal hWnd As LongPtr, _
                                 ByVal Style As enuWindowStyle) As Boolean
#Else
Public Function WindowStyleIsSet(ByVal hWnd As Long, _
                                 ByVal Style As enuWindowStyle) As Boolean
#End If
' ==========================================================================
' Description : Determines if a style is set in the windows long
'
' Parameters  : hWnd        The handle of the window
'               Style       The window style to test for
'
' Returns     : Boolean     Returns True if the style is set
' ==========================================================================

    Const sPROC As String = "WindowStyleIsSet"

    Dim bRtn    As Boolean
    Dim lStyle  As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Quit if the handle is not available
    ' -----------------------------------
    If (hWnd = 0) Then
        GoTo PROC_EXIT
    End If

    ' Get the window style
    ' --------------------
    lStyle = GetWindowLong(hWnd, GWL_STYLE)

    ' Get the style bits
    ' ------------------
    If (lStyle And Style) Then
        bRtn = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WindowStyleIsSet = bRtn

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
