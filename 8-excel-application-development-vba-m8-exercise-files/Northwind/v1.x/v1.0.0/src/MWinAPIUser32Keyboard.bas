Attribute VB_Name = "MWinAPIUser32Keyboard"
' ==========================================================================
' Module      : MWinAPIUser32Keyboard
' Type        : Module
' Description : Support for keyboard emulation
' --------------------------------------------------------------------------
' Procedures  : SendVirtualKey
' --------------------------------------------------------------------------
' Dependencies: MVBALogic
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

Private Const msMODULE                  As String = "MWinAPIUser32Keyboard"

Private Const KEYEVENTF_EXTENDEDKEY     As Long = &H1
Private Const KEYEVENTF_KEYUP           As Long = &H2


' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

' Virtual-Key codes are defined in WinUser.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd375731(v=vs.85).aspx
' -----------------------------------
Public Enum enuVirtualKeyCode
    VK_LBUTTON = &H1
    VK_RBUTTON = &H2
    VK_CANCEL = &H3
    VK_MBUTTON = &H4
    VK_XBUTTON1 = &H5
    VK_XBUTTON2 = &H6
    VK_BACK = &H8
    VK_TAB = &H9
    VK_CLEAR = &HC
    VK_RETURN = &HD
    VK_SHIFT = &H10
    VK_CONTROL = &H11
    VK_MENU = &H12
    VK_PAUSE = &H13
    VK_CAPITAL = &H14
    VK_KANA = &H15
    VK_JUNJA = &H17
    VK_FINAL = &H18
    VK_HANJA = &H19
    VK_KANJI = &H19
    VK_ESCAPE = &H1B
    VK_CONVERT = &H1C
    VK_NONCONVERT = &H1D
    VK_ACCEPT = &H1E
    VK_MODECHANGE = &H1F
    VK_SPACE = &H20
    VK_PRIOR = &H21
    VK_NEXT = &H22
    VK_END = &H23
    VK_HOME = &H24
    VK_LEFT = &H25
    VK_UP = &H26
    VK_RIGHT = &H27
    VK_DOWN = &H28
    VK_SELECT = &H29
    VK_PRINT = &H2A
    VK_EXECUTE = &H2B
    VK_SNAPSHOT = &H2C
    VK_INSERT = &H2D
    VK_DELETE = &H2E
    VK_HELP = &H2F

    VK_0 = &H30
    VK_1 = &H31
    VK_2 = &H32
    VK_3 = &H33
    VK_4 = &H34
    VK_5 = &H35
    VK_6 = &H36
    VK_7 = &H37
    VK_8 = &H38
    VK_9 = &H39

    VK_A = &H41
    VK_B = &H42
    VK_C = &H43
    VK_D = &H44
    VK_E = &H45
    VK_F = &H46
    VK_G = &H47
    VK_H = &H48
    VK_I = &H49
    VK_J = &H4A
    VK_K = &H4B
    VK_L = &H4C
    VK_M = &H4D
    VK_N = &H4E
    VK_O = &H4F
    VK_P = &H50
    VK_Q = &H51
    VK_R = &H52
    VK_S = &H53
    VK_T = &H54
    VK_U = &H55
    VK_V = &H56
    VK_W = &H57
    VK_X = &H58
    VK_Y = &H59
    VK_Z = &H5A

    VK_LWIN = &H5B
    VK_RWIN = &H5C
    VK_APPS = &H5D
    VK_SLEEP = &H5F

    VK_NUMPAD0 = &H60
    VK_NUMPAD1 = &H61
    VK_NUMPAD2 = &H62
    VK_NUMPAD3 = &H63
    VK_NUMPAD4 = &H64
    VK_NUMPAD5 = &H65
    VK_NUMPAD6 = &H66
    VK_NUMPAD7 = &H67
    VK_NUMPAD8 = &H68
    VK_NUMPAD9 = &H69

    VK_MULTIPLY = &H6A
    VK_ADD = &H6B
    VK_SEPARATOR = &H6C
    VK_SUBTRACT = &H6D
    VK_DECIMAL = &H6E
    VK_DIVIDE = &H6F

    VK_F1 = &H70
    VK_F2 = &H71
    VK_F3 = &H72
    VK_F4 = &H73
    VK_F5 = &H74
    VK_F6 = &H75
    VK_F7 = &H76
    VK_F8 = &H77
    VK_F9 = &H78
    VK_F10 = &H79
    VK_F11 = &H7A
    VK_F12 = &H7B
    VK_F13 = &H7C
    VK_F14 = &H7D
    VK_F15 = &H7E
    VK_F16 = &H7F
    VK_F17 = &H80
    VK_F18 = &H81
    VK_F19 = &H82
    VK_F20 = &H83
    VK_F21 = &H84
    VK_F22 = &H85
    VK_F23 = &H86
    VK_F24 = &H87

    VK_LSHIFT = &HA0
    VK_RSHIFT = &HA1
    VK_LCONTROL = &HA2
    VK_RCONTROL = &HA3
    VK_LMENU = &HA4
    VK_RMENU = &HA5

    VK_BROWSER_BACK = &HA6
    VK_BROWSER_FORWARD = &HA7
    VK_BROWSER_REFRESH = &HA8
    VK_BROWSER_STOP = &HA9
    VK_BROWSER_SEARCH = &HAA
    VK_BROWSER_FAVORITES = &HAB
    VK_BROWSER_HOME = &HAC

    VK_VOLUME_MUTE = &HAD
    VK_VOLUME_DOWN = &HAE
    VK_VOLUME_UP = &HAF

    VK_MEDIA_NEXT_TRACK = &HB0
    VK_MEDIA_PREV_TRACK = &HB1
    VK_MEDIA_STOP = &HB2
    VK_MEDIA_PLAY_PAUSE = &HB3

    VK_LAUNCH_MAIL = &HB4
    VK_LAUNCH_MEDIA_SELECT = &HB5
    VK_LAUNCH_APP1 = &HB6
    VK_LAUNCH_APP2 = &HB7
End Enum

' -----------------------------------
' Type declarations
' -----------------------------------
' Global Level
' ----------------

' The MOUSEINPUT structure is defined in WinUser.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms646273(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Public Type TMOUSEINPUT
        dx          As Long
        dy          As Long
        mouseData   As Long
        dwFlags     As Long
        time        As Long
        dwExtraInfo As LongPtr
    End Type
#Else
    Public Type TMOUSEINPUT
        dx          As Long
        dy          As Long
        mouseData   As Long
        dwFlags     As Long
        time        As Long
        dwExtraInfo As Long
    End Type
#End If

' The KEYBDINPUT structure is defined in WinUser.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms646271(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Public Type TKEYBDINPUT
        wVk         As Integer
        wScan       As Integer
        dwFlags     As Long
        time        As Long
        dwExtraInfo As LongPtr
    End Type
#Else
    Public Type TKEYBDINPUT
        wVk         As Integer
        wScan       As Integer
        dwFlags     As Long
        time        As Long
        dwExtraInfo As Long
    End Type
#End If

' The HARDWAREINPUT structure is defined in WinUser.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms646269(v=vs.85).aspx
' -----------------------------------
Public Type THARDWAREINPUT
    uMsg    As Long
    wPramL  As Integer
    wPramH  As Integer
End Type

' The INPUT structure is defined in WinUser.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms646270(v=vs.85).aspx
' -----------------------------------
Public Type TINPUT
    type    As Long
    mi      As TMOUSEINPUT
    ki      As TKEYBDINPUT
    hi      As THARDWAREINPUT
End Type

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The keybd_event function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms646304(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Sub keybd_event _
            Lib "user32.dll" (ByVal bVk As Byte, _
                          ByVal bScan As Byte, _
                          ByVal dwFlags As Long, _
                          ByVal dwExtraInfo As Long)
#Else
    Private Declare _
            Sub keybd_event _
            Lib "user32.dll" (ByVal bVk As Byte, _
                          ByVal bScan As Byte, _
                          ByVal dwFlags As Long, _
                          ByVal dwExtraInfo As Long)
#End If

' The SendInput function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms646310(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SendInput _
            Lib "user32.dll" (ByVal nInputs As Long, _
                          ByRef pInputs() As TINPUT, _
                          ByVal cbSize As Integer) _
            As Long
#Else
    Private Declare _
            Function SendInput _
            Lib "user32.dll" (ByVal nInputs As Long, _
                          ByRef pInputs() As TINPUT, _
                          ByVal cbSize As Integer) _
            As Long
#End If

Public Sub SendVirtualKey(ByVal VirtualKeyCode As enuVirtualKeyCode, _
                 Optional ByVal CombinationKey As enuVirtualKeyCode)
' ==========================================================================
' Description : Initiate a keyboard event
'
' Parameters  : VirtualKeyCode  The key to send
'               CombinationKey  An additional key to send (e.g., vbKeyA)
' ==========================================================================

    If ((VirtualKeyCode < 1) Or (VirtualKeyCode > 254)) Then
        GoTo PROC_EXIT
    End If

    Select Case VirtualKeyCode
    Case VK_LWIN, VK_RWIN
        If IsBetween(CombinationKey, VK_A, VK_Z) Then
            Call keybd_event(VirtualKeyCode, 0, KEYEVENTF_EXTENDEDKEY, 0)
            Call keybd_event(CombinationKey, 0, 0, 0)
            Call keybd_event(VirtualKeyCode, 0, KEYEVENTF_KEYUP, 0)
        Else
            Call keybd_event(VirtualKeyCode, 0, KEYEVENTF_EXTENDEDKEY, 0)
            Call keybd_event(VirtualKeyCode, 0, KEYEVENTF_KEYUP, 0)
        End If
    Case Else
        Call keybd_event(VirtualKeyCode, 0, KEYEVENTF_EXTENDEDKEY, 0)
        Call keybd_event(VirtualKeyCode, 0, KEYEVENTF_KEYUP, 0)
    End Select

PROC_EXIT:

    Exit Sub

End Sub
