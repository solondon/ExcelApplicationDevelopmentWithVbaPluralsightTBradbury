Attribute VB_Name = "MWinAPIUser32"
' ==========================================================================
' Module      : MWinAPIUser32
' Type        : Module
' Description : Support for window API functions
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

Private Const msMODULE  As String = "MWinAPIUser32"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

' Load resource flags are defined in WinUser.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms648045(v=vs.85).aspx
' -----------------------------------
Public Enum enuLoadResourceFlag
    LR_DEFAULTCOLOR = &H0
    LR_MONOCHROME = &H1
    LR_COLOR = &H2
    LR_COPYRETURNORG = &H4
    LR_COPYDELETEORG = &H8
    LR_LOADFROMFILE = &H10
    LR_LOADTRANSPARENT = &H20
    LR_DEFAULTSIZE = &H40
    LR_VGACOLOR = &H80
    LR_LOADMAP3DCOLORS = &H1000
    LR_CREATEDIBSECTION = &H2000
    LR_COPYFROMRESOURCE = &H4000
    LR_SHARED = &H8000
End Enum

' Menu flags are defined in WinUser.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms647636(v=vs.85).aspx
' -----------------------------------
Public Enum enuMenuFlag
    MF_INSERT = &H0
    MF_CHANGE = &H80
    MF_APPEND = &H100
    MF_DELETE = &H200
    MF_REMOVE = &H1000

    MF_BYCOMMAND = &H0
    MF_BYPOSITION = &H400

    MF_SEPARATOR = &H800

    MF_ENABLED = &H0
    MF_GRAYED = &H1
    MF_DISABLED = &H2

    MF_UNCHECKED = &H0
    MF_CHECKED = &H8
    MF_USECHECKBITMAPS = &H200

    MF_STRING = &H0
    MF_BITMAP = &H4
    MF_OWNERDRAW = &H100

    MF_POPUP = &H10
    MF_MENUBARBREAK = &H20
    MF_MENUBREAK = &H40

    MF_UNHILITE = &H0
    MF_HILITE = &H80

    MF_DEFAULT = &H1000
    
    MF_SYSMENU = &H2000
    MF_HELP = &H4000
    MF_RIGHTJUSTIFY = &H4000
    MF_MOUSESELECT = &H8000
    
    MF_END = &H80   ' Obsolete -- only used by old RES files

    MFT_STRING = MF_STRING
    MFT_BITMAP = MF_BITMAP
    MFT_MENUBARBREAK = MF_MENUBARBREAK
    MFT_MENUBREAK = MF_MENUBREAK
    MFT_OWNERDRAW = MF_OWNERDRAW
    MFT_RADIOCHECK = &H200
    MFT_SEPARATOR = MF_SEPARATOR
    MFT_RIGHTORDER = &H2000
    MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY

    MFS_GRAYED = &H3
    MFS_DISABLED = MFS_GRAYED
    MFS_CHECKED = MF_CHECKED
    MFS_HILITE = MF_HILITE
    MFS_ENABLED = MF_ENABLED
    MFS_UNCHECKED = MF_UNCHECKED
    MFS_UNHILITE = MF_UNHILITE
    MFS_DEFAULT = MF_DEFAULT
End Enum

' ShowWindow commands are defined in WinUser.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633548(v=vs.85).aspx
' -----------------------------------
Public Enum enuShowWindowCommand
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_FORCEMINIMIZE = 11
    SW_MAX = 11
End Enum

' Window messages are defined in WinUser.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ff381405(v=vs.85).aspx
' -----------------------------------
Public Enum enuWindowMessage
    WM_NULL = &H0
    WM_CREATE = &H1
    WM_DESTROY = &H2
    WM_MOVE = &H3
    WM_SIZE = &H5

    WM_ACTIVATE = &H6

    WM_SETFOCUS = &H7
    WM_KILLFOCUS = &H8
    WM_ENABLE = &HA
    WM_SETREDRAW = &HB
    WM_SETTEXT = &HC
    WM_GETTEXT = &HD
    WM_GETTEXTLENGTH = &HE
    WM_PAINT = &HF
    WM_CLOSE = &H10

    WM_QUERYENDSESSION = &H11
    WM_QUERYOPEN = &H13
    WM_ENDSESSION = &H16

    WM_QUIT = &H12
    WM_ERASEBKGND = &H14
    WM_SYSCOLORCHANGE = &H15
    WM_SHOWWINDOW = &H18
    WM_WININICHANGE = &H1A

    WM_DEVMODECHANGE = &H1B
    WM_ACTIVATEAPP = &H1C
    WM_FONTCHANGE = &H1D
    WM_TIMECHANGE = &H1E
    WM_CANCELMODE = &H1F
    WM_SETCURSOR = &H20
    WM_MOUSEACTIVATE = &H21
    WM_CHILDACTIVATE = &H22
    WM_QUEUESYNC = &H23

    WM_GETMINMAXINFO = &H24

    WM_PAINTICON = &H26
    WM_ICONERASEBKGND = &H27
    WM_NEXTDLGCTL = &H28
    WM_SPOOLERSTATUS = &H2A
    WM_DRAWITEM = &H2B
    WM_MEASUREITEM = &H2C
    WM_DELETEITEM = &H2D
    WM_VKEYTOITEM = &H2E
    WM_CHARTOITEM = &H2F
    WM_SETFONT = &H30
    WM_GETFONT = &H31
    WM_SETHOTKEY = &H32
    WM_GETHOTKEY = &H33
    WM_QUERYDRAGICON = &H37
    WM_COMPAREITEM = &H39

    WM_GETOBJECT = &H3D

    WM_COMPACTING = &H41
    WM_COMMNOTIFY = &H44    ' no longer suported
    WM_WINDOWPOSCHANGING = &H46
    WM_WINDOWPOSCHANGED = &H47

    WM_POWER = &H48

    WM_COPYDATA = &H4A
    WM_CANCELJOURNAL = &H4B

    WM_NOTIFY = &H4E
    WM_INPUTLANGCHANGEREQUEST = &H50
    WM_INPUTLANGCHANGE = &H51
    WM_TCARD = &H52
    WM_HELP = &H53
    WM_USERCHANGED = &H54
    WM_NOTIFYFORMAT = &H55

    WM_CONTEXTMENU = &H7B
    WM_STYLECHANGING = &H7C
    WM_STYLECHANGED = &H7D
    WM_DISPLAYCHANGE = &H7E
    WM_GETICON = &H7F
    WM_SETICON = &H80

    WM_NCCREATE = &H81
    WM_NCDESTROY = &H82
    WM_NCCALCSIZE = &H83
    WM_NCHITTEST = &H84
    WM_NCPAINT = &H85
    WM_NCACTIVATE = &H86
    WM_GETDLGCODE = &H87

    WM_SYNCPAINT = &H88

    WM_NCMOUSEMOVE = &HA0
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCLBUTTONDBLCLK = &HA3
    WM_NCRBUTTONDOWN = &HA4
    WM_NCRBUTTONUP = &HA5
    WM_NCRBUTTONDBLCLK = &HA6
    WM_NCMBUTTONDOWN = &HA7
    WM_NCMBUTTONUP = &HA8
    WM_NCMBUTTONDBLCLK = &HA9
    WM_NCXBUTTONDOWN = &HAB
    WM_NCXBUTTONUP = &HAC
    WM_NCXBUTTONDBLCLK = &HAD

    WM_INPUT = &HFF

    WM_KEYFIRST = &H100
    WM_KEYDOWN = &H100
    WM_KEYUP = &H101
    WM_CHAR = &H102
    WM_DEADCHAR = &H103
    WM_SYSKEYDOWN = &H104
    WM_SYSKEYUP = &H105
    WM_SYSCHAR = &H106
    WM_SYSDEADCHAR = &H107

    WM_UNICHAR = &H109
    WM_KEYLAST = &H109
    UNICODE_NOCHAR = &HFFFF

    WM_IME_STARTCOMPOSITION = &H10D
    WM_IME_ENDCOMPOSITION = &H10E
    WM_IME_COMPOSITION = &H10F
    WM_IME_KEYLAST = &H10F

    WM_INITDIALOG = &H110
    WM_COMMAND = &H111
    WM_SYSCOMMAND = &H112
    WM_TIMER = &H113
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
    WM_INITMENU = &H116
    WM_INITMENUPOPUP = &H117
    WM_MENUSELECT = &H11F
    WM_MENUCHAR = &H120
    WM_ENTERIDLE = &H121

    WM_MENURBUTTONUP = &H122
    WM_MENUDRAG = &H123
    WM_MENUGETOBJECT = &H124
    WM_UNINITMENUPOPUP = &H125
    WM_MENUCOMMAND = &H126

    WM_CHANGEUISTATE = &H127
    WM_UPDATEUISTATE = &H128
    WM_QUERYUISTATE = &H129

    WM_CTLCOLORMSGBOX = &H132
    WM_CTLCOLOREDIT = &H133
    WM_CTLCOLORLISTBOX = &H134
    WM_CTLCOLORBTN = &H135
    WM_CTLCOLORDLG = &H136
    WM_CTLCOLORSCROLLBAR = &H137
    WM_CTLCOLORSTATIC = &H138

    WM_MOUSEFIRST = &H200
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_LBUTTONDBLCLK = &H203
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RBUTTONDBLCLK = &H206
    WM_MBUTTONDOWN = &H207
    WM_MBUTTONUP = &H208
    WM_MBUTTONDBLCLK = &H209
    WM_MOUSEWHEEL = &H20A
    WM_XBUTTONDOWN = &H20B
    WM_XBUTTONUP = &H20C
    WM_XBUTTONDBLCLK = &H20D

    WM_CUT = &H300
    WM_COPY = &H301
    WM_PASTE = &H302
    WM_CLEAR = &H303
    WM_UNDO = &H304
    WM_RENDERFORMAT = &H305
    WM_RENDERALLFORMATS = &H306
    WM_DESTROYCLIPBOARD = &H307
    WM_DRAWCLIPBOARD = &H308
    WM_PAINTCLIPBOARD = &H309
    WM_VSCROLLCLIPBOARD = &H30A
    WM_SIZECLIPBOARD = &H30B
    WM_ASKCBFORMATNAME = &H30C
    WM_CHANGECBCHAIN = &H30D
    WM_HSCROLLCLIPBOARD = &H30E
    WM_QUERYNEWPALETTE = &H30F
    WM_PALETTEISCHANGING = &H310
    WM_PALETTECHANGED = &H311
    WM_HOTKEY = &H312

    WM_USER = &H400
End Enum
