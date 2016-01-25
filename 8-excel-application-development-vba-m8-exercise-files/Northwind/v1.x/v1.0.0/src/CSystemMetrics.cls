VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSystemMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' Module      : CSystemMetrics
' Type        : Class Module
' Description : Encapsulates system metrics.
' --------------------------------------------------------------------------
' Properties  : CaptionHeight       (Get)   Long
'               DialogFrameHeight   (Get)   Long
'               DialogFrameWidth    (Get)   Long
'               IconHeight          (Get)   Long
'               IconWidth           (Get)   Long
'               Monitors            (Get)   Long
'               MouseButtons        (Get)   Long
'               MouseIsPresent      (Get)   Boolean
'               MouseWheelIsPresent (Get)   Boolean
'               ScreenHeight        (Get)   Long
'               ScreenWidth         (Get)   Long
'               ShowSounds          (Get)   Boolean
'               SmallCaptionHeight  (Get)   Long
'               VirtualScreenHeight (Get)   Long
'               VirtualScreenWidth  (Get)   Long
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

Private Const msMODULE                  As String = "CSystemMetrics"

Public Property Get CaptionHeight() As Long
' ==========================================================================

    CaptionHeight = GetSystemMetric(SM_CYCAPTION)

End Property

Public Property Get DialogFrameHeight() As Long
' ==========================================================================

    DialogFrameHeight = GetSystemMetric(SM_CYDLGFRAME)

End Property

Public Property Get DialogFrameWidth() As Long
' ==========================================================================

    DialogFrameWidth = GetSystemMetric(SM_CXDLGFRAME)

End Property

Public Property Get IconHeight() As Long
' ==========================================================================

    IconHeight = GetSystemMetric(SM_CYICON)

End Property

Public Property Get IconWidth() As Long
' ==========================================================================

    IconWidth = GetSystemMetric(SM_CXICON)

End Property

Public Property Get Monitors() As Long
' ==========================================================================

    Monitors = GetSystemMetric(SM_CMONITORS)

End Property

Public Property Get MouseButtons() As Long
' ==========================================================================

    MouseButtons = GetSystemMetric(SM_CMOUSEBUTTONS)

End Property

Public Property Get MouseIsPresent() As Boolean
' ==========================================================================

    MouseIsPresent = CBool(GetSystemMetric(SM_MOUSEPRESENT))

End Property

Public Property Get MouseWheelIsPresent() As Boolean
' ==========================================================================

    MouseWheelIsPresent = CBool(GetSystemMetric(SM_MOUSEWHEELPRESENT))

End Property

Public Property Get ScreenHeight() As Long
' ==========================================================================

    ScreenHeight = GetSystemMetric(SM_CYSCREEN)

End Property

Public Property Get ScreenWidth() As Long
' ==========================================================================

    ScreenWidth = GetSystemMetric(SM_CXSCREEN)

End Property

Public Property Get ShowSounds() As Boolean
' ==========================================================================

    ShowSounds = CBool(GetSystemMetric(SM_SHOWSOUNDS))

End Property

Public Property Get SmallCaptionHeight() As Long
' ==========================================================================

    SmallCaptionHeight = GetSystemMetric(SM_CYSMCAPTION)

End Property

Public Property Get VirtualScreenHeight() As Long
' ==========================================================================

    VirtualScreenHeight = GetSystemMetric(SM_CYVIRTUALSCREEN)

End Property

Public Property Get VirtualScreenWidth() As Long
' ==========================================================================

    VirtualScreenWidth = GetSystemMetric(SM_CXVIRTUALSCREEN)

End Property
