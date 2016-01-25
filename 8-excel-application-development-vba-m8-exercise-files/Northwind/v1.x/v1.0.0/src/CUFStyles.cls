VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUFStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' Module      : CUFStyles
' Type        : Class Module
' Description : Control visual styles for UserForms
' --------------------------------------------------------------------------
' Properties  : CloseButtonEnabled      (Let)   Boolean
'               CloseButtonVisible      (Get)   Boolean
'               CloseButtonVisible      (Let)   Boolean
'               hWnd                    (Get)   LongPtr
'               MaximizeButtonVisible   (Get)   Boolean
'               MaximizeButtonVisible   (Let)   Boolean
'               Maximized               (Get)   Boolean
'               MinimizeButtonVisible   (Get)   Boolean
'               MinimizeButtonVisible   (Let)   Boolean
'               Minimized               (Get)   Boolean
'               Opacity                 (Get)   Byte
'               Opacity                 (Let)   Byte
'               Resizable               (Get)   Boolean
'               Resizable               (Let)   Boolean
'               TitleBarHeight          (Get)   Single
'               TitleBarVisible         (Get)   Boolean
'               TitleBarVisible         (Let)   Boolean
'               UserForm                (Set)   MSForms.UserForm
' --------------------------------------------------------------------------
' Procedures  : SetResizerVisible
' --------------------------------------------------------------------------
' Events      : OnUserFormChanged
'               mfrm_UserForm_Layout
'               Class_Initialize
'               Class_Terminate
' --------------------------------------------------------------------------
' Dependencies: MMSFormsAPI
'               MWinAPIUser32
'               MWinAPIUser32Style
'               MWinAPIUser32Window
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

Private Const msMODULE              As String = "CUFStyles"

Private Const msFRA_RESIZER         As String = "fraResizer"
Private Const msLBL_GRABBER         As String = "lblGrabber"

Private Const SYMBOL_CHARSET        As Long = 2 ' 0x02

' -----------------------------------
' Event declarations
' -----------------------------------

Public Event OnUserFormChanged(ByRef UF As MSForms.UserForm)

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private mb_Top                      As Boolean
Private mb_Topmost                  As Boolean

Private msng_CaptionHeight          As Single
Private msng_FrameHeight            As Single
Private msng_FrameWidth             As Single
Private msng_TitleBarHeight         As Single

#If VBA7 Then
    Private ml_hWnd                 As LongPtr
#Else
    Private ml_hWnd                 As Long
#End If

Private moSM                        As CSystemMetrics
Private mfra_Resizer                As MSForms.Frame
Attribute mfra_Resizer.VB_VarHelpID = -1
Private mobj_UserForm               As Object
Private WithEvents mfrm_UserForm    As MSForms.UserForm
Attribute mfrm_UserForm.VB_VarHelpID = -1

Public Property Get CaptionHeight() As Single
' ==========================================================================

    Dim sngPPP As Single

    If (msng_CaptionHeight = 0) Then
        sngPPP = PointsPerPixel(saY)
        msng_CaptionHeight = (moSM.CaptionHeight * sngPPP)
    End If

    CaptionHeight = msng_CaptionHeight

End Property

Public Property Let CloseButtonEnabled(ByVal Enabled As Boolean)
' ==========================================================================

    Call SetUserFormCloseButtonState(mfrm_UserForm, Enabled)

End Property

Public Property Get CloseButtonVisible() As Boolean
' ==========================================================================

    CloseButtonVisible = WindowStyleIsSet(ml_hWnd, WS_SYSMENU)

End Property

Public Property Let CloseButtonVisible(ByVal Visible As Boolean)
' ==========================================================================

    Call SetWindowStyle(ml_hWnd, WS_SYSMENU, Visible)

End Property

Public Property Get FrameHeight() As Single
' ==========================================================================

    Dim sngPPP As Single

    If (msng_FrameHeight = 0) Then
        sngPPP = PointsPerPixel(saY)
        msng_FrameHeight = (moSM.DialogFrameHeight * sngPPP)
    End If

    FrameHeight = msng_FrameHeight

End Property

Public Property Get FrameWidth() As Single
' ==========================================================================

    Dim sngPPP As Single

    If (msng_FrameWidth = 0) Then
        sngPPP = PointsPerPixel(saY)
        msng_FrameWidth = (moSM.DialogFrameWidth * sngPPP)
    End If

    FrameWidth = msng_FrameWidth

End Property

#If VBA7 Then
Public Property Get hWnd() As LongPtr
#Else
Public Property Get hWnd() As Long
#End If
' ==========================================================================

    hWnd = ml_hWnd

End Property

Public Property Get MaximizeButtonVisible() As Boolean
' ==========================================================================

    MaximizeButtonVisible = WindowStyleIsSet(ml_hWnd, WS_MAXIMIZEBOX)

End Property

Public Property Let MaximizeButtonVisible(ByVal Visible As Boolean)
' ==========================================================================

    Call SetWindowStyle(hWnd, WS_MAXIMIZEBOX, Visible)

End Property

Public Property Get Maximized() As Boolean
' ==========================================================================

    Maximized = UserFormIsMaximized(mfrm_UserForm)

End Property

Public Property Get MinimizeButtonVisible() As Boolean
' ==========================================================================

    MinimizeButtonVisible = WindowStyleIsSet(ml_hWnd, WS_MINIMIZEBOX)

End Property

Public Property Let MinimizeButtonVisible(ByVal Visible As Boolean)
' ==========================================================================

    Call SetWindowStyle(ml_hWnd, WS_MINIMIZEBOX, Visible)

End Property

Public Property Get Minimized() As Boolean
' ==========================================================================

    Minimized = UserFormIsMinimized(mfrm_UserForm)

End Property

Public Property Get Opacity() As Byte
' ==========================================================================

    Opacity = GetWindowOpacity(ml_hWnd)

End Property

Public Property Let Opacity(ByVal Opacity As Byte)
' ==========================================================================

    Call SetWindowOpacity(ml_hWnd, Opacity)

End Property

Public Property Get Resizable() As Boolean
' ==========================================================================

    Resizable = WindowStyleIsSet(ml_hWnd, WS_SIZEBOX)

End Property

Public Property Let Resizable(ByVal Sizable As Boolean)
' ==========================================================================

    Call SetWindowStyle(ml_hWnd, WS_SIZEBOX, Sizable)
    Call SetResizerVisible(Sizable)

End Property

Public Property Get TitleBarHeight() As Single
' ==========================================================================

    Dim sngPPP As Single

    If (msng_TitleBarHeight = 0) Then
        sngPPP = PointsPerPixel(saY)
        msng_TitleBarHeight = (moSM.CaptionHeight * sngPPP) _
                            + (moSM.DialogFrameHeight * sngPPP)
    End If

    TitleBarHeight = msng_TitleBarHeight

End Property

Public Property Get TitleBarVisible() As Boolean
' ==========================================================================

    TitleBarVisible = WindowStyleIsSet(ml_hWnd, WS_CAPTION)

End Property

Public Property Let TitleBarVisible(ByVal Visible As Boolean)
' ==========================================================================

    If WindowStyleIsSet(ml_hWnd, WS_CAPTION) Then
        Call SetWindowStyle(ml_hWnd, WS_CAPTION, Visible)
        If (Not Visible) Then
            mobj_UserForm.Height = mobj_UserForm.Height _
                                 - Me.CaptionHeight * PointsPerPixel(saY)
        End If
    Else
        Call SetWindowStyle(ml_hWnd, WS_CAPTION, Visible)
        If (Visible) Then
            mobj_UserForm.Height = mobj_UserForm.Height _
                                 + Me.CaptionHeight * PointsPerPixel(saY)
        End If
    End If

End Property

Public Property Get UserForm() As MSForms.UserForm
' ==========================================================================

    Set UserForm = mfrm_UserForm

End Property

Public Property Set UserForm(ByRef UF As Object)
' ==========================================================================

    Set mobj_UserForm = UF
    Set mfrm_UserForm = UF
    RaiseEvent Me.OnUserFormChanged(UF)

End Property

Public Property Get Top() As Boolean
' ==========================================================================

    Top = mb_Top

End Property

Public Property Let Top(ByVal bTop As Boolean)
' ==========================================================================

    If bTop Then
        mb_Top = SetTopWindow(ml_hWnd, bTop)
        mb_Topmost = False
    Else
        mb_Top = SetTopWindow(ml_hWnd, bTop)
        mb_Top = bTop
    End If

End Property

Public Property Get Topmost() As Boolean
' ==========================================================================

    Topmost = mb_Topmost

End Property

Public Property Let Topmost(ByVal bTopmost As Boolean)
' ==========================================================================

    If bTopmost Then
        mb_Topmost = SetTopmostWindow(ml_hWnd, bTopmost)
        mb_Top = False
    Else
        mb_Topmost = SetTopmostWindow(ml_hWnd, bTopmost)
        mb_Topmost = bTopmost
    End If

End Property

Private Sub SetResizerVisible(ByVal Visible As Boolean)
' ==========================================================================
' Description : Set the visibility of the resize "grabber"
'
' Parameters  : Visible     Determines if the grabber
'                           is being added or deleted
' ==========================================================================

    Const sPROC     As String = "SetResizerVisible"

    Dim lblGrabber  As MSForms.Label


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (Me.UserForm Is Nothing) Then
        GoTo PROC_EXIT
    ElseIf (Not Visible) Then
        GoTo HIDE_RESIZER
    End If

    ' ----------------------------------------------------------------------
    ' Add the frame
    ' -------------
    Set mfra_Resizer = Me.UserForm.Controls.Add(gsFORMS_PROGID_FRAME, _
                                                msFRA_RESIZER, _
                                                True)

    ' Add the label
    ' -------------
    Set lblGrabber = mfra_Resizer.Add(gsFORMS_PROGID_LABEL, _
                                      msLBL_GRABBER, _
                                      True)

    ' Set the label properties
    ' ------------------------
    With lblGrabber
        With .Font
            .Name = "Marlett"
            .Charset = SYMBOL_CHARSET
            .Size = 14
            .Bold = True
        End With
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .Caption = Chr(111) ' o - provides the /// grabber
        .AutoSize = True
        .Top = 0
        .Left = 0
        .ZOrder
        .Enabled = False
    End With

    ' Set the frame properties
    ' ------------------------
    With mfra_Resizer
        .BackColor = mfrm_UserForm.BackColor
        .BorderStyle = fmBorderStyleNone
        .Caption = vbNullString
        .SpecialEffect = fmSpecialEffectFlat
        .Height = lblGrabber.Height + 1
        .Width = lblGrabber.Width + 1
        .Top = mfrm_UserForm.InsideHeight - .Height
        .Left = mfrm_UserForm.InsideWidth - .Width
        .ZOrder (fmZOrderBack)
        .Enabled = False
    End With

    mfrm_UserForm.Repaint

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

HIDE_RESIZER:

    If (Not (mfra_Resizer Is Nothing)) Then
        mfra_Resizer.Controls.Remove (msLBL_GRABBER)
        mfrm_UserForm.Controls.Remove (msFRA_RESIZER)
        Set mfra_Resizer = Nothing
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

Public Sub OnUserFormChanged(ByRef UF As MSForms.UserForm)
' ==========================================================================

    ml_hWnd = GetUserFormHandle(UF)

End Sub

Private Sub mfrm_UserForm_Layout()
' ==========================================================================

    If (Not (mfra_Resizer Is Nothing)) Then
        With mfra_Resizer
            .Top = mfrm_UserForm.InsideHeight - .Height
            .Left = mfrm_UserForm.InsideWidth - .Width
        End With
    End If

End Sub

Private Sub Class_Initialize()
' ==========================================================================

    Set moSM = New CSystemMetrics

End Sub

Private Sub Class_Terminate()
' ==========================================================================

    Call SetResizerVisible(False)
    
    Set mfra_Resizer = Nothing
    Set mfrm_UserForm = Nothing
    Set moSM = Nothing

End Sub
