VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FLightbox 
   Caption         =   "Lightbox"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "FLightbox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FLightbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==========================================================================
' Module      : FLightbox
' Type        : Form
' Description : Provide a lightbox effect behind dialogs
' --------------------------------------------------------------------------
' Properties  : IFLightbox_FadeIn       (Get)   Boolean
'               IFLightbox_FadeIn       (Let)   Boolean
'               IFLightbox_FadeOut      (Get)   Boolean
'               IFLightbox_FadeOut      (Let)   Boolean
'               IFLightbox_FadeSpeed    (Get)   Byte
'               IFLightbox_FadeSpeed    (Let)   Byte
'               IFLightbox_Opacity      (Get)   Byte
'               IFLightbox_Opacity      (Let)   Byte
'               IFLightbox_Tag          (Get)   String
'               IFLightbox_Tag          (Let)   String
'               Styles                  (Get)   CUFStyles
' --------------------------------------------------------------------------
' Procedures  : IFLightbox_Hide
'               IFLightbox_Repaint
'               IFLightbox_Show
' --------------------------------------------------------------------------
' Events      : UserForm_Activate
'               UserForm_Initialize
'               UserForm_Terminate
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Interface declarations
' -----------------------------------

Implements IFLightbox

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE  As String = "FLightbox"
Private Const mlPAUSE   As Long = 25

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private mb_FadeIn       As Boolean
Private mb_FadeOut      As Boolean
Private mbyt_FadeSpeed  As Byte
Private mbyt_Opacity    As Byte

Private mo_Styles       As CUFStyles

Private Property Get IFLightbox_FadeIn() As Boolean
' ==========================================================================

    IFLightbox_FadeIn = mb_FadeIn

End Property

Private Property Let IFLightbox_FadeIn(ByVal RHS As Boolean)
' ==========================================================================

    mb_FadeIn = RHS

End Property

Private Property Get IFLightbox_FadeOut() As Boolean
' ==========================================================================

    IFLightbox_FadeOut = mb_FadeOut

End Property

Private Property Let IFLightbox_FadeOut(ByVal RHS As Boolean)
' ==========================================================================

    mb_FadeOut = RHS

End Property

Private Property Get IFLightbox_FadeSpeed() As Byte
' ==========================================================================

    IFLightbox_FadeSpeed = mbyt_FadeSpeed

End Property

Private Property Let IFLightbox_FadeSpeed(ByVal RHS As Byte)
' ==========================================================================

    mbyt_FadeSpeed = RHS

End Property

Private Property Get IFLightbox_Opacity() As Byte
' ==========================================================================

    IFLightbox_Opacity = mbyt_Opacity

End Property

Private Property Let IFLightbox_Opacity(ByVal RHS As Byte)
' ==========================================================================

    mbyt_Opacity = RHS

End Property

Private Property Get IFLightbox_Tag() As String
' ==========================================================================

    IFLightbox_Tag = Me.Tag

End Property

Private Property Let IFLightbox_Tag(ByVal RHS As String)
' ==========================================================================

    Me.Tag = RHS

End Property

Private Sub IFLightbox_Hide()
' ==========================================================================

    Const sPROC     As String = "IFLightbox_Hide"

    Dim bytOpacity  As Byte
    Dim iInterval   As Integer


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (Not mb_FadeOut) Then
        GoTo HIDE_IT
    End If

    ' Fade-out effect
    ' ---------------
    For iInterval = IIf(mbyt_FadeSpeed = 255, _
                        mbyt_FadeSpeed - 1, _
                        mbyt_FadeSpeed) To 1 Step -1

        bytOpacity = mbyt_Opacity * (iInterval / mbyt_FadeSpeed)
        Pause mlPAUSE
        Me.Styles.Opacity = bytOpacity
        Me.Repaint
    
    Next iInterval

HIDE_IT:

    Me.Hide

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

Private Sub IFLightbox_Repaint()
' ==========================================================================

    Me.Repaint

End Sub

Private Sub IFLightbox_Show()
' ==========================================================================
' Description : Display the lightbox with optional Fade-in
' ==========================================================================

    Const sPROC     As String = "IFLightbox_Show"

    Dim bytOpacity  As Byte
    Dim bytInterval As Byte


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Don't show if debugging
    ' -----------------------
    If gbDEBUG_MODE Then
        GoTo PROC_EXIT
    End If

    With Me
        .Show vbModeless
        .Top = Application.Top
        .Left = Application.Left
    End With

    ' Fade-in effect
    ' --------------
    If (Not mb_FadeIn) Then
        GoTo NO_FADE
    End If

    For bytInterval = 0 To IIf(mbyt_FadeSpeed = 255, _
                               254, _
                               mbyt_FadeSpeed)

        bytOpacity = mbyt_Opacity * (bytInterval / mbyt_FadeSpeed)
        Pause mlPAUSE
        Me.Styles.Opacity = bytOpacity
        Me.Repaint

    Next bytInterval

NO_FADE:

    ' Make sure it is at full setting
    ' -------------------------------
    Me.Styles.Opacity = mbyt_Opacity
    Me.Repaint

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

Public Property Get Styles() As CUFStyles
' ==========================================================================

    Set Styles = mo_Styles

End Property

Private Sub UserForm_Activate()
' ==========================================================================

    ' This form can't be properly positioned to
    ' overlap the application window until activated
    ' ----------------------------------------------
    With Me
        .Top = Application.Top - 2
        .Left = Application.Left - 3
    End With

End Sub

Private Sub UserForm_Initialize()
' ==========================================================================

    Const sPROC As String = "UserForm_Initialize"


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Start the services
    ' ------------------
    Set mo_Styles = New CUFStyles
    Set Me.Styles.UserForm = Me

    Me.Styles.TitleBarVisible = False

    ' Set the defaults
    ' ----------------
    mb_FadeIn = True
    mb_FadeOut = True
    mbyt_FadeSpeed = 8
    mbyt_Opacity = 128  ' 50%

    ' Size and position
    ' -----------------
    With Me
        .StartUpPosition = supManual
        .Top = Application.Top
        .Left = Application.Left
        .Height = Application.Height + 4
        .Width = Application.Width + 4
    End With

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

Private Sub UserForm_Terminate()
' ==========================================================================

    Set mo_Styles = Nothing

End Sub
