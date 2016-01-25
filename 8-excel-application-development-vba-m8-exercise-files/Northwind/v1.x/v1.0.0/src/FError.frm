VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FError 
   Caption         =   "System Error"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7050
   OleObjectBlob   =   "FError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==========================================================================
' Module      : FError
' Type        : Form
' Description : Display top-level error messages
' --------------------------------------------------------------------------
' Properties  : IFError_Caption           (Get)  String
'               IFError_Caption           (Let)  String
'               IFError_DialogResult      (Get)  VbMsgBoxResult
'               IFError_ErrorDescription  (Let)  String
'               IFError_ErrorNumber       (Let)  Long
'               IFError_MsgBoxStyle       (Let)  VbMsgBoxStyle
'               IFError_Procedure         (Let)  String
'               IFError_SendEmail         (Get)  Boolean
'               IFError_Tag               (Get)  String
'               IFError_Tag               (Let)  String
'               IFError_UserComments      (Get)  String
' --------------------------------------------------------------------------
' Procedures  : IFError_Hide
'               IFError_Show
' --------------------------------------------------------------------------
' Events      : OnButtonsChanged
'               OnDefaultButtonChanged
'               OnIconChanged
'               OnStyleChanged
'               cmd1_Click
'               cmd2_Click
'               cmd3_Click
'               cmd4_Click
'               cmdSendEmail_Click
'               txtUserComments_Change
'               UserForm_Initialize
'               UserForm_QueryClose
' --------------------------------------------------------------------------
' Dependencies: MVBAError, MWinAPIMM
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Interface declarations
' -----------------------------------

Implements IFError

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE          As String = "FError"
Private Const msngDEBUG_HEIGHT  As Single = 135

' -----------------------------------
' Event declarations
' -----------------------------------

Public Event OnButtonsChanged(ByVal Buttons As VbMsgBoxStyle)
Public Event OnDefaultButtonChanged(ByVal DefaultButton As VbMsgBoxStyle)
Public Event OnIconChanged(ByVal Icon As VbMsgBoxStyle)
Public Event OnStyleChanged(ByVal Style As VbMsgBoxStyle)

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private me_DialogResult         As VbMsgBoxResult
Private me_Style                As VbMsgBoxStyle
Private mb_SendEmail            As Boolean

Private meButtons               As VbMsgBoxStyle
Private meIcon                  As VbMsgBoxStyle
Private meDefault               As VbMsgBoxStyle
Private meModal                 As VbMsgBoxStyle

Private mo_Position             As CUFPosition

Private Property Get IFError_Caption() As String
' ==========================================================================

    IFError_Caption = Me.Caption

End Property

Private Property Let IFError_Caption(ByVal RHS As String)
' ==========================================================================

    Me.Caption = IFError_Caption

End Property

Private Property Get IFError_DialogResult() As VbMsgBoxResult
' ==========================================================================

    IFError_DialogResult = me_DialogResult

End Property

Private Property Let IFError_ErrorDescription(ByVal RHS As String)
' ==========================================================================

    txtErrDesc.Value = RHS

End Property

Private Property Let IFError_ErrorNumber(ByVal RHS As Long)
' ==========================================================================

    txtErrNum.Value = RHS

End Property

Private Property Let IFError_MsgBoxStyle(ByVal RHS As VbMsgBoxStyle)
' ==========================================================================

    me_Style = RHS

    RaiseEvent Me.OnStyleChanged(me_Style)

End Property

Private Property Let IFError_Procedure(ByVal RHS As String)
' ==========================================================================

    txtProc.Value = RHS

End Property

Private Property Get IFError_SendEmail() As Boolean
' ==========================================================================

    IFError_SendEmail = mb_SendEmail

End Property

Private Property Get IFError_Tag() As String
' ==========================================================================

    IFError_Tag = Me.Tag

End Property

Private Property Let IFError_Tag(ByVal RHS As String)
' ==========================================================================

    Me.Tag = RHS

End Property

Private Property Get IFError_UserComments() As String
' ==========================================================================

    IFError_UserComments = txtUserComments.Value

End Property

Private Sub IFError_Hide()
' ==========================================================================

    Me.Hide

End Sub

Private Sub IFError_Show()
' ==========================================================================

    Select Case meIcon
    Case vbCritical
        Call PlayEventSound(SND_ALIAS_SYSTEMHAND)
    Case vbExclamation
        Call PlayEventSound(SND_ALIAS_SYSTEMEXCLAMATION)
    Case vbInformation
        Call PlayEventSound(SND_ALIAS_SYSTEMDEFAULT)
    Case vbQuestion
        Call PlayEventSound(SND_ALIAS_SYSTEMQUESTION)
    Case Else
        Call PlayEventSound(SND_ALIAS_SYSTEMDEFAULT)
    End Select

    Me.Show

End Sub

Public Property Get Position() As CUFPosition
' ==========================================================================

    Set Position = mo_Position

End Property

Public Sub OnButtonsChanged(ByVal Buttons As VbMsgBoxStyle)
' ==========================================================================

    cmd2.Visible = False
    cmd3.Visible = False
    cmd4.Visible = False

    Select Case Buttons
    Case vbRetryCancel
        With cmd1
            .Caption = "Retry"
        End With
        With cmd2
            .Visible = True
            .Caption = "Cancel"
            .Cancel = True
        End With
    
    Case vbYesNo
        With cmd1
            .Caption = "Yes"
        End With
        With cmd2
            .Visible = True
            .Caption = "No"
            .Cancel = True
        End With

    Case vbYesNoCancel
        With cmd1
            .Caption = "Yes"
        End With
        With cmd2
            .Caption = "No"
            .Visible = True
        End With
        With cmd3
            .Visible = True
            .Caption = "Cancel"
            .Cancel = True
        End With
    
    Case vbAbortRetryIgnore
        With cmd1
            .Caption = "Abort"
            .Cancel = True
        End With
        With cmd2
            .Visible = True
            .Caption = "Retry"
        End With
        With cmd3
            .Visible = True
            .Caption = "Ignore"
        End With
    
    Case vbOKCancel
        With cmd1
            .Caption = "OK"
        End With
        With cmd2
            .Visible = True
            .Caption = "Cancel"
            .Cancel = True
        End With
    
    Case Else
        With cmd1
            .Caption = "OK"
            .Cancel = True
        End With
    End Select

End Sub

Public Sub OnDefaultButtonChanged(ByVal DefaultButton As VbMsgBoxStyle)
' ==========================================================================

    cmd1.Default = False
    cmd2.Default = False
    cmd3.Default = False
    cmd4.Default = False

    Select Case DefaultButton
    Case vbDefaultButton1
        cmd1.Default = True
    Case vbDefaultButton2
        cmd2.Default = True
    Case vbDefaultButton3
        cmd3.Default = True
    Case vbDefaultButton4
        cmd4.Default = True
    End Select

End Sub

Public Sub OnIconChanged(ByVal Icon As VbMsgBoxStyle)
' ==========================================================================

    Dim ctl As MSForms.Control

    ' Hide them all
    ' -------------
    imgCritical.Visible = False
    imgQuestion.Visible = False
    imgExclamation.Visible = False
    imgInformation.Visible = False

    ' Select which one should be visible
    ' ----------------------------------
    Select Case Icon
    Case vbCritical
        Set ctl = imgCritical
    Case vbQuestion
        Set ctl = imgQuestion
    Case vbExclamation
        Set ctl = imgExclamation
    Case vbInformation
        Set ctl = imgInformation
    End Select

    ' Display it
    ' ----------
    ctl.Visible = True
'    ctl.TabStop = False
    
    Set ctl = Nothing

End Sub

Public Sub OnStyleChanged(ByVal Style As VbMsgBoxStyle)
' ==========================================================================
    
    meButtons = ParseVbMsgBoxStyle(Style, mbspButtons)
    meDefault = ParseVbMsgBoxStyle(Style, mbspDefaultButton)
    meIcon = ParseVbMsgBoxStyle(Style, mbspIcon)

    RaiseEvent Me.OnButtonsChanged(meButtons)
    RaiseEvent Me.OnDefaultButtonChanged(meDefault)
    RaiseEvent Me.OnIconChanged(meIcon)
    
End Sub

Private Sub cmd1_Click()
' ==========================================================================
    
    Select Case meButtons
    Case vbOKOnly
        me_DialogResult = vbOK
    Case vbRetryCancel
        me_DialogResult = vbRetry
    Case vbYesNo
        me_DialogResult = vbYes
    Case vbYesNoCancel
        me_DialogResult = vbYes
    Case vbAbortRetryIgnore
        me_DialogResult = vbAbort
    Case vbOKCancel
        me_DialogResult = vbOK
    Case Else
        me_DialogResult = vbOK
    End Select

    Me.Hide
    
End Sub

Private Sub cmd2_Click()
' ==========================================================================

    Select Case meButtons
    Case vbRetryCancel
        me_DialogResult = vbCancel
    Case vbYesNo
        me_DialogResult = vbNo
    Case vbYesNoCancel
    Case vbAbortRetryIgnore
        me_DialogResult = vbRetry
    Case vbOKCancel
        me_DialogResult = vbCancel
    Case Else
    End Select

    Me.Hide

End Sub

Private Sub cmd3_Click()
' ==========================================================================

    Select Case meButtons
    Case vbYesNoCancel
        me_DialogResult = vbCancel
    Case vbAbortRetryIgnore
        me_DialogResult = vbIgnore
    End Select

    Me.Hide

End Sub

Private Sub cmdSendEmail_Click()
' ==========================================================================

    mb_SendEmail = True
    me_DialogResult = vbOK

    Me.Hide

End Sub

Private Sub txtUserComments_Change()
' ==========================================================================

    cmdSendEmail.Enabled = Len(Trim$(txtUserComments))

End Sub

Private Sub UserForm_Initialize()
' ==========================================================================

    Me.Caption = gsAPP_NAME & " Error"

    If gbDEBUG_MODE Then
        cmdSendEmail.Visible = False
        lblUserComments.Visible = False
        txtUserComments.Visible = False
        fraUserComments.Visible = False
        Me.Height = msngDEBUG_HEIGHT
        If (Not Application.VBE.MainWindow.Visible) Then
            Application.VBE.MainWindow.Visible = True
        End If
    End If

    Set mo_Position = New CUFPosition
    Set Me.Position.UserForm = Me
    Call Me.Position.SetPosition

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ==========================================================================
' Description : Clean up after use, and check for improper exit
'
' Parameters  : Cancel      Setting this argument to any value other than 0
'                           stops the QueryClose event in all loaded user
'                           forms and prevents the UserForm from closing.
'               CloseMode   Indicates the cause of the QueryClose event.
' ==========================================================================

    If (CloseMode <> vbFormCode) Then

        ' Prevent the form from unloading
        ' -------------------------------
        Cancel = True

        ' Use default processing
        ' ----------------------
        Select Case meButtons
        Case vbRetryCancel
            Call cmd2_Click
        Case vbYesNo
            Call cmd2_Click
        Case vbYesNoCancel
            Call cmd3_Click
        Case vbAbortRetryIgnore
            Call cmd1_Click
        Case vbOKCancel
            Call cmd2_Click
        Case Else
            me_DialogResult = vbCancel
            Cancel = False
        End Select
    End If

End Sub
