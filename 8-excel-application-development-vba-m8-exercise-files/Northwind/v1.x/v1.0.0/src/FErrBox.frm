VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FErrBox 
   Caption         =   "System Error"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   OleObjectBlob   =   "FErrBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FErrBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==========================================================================
' Module      : FError
' Type        : Form
' Description : Display top-level error messages
' --------------------------------------------------------------------------
' Properties  : IFErrBox_Caption            (Get)   String
'               IFErrBox_Caption            (Let)   String
'               IFErrBox_DialogResult       (Get)   VbMsgBoxResult
'               IFErrBox_ErrorDescription   (Let)   String
'               IFErrBox_ErrorNumber        (Let)   Long
'               IFErrBox_MsgBoxStyle        (Let)   VbMsgBoxStyle
'               IFErrBox_Procedure          (Let)   String
'               IFErrBox_SendEmail          (Get)   Boolean
'               IFErrBox_Tag                (Get)   String
'               IFErrBox_Tag                (Let)   String
'               IFErrBox_UserComments       (Get)   String
'               Position                    (Get)   CUFPosition
' --------------------------------------------------------------------------
' Procedures  : IFErrBox_Hide
'               IFErrBox_Show
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
'               UserForm_Terminate
' --------------------------------------------------------------------------
' Dependencies: MVBAError
'               MWinAPIMM
'               CSystemMetrics
'               CUFPosition
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Interface declarations
' -----------------------------------

Implements IFErrBox

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE          As String = "FError"

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
Private mo_SM                   As CSystemMetrics

Private Property Get IFErrBox_Caption() As String
' ==========================================================================

    IFErrBox_Caption = Me.Caption

End Property

Private Property Let IFErrBox_Caption(ByVal RHS As String)
' ==========================================================================

    Me.Caption = IFErrBox_Caption

End Property

Private Property Get IFErrBox_DialogResult() As VbMsgBoxResult
' ==========================================================================

    IFErrBox_DialogResult = me_DialogResult

End Property

Private Property Let IFErrBox_ErrorDescription(ByVal RHS As String)
' ==========================================================================

    txtErrDesc.Value = RHS

End Property

Private Property Let IFErrBox_ErrorNumber(ByVal RHS As Long)
' ==========================================================================

    txtErrNum.Value = RHS

End Property

Private Property Let IFErrBox_MsgBoxStyle(ByVal RHS As VbMsgBoxStyle)
' ==========================================================================

    me_Style = RHS

    RaiseEvent Me.OnStyleChanged(me_Style)

End Property

Private Property Let IFErrBox_Procedure(ByVal RHS As String)
' ==========================================================================

    txtProc.Value = RHS

End Property

Private Property Get IFErrBox_SendEmail() As Boolean
' ==========================================================================

    IFErrBox_SendEmail = mb_SendEmail

End Property

Private Property Get IFErrBox_Tag() As String
' ==========================================================================

    IFErrBox_Tag = Me.Tag

End Property

Private Property Let IFErrBox_Tag(ByVal RHS As String)
' ==========================================================================

    Me.Tag = RHS

End Property

Private Property Get IFErrBox_UserComments() As String
' ==========================================================================

    IFErrBox_UserComments = txtUserComments.Value

End Property

Private Sub IFErrBox_Hide()
' ==========================================================================

    Me.Hide

End Sub

Private Sub IFErrBox_Show()
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

    Dim oCtl    As MSForms.Control

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
        Set oCtl = imgCritical
    Case vbQuestion
        Set oCtl = imgQuestion
    Case vbExclamation
        Set oCtl = imgExclamation
    Case vbInformation
        Set oCtl = imgInformation
    End Select

    ' Display it
    ' ----------
    oCtl.Visible = True
    oCtl.TabStop = False

    Set oCtl = Nothing

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
        If (Not Application.VBE.MainWindow.Visible) Then
            Application.VBE.MainWindow.Visible = True
        End If
    End If

    Set mo_Position = New CUFPosition
    Set Me.Position.UserForm = Me
    Call Me.Position.SetPosition

    Set mo_SM = New CSystemMetrics

End Sub

Private Sub UserForm_Layout()
' ==========================================================================

    Dim sngHeight   As Single
    Dim sngWidth    As Single

    ' ----------------------------------------------------------------------

    If fraUserComments.Visible Then
        sngHeight = fraUserComments.Top _
                  + fraUserComments.Height
    ElseIf cmd4.Visible Then
        sngHeight = cmd4.Top _
                  + cmd4.Height
    Else
        sngHeight = txtErrDesc.Top _
                  + txtErrDesc.Height
    End If

    sngHeight = sngHeight _
              + txtProc.Top _
              + PixelsToPoints(saY, (2 * mo_SM.DialogFrameHeight)) _
              + PixelsToPoints(saY, mo_SM.CaptionHeight)

    sngWidth = cmd1.Left _
             + cmd1.Width _
             + imgCritical.Left _
             + PixelsToPoints(saX, (2 * mo_SM.DialogFrameWidth))

    ' ----------------------------------------------------------------------

    With Me
        .Height = sngHeight
        .Width = sngWidth
    End With

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

    If (CloseMode = vbFormControlMenu) Then

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

Private Sub UserForm_Terminate()
' ==========================================================================

    Set mo_Position = Nothing
    Set mo_SM = Nothing

End Sub
