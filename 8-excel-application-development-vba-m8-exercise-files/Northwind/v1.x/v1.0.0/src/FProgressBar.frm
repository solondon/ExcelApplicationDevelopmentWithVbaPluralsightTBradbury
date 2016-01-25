VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FProgressBar 
   Caption         =   "ProgressBar"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6150
   OleObjectBlob   =   "FProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==========================================================================
' Module      : FProgressBar
' Type        : Form
' Description : Dialog implementation of IProgressBar
' --------------------------------------------------------------------------
' Properties  : IProgressBar_Canceled           (Get)   Boolean
'               IProgressBar_Canceled           (Let)   Boolean
'               IProgressBar_CancelVisible      (Get)   Boolean
'               IProgressBar_CancelVisible      (Let)   Boolean
'               IProgressBar_Caption            (Get)   String
'               IProgressBar_Caption            (Let)   String
'               IProgressBar_ChangeRate         (Get)   Double
'               IProgressBar_ChangeRate         (Let)   Double
'               IProgressBar_Max                (Get)   Long
'               IProgressBar_Max                (Let)   Long
'               IProgressBar_Min                (Get)   Long
'               IProgressBar_Min                (Let)   Long
'               IProgressBar_OverallCaption     (Get)   String
'               IProgressBar_OverallCaption     (Let)   String
'               IProgressBar_OverallChangeRate  (Get)   Double
'               IProgressBar_OverallChangeRate  (Let)   Double
'               IProgressBar_OverallMax         (Get)   Long
'               IProgressBar_OverallMax         (Let)   Long
'               IProgressBar_OverallMin         (Get)   Long
'               IProgressBar_OverallMin         (Let)   Long
'               IProgressBar_OverallPercent     (Get)   Double
'               IProgressBar_OverallValue       (Get)   Long
'               IProgressBar_OverallValue       (Let)   Long
'               IProgressBar_OverallVisible     (Get)   Boolean
'               IProgressBar_OverallVisible     (Let)   Boolean
'               IProgressBar_Percent            (Get)   Double
'               IProgressBar_Title              (Get)   String
'               IProgressBar_Title              (Let)   String
'               IProgressBar_Value              (Get)   Long
'               IProgressBar_Value              (Let)   Long
'               Position                        (Get)   CUFPosition
'               Styles                          (Get)   CUFStyles
' --------------------------------------------------------------------------
' Procedures  : IProgressBar_Complete
'               IProgressBar_Decrement
'               IProgressBar_Hide
'               IProgressBar_Increment
'               IProgressBar_OverallComplete
'               IProgressBar_OverallDecrement
'               IProgressBar_OverallIncrement
'               IProgressBar_OverallReset
'               IProgressBar_Refresh
'               IProgressBar_Reset
'               IProgressBar_Show
'               GetPercentage                           Double
' --------------------------------------------------------------------------
' Events      : OnOverallValueChanged
'               OnValueChanged
'               cmdCancel_Click
'               UserForm_Activate
'               UserForm_Initialize
'               UserForm_Layout
'               UserForm_Terminate
' --------------------------------------------------------------------------
' Dependencies: CUFPosition
'               IFProgressBar
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Interface declarations
' -----------------------------------

Implements IProgressBar

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE              As String = "FProgressBar"

Private Const msDEFAULT_CAP         As String = vbNullString
Private Const mlDEFAULT_MIN         As Long = 0
Private Const mlDEFAULT_MAX         As Long = 100
Private Const mlDEFAULT_VAL         As Long = -1
Private Const mdblDEFAULT_CHG       As Double = 0.05

' Form settings
' -------------
Private Const msngHEIGHT_DEFAULT    As Single = 64.5
Private Const msngHEIGHT_OVERALL    As Single = 100
Private Const msngWIDTH_DEFAULT     As Single = 237

' -----------------------------------
' Event declarations
' -----------------------------------
' Module Level
' ----------------

Public Event OnOverallValueChanged(ByVal Value As Long)
Public Event OnValueChanged(ByVal Value As Long)

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private mb_Canceled                 As Boolean
Private mb_CancelVisible            As Boolean

' ProgressBar
Private ms_Caption                  As String
Private mb_Visible                  As Boolean
Private ml_Min                      As Long
Private ml_Max                      As Long
Private ml_Value                    As Long
Private mdbl_ChangeRate             As Double
Private mdbl_Percent                As Double
Private mdblLastPct                 As Double

' OverallBar
Private ms_OverallCaption           As String
Private mb_OverallVisible           As Boolean
Private ml_OverallMin               As Long
Private ml_OverallMax               As Long
Private ml_OverallValue             As Long
Private mdbl_OverallChangeRate      As Double
Private mdbl_OverallPercent         As Double
Private mdblOverallLastPct          As Double

Private mo_Position                 As CUFPosition
Private mo_Styles                   As CUFStyles

Private Property Get IProgressBar_Canceled() As Boolean
' ==========================================================================

    IProgressBar_Canceled = mb_Canceled

End Property

Private Property Let IProgressBar_Canceled(ByVal RHS As Boolean)
' ==========================================================================

    mb_Canceled = RHS

End Property

Private Property Get IProgressBar_CancelVisible() As Boolean
' ==========================================================================

    IProgressBar_CancelVisible = mb_CancelVisible

End Property

Private Property Let IProgressBar_CancelVisible(ByVal RHS As Boolean)
' ==========================================================================

    mb_CancelVisible = RHS

    cmdCancel.Cancel = mb_CancelVisible
    cmdCancel.Visible = mb_CancelVisible

    UserForm_Layout

End Property

Private Property Get IProgressBar_Caption() As String
' ==========================================================================

    IProgressBar_Caption = ms_Caption

End Property

Private Property Let IProgressBar_Caption(ByVal RHS As String)
' ==========================================================================
' Description : This is the base caption string.
'               The percentage is calculated during the
'               progress change and appended automatically.
' ==========================================================================

    ms_Caption = RHS

    IProgressBar_Refresh

End Property

Private Property Get IProgressBar_ChangeRate() As Double
' ==========================================================================

    IProgressBar_ChangeRate = mdbl_ChangeRate

End Property

Private Property Let IProgressBar_ChangeRate(ByVal RHS As Double)
' ==========================================================================

    mdbl_ChangeRate = RHS

End Property

Private Property Get IProgressBar_Max() As Long
' ==========================================================================

    IProgressBar_Max = ml_Max

End Property

Private Property Let IProgressBar_Max(ByVal RHS As Long)
' ==========================================================================

    ml_Max = RHS

End Property

Private Property Get IProgressBar_Min() As Long
' ==========================================================================

    IProgressBar_Min = ml_Min

End Property

Private Property Let IProgressBar_Min(ByVal RHS As Long)
' ==========================================================================

    ml_Min = RHS

End Property

Private Property Get IProgressBar_OverallCaption() As String
' ==========================================================================

    IProgressBar_OverallCaption = ms_OverallCaption

End Property

Private Property Let IProgressBar_OverallCaption(ByVal RHS As String)
' ==========================================================================
' Description : This is the base caption string.
'               The percentage is calculated during the
'               progress change and appended automatically.
' ==========================================================================

    ms_OverallCaption = RHS

    IProgressBar_Refresh

End Property

Private Property Get IProgressBar_OverallChangeRate() As Double
' ==========================================================================

    IProgressBar_OverallChangeRate = mdbl_OverallChangeRate

End Property

Private Property Let IProgressBar_OverallChangeRate(ByVal RHS As Double)
' ==========================================================================

    mdbl_OverallChangeRate = RHS

End Property

Private Property Get IProgressBar_OverallMax() As Long
' ==========================================================================

    IProgressBar_OverallMax = ml_OverallMax

End Property

Private Property Let IProgressBar_OverallMax(ByVal RHS As Long)
' ==========================================================================

    ml_OverallMax = RHS

End Property

Private Property Get IProgressBar_OverallMin() As Long
' ==========================================================================

    IProgressBar_OverallMin = ml_OverallMin

End Property

Private Property Let IProgressBar_OverallMin(ByVal RHS As Long)
' ==========================================================================

    ml_OverallMin = RHS

End Property

Private Property Get IProgressBar_OverallPercent() As Double
' ==========================================================================

    IProgressBar_OverallPercent = mdbl_OverallPercent

End Property

Private Property Get IProgressBar_OverallValue() As Long
' ==========================================================================

    IProgressBar_OverallValue = ml_OverallValue

End Property

Private Property Let IProgressBar_OverallValue(ByVal RHS As Long)
' ==========================================================================

    If (RHS <> ml_OverallValue) Then
        ml_OverallValue = Within(RHS, ml_OverallMin, ml_OverallMax)
        RaiseEvent Me.OnOverallValueChanged(ml_OverallValue)
    End If

End Property

Private Property Get IProgressBar_OverallVisible() As Boolean
' ==========================================================================

    IProgressBar_OverallVisible = mb_OverallVisible

End Property

Private Property Let IProgressBar_OverallVisible(ByVal RHS As Boolean)
' ==========================================================================

    mb_OverallVisible = RHS
    lblOverall.Visible = RHS

    UserForm_Layout

End Property

Private Property Get IProgressBar_Percent() As Double
' ==========================================================================

    IProgressBar_Percent = mdbl_Percent

End Property

Private Property Get IProgressBar_Title() As String
' ==========================================================================

    IProgressBar_Title = Me.Caption

End Property

Private Property Let IProgressBar_Title(ByVal RHS As String)
' ==========================================================================

    Me.Caption = RHS

End Property

Private Property Get IProgressBar_Value() As Long
' ==========================================================================

    IProgressBar_Value = ml_Value

End Property

Private Property Let IProgressBar_Value(ByVal RHS As Long)
' ==========================================================================

    If (RHS <> ml_Value) Then
        ml_Value = Within(RHS, ml_Min, ml_Max)
        RaiseEvent Me.OnValueChanged(ml_Value)
    End If

End Property

Private Sub IProgressBar_Complete()
' ==========================================================================

    IProgressBar_Value = IProgressBar_Max

End Sub

Private Sub IProgressBar_Decrement()
' ==========================================================================

    If (IProgressBar_Value > IProgressBar_Min) Then
        IProgressBar_Value = IProgressBar_Value - 1
    End If

End Sub

Private Sub IProgressBar_Hide()
' ==========================================================================

    mb_Visible = False
    Me.Hide

End Sub

Private Sub IProgressBar_Increment()
' ==========================================================================

    If (IProgressBar_Value < IProgressBar_Max) Then
        IProgressBar_Value = IProgressBar_Value + 1
    End If

End Sub

Private Sub IProgressBar_OverallComplete()
' ==========================================================================

    IProgressBar_OverallValue = IProgressBar_OverallMax

End Sub

Private Sub IProgressBar_OverallDecrement()
' ==========================================================================

    If (IProgressBar_OverallValue > IProgressBar_OverallMin) Then
        IProgressBar_OverallValue = IProgressBar_OverallValue - 1
    End If

End Sub

Private Sub IProgressBar_OverallIncrement()
' ==========================================================================

    If (IProgressBar_OverallValue < IProgressBar_OverallMax) Then
        IProgressBar_OverallValue = IProgressBar_OverallValue + 1
    End If

End Sub

Private Sub IProgressBar_OverallReset()
' ==========================================================================

    IProgressBar_ChangeRate = mdblDEFAULT_CHG
    IProgressBar_OverallCaption = msDEFAULT_CAP

    IProgressBar_OverallMin = mlDEFAULT_MIN
    IProgressBar_OverallMax = mlDEFAULT_MAX
    IProgressBar_OverallValue = mlDEFAULT_VAL

End Sub

Private Sub IProgressBar_Refresh()
' ==========================================================================
' Description : Refresh the display
' ==========================================================================

    Dim sngWidth As Single
    Dim sngWidthOverall As Single

    Dim sCap    As String
    Dim sCapOverall As String
    Dim sPct    As String
    Dim sPctOverall As String

    ' Build the caption
    ' -----------------
    If (Len(ms_Caption) > 0) Then
        sCap = ms_Caption & "..."
    End If
    sPct = Format(mdbl_Percent, "0%")
    sngWidth = mdbl_Percent * 200

    If mb_OverallVisible Then
        If (Len(ms_OverallCaption) > 0) Then
            sCapOverall = ms_OverallCaption & "..."
        End If
        sPctOverall = Format(mdbl_OverallPercent, "0%")
        sngWidthOverall = mdbl_OverallPercent * 200
    End If

    ' Refresh the display
    ' -------------------
    If Me.Visible Then
        lblProgress.Caption = sCap
        lblProgressBar.Width = sngWidth
        lblProgressPct.Caption = sPct

        If mb_OverallVisible Then
            lblOverall.Caption = sCapOverall
            lblOverallBar.Width = sngWidthOverall
            lblOverallPct.Caption = sPctOverall
        End If
    End If

End Sub

Private Sub IProgressBar_Reset()
' ==========================================================================

    IProgressBar_ChangeRate = mdblDEFAULT_CHG
    IProgressBar_Caption = msDEFAULT_CAP

    IProgressBar_Min = mlDEFAULT_MIN
    IProgressBar_Max = mlDEFAULT_MAX
    IProgressBar_Value = mlDEFAULT_VAL

End Sub

Private Sub IProgressBar_Show()
' ==========================================================================

    mb_Visible = True
    Me.Show vbModeless

End Sub

Public Property Get Position() As CUFPosition
' ==========================================================================

    Set Position = mo_Position

End Property

Public Property Get Styles() As CUFStyles
' ==========================================================================

    Set Styles = mo_Styles

End Property

Private Function GetPercentage(ByVal Min As Long, _
                               ByVal Max As Long, _
                               ByVal Progress As Long) As Double
' ==========================================================================
' Description : Calculate the percentage for a progress bar
'
' Parameters  : Min         The minimum value for the bar
'               Max         The maximum value for the bar
'               Progress    The current value of the bar
'
' Returns     : Double      The current progress percentage
' ==========================================================================

    Const sPROC As String = "GetPercentage"

    Dim dblRtn  As Double


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Calculate the progress percentage
    ' ---------------------------------
    If (Max <= Min) Then
        dblRtn = 0
    Else
        dblRtn = Abs((Progress - Min) / (Max - Min))
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetPercentage = dblRtn

    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If (Err.Number = ERR_USER_INTERRUPT) Then
        Resume PROC_EXIT
    ElseIf ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Sub OnOverallValueChanged(ByVal Value As Long)
' ==========================================================================
' Description : Recalculate overall percentage based on the new value
' ==========================================================================

    Const sPROC As String = "OnOverallValueChanged"


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Get the current percentage
    ' --------------------------
    mdbl_OverallPercent = GetPercentage(ml_OverallMin, ml_OverallMax, Value)

    ' Refresh if needed
    ' -----------------
    If (Abs(mdbl_OverallPercent - mdblOverallLastPct) _
      > mdbl_OverallChangeRate) Then
        mdblOverallLastPct = mdbl_OverallPercent
        If mb_Visible Then
            IProgressBar_Refresh

            ' Process events (allow cancel) at each update
            ' --------------------------------------------
            DoEvents

        End If
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

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

Public Sub OnValueChanged(ByVal Value As Long)
' ==========================================================================
' Description : Recalculate percentage based on the new value
' ==========================================================================

    Const sPROC As String = "OnValueChanged"

    Dim bRefresh As Boolean

    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Get the current percentage
    ' --------------------------
    mdbl_Percent = GetPercentage(ml_Min, ml_Max, Value)

    ' Refresh if needed
    ' -----------------
    bRefresh = (Abs(mdbl_Percent - mdblLastPct) > mdbl_ChangeRate)
    If bRefresh Then
        mdblLastPct = mdbl_Percent
        If mb_Visible Then
            IProgressBar_Refresh

            ' Process events (allow cancel) at each update
            ' --------------------------------------------
            DoEvents

        End If
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

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

Private Sub cmdCancel_Click()
' ==========================================================================

    Me.Styles.Topmost = False
    IProgressBar_Canceled = True

End Sub

Private Sub UserForm_Activate()
' ==========================================================================

    ' Size the form
    ' -------------
    UserForm_Layout

    ' Center the form
    ' ---------------
    Call Me.Position.SetPosition

End Sub

Private Sub UserForm_Initialize()
' ==========================================================================
' Description : Set initial values
' ==========================================================================

    Const sPROC As String = "UserForm_Initialize"


    On Error GoTo PROC_ERR
    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Me.Caption = gsAPP_NAME

    lblProgressBar.Width = 1
    lblProgressPct.Caption = "0%"

    lblOverallBar.Width = 1
    lblOverallPct.Caption = "0%"

    Call IProgressBar_Reset
    Call IProgressBar_OverallReset

    Me.Height = msngHEIGHT_DEFAULT
    Me.Width = msngWIDTH_DEFAULT

    ' Set the form styles
    ' -------------------
    Set mo_Styles = New CUFStyles
    Set Me.Styles.UserForm = Me

    Me.Styles.Topmost = True
    Me.Styles.CloseButtonVisible = gbDEBUG_MODE

    ' Position the form
    ' -----------------
    Set mo_Position = New CUFPosition
    Set Me.Position.UserForm = Me
    
    Call Me.Position.SetPosition

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

Private Sub UserForm_Layout()
' ==========================================================================
' Description : Manage the resizing of the dialog and its controls
' ==========================================================================

    Dim sngCaptionHeight    As Single
    Dim sngFrameHeight      As Single
    Dim sngFrameWidth       As Single
    Dim sngInsideHeight     As Single
    Dim sngInsideWidth      As Single

    Dim oSM                 As CSystemMetrics

    ' ----------------------------------------------------------------------
    
    Set oSM = New CSystemMetrics

    sngCaptionHeight = oSM.CaptionHeight * PointsPerPixel(saY)
    sngFrameHeight = oSM.DialogFrameHeight * PointsPerPixel(saY)
    sngFrameWidth = oSM.DialogFrameWidth * PointsPerPixel(saX)

    ' Set the height
    ' --------------
    If mb_OverallVisible Then
        sngInsideHeight = lblOverallFrame.Top _
                        + lblOverallFrame.Height _
                        + lblProgress.Top
    Else
        sngInsideHeight = lblOverall.Top
    End If

    Me.Height = sngInsideHeight _
              + (2 * sngFrameHeight) _
              + sngCaptionHeight

    ' Set the width
    ' -------------
    If mb_CancelVisible Then
        sngInsideWidth = cmdCancel.Left _
                       + cmdCancel.Width _
                       + lblProgressFrame.Left
        cmdCancel.Visible = True
    Else
        sngInsideWidth = cmdCancel.Left
        cmdCancel.Visible = False
    End If

    Me.Width = sngInsideWidth _
             + (2 * sngFrameWidth)

    ' ----------------------------------------------------------------------

    Set oSM = Nothing

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

        If mb_CancelVisible Then

            ' Prevent the form from unloading
            ' -------------------------------
            Cancel = True
    
            ' Use deault processing
            ' ---------------------
            Call cmdCancel_Click
        
        End If

    End If

End Sub

Private Sub UserForm_Terminate()
' ==========================================================================

    Set mo_Position = Nothing
    Set mo_Styles = Nothing

End Sub
