VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUFPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' Module      : CUFPosition
' Type        : Class Module
' Description : Set the position of a UserForm
' --------------------------------------------------------------------------
' Properties  : UserForm            (Get)   MSForms.UserForm
'               UserForm            (Set)   Object
' --------------------------------------------------------------------------
' Procedures  : SetPosition
' --------------------------------------------------------------------------
' Events      : Class_Initialize
'               Class_Terminate
' --------------------------------------------------------------------------
' Dependencies: MMSFormsUserForm
'               MWinAPIGDI32
'               CSystemMetrics
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

Private Const msMODULE  As String = "CUFPosition"

' -----------------------------------
' Event declarations
' -----------------------------------

'Public Event OnUserFormChanged(ByRef UF As MSForms.UserForm)

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private mfrm_UserForm   As MSForms.UserForm
Private mobj_UserForm   As Object

Private mo_SM           As CSystemMetrics

Public Property Get UserForm() As MSForms.UserForm
' ==========================================================================

    Set UserForm = mfrm_UserForm

End Property

Public Property Set UserForm(ByRef UF As Object)
' ==========================================================================

    Set mfrm_UserForm = UF
    Set mobj_UserForm = UF

End Property

Public Sub SetPosition( _
            Optional ByVal PositionTo As enuUserFormPositionTo = ufptAuto, _
            Optional ByVal HPos As enuUserFormHPos = ufhpAuto, _
            Optional ByVal VPos As enuUserFormVPos = ufvpAuto, _
            Optional ByVal XPixels As Long, _
            Optional ByVal YPixels As Long)
' ==========================================================================
' Description : Set the position of a UserForm
'
' Parameters  : PositionTo      Position refrerence for the UserForm
'               HPos            The horizontal position
'               VPos            The vertical position
'               XPixels         Additional pixels on the X-axis
'               YPixels         Additional pixels on the Y-axis
' ==========================================================================

    Const sPROC As String = "SetPosition"

    Dim sngL    As Single
    Dim sngT    As Single
    Dim sngX    As Single
    Dim sngY    As Single

    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Set the form to allow manual positioning
    ' ----------------------------------------
    mobj_UserForm.StartUpPosition = supManual

    ' ----------------------------------------------------------------------
    ' A UserForm should not be displayed
    ' when the application is minimized.
    ' ----------------------------------
    If (Application.WindowState = xlMinimized) Then
        With mobj_UserForm
            .Top = 0
            .Left = 0
        End With
        GoTo PROC_EXIT
    End If

    ' ----------------------------------------------------------------------

    sngX = XPixels * PointsPerPixel(saX)
    sngY = YPixels * PointsPerPixel(saY)

    ' ----------------------------------------------------------------------
    ' Set the defaults
    ' ----------------
    If (HPos = ufhpAuto) Then
        HPos = ufhpCenter
    End If

    If (VPos = ufvpAuto) Then
        VPos = ufvpCenter
    End If

    If (PositionTo = ufptAuto) Then
        PositionTo = ufptApp
    ElseIf (PositionTo = ufptScreen) Then
        GoTo POS_SCREEN
    ElseIf (PositionTo = ufptVScreen) Then
        GoTo POS_VSCREEN
    End If

    ' ----------------------------------------------------------------------

POS_APP:

    Select Case HPos
    Case ufhpLeft
        sngL = Application.Left
    Case ufhpCenter
        sngL = Application.Left _
             + (Application.Width / 2) _
             - (mobj_UserForm.Width / 2)
    Case ufhpRight
        sngL = Application.Left _
             + Application.Width _
             - mobj_UserForm.Width
    End Select

    Select Case VPos
    Case ufvpTop
        sngT = Application.Top
    Case ufvpCenter
        sngT = Application.Top _
             + (Application.Height / 2) _
             - (mobj_UserForm.Height / 2)
    Case ufvpBottom
        sngT = Application.Top _
             + Application.Height _
             - mobj_UserForm.Height
    End Select

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

POS_SCREEN:

    Select Case HPos
    Case ufhpLeft
        sngL = 0
    Case ufhpCenter
        sngL = ((mo_SM.ScreenWidth * PointsPerPixel(saX)) / 2) _
             - (mobj_UserForm.Width / 2)
    Case ufhpRight
        sngL = (mo_SM.ScreenWidth * PointsPerPixel(saX)) _
             - mobj_UserForm.Width
    End Select

    Select Case VPos
    Case ufvpTop
        sngT = 0
    Case ufvpCenter
        sngT = ((mo_SM.ScreenHeight * PointsPerPixel(saY)) / 2) _
             - (mobj_UserForm.Height / 2)
    Case ufvpBottom
        sngT = (mo_SM.ScreenHeight * PointsPerPixel(saY)) _
             - mobj_UserForm.Height
    End Select

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

POS_VSCREEN:

    Select Case HPos
    Case ufhpLeft
        sngL = 0
    Case ufhpCenter
        sngL = ((mo_SM.VirtualScreenWidth * PointsPerPixel(saX)) / 2) _
             - (mobj_UserForm.Width / 2)
    Case ufhpRight
        sngL = (mo_SM.VirtualScreenWidth * PointsPerPixel(saX)) _
             - mobj_UserForm.Width
    End Select

    Select Case VPos
    Case ufvpTop
        sngT = 0
    Case ufvpCenter
        sngT = ((mo_SM.VirtualScreenHeight * PointsPerPixel(saY)) / 2) _
             - (mobj_UserForm.Height / 2)
    Case ufvpBottom
        sngT = (mo_SM.VirtualScreenHeight * PointsPerPixel(saY)) _
             - mobj_UserForm.Height
    End Select

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call mobj_UserForm.Move(sngL + sngX, sngT + sngY)

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

Private Sub Class_Initialize()
' ==========================================================================

    Set mo_SM = New CSystemMetrics

End Sub

Private Sub Class_Terminate()
' ==========================================================================

    Set mo_SM = Nothing

    Set mfrm_UserForm = Nothing
    Set mobj_UserForm = Nothing
    
End Sub
