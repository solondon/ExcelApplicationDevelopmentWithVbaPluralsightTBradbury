VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FAbout 
   Caption         =   "About"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   OleObjectBlob   =   "FAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==========================================================================
' Module      : FAbout
' Type        : Form
' Description : Display information about the application
' --------------------------------------------------------------------------
' Properties  : Version     (Get)   CVersion
' --------------------------------------------------------------------------
' Events      : cmdOK_Click
'               cmdSysInfo_Click
'               lblApp_Copyright_Click
'               lblApp_Date_Click
'               lblApp_Description_Click
'               lblApp_Name_Click
'               UserForm_Click
'               UserForm_Initialize
'               UserForm_QueryClose
'               UserForm_Terminate
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private msMSInfo32  As String
Private mo_Version  As CVersion
Private mo_Position As CUFPosition

Public Property Get Position() As CUFPosition
' ==========================================================================

    Set Position = mo_Position

End Property

Public Property Get Version() As CVersion
' ==========================================================================

    If (mo_Version Is Nothing) Then
        Set mo_Version = New CVersion
    End If

    Set Version = mo_Version

End Property

Private Sub cmdOK_Click()
' ==========================================================================

    Me.Hide

End Sub

Private Sub cmdSysInfo_Click()
' ==========================================================================

    Dim dblRtn  As Double

    dblRtn = Shell(msMSInfo32, vbNormalFocus)

    dblRtn = 0

End Sub

Private Sub lblApp_Copyright_Click()
' ==========================================================================

    Call cmdOK_Click

End Sub

Private Sub lblApp_Date_Click()
' ==========================================================================

    Call cmdOK_Click

End Sub

Private Sub lblApp_Description_Click()
' ==========================================================================

    Call cmdOK_Click

End Sub

Private Sub lblApp_Name_Click()
' ==========================================================================

    Call cmdOK_Click

End Sub

Private Sub UserForm_Click()
' ==========================================================================

    Call cmdOK_Click

End Sub

Private Sub UserForm_Initialize()
' ==========================================================================

' Locate the MSInfo application
' -----------------------------
    msMSInfo32 = ShellGetFolderPath(sfProgramFilesCommon) _
                 & "Microsoft Shared\MSInfo\msinfo32.exe"

    ' Hide the button if it is not found
    ' ----------------------------------
    cmdSysInfo.Visible = FileExists(msMSInfo32)


    ' Use the Version info to populate the form
    ' -----------------------------------------
    Me.Caption = "About " & Me.Version.ProductName

    lblName = Me.Version.ProductName _
            & " version " _
            & Me.Version.ProductVersion
    #If Win64 Then
        lblName = lblName & " (64-bit)"
    #Else
        lblName = lblName & " (32-bit)"
    #End If

    lblCopyright = Me.Version.Copyright
    lblDescription = Me.Version.Description
    lblBuildDate = "Last revised " _
                   & Format(Me.Version.BuildDate, gsVBA_FMTDTM_LONGDATE)
    lblSupport = Me.Version.Support
    lblWarning = Me.Version.Warning


    ' Adjust the size if additional info is not available
    ' ---------------------------------------------------
    If ((Me.Version.Support = vbNullString) _
    And (Me.Version.Warning = vbNullString)) Then
        lblSupport.Visible = False
        fraWarning.Visible = False
        Me.Height = lblSupport.Top + (lblName.Top * 2)

    ElseIf (lblSupport = vbNullString) Then
        lblSupport.Visible = False
        fraWarning.Top = lblSupport.Top
        fraWarning.Height = lblWarning.Height + lblName.Top
        Me.Height = fraWarning.Top + fraWarning.Height + (lblName.Top * 2)

    ElseIf (lblWarning = vbNullString) Then
        fraWarning.Visible = False
        Me.Height = lblSupport.Top + lblSupport.Height + (lblName.Top * 2)

    Else
        fraWarning.Top = lblSupport.Top + lblSupport.Height + lblName.Top
        fraWarning.Height = lblWarning.Height + lblName.Top
        Me.Height = fraWarning.Top + fraWarning.Height + (lblName.Top * 2)
    End If

    ' Center the form
    ' ---------------
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
        Call cmdOK_Click

    End If

End Sub

Private Sub UserForm_Terminate()
' ==========================================================================

    Set mo_Position = Nothing
    Set mo_Version = Nothing

End Sub
