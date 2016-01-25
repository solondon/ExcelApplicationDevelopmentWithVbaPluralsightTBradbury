Attribute VB_Name = "MNWBuild"
' ==========================================================================
' Module      : MNWSheets
' Type        : Module
' Description :
' --------------------------------------------------------------------------
' Procedures  : BuildWorkbook
'               SetDocumentProperties
'               SetVBAProjectProperties
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

Private Const msMODULE                  As String = "MNWSheets"

Public Sub BuildWorkbook(Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Add sheets to the workbook and format them for use
' ==========================================================================

    Const sPROC     As String = "BuildWorkbook"
    Const sTHEME    As String = "My Sample Theme.thmx"

    Dim sFileName   As String

    Dim oPB         As IProgressBar
    Dim udtProps    As TApplicationProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE

    ' Set properties
    ' --------------
    Call SetVBAProjectProperties
    Call SetDocumentProperties

    ' Add the ECM document theme
    ' --------------------------
    sFileName = Environ("APPDATA") _
              & "\Microsoft\Templates\Document Themes\" _
              & sTHEME

    If FileExists(sFileName) Then
        ThisWorkbook.ApplyTheme (sFileName)

    Else
        sFileName = ThisWorkbook.Path & "\" & sTHEME

        If FileExists(sFileName) Then
            ThisWorkbook.ApplyTheme (sFileName)
        End If
    End If
    
    Set oPB = New FProgressBar

    oPB.Show

    Call DeleteSheets(oPB)
    Call AddSheets(oPB)

    If LoadData Then
        Call FormatSheets(oPB, True, True)
    Else
        Call FormatSheets(oPB)
    End If

    Call ActivateTab(1)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    oPB.Hide
    Set oPB = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

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

Public Sub SetDocumentProperties()
' ==========================================================================
' Description : Set the document properties from the version info
' ==========================================================================

    Const sPROC As String = "SetDocumentProperties"

    Dim oVer    As CVersion


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Set oVer = New CVersion

    Call SetDocumentProperty(gsDOCPROP_AUTHOR, _
                             dpgBuiltIn, _
                             gsAPP_NAME & " Team")
    Call SetDocumentProperty(gsDOCPROP_COMPANY, _
                             dpgBuiltIn, _
                             gsAPP_COMPANY)

    Call SetDocumentProperty(gsDOCPROP_KEYWORDS, _
                             dpgBuiltIn, _
                             gsAPP_CODE)
    Call SetDocumentProperty(gsDOCPROP_TITLE, _
                             dpgBuiltIn, _
                             gsAPP_NAME)

    Call SetDocumentProperty(gsDOCPROP_VERSION, _
                             dpgBuiltIn, _
                             oVer.ProductVersion)
    Call SetDocumentProperty(gsDOCPROP_COMMENTS, _
                             dpgBoth, _
                             "Version " & oVer.ProductVersion & vbNewLine _
                             & oVer.BuildDate)

    Call SetDocumentProperty(oVer.ProductName, _
                             dpgCustom, _
                             True)
    Call SetDocumentProperty(gsDOCPROP_CST_VERSION, _
                             dpgCustom, _
                             oVer.ProductVersion)
    Call SetDocumentProperty(gsDOCPROP_CST_BUILDDATE, _
                             dpgCustom, _
                             oVer.BuildDate)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set oVer = Nothing

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

Public Sub SetVBAProjectProperties()
' ==========================================================================
' Description : Set the properties for the VBA project
' ==========================================================================

    Const sPROC As String = "SetVBAProjectProperties"


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    With ThisWorkbook.VBProject
        .Name = gsVBA_PROJ
        .Description = gsAPP_NAME & " automation code"
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
