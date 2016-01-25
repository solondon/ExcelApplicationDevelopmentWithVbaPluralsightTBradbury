Attribute VB_Name = "MVBAError"
' ==========================================================================
' Module      : MVBAError
' Type        : Module
' Description : Constants and methods to support error handling
' --------------------------------------------------------------------------
' Procedures  : ErrBox                  VbMsgBoxResult
'               ErrorHandler            Boolean
'               GetErrProperties
'               MsgBoxErr
' --------------------------------------------------------------------------
' Dependencies: CMSOutlookMail
'               FError
'               IFError
'               MMSOffice
'               MMSOfficeDocuments
'               MMSOutlook
'               MVBABitwise
'               MVBADebug
'               MVBALogFile
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuErrorOutput
    eoNull = 0
    eoImmediate = 16
    eoLogFile = 32
    eoEventLog = 64
End Enum

' -----------------------------------
' Type declarations
' -----------------------------------
' Global Level
' ----------------

' The TErrProperties type stores all of
' the scalar properties of the Err object.
' ----------------------------------------
Public Type TErrProperties
    Description                         As String
    HelpContext                         As Long
    HelpFile                            As String
    LastDllError                        As Long
    Number                              As Long
    Source                              As String
End Type

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------
' VBA Error Codes
' http://support.microsoft.com/kb/146864
' --------------------------------------
Public Const ERR_SUCCESS                As Long = 0   ' Generic success code
Public Const ERR_FAILURE                As Long = 1   ' Generic failure code
Public Const ERR_INVALID_PROCEDURE_CALL As Long = 5
Public Const ERR_SUBSCRIPT_OUT_OF_RANGE As Long = 9
Public Const ERR_TYPE_MISMATCH          As Long = 13
Public Const ERR_USER_INTERRUPT         As Long = 18
Public Const ERR_FILE_NOT_FOUND         As Long = 53
Public Const ERR_FILE_EXISTS            As Long = 58
Public Const ERR_OBJ_OR_WITH_NOT_SET    As Long = 91
Public Const ERR_INVALID_PATTERN_STRING As Long = 93
Public Const ERR_OLE_AUTOMATION         As Long = 440
Public Const ERR_ARGUMENT_NOT_OPTIONAL  As Long = 449
Public Const ERR_KEY_EXISTS             As Long = 457

' Run-time error number for custom errors
' ---------------------------------------
Public Const ERR_HANDLED                As Long = vbObjectError + 1
Public Const ERR_INVALID_ASSERTION      As Long = vbObjectError + 2

' ----------------
' Module Level
' ----------------

Private Const msMODULE                  As String = "MVBAError"

Private Const msERR_SILENT              As String = "UserCancel"
Private Const mlERR_OUTPUT              As Long _
                                         = enuErrorOutput.eoImmediate _
                                        Or enuErrorOutput.eoLogFile

Public Function ErrBox(ByRef ErrProps As TErrProperties, _
              Optional ByVal Buttons As VbMsgBoxStyle _
                                      = vbOKOnly _
                                     Or vbDefaultButton1 _
                                     Or vbCritical, _
              Optional ByVal Title As String) As VbMsgBoxResult
' ==========================================================================
' Description : Show the error dialog
'
' Parameters  : ErrProps    The error values
'               Buttons     The dialog style settings
'
' Returns     : VbMsgBoxResult
' ==========================================================================

    Const sPROC     As String = "ErrBox"

    Dim sSubject    As String
    Dim sBody       As String
    Dim sErrFile    As String: sErrFile = GetLogFileName(lftError)

    Dim eRtn        As VbMsgBoxResult
    Dim frm         As IFError


    On Error Resume Next
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Set frm = New FError

    frm.MsgBoxStyle = Buttons
    frm.Procedure = ErrProps.Source
    frm.ErrorNumber = ErrProps.Number
    frm.ErrorDescription = ErrProps.Description

    If (Len(Title) > 0) Then
        frm.Caption = Title
    End If

    frm.Show
    eRtn = frm.DialogResult

    If frm.SendEmail Then
        sSubject = gsAPP_NAME _
                 & " Error " _
                 & ErrProps.Number

        sBody = "Source  : " & ErrProps.Source & vbNewLine _
              & "Error   : " & CStr(ErrProps.Number) & vbNewLine _
              & "Descr   : " & ErrProps.Description & vbNewLine _
              & vbNewLine _
              & "Computer: " & Environ("COMPUTERNAME") & vbNewLine _
              & "User    : " & Environ("USERNAME") & vbNewLine _
              & "File    : " & GetDocumentName(True) & vbNewLine _
              & vbNewLine _
              & "Comments:" & vbNewLine _
              & frm.UserComments

'        Set oOl = New CMSOutlookMail
'        oOl.OpenOutlook
'        Call oOl.SendMail(Recipients:=gsAPP_EMAIL, _
'                          Subject:=sSubject, _
'                          Body:=sBody, _
'                          Attachments:=sErrFile, _
'                          ShowDialog:=True)
'        oOl.CloseOutlook
        Call SendEmailOutlook(gsAPP_EMAIL, _
                              sSubject, _
                              sBody, _
                              olbfPlain, _
                              sErrFile)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ErrBox = eRtn

'    Set oOl = Nothing
    Set frm = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

End Function

Public Function ErrorHandler(ByVal Module As String, _
                             ByVal Procedure As String, _
                    Optional ByVal EntryPoint As Boolean, _
                    Optional ByVal Source As String) As Boolean
' ==========================================================================
' Description : This is the central error handler for the application.
'               Runtime errors that occur during program execution are
'               logged. If the error occurs in an entry-point procedure,
'               an error dialog is presented to the user.
'
' Parameters  : Module      The module where the error occurred.
'
'               Procedure   The procedure in which the error occurred.
'
'               EntryPoint  (Optional) Set to True if this call is being
'                           made from an entry point procedure, such as a
'                           menu item. If True, an error message will
'                           be displayed to the user.
'
'               Source      (Optional) For multiple-document projects,
'                           this is the name of the document in which
'                           the error occurred.
'
' Returns     : Boolean
' ==========================================================================

    Static ssErrMsg As String

    Dim bRtn        As Boolean
    Dim eMBR        As VbMsgBoxResult
    Dim udtErr      As TErrProperties

    ' ----------------------------------------------------------------------
    ' Store the error information before it is
    ' cleared by 'On Error Resume Next' below
    ' ----------------------------------------
    Call GetErrProperties(udtErr, Module, Procedure)

    ' Errors cannot be allowed in the central error handler
    ' -----------------------------------------------------
    On Error Resume Next

    ' If a Source was not provided, use this document
    ' -----------------------------------------------
    If (Len(Source) = 0) Then
        Source = GetDocumentName()
    End If

    ' Clear the message on user interrupt
    ' -----------------------------------
    If (udtErr.Number = ERR_USER_INTERRUPT) Then
        ssErrMsg = msERR_SILENT
    End If

    ' Store the error message in the static variable
    ' ----------------------------------------------
    If (Len(ssErrMsg) = 0) Then
        ssErrMsg = udtErr.Description
    End If

    ' ----------------------------------------------------------------------
    ' Send to logged outputs
    ' ----------------------
    If BitIsSet(mlERR_OUTPUT, enuErrorOutput.eoImmediate) Then
        Call DebugPrint(Module, _
                        Procedure, _
                        "Error " & udtErr.Number & ": " & ssErrMsg)
    End If

    If BitIsSet(mlERR_OUTPUT, enuErrorOutput.eoLogFile) Then
        Call LogFileEntry(lftError, _
                          Module, _
                          Procedure, _
                          ssErrMsg, _
                          udtErr.Number, _
                          Source)
    End If

    If BitIsSet(mlERR_OUTPUT, enuErrorOutput.eoEventLog) Then
        Call LogEvent(Module, _
                      Procedure, _
                      ssErrMsg, _
                      EVENTLOG_ERROR_TYPE, _
                      udtErr.Number)
    End If

    ' Don't display or debug silent errors
    ' ------------------------------------
    If (ssErrMsg <> msERR_SILENT) Then

        ' Show the error message if it is an
        ' entry point procedure or in debug mode
        ' --------------------------------------
        If (EntryPoint Or gbDEBUG_MODE) Then
            Select Case Application.Name
            Case gsOFFICE_APPNAME_EXCEL, gsOFFICE_APPNAME_WORD
                Application.ScreenUpdating = True
            End Select

            eMBR = ErrBox(udtErr)

            ' Clear the error message variable
            ' --------------------------------
            ssErrMsg = vbNullString
        End If

        ' The return value is the debug mode
        ' ----------------------------------
        bRtn = gbDEBUG_MODE

    Else
        ' Clear the error message variable
        ' --------------------------------
        If EntryPoint Then
            ssErrMsg = vbNullString
        End If

        bRtn = False
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ErrorHandler = bRtn

End Function

Public Sub GetErrProperties(ByRef Properties As TErrProperties, _
                   Optional ByVal Module As String, _
                   Optional ByVal Procedure As String)
' ==========================================================================
' Description : Copy the current Err object values to a structure.
'
' Parameters  : Properties  The structure to populate
' ==========================================================================

    With Properties
        .Description = Err.Description
        .HelpContext = Err.HelpContext
        .HelpFile = Err.HelpFile
        .LastDllError = Err.LastDllError
        .Number = Err.Number

        If ((Len(Err.Source) > 0) And (Err.Source <> Application.Name)) Then
            .Source = Err.Source
        ElseIf ((Len(Module) > 0) And (Len(Procedure) > 0)) Then
            .Source = Concat(".", Module, Procedure)
        ElseIf (Len(Module) > 0) Then
            .Source = Module
        ElseIf (Len(Procedure) > 0) Then
            .Source = Procedure
        Else
            .Source = Err.Source
        End If
    End With

End Sub

Public Sub MsgBoxErr(ByRef ErrProperties As TErrProperties)
' ==========================================================================
' Description : Display the error in a message box
'
' Parameters  : ErrProperties   The contents of the error
' ==========================================================================

    Const sPROC     As String = "MsgBoxErr"

    Dim sTitle      As String: sTitle = gsAPP_NAME & " Error"
    Dim sPrompt     As String
    Dim eButtons    As VbMsgBoxStyle: eButtons = vbCritical Or vbOKOnly
    Dim eMBR        As VbMsgBoxResult

    ' ----------------------------------------------------------------------

    With ErrProperties
        sPrompt = "The following error occurred in " _
                  & .Source & ":" & vbNewLine _
                  & vbNewLine _
                  & "Error: " & .Number & vbNewLine _
                  & "Description: " & .Description

        If (Right$(.Description, 1) <> ".") Then
            sPrompt = sPrompt & "."
        End If
    End With

    eMBR = MsgBox(sPrompt, eButtons, sTitle)

End Sub
