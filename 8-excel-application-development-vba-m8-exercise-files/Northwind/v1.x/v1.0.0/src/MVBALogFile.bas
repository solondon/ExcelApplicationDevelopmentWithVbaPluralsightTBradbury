Attribute VB_Name = "MVBALogFile"
' ==========================================================================
' Module      : MVBALogFile
' Type        : Module
' Description : Support for creating log files.
' --------------------------------------------------------------------------
' Procedures  : GetLogFileName      String
'               LogFileEntry
' --------------------------------------------------------------------------
' Dependencies: MMSOfficeDocuments
'               MVBABitwise
'               MVBADebug
'               MVBAFile
'               MVBAStrings
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

Public Const gsLOGFILE_MLSEP    As String = "|"    ' Multi-line separator

' ----------------
' Module Level
' ----------------

Private Const msMODULE          As String = "MVBALogFile"

Private Const msLOGFILE_APP     As String = gsAPP_NAME & ".log"
Private Const msLOGFILE_DB      As String = gsAPP_NAME & " Database.log"
Private Const msLOGFILE_ERROR   As String = gsAPP_NAME & " Error.log"
Private Const msLOGFILE_SETUP   As String = gsAPP_NAME & " Setup.log"
Private Const msLOGFILE_TRACE   As String = gsAPP_NAME & " Trace.log"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuLogFileType
    lftUnknown = 0
    lftApp = 1
    lftError = 2
    lftTrace = 4
    lftDatabase = 16
    lftSetup = 32
End Enum

Public Function GetLogFileName(ByVal LogFileType As enuLogFileType) _
       As String
' ==========================================================================
' Description : Build the fully-qulified file name for the error log.
'
' Parameters  : LogFileType     The type of log file to use.
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "GetLogFileName"

    Dim sMsg    As String
    Dim sPath   As String
    Dim sRtn    As String

    ' Get the application directory
    ' -----------------------------
    sPath = ThisWorkbook.Path

    ' Make sure the path ends with a delimiter
    ' ----------------------------------------
    If (Right$(sPath, 1) <> "\") Then
        sPath = sPath & "\"
    End If

    ' Build the final name by LogFileType
    ' -------------------------------
    Select Case LogFileType
    Case lftApp
        sRtn = sPath & msLOGFILE_APP

    Case lftDatabase
        sRtn = sPath & msLOGFILE_DB

    Case lftError
        sRtn = sPath & msLOGFILE_ERROR

    Case lftTrace
        sRtn = sPath & msLOGFILE_TRACE

    Case lftSetup
        sRtn = sPath & msLOGFILE_SETUP

    Case Else
        sMsg = "Invlalid LogFileType (" & LogFileType & ")."
        Call LogFileEntry(lftError, _
                          msMODULE, _
                          sPROC, _
                          sMsg, _
                          ERR_INVALID_PROCEDURE_CALL)
        Call DebugAssert(False, _
                         msMODULE, _
                         sPROC, _
                         sMsg)
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetLogFileName = sRtn

End Function

Public Sub LogFileEntry(ByVal LogFileType As enuLogFileType, _
                        ByVal Module As String, _
                        ByVal Procedure As String, _
                        ByVal Message As String, _
               Optional ByVal ErrNumber As Long, _
               Optional ByVal Source As String)
' ==========================================================================
' Description : Write a log entry
'
' Parameters  : LogFileType Identifies which log file to write to
'               Module      The name of the calling module.
'               Procedure   The name of the calling procedure.
'               Message     The message to log.
'               ErrNumber   The error number associated with the message.
'                           This value is only used for error log entries.
'               Source      The name of the document in error.
'
' Notes       : There are no trace calls in this method to prevent
'               circular references.
'
'               The global value gbTRACE_STACK will be used to indicate
'               if a simulated stack operation is desired for trace logs.
' ==========================================================================

    Const sPROC     As String = "LogFileEntry"

    Dim bFileOpen   As Boolean

    Dim iFileNum    As Integer

    Dim lPos        As Long

    Dim sComputer   As String
    Dim sUser       As String
    Dim sHeader     As String
    Dim sLogFile    As String
    Dim sLogEntry   As String


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Get the log file name
    ' ---------------------
    sLogFile = GetLogFileName(LogFileType)

    ' Get the computer name
    ' ---------------------
    sComputer = Environ("COMPUTERNAME")

    ' Get the user name
    ' -----------------
    sUser = Environ("USERNAME")

    ' Get the default Workbook name if needed
    ' ---------------------------------------
    If (Len(Source) = 0) Then
        Source = GetDocumentName()
    End If

    ' If the log does not exist create
    ' an empty one with a header row
    ' --------------------------------
    iFileNum = FreeFile()

    If Not FileExists(sLogFile) Then

        ' Build the header
        ' ----------------
        If BitIsSet(LogFileType, lftError) Then

            ' Only error logs have an error number column
            ' -------------------------------------------
            sHeader = Concat(vbTab, _
                             "Date", _
                             "Time", _
                             "Computer", _
                             "User", _
                             "Source", _
                             "Module", _
                             "Procedure", _
                             "Message", _
                             "Error Number")

        ElseIf BitIsSet(LogFileType, lftTrace) Then
            If gbTRACE_STACK Then
                sHeader = Concat(vbTab, _
                                 "Date", _
                                 "Time", _
                                 "Computer", _
                                 "User", _
                                 "Source", _
                                 "Module", _
                                 "Procedure", _
                                 "Message", _
                                 "Level")
            Else
                sHeader = Concat(vbTab, _
                                 "Date", _
                                 "Time", _
                                 "Computer", _
                                 "User", _
                                 "Source", _
                                 "Module", _
                                 "Procedure", _
                                 "Message")
            End If
        Else
            sHeader = Concat(vbTab, _
                             "Date", _
                             "Time", _
                             "Computer", _
                             "User", _
                             "Source", _
                             "Module", _
                             "Procedure", _
                             "Message")
        End If

        Open sLogFile For Append As #iFileNum
        bFileOpen = True
        Print #iFileNum, sHeader
    End If

    ' File previously existed
    ' Open and add the log entry
    ' --------------------------
    If (Not bFileOpen) Then
        Open sLogFile For Append As #iFileNum
    End If

    ' Get rid of any embedded CR+LF combinations
    ' ------------------------------------------
    lPos = InStr(1, Message, vbCrLf, vbTextCompare)
    Do While (lPos > 0)
        Message = Replace(Message, vbCrLf, gsLOGFILE_MLSEP)
        lPos = InStr(1, Message, vbCrLf, vbTextCompare)
    Loop

    ' Get rid of any singleton LF characters
    ' --------------------------------------
    lPos = InStr(1, Message, vbLf, vbTextCompare)
    Do While (lPos > 0)
        Message = Replace(Message, vbLf, gsLOGFILE_MLSEP)
        lPos = InStr(1, Message, vbLf, vbTextCompare)
    Loop

    ' Get rid of any singleton CR characters
    ' --------------------------------------
    lPos = InStr(1, Message, vbCr, vbTextCompare)
    Do While (lPos > 0)
        Message = Replace(Message, vbCr, gsLOGFILE_MLSEP)
        lPos = InStr(1, Message, vbCr, vbTextCompare)
    Loop

    ' Build the log entry
    ' -------------------
    If BitIsSet(LogFileType, lftError) Then
        sLogEntry = Concat(vbTab, _
                           Format$(Now(), "mm/dd/yy"), _
                           Format$(Now(), "hh:mm:ss AM/PM"), _
                           sComputer, _
                           sUser, _
                           Source, _
                           Module, _
                           Procedure, _
                           Message, _
                           ErrNumber)

    ElseIf BitIsSet(LogFileType, lftTrace) Then
        If gbTRACE_STACK Then
            sLogEntry = Concat(vbTab, _
                               Format$(Now(), "mm/dd/yy"), _
                               Format$(Now(), "hh:mm:ss AM/PM"), _
                               sComputer, _
                               sUser, _
                               Source, _
                               Module, _
                               Procedure, _
                               Message, _
                               ErrNumber)
        Else
            sLogEntry = Concat(vbTab, _
                               Format$(Now(), "mm/dd/yy"), _
                               Format$(Now(), "hh:mm:ss AM/PM"), _
                               sComputer, _
                               sUser, _
                               Source, _
                               Module, _
                               Procedure, _
                               Message)
        End If
    Else
        sLogEntry = Concat(vbTab, _
                           Format$(Now(), "mm/dd/yy"), _
                           Format$(Now(), "hh:mm:ss AM/PM"), _
                           sComputer, _
                           sUser, _
                           Source, _
                           Module, _
                           Procedure, _
                           Message)
    End If

    ' Add it to the log
    ' -----------------
    Print #iFileNum, sLogEntry

    ' Close the log file
    ' ------------------
    Close #iFileNum

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
