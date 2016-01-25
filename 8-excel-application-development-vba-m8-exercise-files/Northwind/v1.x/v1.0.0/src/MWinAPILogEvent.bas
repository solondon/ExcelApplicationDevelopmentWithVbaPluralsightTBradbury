Attribute VB_Name = "MWinAPILogEvent"
' ==========================================================================
' Module      : MVBALogEvent
' Type        : Module
' Description : Support for the Event Log
' --------------------------------------------------------------------------
' Procedures  : LogEvent
' --------------------------------------------------------------------------
' Dependencies: MWinAPIKernel32
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

Private Const msMODULE                  As String = "MVBALogEvent"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuEventLogType
    EVENTLOG_SUCCESS = &H0
    EVENTLOG_ERROR_TYPE = &H1
    EVENTLOG_WARNING_TYPE = &H2
    EVENTLOG_INFORMATION_TYPE = &H4
    EVENTLOG_AUDIT_SUCCESS = &H8
    EVENTLOG_AUDIT_FAILURE = &HA    ' EVENTLOG_WARNING_TYPE Or EVENTLOG_AUDIT_SUCCESS
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The RegisterEventSource function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa363654(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegisterEventSource _
            Lib "advapi32" _
            Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, _
                                          ByVal lpSourceName As String) _
            As LongPtr
#Else
    Private Declare _
            Function RegisterEventSource _
            Lib "advapi32" _
            Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, _
                                          ByVal lpSourceName As String) _
            As Long
#End If

' The DeregisterEventSource function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa363642(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function DeregisterEventSource _
            Lib "advapi32" (ByVal hEventLog As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function DeregisterEventSource _
            Lib "advapi32" (ByVal hEventLog As Long) _
            As Boolean
#End If

' The ReportEvent function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa363679(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function ReportEvent _
            Lib "advapi32" _
            Alias "ReportEventA" (ByVal hEventLog As LongPtr, _
                                  ByVal wType As enuEventLogType, _
                                  ByVal wCategory As Integer, _
                                  ByVal dwEventID As LongPtr, _
                                  ByVal lpUserSid As Any, _
                                  ByVal wNumStrings As Integer, _
                                  ByVal dwDataSize As LongPtr, _
                                  ByRef plpStrings As LongPtr, _
                                  ByRef lpRawData As Any) _
            As Boolean
#Else
    Private Declare _
            Function ReportEvent _
            Lib "advapi32" _
            Alias "ReportEventA" (ByVal hEventLog As Long, _
                                  ByVal wType As enuEventLogType, _
                                  ByVal wCategory As Integer, _
                                  ByVal dwEventID As Long, _
                                  ByVal lpUserSid As Any, _
                                  ByVal wNumStrings As Integer, _
                                  ByVal dwDataSize As Long, _
                                  ByRef plpStrings As Long, _
                                  ByRef lpRawData As Any) _
            As Boolean
#End If

' The RtlMoveMemory routine is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ff562030(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Sub CopyMemory _
            Lib "Kernel32" _
            Alias "RtlMoveMemory" (ByRef hpvDest As Any, _
                                   ByRef hpvSource As Any, _
                                   ByVal cbCopy As Long)
#Else
    Private Declare _
            Sub CopyMemory _
            Lib "Kernel32" _
            Alias "RtlMoveMemory" (ByRef hpvDest As Any, _
                                   ByRef hpvSource As Any, _
                                   ByVal cbCopy As Long)
#End If

Public Sub LogEvent(ByVal Module As String, _
                    ByVal Procedure As String, _
                    ByVal Message As String, _
                    ByVal LogType As enuEventLogType, _
                    ByVal EventID As Long)
' ==========================================================================
' Description : Write an entry to the Event Log
'
' Parameters  : Module      The name of the originating module
'               Procedure   The name of the originating procedure
' ==========================================================================

    Const sPROC         As String = "LogNTEvent"

    #If VBA7 Then
        Dim hHeap       As LongPtr
        Dim hEventLog   As LongPtr
        Dim hMem        As LongPtr
        Dim hMsgs       As LongPtr

        Dim lDataSize   As LongPtr
        Dim lUserSid    As LongPtr
    #Else
        Dim hHeap       As Long
        Dim hEventLog   As Long
        Dim hMem        As Long
        Dim hMsgs       As Long

        Dim lDataSize   As Long
        Dim lUserSid    As Long
    #End If


    Dim bRtn            As Boolean

    Dim iCategory       As Integer
    Dim iNumStrings     As Integer

    Dim sUNCServerName  As String: sUNCServerName = vbNullString
    Dim sSourceName     As String: sSourceName = gsAPP_NAME


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Build the message
    ' -----------------

    Message = Concat(".", Module, Procedure) & vbNewLine & Message

    ' Get the size of the message + 1 (null terminator)
    ' -------------------------------------------------
    lDataSize = Len(Message) + 1

    ' Allocate memory
    ' ---------------
    hHeap = GetHeapHandle()
    hMsgs = MemoryAllocate(hHeap, lDataSize)

    ' Copy the message to memory
    ' --------------------------
    CopyMemory ByVal hMsgs, ByVal Message, lDataSize

    ' Identify that only 1 message will be passed
    ' -------------------------------------------
    iNumStrings = 1

    ' Open the Event Log
    ' ------------------
    hEventLog = RegisterEventSource(sUNCServerName, sSourceName)

    ' Write to the Event Log
    ' ----------------------
    bRtn = ReportEvent(hEventLog, _
                       LogType, _
                       iCategory, _
                       EventID, _
                       0&, _
                       iNumStrings, _
                       lDataSize, _
                       hMsgs, _
                       hMsgs)

    ' Close the Event Log
    ' -------------------
    bRtn = DeregisterEventSource(hEventLog)

    ' Free memory
    ' -----------
    bRtn = MemoryFree(hHeap, hMsgs)

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
