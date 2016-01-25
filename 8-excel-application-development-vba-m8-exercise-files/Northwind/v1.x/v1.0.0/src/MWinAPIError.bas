Attribute VB_Name = "MWinAPIError"
' ==========================================================================
' Module      : MWinAPIError
' Type        : Module
' Description : Support for working with Windows errors
' --------------------------------------------------------------------------
' Procedures  : HResultErrorToString
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

' HRESULT return values from WinError.h
' http:'msdn.microsoft.com/en-us/office/aa378137(v=vs.90).aspx
' http:'support.microsoft.com/kb/186063
'  -------------------------------------
Public Const S_OK                           As Long = &H0    ' Operation successful
Public Const S_FALSE                        As Long = &H1
Public Const E_NOTIMPL                      As Long = &H80004001    ' Not implemented
Public Const E_NOINTERFACE                  As Long = &H80004002    ' No such interface supported
Public Const E_POINTER                      As Long = &H80004003    ' Pointer that is not valid
Public Const E_ABORT                        As Long = &H80004004    ' Operation aborted
Public Const E_FAIL                         As Long = &H80004005    ' Unspecified failure
Public Const E_UNEXPECTED                   As Long = &H8000FFFF    ' Unexpected failure
Public Const E_ACCESSDENIED                 As Long = &H80070005    ' General access denied error
Public Const E_HANDLE                       As Long = &H80070006    ' Handle that is not valid
Public Const E_OUTOFMEMORY                  As Long = &H8007000E    ' Failed to allocate necessary memory
Public Const E_INVALIDARG                   As Long = &H80070057    ' One or more arguments are not valid

' General error codes from WinError.h
' http://msdn.microsoft.com/en-us/library/ms681381(v=vs.85).aspx
'  -------------------------------------
Public Const ERROR_SUCCESS                  As Long = 0
Public Const ERROR_INVALID_FUNCTION         As Long = 1
Public Const ERROR_FILE_NOT_FOUND           As Long = 2
Public Const ERROR_PATH_NOT_FOUND           As Long = 3
Public Const ERROR_TOO_MANY_OPEN_FILES      As Long = 4
Public Const ERROR_ACCESS_DENIED            As Long = 5
Public Const ERROR_INVALID_HANDLE           As Long = 6
Public Const ERROR_ARENA_TRASHED            As Long = 7
Public Const ERROR_NOT_ENOUGH_MEMORY        As Long = 8
Public Const ERROR_INVALID_BLOCK            As Long = 9
Public Const ERROR_BAD_ENVIRONMENT          As Long = 10
Public Const ERROR_BAD_FORMAT               As Long = 11
Public Const ERROR_INVALID_ACCESS           As Long = 12
Public Const ERROR_INVALID_DATA             As Long = 13
Public Const ERROR_OUTOFMEMORY              As Long = 14
Public Const ERROR_INVALID_DRIVE            As Long = 15
Public Const ERROR_CURRENT_DIRECTORY        As Long = 16
Public Const ERROR_NOT_SAME_DEVICE          As Long = 17
Public Const ERROR_NO_MORE_FILES            As Long = 18
Public Const ERROR_WRITE_PROTECT            As Long = 19
Public Const ERROR_BAD_UNIT                 As Long = 20
Public Const ERROR_NOT_READY                As Long = 21
Public Const ERROR_BAD_COMMAND              As Long = 22
Public Const ERROR_CRC                      As Long = 23
Public Const ERROR_BAD_LENGTH               As Long = 24
Public Const ERROR_SEEK                     As Long = 25
Public Const ERROR_NOT_DOS_DISK             As Long = 26
Public Const ERROR_SECTOR_NOT_FOUND         As Long = 27
Public Const ERROR_OUT_OF_PAPER             As Long = 28
Public Const ERROR_WRITE_FAULT              As Long = 29
Public Const ERROR_READ_FAULT               As Long = 30
Public Const ERROR_GEN_FAILURE              As Long = 31
Public Const ERROR_SHARING_VIOLATION        As Long = 32
Public Const ERROR_LOCK_VIOLATION           As Long = 33
Public Const ERROR_WRONG_DISK               As Long = 34
Public Const ERROR_SHARING_BUFFER_EXCEEDED  As Long = 36
Public Const ERROR_HANDLE_EOF               As Long = 38
Public Const ERROR_HANDLE_DISK_FULL         As Long = 39
Public Const ERROR_NOT_SUPPORTED            As Long = 50
Public Const ERROR_REM_NOT_LIST             As Long = 51
Public Const ERROR_DUP_NAME                 As Long = 52
Public Const ERROR_BAD_NETPATH              As Long = 53
Public Const ERROR_NETWORK_BUSY             As Long = 54
Public Const ERROR_DEV_NOT_EXIST            As Long = 55
Public Const ERROR_TOO_MANY_CMDS            As Long = 56
Public Const ERROR_ADAP_HDW_ERR             As Long = 57
Public Const ERROR_BAD_NET_RESP             As Long = 58
Public Const ERROR_UNEXP_NET_ERR            As Long = 59
Public Const ERROR_BAD_REM_ADAP             As Long = 60
Public Const ERROR_PRINTQ_FULL              As Long = 61
Public Const ERROR_NO_SPOOL_SPACE           As Long = 62
Public Const ERROR_PRINT_CANCELLED          As Long = 63
Public Const ERROR_NETNAME_DELETED          As Long = 64
Public Const ERROR_NETWORK_ACCESS_DENIED    As Long = 65
Public Const ERROR_BAD_DEV_TYPE             As Long = 66
Public Const ERROR_BAD_NET_NAME             As Long = 67
Public Const ERROR_TOO_MANY_NAMES           As Long = 68
Public Const ERROR_TOO_MANY_SESS            As Long = 69
Public Const ERROR_SHARING_PAUSED           As Long = 70
Public Const ERROR_REQ_NOT_ACCEP            As Long = 71
Public Const ERROR_REDIR_PAUSED             As Long = 72
Public Const ERROR_FILE_EXISTS              As Long = 80
Public Const ERROR_CANNOT_MAKE              As Long = 82
Public Const ERROR_FAIL_I24                 As Long = 83
Public Const ERROR_OUT_OF_STRUCTURES        As Long = 84
Public Const ERROR_ALREADY_ASSIGNED         As Long = 85
Public Const ERROR_INVALID_PASSWORD         As Long = 86
Public Const ERROR_INVALID_PARAMETER        As Long = 87
Public Const ERROR_NET_WRITE_FAULT          As Long = 88
Public Const ERROR_NO_PROC_SLOTS            As Long = 89

' ----------------
' Module Level
' ----------------

Private Const msMODULE                      As String = "MWinAPIError"

Public Function HResultErrorToString(ByVal ErrNumber As Long) As String
' ==========================================================================
' Description : Gets the message text for HRESULT errors
'
' Parameters  : ErrNumber   The error number returned by the OLE interface
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "HResultErrorToString"

    Dim sRtn    As String


    Call Trace(tlNormal, msMODULE, sPROC, ErrNumber)

    ' ----------------------------------------------------------------------

    Select Case ErrNumber
    Case S_OK
        sRtn = "Operation successful"

    Case E_ABORT
        sRtn = "Operation aborted"

    Case E_ACCESSDENIED
        sRtn = "Access denied"

    Case E_FAIL
        sRtn = "General failure"

    Case E_HANDLE
        sRtn = "Invalid Handle"

    Case E_INVALIDARG
        sRtn = "Invalid Argument"

    Case E_NOINTERFACE
        sRtn = "The object does not support the " _
             & "interface specified in riid."

    Case E_NOTIMPL
        sRtn = "Not implemented"

    Case E_OUTOFMEMORY
        sRtn = "Out of memory"

    Case E_POINTER
        sRtn = "Invalid Pointer"

    Case E_UNEXPECTED
        sRtn = "Unknown error"

    Case Else
        sRtn = "Unhandled error"

    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    HResultErrorToString = sRtn

    On Error GoTo 0
    Call Trace(tlNormal, msMODULE, sPROC, sRtn)

End Function
