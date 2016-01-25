Attribute VB_Name = "MWinAPIKernel32"
' ==========================================================================
' Module      : MWinAPIKernel32
' Type        : Module
' Description : Support for Windows API kernel functions
' --------------------------------------------------------------------------
' Procedures  : CloseHandle         Boolean
'               GetHeapHandle       LongPtr
'               LastDLLErrText      String
'               MemoryAllocate      LongPtr
'               MemoryFree          Boolean
'               Pause
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

Private Const msMODULE                          As String = "MWinAPIKernel32"


' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

' FormatMessage flags
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms679351(v=vs.85).aspx
' -----------------------------------
Public Enum enuFormatMessageFlag
    FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100  '  256
    FORMAT_MESSAGE_IGNORE_INSERTS = &H200   '  512
    FORMAT_MESSAGE_FROM_STRING = &H400      ' 1024
    FORMAT_MESSAGE_FROM_HMODULE = &H800     ' 2048
    FORMAT_MESSAGE_FROM_SYSTEM = &H1000     ' 4096
    FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000  ' 8192
    FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF    '  255
End Enum

Public Enum enuGlobalAllocFlag
    GMEM_FIXED = &H0
    GMEM_MOVEABLE = &H2
    GMEM_ZEROINIT = &H40
    GPTR = &H40         ' GMEM_FIXED OR GMEM_ZEROINIT
    GHND = &H42         ' GMEM_MOVEABLE OR GMEM_ZEROINIT
End Enum

' ----------------
' Module Level
' ----------------

Private Enum enuHeapAllocFlag
    HEAP_NO_SERIALIZE = &H1&
    HEAP_GENERATE_EXCEPTIONS = &H4&
    HEAP_ZERO_MEMORY = &H8&
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The CloseHandle function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724211(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function CloseByHandle _
            Lib "Kernel32" _
            Alias "CloseHandle" (ByVal hObject As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function CloseByHandle _
            Lib "Kernel32" _
            Alias "CloseHandle" (ByVal hObject As Long) _
            As Boolean
#End If

' The FormatMessage function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms679351(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
Private Declare PtrSafe _
        Function FormatMessage _
        Lib "Kernel32" _
        Alias "FormatMessageA" (ByVal dwFlags As enuFormatMessageFlag, _
                                ByVal lpSource As enuFormatMessageFlag, _
                                ByVal dwMessageId As Long, _
                                ByVal dwLanguageId As Long, _
                                ByVal lpBuffer As String, _
                                ByVal nSize As Long, _
                                ByRef Arguments As Long) _
        As Long
#Else
Private Declare _
        Function FormatMessage _
        Lib "Kernel32" _
        Alias "FormatMessageA" (ByVal dwFlags As enuFormatMessageFlag, _
                                ByVal lpSource As enuFormatMessageFlag, _
                                ByVal dwMessageId As Long, _
                                ByVal dwLanguageId As Long, _
                                ByVal lpBuffer As String, _
                                ByVal nSize As Long, _
                                ByRef Arguments As Long) _
        As Long
#End If

' The GetProcessHeap function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366711(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetProcessHeap _
            Lib "Kernel32" () _
            As LongPtr
#Else
    Private Declare _
            Function GetProcessHeap _
            Lib "Kernel32" () _
            As Long
#End If

' The HeapAlloc function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366597(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function HeapAlloc _
            Lib "Kernel32" (ByVal hHeap As LongPtr, _
                            ByVal dwFlags As enuHeapAllocFlag, _
                            ByVal dwBytes As Long) _
            As LongPtr
#Else
    Private Declare _
            Function HeapAlloc _
            Lib "Kernel32" (ByVal hHeap As Long, _
                            ByVal dwFlags As enuHeapAllocFlag, _
                            ByVal dwBytes As Long) _
            As Long
#End If

' The HeapFree function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366701(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function HeapFree _
            Lib "Kernel32" (ByVal hHeap As LongPtr, _
                            ByVal dwFlags As LongPtr, _
                            ByVal lpMem As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function HeapFree _
            Lib "Kernel32" (ByVal hHeap As Long, _
                            ByVal dwFlags As Long, _
                            ByVal lpMem As Long) _
            As Boolean
#End If

' The HeapUnlock function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366707(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function HeapUnlock _
            Lib "Kernel32" (ByVal hHeap As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function HeapUnlock _
            Lib "Kernel32" (ByVal hHeap As Long) _
            As Boolean
#End If

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

' The Sleep function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms686298(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Sub Sleep _
            Lib "Kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare _
            Sub Sleep _
            Lib "Kernel32" (ByVal dwMilliseconds As Long)
#End If

#If VBA7 Then
Public Function CloseHandle(ByVal hObject As LongPtr) As Boolean
#Else
Public Function CloseHandle(ByVal hObject As Long) As Boolean
#End If
' ==========================================================================
' Description : Close an object using its handle
'
' Parameters  : hObject     The handle of the object to close
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "CloseHandle"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = CloseByHandle(hObject)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    CloseHandle = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
      Stop
      Resume
    Else
      Resume PROC_EXIT
    End If

End Function

#If VBA7 Then
Public Function GetHeapHandle() As LongPtr
#Else
Public Function GetHeapHandle() As Long
#End If
' ==========================================================================
' Description : Get the handle of the process heap
'
' Returns     : Long
' ==========================================================================

    Const sPROC     As String = "GetHeapHandle"

    #If VBA7 Then
        Dim lRtn    As LongPtr
    #Else
        Dim lRtn    As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    lRtn = GetProcessHeap()

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetHeapHandle = lRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function LastDLLErrText(ByVal ErrNum As Long) As String
' ==========================================================================
' Description : Retrieve the Windows error text
'
' Parameters  : ErrNum      A Windows error number
'
' Returns     : String
'
' Comments    : There is no error handling due to the likelihood of
'               this function being called from the error handler
' ==========================================================================

    Const sPROC         As String = "LastDLLErrText"
    Const lBUFFER_SIZE  As Long = 255

    Dim sBuffer         As String * lBUFFER_SIZE
    Dim lRtn            As Long

    Dim sRtn            As String

    ' ----------------------------------------------------------------------
    ' Retrieve the error text
    ' -----------------------

    lRtn = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, _
                         0&, _
                         ErrNum, _
                         0, _
                         sBuffer, _
                         lBUFFER_SIZE, _
                         0)

    sRtn = Left$(sBuffer, lRtn)

    ' ----------------------------------------------------------------------

    LastDLLErrText = sRtn

End Function

#If Win64 Then
Public Function MemoryAllocate(ByVal hHeap As LongPtr, _
                               ByVal Bytes As LongLong) As LongPtr
#Else
Public Function MemoryAllocate(ByVal hHeap As Long, _
                               ByVal Bytes As Long) As Long
#End If
' ==========================================================================
' Description : Allocates a block of memory from a heap.
'               The allocated memory is not movable.
'
' Parameters  : hHeap       A handle to the heap from which
'                           the memory will be allocated.
'               Bytes       The number of bytes to be allocated.
'
' Returns     : LongPtr     A pointer to the allocated memory block.
' ==========================================================================

    Const sPROC     As String = "MemoryAllocate"

    #If Win64 Then
        Dim lRtn    As LongPtr
    #Else
        Dim lRtn    As Long
    #End If

    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    lRtn = HeapAlloc(hHeap, 0&, Bytes)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    MemoryAllocate = lRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

#If VBA7 Then
Public Sub MemoryCopy(ByVal Source As LongPtr, _
                      ByVal Destination As LongPtr, _
                      ByVal Length As Long)
#Else
Public Sub MemoryCopy(ByVal Source As Long, _
                      ByVal Destination As Long, _
                      ByVal Length As Long)
#End If
' ==========================================================================
' Description : Copies a block of memory from one location to another.
'
' Parameters  : Source          A pointer to the starting address
'                               of the block of memory to copy.
'               Destination     A pointer to the starting address
'                               of the copied block's destination.
'               Length          The size of the block of memory
'                               to copy, in bytes.
' ==========================================================================

    Const sPROC As String = "MemoryCopy"


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Call CopyMemory(Destination, Source, Length)

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

#If VBA7 Then
Public Function MemoryFree(ByRef hHeap As LongPtr, _
                           ByRef hMemory As LongPtr) As Boolean
#Else
Public Function MemoryFree(ByRef hHeap As Long, _
                           ByRef hMemory As Long) As Boolean
#End If
' ==========================================================================
' Description : Frees a memory block allocated from a heap
'               by the HeapAlloc or HeapReAlloc function.
'
' Parameters  : hHeap       A handle to the heap whose memory block is to
'                           be freed. This handle is returned by either
'                           the HeapCreate or GetProcessHeap function.
'               hMemory     A pointer to the memory block to be freed.
'                           This pointer is returned by the HeapAlloc or
'                           HeapReAlloc function. If this pointer is NULL,
'                           the behavior is undefined.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "MemoryFree"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = HeapFree(hHeap, 0&, hMemory)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    MemoryFree = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

#If VBA7 Then
Public Function MemoryUnlock(ByRef hHeap As LongPtr) As Boolean
#Else
Public Function MemoryUnlock(ByRef hHeap As Long) As Boolean
#End If
' ==========================================================================
' Description : Releases ownership of the critical section object, or lock,
'               that is associated with a specified heap. It reverses the
'               action of the HeapLock function.
'
' Parameters  : hHeap       A handle to the heap to be unlocked.
'                           This handle is returned by either the
'                           HeapCreate or GetProcessHeap function.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "MemoryUnlock"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = HeapUnlock(hHeap)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    MemoryUnlock = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function


Public Sub Pause(ByVal Milliseconds As Long)
' ==========================================================================
' Description : Pause program execution for a number of milliseconds
'
' Parameters  : Milliseconds    The number of milliseconds to pause
' ==========================================================================

    Const sPROC As String = "Pause"

    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Milliseconds)

    ' ----------------------------------------------------------------------

    If (Milliseconds > 0) Then
        Call Sleep(Milliseconds)
    End If

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
