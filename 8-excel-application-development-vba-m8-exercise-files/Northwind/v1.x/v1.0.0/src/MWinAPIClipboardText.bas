Attribute VB_Name = "MWinAPIClipboardText"
' ==========================================================================
' Module      : MWinAPIClipboardText
' Type        : Module
' Description : Support for text on the Windows clipboard
' --------------------------------------------------------------------------
' Procedures  : SetClipboardText
' --------------------------------------------------------------------------
' Dependencies: MWinAPIClipboard
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

Private Const msMODULE  As String = "MWinAPIClipboardText"

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The GlobalAlloc function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366574(v=vs.85).aspx
' -----------------------------------
#If VAB7 Then
Private Declare PtrSafe _
        Function GlobalAlloc _
        Lib "Kernel32" (ByVal wFlags As enuGlobalAllocFlag, _
                        ByVal dwBytes As Long) _
        As LongPtr
#Else
Private Declare _
        Function GlobalAlloc _
        Lib "Kernel32" (ByVal wFlags As enuGlobalAllocFlag, _
                        ByVal dwBytes As Long) _
        As Long
#End If

' The GlobalLock function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366584(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
Private Declare PtrSafe _
        Function GlobalLock _
        Lib "Kernel32" (ByVal hMem As LongPtr) _
        As LongPtr
#Else
Private Declare _
        Function GlobalLock _
        Lib "Kernel32" (ByVal hMem As Long) _
        As Long
#End If

' The GlobalUnlock function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366595(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
Private Declare PtrSafe _
        Function GlobalUnlock _
        Lib "Kernel32" (ByVal hMem As LongPtr) _
        As Long
#Else
Private Declare _
        Function GlobalUnlock _
        Lib "Kernel32" (ByVal hMem As Long) _
        As Long
#End If

' The lstrcpy function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms647490(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
Private Declare PtrSafe _
        Function lstrcpy _
        Lib "Kernel32" (ByVal lpString1 As Any, _
                        ByVal lpString2 As Any) _
        As LongPtr
#Else
Private Declare _
        Function lstrcpy _
        Lib "Kernel32" (ByVal lpString1 As Any, _
                        ByVal lpString2 As Any) _
        As Long
#End If

' The OpenClipboard function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms649048(v=vs.90).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function OpenClipboard _
            Lib "User32" (ByVal hWnd As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function OpenClipboard _
            Lib "User32" (ByVal hWnd As Long) _
            As Boolean
#End If

' The EmptyClipboard function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms649037(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function EmptyClipboard _
            Lib "User32" () _
            As Boolean
#Else
    Private Declare _
            Function EmptyClipboard _
            Lib "User32" () _
            As Boolean
#End If

' The SetClipboardData function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms649051(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SetClipboardData _
            Lib "User32" (ByVal wFormat As enuClipboardFormat, _
                          ByVal hMem As LongPtr) As LongPtr
#Else
    Private Declare _
            Function SetClipboardData _
            Lib "User32" (ByVal wFormat As enuClipboardFormat, _
                          ByVal hMem As Long) As Long
#End If

' The CloseClipboard function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms649035(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function CloseClipboard _
            Lib "User32" () _
            As Boolean
#Else
    Private Declare _
            Function CloseClipboard _
            Lib "User32" () _
            As Boolean
#End If

Public Function SetClipboardText(ByRef Text As String) As Boolean
' ==========================================================================
' Description : Copy a text string to the Windows clipboard
'
' Parameters  : Text        The string to copy
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC         As String = "SetClipboardText"

    Dim bRtn            As Boolean

    #If VBA7 Then
        Dim hMemGlob    As LongPtr
        Dim lpMemGlob   As LongPtr
        Dim hMemClip    As LongPtr
    #Else
        Dim hMemGlob    As Long
        Dim lpMemGlob   As Long
        Dim hMemClip    As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Allocate moveable global memory
    '-------------------------------------------
    hMemGlob = GlobalAlloc(GHND, Len(Text) + 1)

    ' Get a far pointer to this memory
    ' --------------------------------
    lpMemGlob = GlobalLock(hMemGlob)

    ' Copy the string to global memory
    ' --------------------------------
    lpMemGlob = lstrcpy(lpMemGlob, Text)

    ' Unlock the memory
    ' -----------------
    If (GlobalUnlock(hMemGlob) <> 0) Then
        GoTo CLOSE_CLIPBOARD
    End If

    ' Open the Clipboard
    ' ------------------
    bRtn = OpenClipboard(0&)

    ' Quit if unsuccessful
    ' --------------------
    If (bRtn = False) Then
        GoTo PROC_EXIT
    End If

    ' Clear the contents
    ' ------------------
    bRtn = EmptyClipboard()

    ' Put the string on the clipboard
    ' -------------------------------
    hMemClip = SetClipboardData(CF_TEXT, hMemGlob)

CLOSE_CLIPBOARD:

    bRtn = CloseClipboard

    ' ----------------------------------------------------------------------

PROC_EXIT:

    SetClipboardText = bRtn

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
