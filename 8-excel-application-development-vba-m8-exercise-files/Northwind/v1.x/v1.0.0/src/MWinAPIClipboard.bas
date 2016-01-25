Attribute VB_Name = "MWinAPIClipboard"
' ==========================================================================
' Module      : MWinAPIClipboard
' Type        : Module
' Description : Support for clipboard operations
' ------------------------------------------------------------------------
' Procedures  : ClearClipboard
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

Private Const msMODULE  As String = "MWinAPIClipboard"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuClipboardFormat
    CF_TEXT = 1
    CF_BITMAP = 2
    CF_METAFILEPICT = 3
    CF_SYLK = 4
    CF_DIF = 5
    CF_TIFF = 6
    CF_OEMTEXT = 7
    CF_DIB = 8
    CF_PALETTE = 9
    CF_PENDATA = 10
    CF_RIFF = 11
    CF_WAVE = 12
    CF_UNICODETEXT = 13
    CF_ENHMETAFILE = 14
    CF_HDROP = 15
    CF_LOCALE = 16
    CF_DIBV5 = 17
    CF_MAX = 18
    CF_OWNERDISPLAY = &H80
    CF_DSPTEXT = &H81
    CF_DSPBITMAP = &H82
    CF_DSPMETAFILEPICT = &H83
    CF_DSPENHMETAFILE = &H8E
    CF_PRIVATEFIRST = &H200     ' "Private" formats don't get GlobalFree()'d
    CF_PRIVATELAST = &H2FF
    CF_GDIOBJFIRST = &H300      ' "GDIOBJ" formats do get DeleteObject()'d
    CF_GDIOBJLAST = &H3FF
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------
' The OpenClipboard function is described on MSDN at
' http://msdn.microsoft.com/en-us/office/ms649048(v=vs.90).aspx
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

Public Function ClearClipboard() As Boolean
' ==========================================================================
' Description : Empty the Windows clipboard.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "ClearClipboard"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = OpenClipboard(0&)
    bRtn = EmptyClipboard
    CloseClipboard

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ClearClipboard = bRtn

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
