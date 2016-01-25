Attribute VB_Name = "MWinAPIRegistryFile"
' ==========================================================================
' Module      : MWinAPIRegistryFile
' Type        : Module
' Description : Support for file information in the Windows Registry
' --------------------------------------------------------------------------
' Procedures  : WindowsFileExtensionsAreHidden      Boolean
' --------------------------------------------------------------------------
' Dependencies: MWinAPIRegistry
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

Private Const msMODULE As String = "MWinAPIRegistryFile"

Public Function WindowsFileExtensionsAreHidden() As Boolean
' ==========================================================================
' Description : Determine whether the Windows setting
'               "Hide extensions for known file types" is enabled.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC         As String = "WindowsFileExtensionsAreHidden"
    Const sSUB_KEY      As String = "Software\Microsoft\Windows\" _
                                  & "CurrentVersion\Explorer\Advanced"
    Const sVALUE_NAME   As String = "HideFileExt"

    Dim bRtn            As Boolean
    Dim lRtn            As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    lRtn = RegGetKeyValue(HKCU, sSUB_KEY, sVALUE_NAME)
    bRtn = (lRtn <> 0)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WindowsFileExtensionsAreHidden = bRtn

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
