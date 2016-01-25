Attribute VB_Name = "MWinAPI"
' ==========================================================================
' Module      : MWinAPI
' Type        : Module
' Description : Support for the Windows API
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

Public Const MAX_PATH   As Long = 260

' ----------------
' Module Level
' ----------------

Private Const msMODULE  As String = "MWinAPI"

' -----------------------------------
' Type declarations
' -----------------------------------
' Global Level
' ----------------

' The POINT structure is defined in Windef.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd162805(v=vs.85).aspx
' -----------------------------------
Public Type TPOINT
    X                   As Long
    Y                   As Long
End Type

' The RECT structure is defined in Windef.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd162897(v=vs.85).aspx
' -----------------------------------
Public Type TRECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type
