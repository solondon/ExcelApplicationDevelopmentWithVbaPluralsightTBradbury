Attribute VB_Name = "MWinAPIGUID"
' ==========================================================================
' Module      : MWinAPIGUID
' Type        : Module
' Description : Support for GUIDs
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

Private Const msMODULE  As String = "MWinAPIGUID"

' -----------------------------------
' Type declarations
' -----------------------------------
' Global Level
' ----------------

' The GUID structure is defined in Guiddef.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa373931(v=vs.85).aspx
' -----------------------------------
Public Type TGUID
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(7)            As Byte
End Type
