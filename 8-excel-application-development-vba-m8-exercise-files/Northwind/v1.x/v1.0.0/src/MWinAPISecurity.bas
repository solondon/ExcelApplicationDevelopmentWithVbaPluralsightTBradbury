Attribute VB_Name = "MWinAPISecurity"
' ==========================================================================
' Module      : MWinAPISecurity
' Type        : Module
' Description : Support for security settings
' --------------------------------------------------------------------------
' Comments    : This module is primarily used by MWinAPIRegistry
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

Public Const ACL_REVISION                   As Byte = (2)

Public Const SECURITY_DESCRIPTOR_REVISION   As Byte = (1)
Public Const SECURITY_DESCRIPTOR_MIN_LENGTH As Long = (20)

' Standard access types are defined in WinNT.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa379607(v=vs.85).aspx and
' http://msdn.microsoft.com/en-us/library/aa374892(v=vs.85).aspx
' -----------------------------------
Public Const Delete                         As Long = &H10000
Public Const READ_CONTROL                   As Long = &H20000
Public Const WRITE_DAC                      As Long = &H40000
Public Const WRITE_OWNER                    As Long = &H80000
Public Const SYNCHRONIZE                    As Long = &H100000

Public Const STANDARD_RIGHTS_REQUIRED       As Long = &HF0000

Public Const STANDARD_RIGHTS_READ           As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE          As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE        As Long = READ_CONTROL

Public Const STANDARD_RIGHTS_ALL            As Long = Delete Or READ_CONTROL Or WRITE_DAC Or WRITE_OWNER Or SYNCHRONIZE
Public Const SPECIFIC_RIGHTS_ALL            As Long = &HFFFF

Public Const GENERIC_READ                   As Long = &H80000000
Public Const GENERIC_WRITE                  As Long = &H40000000
Public Const GENERIC_EXECUTE                As Long = &H20000000
Public Const GENERIC_ALL                    As Long = &H10000000

' ----------------
' Module Level
' ----------------

Private Const msMODULE                      As String = "MWinAPISecurity"

' -----------------------------------
' Type declarations
' -----------------------------------
' Global Level
' ----------------

' The ACL structure is the header of an access control list (ACL).
' A complete ACL consists of an ACL structure followed by an
' ordered list of zero or more access control entries (ACEs).
' The ACL structure is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa374931(v=vs.85).aspx
' -----------------------------------
Type TACL
    AclRevision                 As Byte
    Sbz1                        As Byte
    AclSize                     As Integer
    AceCount                    As Integer
    Sbz2                        As Integer
End Type

' The SECURITY_ATTRIBUTES structure contains the
' security descriptor for an object and specifies whether the
' handle retrieved by specifying this structure is inheritable.
' The SECURITY_ATTRIBUTES structure is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa379560(v=vs.85).aspx
' --------------------------------------------------------------
#If VBA7 Then
    Type TSECURITY_ATTRIBUTES
        nLength                 As Long
        lpSecurityDescriptor    As LongPtr
        bInheritHandle          As Boolean
    End Type
#Else
    Type TSECURITY_ATTRIBUTES
        nLength                 As Long
        lpSecurityDescriptor    As Long
        bInheritHandle          As Boolean
    End Type
#End If

' The SECURITY_DESCRIPTOR structure contains the
' security information associated with an object.
' The SECURITY_DESCRIPTOR structure is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa379561(v=vs.85).aspx
' --------------------------------------------------------------
#If VBA7 Then
    Type TSECURITY_DESCRIPTOR
        Revision                As Byte
        Sbz1                    As Byte
        Control                 As Integer
        Owner                   As LongPtr
        Group                   As LongPtr
        SACL                    As TACL
        DACL                    As TACL
    End Type
#Else
    Type TSECURITY_DESCRIPTOR
        Revision                As Byte
        Sbz1                    As Byte
        Control                 As Integer
        Owner                   As Long
        Group                   As Long
        SACL                    As TACL
        DACL                    As TACL
    End Type
#End If
