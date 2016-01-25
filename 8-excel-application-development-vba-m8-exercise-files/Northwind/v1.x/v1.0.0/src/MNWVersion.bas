Attribute VB_Name = "MNWVersion"
' ==========================================================================
' Module      : MNWVersion
' Type        : Module
' Description : Version constants for use by the CVersion object.
'               These values are abstracted here to allow multiple
'               systems to use CVersion without modification.
' ==========================================================================

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

Public Const gdVER_BUILDDATE    As Date = #4/13/2015#

Public Const glVER_MAJOR        As Long = 1
Public Const glVER_MINOR        As Long = 0
Public Const glVER_PATCH        As Long = 0
Public Const glVER_BUILD        As Long = 1

Public Const gbVER_HOTFIX       As Boolean = False

Public Const gsVER_COPYRIGHT    As String _
       = "Copyright © 2015 Terry L. Bradbury"

Public Const gsVER_DESCR        As String _
       = "Demo application to demonstrate source management."

Public Const gsVER_SUPPORT As String _
             = "This is an unsupported application for Microsoft Excel." _
             & vbNewLine _
             & "This product is provided as-is, without warranty," _
             & vbNewLine _
             & "and any use of this product does not constitute a" _
             & vbNewLine _
             & "contract, real or implied, on the part of the developer."

Public Const gsVER_WARNING As String = vbNullString

'Public Const gsVER_WARNING As String _
'             = "WARNING: This product is protected by copyright law " _
'             & "and international treaties." & vbNewLine _
'             & "Unauthorized reproduction, distribution, decompiling of " _
'             & "this program, or any part of it, may result in severe " _
'             & "civil and criminal penalties, and will be prosecuted " _
'             & "to the maximum extend possible under law."
