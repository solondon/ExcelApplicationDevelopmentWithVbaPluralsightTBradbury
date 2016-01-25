Attribute VB_Name = "MMSFormsUserForm"
' ==========================================================================
' Module      : MMSFormsUserForm
' Type        : Module
' Description :
' --------------------------------------------------------------------------
' Properties  : XXX
' --------------------------------------------------------------------------
' Procedures  : XXX
' --------------------------------------------------------------------------
' Events      : XXX
' --------------------------------------------------------------------------
' Dependencies: XXX
' --------------------------------------------------------------------------
' References  : XXX
' --------------------------------------------------------------------------
' Comments    :
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

Public Const GLOBAL_CONST As String = "MMSFormsUserForm"

' ----------------
' Module Level
' ----------------

Private Const msMODULE As String = "MMSFormsUserForm"

Private Const mlPOS_AUTO    As Long = 0

Private Const mlPOS_APP     As Long = 1
Private Const mlPOS_SCREEN  As Long = 2
Private Const mlPOS_VSCREEN As Long = 4

Private Const mlPOS_LEFT    As Long = 256
Private Const mlPOS_RIGHT   As Long = 512
Private Const mlPOS_TOP     As Long = 1024
Private Const mlPOS_BOTTOM  As Long = 2048
Private Const mlPOS_CENTER  As Long = 4096

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuUserFormPositionTo
    ufptAuto = mlPOS_AUTO
    ufptApp = mlPOS_APP
    ufptScreen = mlPOS_SCREEN
    ufptVScreen = mlPOS_VSCREEN
End Enum

Public Enum enuUserFormHPos
    ufhpAuto = mlPOS_AUTO
    ufhpLeft = mlPOS_LEFT
    ufhpCenter = mlPOS_CENTER
    ufhpRight = mlPOS_RIGHT
End Enum

Public Enum enuUserFormVPos
    ufvpAuto = mlPOS_AUTO
    ufvpTop = mlPOS_TOP
    ufvpCenter = mlPOS_CENTER
    ufvpBottom = mlPOS_BOTTOM
End Enum

' -----------------------------------
' Type declarations
' -----------------------------------
' Global Level
' ----------------

'Public Type TPublic
'    PublicID    As Integer
'End Type

' ----------------
' Module Level
' ----------------

'Private Type TPrivate
'    PrivateID   As Integer
'End Type

' -----------------------------------
' Event declarations
' -----------------------------------

'[Public] Event EventName(ByVal Arg As String)

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

'#If VBA7 Then
'    Private Declare PtrSafe _
'            Function <FunctionName> _
'            Lib "user32.dll" _
'            Alias "" (ByVal hWnd As LongPtr) As LongPtr
'#Else
'    Private Declare _
'            Function <FunctionName> _
'            Lib "user32.dll" _
'            Alias "" (ByVal hWnd As Long) As Long
'#End If

' -----------------------------------
' Variable declarations
' -----------------------------------
' Global Level
' ----------------

'Public gsVar    As String

' ----------------
' Module Level
' ----------------

'Private msVar   As String
