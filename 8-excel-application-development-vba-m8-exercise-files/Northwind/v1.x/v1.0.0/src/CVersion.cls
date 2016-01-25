VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' Module      : CVersion
' Type        : Class Module
' Description : Provide version information to the application
' ------------------------------------------------------------------------
' Properties  : Build           (Get)   Long
'               BuildDate       (Get)   Date
'               Copyright       (Get)   String
'               Description     (Get)   String
'               IsHotfix        (Get)   Boolean
'               Major           (Get)   Long
'               Minor           (Get)   Long
'               Patch           (Get)   Long
'               ProductName     (Get)   String
'               ProductVersion  (Get)   String
'               Support         (Get)   String
'               Warning         (Get)   String
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const mb_Hotfix         As Boolean = gbVER_HOTFIX

Private Const ml_Major          As Long = glVER_MAJOR
Private Const ml_Minor          As Long = glVER_MINOR
Private Const ml_Patch          As Long = glVER_PATCH
Private Const ml_Build          As Long = glVER_BUILD

Private Const md_BuildDate      As Date = gdVER_BUILDDATE

Private Const ms_ProductName    As String = gsAPP_NAME
Private Const ms_Copyright      As String = gsVER_COPYRIGHT
Private Const ms_Description    As String = gsVER_DESCR
Private Const ms_Support        As String = gsVER_SUPPORT
Private Const ms_Warning        As String = gsVER_WARNING

Public Property Get Build() As Long
' ==========================================================================

    Build = ml_Build

End Property

Public Property Get BuildDate() As Date
' ==========================================================================

    BuildDate = md_BuildDate

End Property

Public Property Get Copyright() As String
' ==========================================================================

    Copyright = ms_Copyright

End Property

Public Property Get Description() As String
' ==========================================================================

    Description = ms_Description

End Property

Public Property Get IsHotfix() As Boolean
' ==========================================================================

    IsHotfix = mb_Hotfix

End Property

Public Property Get Major() As Long
' ==========================================================================

    Major = ml_Major

End Property

Public Property Get Minor() As Long
' ==========================================================================

    Minor = ml_Minor

End Property

Public Property Get Patch() As Long
' ==========================================================================

    Patch = ml_Patch

End Property

Public Property Get ProductName() As String
' ==========================================================================

    ProductName = ms_ProductName

End Property

Public Property Get ProductVersion() As String
' ==========================================================================

    Dim sPV As String

    sPV = Me.Major & "." _
        & Me.Minor & "." _
        & Me.Patch & "." _
        & Me.Build

    If Me.IsHotfix Then
        sPV = sPV & " [HOTFIX]"
    End If
    
    ProductVersion = sPV

End Property

Public Property Get Support() As String
' ==========================================================================

    Support = ms_Support

End Property

Public Property Get Warning() As String
' ==========================================================================

    Warning = ms_Warning

End Property
