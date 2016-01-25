Attribute VB_Name = "MMSOffice"
' ==========================================================================
' Module      : MMSOffice
' Type        : Module
' Description : Microsoft Office constants
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuOfficeVersion
    OfficeVersion95 = 7
    OfficeVersion97 = 8
    OfficeVersion2000 = 9
    OfficeVersionXP = 10
    OfficeVersion2003 = 11
    OfficeVersion2007 = 12
    OfficeVersion2010 = 14
    OfficeVersion2013 = 15
End Enum

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

Public Const gsOFFICE_APPNAME_ACCESS        As String = "Microsoft Access"
Public Const gsOFFICE_APPNAME_EXCEL         As String = "Microsoft Excel"
Public Const gsOFFICE_APPNAME_OUTLOOK       As String = "Outlook"
Public Const gsOFFICE_APPNAME_POWERPOINT    As String = "Microsoft PowerPoint"
Public Const gsOFFICE_APPNAME_PROJECT       As String = "Microsoft Project"
Public Const gsOFFICE_APPNAME_PUBLISHER     As String = "Microsoft Publisher"
Public Const gsOFFICE_APPNAME_VISIO         As String = "Microsoft Visio"
Public Const gsOFFICE_APPNAME_WORD          As String = "Microsoft Word"
