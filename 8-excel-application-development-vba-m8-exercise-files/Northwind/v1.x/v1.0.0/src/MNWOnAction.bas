Attribute VB_Name = "MNWOnAction"
' ==========================================================================
' Module      : MNWOnAction
' Type        : Module
' Description :
' --------------------------------------------------------------------------
' Procedures  : ChangeSubject
'               GetCountries        Variant
'               GetRegions          Variant
'               NW_About
'               NW_CustomerEdit
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

Private Const msMODULE As String = "MNWOnAction"

Public Sub ChangeSubject(ByVal CodeName As String)

    Dim wks As Excel.Worksheet

    Set wks = GetWorksheet(CodeName)
    
    Select Case CodeName
    Case gsNW_WKSCN_CUSTOMERS
        If (ActiveSheet.CodeName <> CodeName) Then
            wks.Activate
            Application.GoTo wks.Cells(1), True
            Application.GoTo wks.Cells(2, 1)
        End If
    Case Else
        wks.Activate
        Application.GoTo wks.Cells(1), True
        Application.GoTo wks.Cells(2, 1)
    End Select

    Set wks = Nothing

End Sub

Public Function GetCountries() As Variant
' ==========================================================================
' Description : [Enter description]
'
' Parameters  :
'
' Returns     : Variant
'
' Comments    :
' ==========================================================================

    Dim vRtn    As Variant
    Dim rng     As Excel.Range

    ' ----------------------------------------------------------------------

    Set rng = GetRange(GetWorksheet(gsNW_WKSCN_COUNTRIES))

    vRtn = RangeToArray(rng)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetCountries = vRtn

    Set rng = Nothing
    On Error Resume Next
    Erase vRtn
    vRtn = Empty

End Function

Public Function GetRegions(ByVal Country As String) As Variant
' ==========================================================================
' Description : [Enter description]
'
' Parameters  :
'
' Returns     : Variant
'
' Comments    :
' ==========================================================================

    Dim vRtn    As Variant
    Dim rng     As Excel.Range

    ' ----------------------------------------------------------------------

    Set rng = GetRange(GetWorksheet(gsNW_WKSCN_REGIONS), Country)

    If (rng Is Nothing) Then
        vRtn = Array()
        GoTo PROC_EXIT
    End If

    vRtn = RangeToArray(rng.Columns(scRegionsRegion))

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetRegions = vRtn

    Set rng = Nothing
    On Error Resume Next
    Erase vRtn
    vRtn = Empty

End Function

Public Sub NW_About()

    Dim fab As FAbout

    Set fab = New FAbout

    fab.Show

    On Error Resume Next
    Unload fab

    Set fab = Nothing

End Sub

Public Sub NW_CustomerEdit()

    Dim frm As FCustomer
    
    Set frm = New FCustomer
    
    frm.Show vbModeless

    Set frm = Nothing

End Sub
