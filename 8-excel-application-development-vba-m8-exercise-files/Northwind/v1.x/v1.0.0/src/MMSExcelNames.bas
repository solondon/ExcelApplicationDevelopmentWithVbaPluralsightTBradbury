Attribute VB_Name = "MMSExcelNames"
' ==========================================================================
' Module      : MMSExcelNames
' Type        : Module
' Description : Support for working with Excel Names
' --------------------------------------------------------------------------
' Procedures  : IsBuiltInName       Boolean
'               ListNames
'               NameCreator         enuExcelNameCreator
'               NameExists          Boolean
'               NameScope           enuExcelNameScope
'               RemoveNames         Boolean
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

Private Const msMODULE As String = "MMSExcelNames"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuExcelNameCreator
    encUnknown = 0
    encExcel = 1
    encUser = 2
End Enum

Public Enum enuExcelNameScope
    ensUnknown = 0
    ensWorkbook = 1
    ensWorksheet = 2
End Enum

Public Function IsBuiltInName(ByRef Name As Excel.Name) As Boolean
' ==========================================================================
' Description : Determines if a Name object was system-generated
'
' Parameters  : Name    The Name to inspect
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "IsBuiltInName"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = (NameCreator(Name) = encExcel)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsBuiltInName = bRtn

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

Public Sub ListNames(Optional ByVal Scope As enuExcelNameScope _
                                           = ensWorkbook Or ensWorksheet, _
                     Optional ByVal Creator As enuExcelNameCreator _
                                             = encExcel Or encUser)
' ==========================================================================
' Description : List information on names in this project
'
' Parameters  : Scope   Identifies if Workbook or Worksheet
'                       names are to be listed
'               Creator Identifies if user or system-generated
'                       names are to be listed
' ==========================================================================

    Const sPROC As String = "ListNames"

    Dim Nm      As Excel.Name
    Dim wks     As Excel.Worksheet


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If BitIsSet(Scope, ensWorkbook) Then

        Debug.Print String$(glLIST_LINELEN, "=")
        Debug.Print "Workbook Names:"
        Debug.Print String$(glLIST_LINELEN, gsLIST_LINECHAR)
        Debug.Print

        For Each Nm In ThisWorkbook.Names
            If BitIsSet(NameScope(Nm), ensWorkbook) Then
                If BitIsSet(Creator, NameCreator(Nm)) Then
                    Debug.Print Nm.Name, _
                                Nm.RefersTo & IIf(Nm.Visible, _
                                                  vbNullString, _
                                                " (Hidden)")
                End If
            End If
        Next Nm

    End If

    If BitIsSet(Scope, ensWorksheet) Then

        Debug.Print String$(glLIST_LINELEN, "=")
        Debug.Print "Worksheet Names:"
        Debug.Print String$(glLIST_LINELEN, gsLIST_LINECHAR)
        Debug.Print

        For Each wks In ThisWorkbook.Worksheets
            For Each Nm In wks.Names
                If BitIsSet(NameScope(Nm), ensWorksheet) Then
                    If BitIsSet(Creator, NameCreator(Nm)) Then
                        Debug.Print Nm.Name, _
                                    Nm.RefersTo & IIf(Nm.Visible, _
                                                      vbNullString, _
                                                    " (Hidden)")
                    End If
                End If
            Next Nm
        Next wks

    End If

    Debug.Print String$(glLIST_LINELEN, "=")

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set Nm = Nothing
    Set wks = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub

Public Function NameCreator(ByRef Name As Excel.Name) As enuExcelNameCreator
' ==========================================================================
' Description : Determine if an Excel Name is system or user generated
'
' Parameters  : Name    The Name object to examine
'
' Returns     : enuExcelNameCreator
' ==========================================================================

    Const sPROC     As String = "NameCreator"

    Dim lPos        As Long
    Dim sName       As String
    Dim vReserved   As Variant
    Dim eRtn        As enuExcelNameCreator


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Locate the name of the Name
    ' ---------------------------

    lPos = InStr(1, Name.Name, "!_")
    If (lPos > 0) Then
        sName = Mid$(Name.Name, lPos + 2)
    Else
        lPos = InStr(1, Name.Name, "!")
        If (lPos > 0) Then
            sName = Mid$(Name.Name, lPos + 1)
        Else
            eRtn = encUser
            GoTo PROC_EXIT
        End If
    End If

    ' Test the list of reserved names
    ' -------------------------------
    vReserved = Array("Auto_Activate", _
                      "Auto_Close", _
                      "Auto_Deactivate", _
                      "Auto_Open", _
                      "Consolidate_Area", _
                      "Criteria", _
                      "Data_Form", _
                      "Database", _
                      "Extract", _
                      "FilterDatabase", _
                      "Print_Area", _
                      "Print_Titles", _
                      "Recorder", _
                      "Sheet_Title")

    If AnyEqual(sName, vReserved) Then
        eRtn = encExcel
        GoTo PROC_EXIT
    End If

    ' Test for special exceptions
    ' ---------------------------
    If (InStr(1, sName, ".wvu.", vbTextCompare) > 0) Then
        eRtn = encExcel
    ElseIf (InStr(1, sName, "wrn.", vbTextCompare) > 0) Then
        eRtn = encExcel
    Else
        eRtn = encUser
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    NameCreator = eRtn

    On Error Resume Next
    Erase vReserved
    vReserved = Empty

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

Public Function NameExists(ByVal Name As String, _
                  Optional ByRef Book As Excel.Workbook) As Boolean
' ==========================================================================
' Description : Determines if a name has been defined.
'
' Parameters  : Name        The name to check for
'               Book        The workbook to look in
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "NameExists"

    Dim bRtn    As Boolean
    Dim wkb     As Excel.Workbook

    ' Get the workbook
    ' ----------------
    If (Book Is Nothing) Then
        Set wkb = ThisWorkbook
    Else
        Set wkb = Book
    End If

    On Error Resume Next

    bRtn = CBool(Len(wkb.Names(Name).Name))

    ' ----------------------------------------------------------------------

PROC_EXIT:

    NameExists = bRtn

    Set wkb = Nothing

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

Public Function NameScope(ByRef Name As Excel.Name) As enuExcelNameScope
' ==========================================================================
' Description : Determine the scope of a given Name object
'
' Parameters  : Name    The Name to inspect
'
' Returns     : enuExcelNameScope
' ==========================================================================

    Const sPROC As String = "NameScope"

    Dim eRtn    As enuExcelNameScope


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Select Case Name.Parent.Name
    Case ThisWorkbook.Name
        eRtn = ensWorkbook
    Case Else
        eRtn = ensWorksheet
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    NameScope = eRtn

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

Public Function RemoveNames(Optional ByRef Sheet As Excel.Worksheet) _
       As Boolean
' ==========================================================================
' Description : Remove names from a worksheet.
'               If the sheet is not provided, then use the workbook
'
' Parameters  : Sheet   The worksheet to modify
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "RemoveNames"

    Dim bRtn    As Boolean: bRtn = True
    Dim oNm     As Excel.Name


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (Sheet Is Nothing) Then
        For Each oNm In ThisWorkbook.Names
            oNm.Delete
        Next oNm
    Else
        For Each oNm In Sheet.Names
            oNm.Delete
        Next oNm
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    RemoveNames = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    bRtn = False

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function
