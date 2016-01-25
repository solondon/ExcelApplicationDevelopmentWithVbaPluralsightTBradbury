Attribute VB_Name = "MMSFormsComboBox"
' ==========================================================================
' Module      : MMSFormsComboBox
' Type        : Module
' Description : Support routines for working with the ComboBox control
' --------------------------------------------------------------------------
' Procedures  : CBORemoveAll
'               CBOSortList
' --------------------------------------------------------------------------
' Dependencies: MVBAArray
' --------------------------------------------------------------------------
' References  : Microsoft Forms 2.0 Object Library
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

Private Const msMODULE As String = "MMSFormsComboBox"

Public Sub CBORemoveAll(ByRef Combo As MSForms.ComboBox)
' ==========================================================================
' Description : Remove all of the items from a ComboBox
' ==========================================================================

    Const sPROC As String = "CBORemoveAll"

    Dim lCnt    As Long
    Dim lIdx    As Long
    Dim oCtl    As MSForms.Control: Set oCtl = Combo


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, oCtl.Name)

    ' ----------------------------------------------------------------------
    ' Quit if it is a data-bound control
    ' ----------------------------------
    If (Len(oCtl.ControlSource) > 0) Then
        GoTo PROC_EXIT
    End If


    lCnt = Combo.ListCount - 1

    For lIdx = lCnt To 0 Step -1
        Combo.RemoveItem lIdx
    Next

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set oCtl = Nothing

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

Public Sub CBOSortList(ByRef CBO As MSForms.ComboBox, _
              Optional ByVal Column As Long, _
              Optional ByVal Descending As Boolean)
' ==========================================================================
' Description : Sort the items in a ListBox.
'
' Parameters  : CBO         The ComboBox to sort
'               Column      The column to sort on
'               SortOrder   Sort ascending or descending
' ==========================================================================

    Const sPROC As String = "CBOSortList"

    Dim vList   As Variant


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    vList = CBO.List

    Call BubbleSortArray(vList, Column, Descending, True)

    CBO.List = vList

    ' ----------------------------------------------------------------------

PROC_EXIT:

    On Error Resume Next
    Erase vList
    vList = Empty

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
