VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCustomer 
   Caption         =   "Customer"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6930
   OleObjectBlob   =   "FCustomer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==========================================================================
' Module      : FCustomer
' Type        : Form
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

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

'Public Const GLOBAL_CONST As String = ""

' ----------------
' Module Level
' ----------------

Private Const msMODULE As String = "FCustomer"

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private ml_Row                  As Long

Private WithEvents Worksheet    As Excel.Worksheet
Attribute Worksheet.VB_VarHelpID = -1

Private Sub UpdateControls()

    With ActiveSheet.UsedRange.Rows(ActiveCell.Row)
        txtCustomerID = .Cells(scCustomerID)
        txtCompanyName = .Cells(scCustomerCompanyName)
        txtContactName = .Cells(scCustomerContactName)
        txtContactTitle = .Cells(scCustomerContactTitle)
        txtAddress = .Cells(scCustomerAddress)
        txtCity = .Cells(scCustomerCity)
        cboRegion = .Cells(scCustomerRegion)
        txtPostalCode = .Cells(scCustomerPostalCode)
        cboCountry = .Cells(scCustomerCountry)
        txtPhone = .Cells(scCustomerPhone)
        txtFax = .Cells(scCustomerFax)
    End With

End Sub

Private Sub cboRegion_Change()
    Worksheet.Cells(ml_Row, scCustomerRegion) = cboRegion
End Sub

Private Sub txtAddress_Change()
    Worksheet.Cells(ml_Row, scCustomerAddress) = txtAddress
End Sub

Private Sub txtCity_Change()
    Worksheet.Cells(ml_Row, scCustomerCity) = txtCity
End Sub

Private Sub txtCompanyName_Change()
    Worksheet.Cells(ml_Row, scCustomerCompanyName) = txtCompanyName
End Sub

Private Sub txtContactName_Change()
    Worksheet.Cells(ml_Row, scCustomerContactName) = txtContactName
End Sub

Private Sub txtCustomerID_Change()
    Worksheet.Cells(ml_Row, scCustomerID) = txtCustomerID
End Sub

Private Sub txtCustomerID_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call KeyFilter(KeyAscii, True, False)
    Call KeyCapitalize(KeyAscii)
End Sub

Private Sub txtFax_Change()
    Worksheet.Cells(ml_Row, scCustomerFax) = txtFax
End Sub

Private Sub txtPhone_Change()
    Worksheet.Cells(ml_Row, scCustomerPhone) = txtPhone
End Sub

Private Sub txtPostalCode_Change()
    Worksheet.Cells(ml_Row, scCustomerPostalCode) = txtPostalCode
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ml_Row = ActiveCell.Row
    Call UpdateControls
End Sub

Private Sub cboCountry_Change()

    Call CBORemoveAll(cboRegion)

    If (Len(cboCountry) > 0) Then
        cboRegion.List = GetRegions(cboCountry)
    End If

    Worksheet.Cells(ml_Row, scCustomerCountry) = cboCountry

End Sub

Private Sub cmdNext_Click()

    ml_Row = ml_Row + 1
    ActiveSheet.UsedRange.Rows(ml_Row).Select
    If (ml_Row > ActiveSheet.UsedRange.Rows.Count) Then
        cmdNext.Enabled = False
    End If

    cmdPrevious.Enabled = True

End Sub

Private Sub cmdPrevious_Click()

    ml_Row = ml_Row - 1
    ActiveSheet.UsedRange.Rows(ml_Row).Select
    
    If (ml_Row = 2) Then
        cmdPrevious.Enabled = False
    End If

    cmdNext.Enabled = True

End Sub

Private Sub txtFax_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call KeyFilter(KeyAscii, False, True, " ", "(", ")", "-", ".")
End Sub

Private Sub txtPhone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call KeyFilter(KeyAscii, False, True, " ", "(", ")", "-", ".")
End Sub

Private Sub UserForm_Activate()
    If (ActiveCell.Row = 2) Then
        cmdPrevious.Enabled = False
    End If
End Sub

Private Sub UserForm_Initialize()

    cboCountry.List = GetCountries
    Set Worksheet = GetWorksheet(gsNW_WKSCN_CUSTOMERS)
    Worksheet.Activate
    With Worksheet.UsedRange
        If (.Rows.Count > 1) Then
            If ((ActiveCell.Row <= .Rows.Count) _
            And (ActiveCell.Row > 1)) Then
                ml_Row = ActiveCell.Row
                .Rows(ml_Row).Select
            Else
                ml_Row = 2
                Worksheet.UsedRange.Rows(ml_Row).Select
            End If
            Call UpdateControls
        Else
            ml_Row = 2
            .Rows(1).Offset(1).Select
        End If
    End With

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
