Attribute VB_Name = "MNWSheets"
' ==========================================================================
' Module      : MNWFormatSheets
' Type        : Module
' Description : Routines to format the worksheets
' --------------------------------------------------------------------------
' Procedures  : AddSheets
'               DeleteSheets
'               FormatSheet
'               FormatSheetCustomers
'               FormatSheetEmployees
'               FormatSheetOrderDetails
'               FormatSheetCategories
'               FormatSheetSuppliers
'               FormatSheetShippers
'               FormatSheetCountries
'               FormatSheetRegions
'               FormatSheetOrders
'               FormatSheetProducts
'               FormatSheets
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

Public Const gsNW_WKSCN_CUSTOMERS       As String = "wksCustomers"
Public Const gsNW_WKSCN_EMPLOYEES       As String = "wksEmployees"
Public Const gsNW_WKSCN_ORDERS          As String = "wksOrders"
Public Const gsNW_WKSCN_ORDERDETAILS    As String = "wksOrderDetails"
Public Const gsNW_WKSCN_PRODUCTS        As String = "wksProducts"
Public Const gsNW_WKSCN_CATEGORIES      As String = "wksCategories"
Public Const gsNW_WKSCN_SUPPLIERS       As String = "wksSuppliers"
Public Const gsNW_WKSCN_SHIPPERS        As String = "wksShippers"
Public Const gsNW_WKSCN_COUNTRIES       As String = "wksCountries"
Public Const gsNW_WKSCN_REGIONS         As String = "wksRegions"

Public Const gsNW_WKSNM_CUSTOMERS       As String = "Customers"
Public Const gsNW_WKSNM_EMPLOYEES       As String = "Employees"
Public Const gsNW_WKSNM_ORDERS          As String = "Orders"
Public Const gsNW_WKSNM_ORDERDETAILS    As String = "OrderDetails"
Public Const gsNW_WKSNM_PRODUCTS        As String = "Products"
Public Const gsNW_WKSNM_CATEGORIES      As String = "Categories"
Public Const gsNW_WKSNM_SUPPLIERS       As String = "Suppliers"
Public Const gsNW_WKSNM_SHIPPERS        As String = "Shippers"
Public Const gsNW_WKSNM_COUNTRIES       As String = "Countries"
Public Const gsNW_WKSNM_REGIONS         As String = "Regions"

Public Const gsSYS_WKSCN_TEMP           As String = "wksTemp"
Public Const gsSYS_WKSNM_TEMP           As String = "Temp"

' ----------------
' Module Level
' ----------------

Private Const msMODULE                  As String = "MNWFormatSheets"

Private Const msNW_WKSCN_DELETE         As String = "wksDelete"
Private Const msNW_WKSNM_DELETE         As String = gsAPP_NAME

Private Const msNUMFMT_SHORTDATE        As String = "m/d/yyyy"
Private Const msNUMFMT_CURRENCY         As String = "$#,##0.00"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuSheetColumnCategories
    scCategoriesCategoryID = 1
    scCategoriesCategoryName
    scCategoriesDescription
    scCategoriesPicture
    [_First] = scCategoriesCategoryID
    [_Last] = scCategoriesPicture
End Enum

Public Enum enuSheetColumnCountries
    scCountryName = 1
    [_First] = scCountryName
    [_Last] = scCountryName
End Enum

Public Enum enuSheetColumnCustomers
    scCustomerID = 1
    scCustomerCompanyName
    scCustomerContactName
    scCustomerContactTitle
    scCustomerAddress
    scCustomerCity
    scCustomerRegion
    scCustomerPostalCode
    scCustomerCountry
    scCustomerPhone
    scCustomerFax
    [_First] = scCustomerID
    [_Last] = scCustomerFax
End Enum

Public Enum enuSheetColumnEmployees
    scEmployeeID = 1
    scEmployeeLastName
    scEmployeeFirstName
    scEmployeeTitle
    scEmployeeTitleOfCourtesy
    scEmployeeBirthDate
    scEmployeeHireDate
    scEmployeeAddress
    scEmployeeCity
    scEmployeeRegion
    scEmployeePostalCode
    scEmployeeCountry
    scEmployeeHomePhone
    scEmployeeExtension
    scEmployeePhoto
    scEmployeeNotes
    scEmployeeReportsTo
    [_First] = scEmployeeID
    [_Last] = scEmployeeReportsTo
End Enum

Public Enum enuSheetColumnOrderDetails
    scOrderDetailOrderID = 1
    scOrderDetailProduct
    scOrderDetailUnitPrice
    scOrderDetailQuantity
    scOrderDetailDiscount
    [_First] = scOrderDetailOrderID
    [_Last] = scOrderDetailDiscount
End Enum

Public Enum enuSheetColumnOrders
    scOrderID = 1
    scOrderCustomer
    scOrderEmployee
    scOrderOrderDate
    scOrderRequiredDate
    scOrderShippedDate
    scOrderShipVia
    scOrderFreight
    scOrderShipName
    scOrderShipAddress
    scOrderShipCity
    scOrderShipRegion
    scOrderShipPostalCode
    scOrderShipCountry
    [_First] = scOrderID
    [_Last] = scOrderShipCountry
End Enum

Public Enum enuSheetColumnProducts
    scProductsProductID = 1
    scProductsProductName
    scProductsSupplier
    scProductsCategory
    scProductsQuantityPerUnit
    scProductsUnitPrice
    scProductsUnitsInStock
    scProductsUnitsOnOrder
    scProductsReorderLevel
    scProductsDiscontinued
    [_First] = scProductsProductID
    [_Last] = scProductsDiscontinued
End Enum

Public Enum enuSheetColumnRegions
    scRegionsCountry = 1
    scRegionsRegion
    [_First] = scRegionsCountry
    [_Last] = scRegionsRegion
End Enum

Public Enum enuSheetColumnShippers
    scShippersShipperID = 1
    scShippersCompanyName
    scShippersPhone
    [_First] = scShippersShipperID
    [_Last] = scShippersPhone
End Enum

Public Enum enuSheetColumnSuppliers
    scSuppliersSupplierID = 1
    scSuppliersCompanyName
    scSuppliersContactName
    scSuppliersContactTitle
    scSuppliersAddress
    scSuppliersCity
    scSuppliersRegion
    scSuppliersPostalCode
    scSuppliersCountry
    scSuppliersPhone
    scSuppliersFax
    scSuppliersHomePage
    scSuppliersURL
    [_First] = scSuppliersSupplierID
    [_Last] = scSuppliersURL
End Enum

Public Sub AddSheets(Optional ByRef PB As IProgressBar)
' ==========================================================================
' Description : Add the worksheets to the Workbook
' ==========================================================================

    Const sPROC     As String = "AddSheets"

    Dim sTitle      As String: sTitle = gsAPP_NAME
    Dim sPrompt     As String
    Dim eButtons    As VbMsgBoxStyle: eButtons = vbInformation Or vbOKOnly
    Dim eMBR        As VbMsgBoxResult

    Dim lIdx        As Long

    Dim vSheet      As Variant
    Dim vSheets     As Variant
    
    Dim udtProps    As TApplicationProperties
    Dim wks         As Excel.Worksheet


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE
    Application.DisplayAlerts = False

    ' Get the array of sheets to add
    ' ------------------------------
    vSheets = Array(gsNW_WKSCN_REGIONS, _
                    gsNW_WKSCN_COUNTRIES, _
                    gsNW_WKSCN_SHIPPERS, _
                    gsNW_WKSCN_SUPPLIERS, _
                    gsNW_WKSCN_CATEGORIES, _
                    gsNW_WKSCN_PRODUCTS, _
                    gsNW_WKSCN_ORDERDETAILS, _
                    gsNW_WKSCN_ORDERS, _
                    gsNW_WKSCN_EMPLOYEES, _
                    gsNW_WKSCN_CUSTOMERS)

    ' Test if the add is needed
    ' -------------------------
    If (Worksheets.Count > UBound(vSheets)) Then
        sPrompt = "The workbook sheets are already added."
        Call MsgBox(sPrompt, eButtons, sTitle)
        GoTo PROC_EXIT
    End If

    ' Update the ProgressBar settings
    ' -------------------------------
    If (Not PB Is Nothing) Then
        PB.Reset
        PB.Caption = vbNewLine & "Adding worksheets"
        PB.Max = UBound(vSheets) + 1
    End If

    ' Set the CodeNames for all of the sheets. After this point
    ' all references to the sheets will be made by CodeName.
    ' ---------------------------------------------------------
    For Each vSheet In vSheets
        If (Not PB Is Nothing) Then
            PB.Caption = vbNewLine & "Adding " & CStr(vSheet)
            PB.Increment
            DoEvents
        End If
        Set wks = ThisWorkbook.Worksheets.Add
        Call SetCodeName(wks, CStr(vSheet))
    Next vSheet

    ' Delete the leftover worksheet
    ' -----------------------------
    Set wks = GetWorksheet(msNW_WKSCN_DELETE)
    If (Not wks Is Nothing) Then
        wks.Delete
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set wks = Nothing

    On Error Resume Next
    Erase vSheets
    vSheets = Empty
    vSheet = Empty

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

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

Public Sub DeleteSheets(Optional ByRef PB As IProgressBar)
' ==========================================================================
' Description : Remove all of the worksheets.
'
' NOTES       : There must always be at least one sheet, so the last sheet
'               is renamed to identify it for deletion during the build.
' ==========================================================================

    Const sPROC     As String = "DeleteSheets"

    Dim lIdx        As Long

    Dim udtProps    As TApplicationProperties
    Dim wks         As Excel.Worksheet


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)


    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE

    ' Turn off warnings so this process
    ' can run without user intervention
    ' ---------------------------------
    Application.DisplayAlerts = False

    ' Add a dummy sheet
    ' -----------------
    Set wks = GetWorksheet(msNW_WKSCN_DELETE)

    If (wks Is Nothing) Then
        If (ThisWorkbook.Worksheets.Count = 1) Then
            Set wks = ThisWorkbook.Worksheets(1)
        Else
            Set wks = ThisWorkbook.Worksheets.Add
        End If
        Call SetCodeName(wks, msNW_WKSCN_DELETE)
        wks.Name = msNW_WKSNM_DELETE
    End If

    ' Update the ProgressBar settings
    ' -------------------------------
    If (Not PB Is Nothing) Then
        PB.Caption = vbNewLine & "Deleting worksheets"
        PB.Min = 0
        PB.Max = ThisWorkbook.Worksheets.Count
        PB.Value = lIdx
    End If

    ' Remove all of the existing sheets
    ' ---------------------------------
    For Each wks In ThisWorkbook.Worksheets

        lIdx = lIdx + 1
        If (Not PB Is Nothing) Then
            PB.Value = lIdx
        End If

        If (wks.CodeName <> msNW_WKSCN_DELETE) Then
            wks.Visible = xlSheetVisible
            wks.Delete
        End If

    Next wks

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set wks = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

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

Public Sub FormatSheet(Optional ByRef Sheet As Excel.Worksheet, _
                       Optional ByVal Reset As Boolean, _
                       Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Apply default formatting and naming to a worksheet
'
' Params      : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC As String = "FormatSheet"

    Dim bDelete As Boolean
    Dim wksActive As Excel.Worksheet
    Dim udtProps As TApplicationProperties


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Make sure there is a sheet to format
    ' ------------------------------------
    If (Sheet Is Nothing) Then
        Set Sheet = ActiveSheet
        bDelete = True
    End If

    Call Trace(tlMaximum, msMODULE, sPROC, Sheet.CodeName)

    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE

    ' Keep track of the current sheet
    ' -------------------------------
    Set wksActive = ActiveSheet

    ' Determine which sheet is targeted for
    ' formatting and call the format routine
    ' --------------------------------------
    Select Case Sheet.CodeName
    Case gsNW_WKSCN_CUSTOMERS
        Call FormatSheetCustomers(Sheet, Reset, LoadData)

    Case gsNW_WKSCN_EMPLOYEES
        Call FormatSheetEmployees(Sheet, Reset, LoadData)

    Case gsNW_WKSCN_ORDERS
        Call FormatSheetOrders(Sheet, Reset, LoadData)
    
    Case gsNW_WKSCN_ORDERDETAILS
        Call FormatSheetOrderDetails(Sheet, Reset, LoadData)

    Case gsNW_WKSCN_PRODUCTS
        Call FormatSheetProducts(Sheet, Reset, LoadData)

    Case gsNW_WKSCN_CATEGORIES
        Call FormatSheetCategories(Sheet, Reset, LoadData)

    Case gsNW_WKSCN_SUPPLIERS
        Call FormatSheetSuppliers(Sheet, Reset, LoadData)

    Case gsNW_WKSCN_SHIPPERS
        Call FormatSheetShippers(Sheet, Reset, LoadData)

    Case gsNW_WKSCN_COUNTRIES
        Call FormatSheetCountries(Sheet, Reset, LoadData)

    Case gsNW_WKSCN_REGIONS
        Call FormatSheetRegions(Sheet, Reset, LoadData)

    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ' Restore the current sheet
    ' -------------------------
    wksActive.Activate

    Set wksActive = Nothing

    If bDelete Then
        Set Sheet = Nothing
    End If

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

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

Private Sub FormatSheetCustomers(ByRef Sheet As Excel.Worksheet, _
                        Optional ByVal Reset As Boolean, _
                        Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetCustomers"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_CUSTOMERS) Then
            .Name = gsNW_WKSNM_CUSTOMERS
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnCustomers.[_First] _
                To enuSheetColumnCustomers.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column
                Case scCustomerID
                    .Value = "Customer ID"
                    .ColumnWidth = 15

                Case scCustomerCompanyName
                    .Value = "Company Name"
                    .ColumnWidth = 33

                Case scCustomerContactName
                    .Value = "Contact Name"
                    .ColumnWidth = 20

                Case scCustomerContactTitle
                    .Value = "Contact Title"
                    .ColumnWidth = 20

                Case scCustomerAddress
                    .Value = "Address"
                    .ColumnWidth = 40

                Case scCustomerCity
                    .Value = "City"
                    .ColumnWidth = 14

                Case scCustomerRegion
                    .Value = "Region"
                    .ColumnWidth = 13

                Case scCustomerPostalCode
                    .Value = "Postal Code"
                    .ColumnWidth = 14

                Case scCustomerCountry
                    .Value = "Country"
                    .ColumnWidth = 11

                Case scCustomerPhone
                    .Value = "Phone"
                    .ColumnWidth = 15

                    On Error Resume Next
                    .Comment.Delete
                    On Error GoTo PROC_ERR

                    .AddComment ("Phone" & vbNewLine _
                               & "Be sure to add" & vbNewLine _
                               & "international calling codes.")
                    With .Comment
                        .Shape.TextFrame.Characters(1, Len("Phone")).Font.Bold _
                            = True
                        .Shape.Top = .Parent.Offset(RowOffset:=3).Top
                        .Shape.Left = .Parent.Offset(ColumnOffset:=1).Left
                    End With

                Case scCustomerFax
                    .Value = "Fax"
                    .ColumnWidth = 15

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_CUSTOMERS & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_CUSTOMERS & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_CUSTOMERS)

        Call FreezeHeaderRow(Sheet, 1, scCustomerCompanyName)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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

Private Sub FormatSheetEmployees(ByRef Sheet As Excel.Worksheet, _
                        Optional ByVal Reset As Boolean, _
                        Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetEmployees"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_EMPLOYEES) Then
            .Name = gsNW_WKSNM_EMPLOYEES
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnEmployees.[_First] _
                To enuSheetColumnEmployees.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column
                Case scEmployeeID = 1
                    .Value = "Employee ID"
                    .ColumnWidth = 15

                Case scEmployeeLastName
                    .Value = "Last Name"
                    .ColumnWidth = 12

                Case scEmployeeFirstName
                    .Value = "First Name"
                    .ColumnWidth = 12

                Case scEmployeeTitle
                    .Value = "Title"
                    .ColumnWidth = 23

                Case scEmployeeTitleOfCourtesy
                    .Value = "Title Of Courtesy"
                    .ColumnWidth = 16

                Case scEmployeeBirthDate
                    .Value = "Birth Date"
                    .ColumnWidth = 14

                Case scEmployeeHireDate
                    .Value = "Hire Date"
                    .ColumnWidth = 14

                Case scEmployeeAddress
                    .Value = "Address"
                    .ColumnWidth = 32

                Case scEmployeeCity
                    .Value = "City"
                    .ColumnWidth = 10

                Case scEmployeeRegion
                    .Value = "Region"
                    .ColumnWidth = 10

                Case scEmployeePostalCode
                    .Value = "Postal Code"
                    .ColumnWidth = 12

                Case scEmployeeCountry
                    .Value = "Country"
                    .ColumnWidth = 9

                Case scEmployeeHomePhone
                    .Value = "Home Phone"
                    .ColumnWidth = 14

                Case scEmployeeExtension
                    .Value = "Extension"
                    .ColumnWidth = 12

                Case scEmployeePhoto
                    .Value = "Photo"
                    .ColumnWidth = 9

                Case scEmployeeNotes
                    .Value = "Notes"
                    .ColumnWidth = 15

                Case scEmployeeReportsTo
                    .Value = "Reports To"
                    .ColumnWidth = 17

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_EMPLOYEES & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_EMPLOYEES & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
                Call ClearErrors(Sheet.UsedRange.Columns(scEmployeePostalCode))
                Call ClearErrors(Sheet.UsedRange.Columns(scEmployeeExtension))
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_EMPLOYEES)

        Call FreezeHeaderRow(Sheet, 1, scEmployeeLastName)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
        With Sheet.Columns(scEmployeeBirthDate)
            .NumberFormat = msNUMFMT_SHORTDATE
            .HorizontalAlignment = xlCenter
        End With
        With Sheet.Columns(scEmployeeHireDate)
            .NumberFormat = msNUMFMT_SHORTDATE
            .HorizontalAlignment = xlCenter
        End With
        Sheet.Columns(scEmployeeNotes).ColumnWidth = 15
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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

Private Sub FormatSheetOrderDetails(ByRef Sheet As Excel.Worksheet, _
                     Optional ByVal Reset As Boolean, _
                     Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetOrderDetails"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_ORDERDETAILS) Then
            .Name = gsNW_WKSNM_ORDERDETAILS
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnOrderDetails.[_First] _
                To enuSheetColumnOrderDetails.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column
                Case scOrderDetailOrderID
                    .Value = "Order ID"
                    .ColumnWidth = 11

                Case scOrderDetailProduct
                    .Value = "Product"
                    .ColumnWidth = 32

                Case scOrderDetailUnitPrice
                    .Value = "Unit Price"
                    .ColumnWidth = 12

                Case scOrderDetailQuantity
                    .Value = "Quantity"
                    .ColumnWidth = 12

                Case scOrderDetailDiscount
                    .Value = "Discount"
                    .ColumnWidth = 12

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_ORDERDETAILS & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_ORDERDETAILS & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_ORDERDETAILS)

        Call FreezeHeaderRow(Sheet, 1, scOrderDetailProduct)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
        With Sheet.Columns(scOrderDetailUnitPrice)
            .NumberFormat = msNUMFMT_CURRENCY
        End With
        With Sheet.Columns(scOrderDetailDiscount)
            .NumberFormat = msNUMFMT_CURRENCY
        End With
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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

Private Sub FormatSheetCategories(ByRef Sheet As Excel.Worksheet, _
                         Optional ByVal Reset As Boolean, _
                         Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetCategories"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_CATEGORIES) Then
            .Name = gsNW_WKSNM_CATEGORIES
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnCategories.[_First] _
                To enuSheetColumnCategories.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column

                Case scCategoriesCategoryID = 1
                    .Value = "Category ID"
                    .ColumnWidth = 11

                Case scCategoriesCategoryName
                    .Value = "Category Name"
                    .ColumnWidth = 18

                Case scCategoriesDescription
                    .Value = "Description"
                    .ColumnWidth = 53

                Case scCategoriesPicture
                    .Value = "Pictre"
                    .ColumnWidth = 11

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_CATEGORIES & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_CATEGORIES & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_CATEGORIES)

        Call FreezeHeaderRow(Sheet, 1, scCategoriesCategoryName)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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


Private Sub FormatSheetSuppliers(ByRef Sheet As Excel.Worksheet, _
                        Optional ByVal Reset As Boolean, _
                        Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetSuppliers"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_SUPPLIERS) Then
            .Name = gsNW_WKSNM_SUPPLIERS
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnSuppliers.[_First] _
                To enuSheetColumnSuppliers.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column
                Case scSuppliersSupplierID
                    .Value = "Supplier ID"
                    .ColumnWidth = 13

                Case scSuppliersCompanyName
                    .Value = "Company Name"
                    .ColumnWidth = 38

                Case scSuppliersContactName
                    .Value = "Contact Name"
                    .ColumnWidth = 25

                Case scSuppliersContactTitle
                    .Value = "Contact Title"
                    .ColumnWidth = 27

                Case scSuppliersAddress
                    .Value = "Address"
                    .ColumnWidth = 40

                Case scSuppliersCity
                    .Value = "City"
                    .ColumnWidth = 13

                Case scSuppliersRegion
                    .Value = "Region"
                    .ColumnWidth = 9

                Case scSuppliersPostalCode
                    .Value = "Postal Code"
                    .ColumnWidth = 12

                Case scSuppliersCountry
                    .Value = "Country"
                    .ColumnWidth = 12

                Case scSuppliersPhone
                    .Value = "Phone"
                    .ColumnWidth = 15

                Case scSuppliersFax
                    .Value = "Fax"
                    .ColumnWidth = 15

                Case scSuppliersHomePage
                    .Value = "Home Page"
                    .ColumnWidth = 35

                Case scSuppliersURL
                    .Value = "URL"
                    .ColumnWidth = 60

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_SUPPLIERS & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_SUPPLIERS & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
                Call ClearErrors(Sheet.UsedRange.Columns(scSuppliersPostalCode))
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_SUPPLIERS)

        Call FreezeHeaderRow(Sheet, 1, scCategoriesCategoryName)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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



Private Sub FormatSheetShippers(ByRef Sheet As Excel.Worksheet, _
                       Optional ByVal Reset As Boolean, _
                       Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetShippers"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_SHIPPERS) Then
            .Name = gsNW_WKSNM_SHIPPERS
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnShippers.[_First] _
                To enuSheetColumnShippers.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column

                Case scShippersShipperID
                    .Value = "Shipper ID"
                    .ColumnWidth = 15

                Case scShippersCompanyName
                    .Value = "Company Name"
                    .ColumnWidth = 21

                Case scShippersPhone
                    .Value = "Phone"
                    .ColumnWidth = 15

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_SHIPPERS & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_SHIPPERS & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_SHIPPERS)

        Call FreezeHeaderRow(Sheet, 1, scShippersCompanyName)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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




Private Sub FormatSheetCountries(ByRef Sheet As Excel.Worksheet, _
                        Optional ByVal Reset As Boolean, _
                        Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetCountries"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_COUNTRIES) Then
            .Name = gsNW_WKSNM_COUNTRIES
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnCountries.[_First] _
                To enuSheetColumnCountries.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column

                Case scCountryName
                    .Value = "Country"
                    .ColumnWidth = 15

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_COUNTRIES & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_COUNTRIES & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_COUNTRIES)

        Call FreezeHeaderRow(Sheet)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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





Private Sub FormatSheetRegions(ByRef Sheet As Excel.Worksheet, _
                      Optional ByVal Reset As Boolean, _
                      Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetRegions"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_REGIONS) Then
            .Name = gsNW_WKSNM_REGIONS
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnRegions.[_First] _
                To enuSheetColumnRegions.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column
                Case scRegionsCountry
                    .Value = "Country"
                    .ColumnWidth = 15

                Case scRegionsRegion
                    .Value = "Region"
                    .ColumnWidth = 15

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_REGIONS & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_REGIONS & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_REGIONS)

        Call FreezeHeaderRow(Sheet)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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






Private Sub FormatSheetOrders(ByRef Sheet As Excel.Worksheet, _
                     Optional ByVal Reset As Boolean, _
                     Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetOrders"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_ORDERS) Then
            .Name = gsNW_WKSNM_ORDERS
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnEmployees.[_First] _
                To enuSheetColumnEmployees.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column
                Case scOrderID
                    .Value = "Order ID"
                    .ColumnWidth = 11

                Case scOrderCustomer
                    .Value = "Customer"
                    .ColumnWidth = 33

                Case scOrderEmployee
                    .Value = "Employee"
                    .ColumnWidth = 17

                Case scOrderOrderDate
                    .Value = "Order Date"
                    .ColumnWidth = 12

                Case scOrderRequiredDate
                    .Value = "Required Date"
                    .ColumnWidth = 12

                Case scOrderShippedDate
                    .Value = "Shipped Date"
                    .ColumnWidth = 12

                Case scOrderShipVia
                    .Value = "Ship Via"
                    .ColumnWidth = 16

                Case scOrderFreight
                    .Value = "Order Freight"
                    .ColumnWidth = 9

                Case scOrderShipName
                    .Value = "Ship Name"
                    .ColumnWidth = 33

                Case scOrderShipAddress
                    .Value = "Ship Address"
                    .ColumnWidth = 27

                Case scOrderShipCity
                    .Value = "Ship City"
                    .ColumnWidth = 14

                Case scOrderShipRegion
                    .Value = "Ship Region"
                    .ColumnWidth = 13

                Case scOrderShipPostalCode
                    .Value = "Ship Postal Code"
                    .ColumnWidth = 17

                Case scOrderShipCountry
                    .Value = "Ship Country"
                    .ColumnWidth = 13

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_ORDERS & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_ORDERS & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
                Call ClearErrors(Sheet.UsedRange.Columns(scOrderShipPostalCode))
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_ORDERS)

        Call FreezeHeaderRow(Sheet, 1, scOrderCustomer)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
        With Sheet.Columns(scOrderOrderDate)
            .NumberFormat = msNUMFMT_SHORTDATE
            .HorizontalAlignment = xlCenter
        End With
        With Sheet.Columns(scOrderRequiredDate)
            .NumberFormat = msNUMFMT_SHORTDATE
            .HorizontalAlignment = xlCenter
        End With
        With Sheet.Columns(scOrderShippedDate)
            .NumberFormat = msNUMFMT_SHORTDATE
            .HorizontalAlignment = xlCenter
        End With
        With Sheet.Columns(scOrderFreight)
            .NumberFormat = msNUMFMT_CURRENCY
        End With
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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

Private Sub FormatSheetProducts(ByRef Sheet As Excel.Worksheet, _
                       Optional ByVal Reset As Boolean, _
                       Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format the Customers worksheet
'
' Parameters  : Sheet       The worksheet to format
' ==========================================================================

    Const sPROC     As String = "FormatSheetProducts"

    Dim lRtn        As Long
    Dim lCol        As Long                     ' The current column
    Dim sPath       As String                   ' Path to the import file
    Dim sFileName   As String                   ' The import file name
    Dim sFullName   As String                   ' Path and File Name
    
    Dim Cell        As Excel.Range              ' The current cell
    Dim wksActive   As Excel.Worksheet          ' The current worksheet
    Dim eVisible    As XlSheetVisibility        ' The current visibility
    Dim udtProps    As TApplicationProperties   ' The current application state
    Dim oFD         As Office.FileDialog

    Dim sCurUser    As String                   ' The name of the current user


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .ScreenUpdating = gbDEBUG_MODE
        sCurUser = .UserName
        .UserName = gsAPP_NAME
    End With

    ' Save the current worksheet
    ' --------------------------
    Set wksActive = ActiveSheet

    If Reset Then
        Call ResetWorksheet(Sheet)
    End If

    If (Reset And LoadData) Then
    End If

    With Sheet
        ' Set the (Tab) display name
        ' --------------------------
        If (Sheet.CodeName = gsNW_WKSCN_PRODUCTS) Then
            .Name = gsNW_WKSNM_PRODUCTS
        End If

        ' Save the current visibility
        ' ---------------------------
        eVisible = .Visible
        .Visible = xlSheetVisible

        .Activate
        .AutoFilterMode = False

        ' Set the titles for the columns
        ' ------------------------------
        For lCol = enuSheetColumnProducts.[_First] _
                To enuSheetColumnProducts.[_Last]

            Set Cell = .Cells(1, lCol)

            With Cell
                Select Case Cell.Column

                Case scProductsProductID
                    .Value = "Product ID"
                    .ColumnWidth = 14

                Case scProductsProductName
                    .Value = "Product Name"
                    .ColumnWidth = 11

                Case scProductsSupplier
                    .Value = "Supplier"
                    .ColumnWidth = 11

                Case scProductsCategory
                    .Value = "Category"
                    .ColumnWidth = 11

                Case scProductsQuantityPerUnit
                    .Value = "Quantity Per Unit"
                    .ColumnWidth = 20

                Case scProductsUnitPrice
                    .Value = "Unit Price"
                    .ColumnWidth = 14

                Case scProductsUnitsInStock
                    .Value = "Units In Stock"
                    .ColumnWidth = 18

                Case scProductsUnitsOnOrder
                    .Value = "Units On Order"
                    .ColumnWidth = 18

                Case scProductsReorderLevel
                    .Value = "Reorder Level"
                    .ColumnWidth = 18

                Case scProductsDiscontinued
                    .Value = "Discontinued"
                    .ColumnWidth = 16

                End Select

            End With

        Next lCol

        If (Reset And LoadData) Then
            sPath = CurDir
            sFileName = gsNW_WKSNM_PRODUCTS & ".txt"
            sFullName = CurDir & "\" & sFileName
            If (Not FileExists(sFullName)) Then
                Set oFD = Application.FileDialog(msoFileDialogOpen)
                With oFD
                    .AllowMultiSelect = False
                    .ButtonName = "Import"
                    With .Filters
                        .Clear
                        .Add "Text Files", "*.txt"
                        .Add "All Files", "*.*"
                    End With
                    .FilterIndex = 1
                    .InitialFileName = sFullName
                    .InitialView = msoFileDialogViewDetails
                    .Title = "Select " & gsNW_WKSNM_PRODUCTS & " Import File"
                    lRtn = .Show
                    If (lRtn = False) Then
                        GoTo IMPORT_CANCELED
                    Else
                        sFullName = .SelectedItems(1)
                    End If
                End With
                sPath = ParsePath(sFullName, pppFullPath)
                sFileName = ParsePath(sFullName, pppFileOnly)
            End If
            
            If FileExists(sFullName) Then
                Call ImportTextFile(Sheet:=Sheet, _
                                    FileName:=sFullName, _
                                    Destination:="$A$2", _
                                    StartRow:=2, _
                                    HasHeaders:=True)
            End If
        End If

IMPORT_CANCELED:

        Call CreateTable(Sheet, gsNW_WKSNM_PRODUCTS)

        Call FreezeHeaderRow(Sheet, 1, scProductsProductName)
        With .UsedRange
            .Rows(1).AutoFilter
        End With
        With Sheet.Columns(scProductsUnitPrice)
            .NumberFormat = msNUMFMT_CURRENCY
        End With
        With Sheet.Columns(scProductsDiscontinued)
            .HorizontalAlignment = xlCenter
        End With
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    wksActive.Activate
    Set wksActive = Nothing
    Set Cell = Nothing
    Set oFD = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)
    With Application
        .DisplayCommentIndicator = xlCommentIndicatorOnly
        .UserName = sCurUser
    End With

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

Public Sub FormatSheets(Optional ByRef PB As IProgressBar, _
                        Optional ByVal Reset As Boolean, _
                        Optional ByVal LoadData As Boolean)
' ==========================================================================
' Description : Format all of the worksheets in the workbook
'
' Parameters  : PB          Optional ProgressBar
'               Reset       If True, clear existing data
'               LoadData    If True, load sample data (Reset must be True)
' ==========================================================================

    Const sPROC As String = "FormatSheets"

    Dim wks     As Excel.Worksheet


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If ObjectExists(PB) Then
        PB.Reset
        PB.Max = Worksheets.Count
    End If

    For Each wks In ThisWorkbook.Worksheets
        If ObjectExists(PB) Then
            PB.Caption = vbNewLine & "Formatting " & wks.CodeName
            PB.Increment
            DoEvents
        End If
        Call FormatSheet(wks, Reset, LoadData)
    Next wks

    ' ----------------------------------------------------------------------

PROC_EXIT:

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
