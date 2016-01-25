Attribute VB_Name = "MMSExcelWorksheets"
' ==========================================================================
' Module      : MMSExcelWorksheets
' Type        : Module
' Description : Support for working with Excel Worksheets
' --------------------------------------------------------------------------
' Procedures  : ActivateTab
'               GetWorksheet            Excel.Worksheet
'               GetWorksheetIndex       Long
'               HideWorksheet
'               ListWindows
'               ListWorksheets
'               ResetWorksheet
'               ResetWorksheets
'               SetCodeName
'               ShowWorksheet
'               ShowWorksheets
'               WorksheetExists         Boolean
'               WorksheetIsVisible      Boolean
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

Private Const msMODULE As String = "MMSExcelWorksheets"

Public Sub ActivateTab(Optional ByVal TabId As Variant)
' ==========================================================================
' Description : Activate a visible worksheet by tab ID
'
' Parameters  : TabId       The tab to activate. This can either be the
'                           positional index of the tab (left to right)
'                           or the text on the tab (the worksheet name).
' ==========================================================================

    Const sPROC As String = "ActivateTab"

    Dim bFound  As Boolean

    Dim lIdx    As Long
    Dim lTab    As Long
    Dim sVis    As String
    Dim sTAB    As String

    Dim wks     As Excel.Worksheet


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, TabId)

    ' ----------------------------------------------------------------------
    ' Determine which type of processing to use
    ' -----------------------------------------

    If IsEmpty(TabId) Then
        lTab = 1
        GoTo FIND_NUMERIC

    ElseIf IsNumeric(TabId) Then
        lTab = CInt(TabId)
        GoTo FIND_NUMERIC

    Else
        sTAB = CStr(TabId)
    End If

    ' ----------------------------------------------------------------------

FIND_STRING:

    ' Find a match for TabId
    ' ----------------------
    For Each wks In Worksheets

        ' Only look for visible sheets
        ' ----------------------------
        If (wks.Visible = xlSheetVisible) Then

            ' Do a text compare on the tab
            ' ----------------------------
            If (StrComp(wks.Name, sTAB, vbTextCompare) = 0) Then
                wks.Activate
                bFound = True
                Exit For
            End If

        End If

    Next wks

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

FIND_NUMERIC:

    ' Find a match for TabId
    ' ----------------------
    For Each wks In Worksheets

        ' Only look for visible sheets
        ' ----------------------------
        If (wks.Visible = xlSheetVisible) Then
            lIdx = lIdx + 1
            sVis = wks.Name
        End If

        ' Do a compare on the tab index
        ' -----------------------------
        If (lIdx = lTab) Then
            wks.Activate
            bFound = True
            Exit For
        End If

    Next wks

    If ((Not bFound) And (lTab = 0)) Then
        ThisWorkbook.Worksheets(sVis).Activate
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set wks = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, ActiveSheet.CodeName)
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

Public Function GetWorksheet(ByVal Name As String, _
                    Optional ByVal UseCodeName As Boolean = True, _
                    Optional ByRef Workbook As Excel.Workbook) _
       As Excel.Worksheet
' ==========================================================================
' Description : Return a worksheet by name
'
' Parameters  : Name        The name that identifies the worksheet.
'               UseCodeName If True, use the CodeName instead of Name.
'               Workbook    Workbook to look in. Defaults to ThisWorkbook.
'
' Returns     : Excel.Worksheet
' ==========================================================================

    Const sPROC As String = "GetWorksheet"

    Dim bFound  As Boolean

    Dim wkb     As Excel.Workbook
    Dim wks     As Excel.Worksheet

    ' ----------------------------------------------------------------------
    ' Make sure a name was provided
    ' -----------------------------
    If (Name = vbNullString) Then
        Call Err.Raise(ERR_ARGUMENT_NOT_OPTIONAL, _
                       Concat(".", msMODULE, sPROC), _
                       "A valid worksheet name is required.")
    End If

    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------

    If (Workbook Is Nothing) Then    ' Use ThisWorkbook by default
        Set wkb = ThisWorkbook
        Call Trace(tlMaximum, msMODULE, sPROC, Name)
    Else
        Set wkb = Workbook
        Call Trace(tlMaximum, msMODULE, sPROC, wkb.Name & "." & Name)
    End If

    ' ----------------------------------------------------------------------

    If UseCodeName Then           ' Search by CodeName

        For Each wks In wkb.Worksheets
            If (StrComp(wks.CodeName, Name, vbTextCompare) = 0) Then
                bFound = True
                Exit For
            End If
        Next wks

    Else                          ' Search by display (Tab) name

        For Each wks In wkb.Worksheets
            If (StrComp(wks.Name, Name, vbTextCompare) = 0) Then
                bFound = True
                Exit For
            End If
        Next wks

    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    If bFound Then
        Set GetWorksheet = wks
    End If

    Call Trace(tlMaximum, _
               msMODULE, _
               sPROC, _
               "Found=" & (CStr(Not (wks Is Nothing))))

    Set wkb = Nothing
    Set wks = Nothing

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

Public Function GetWorksheetIndex(ByVal Name As String, _
                         Optional ByVal UseCodeName As Boolean = True, _
                         Optional ByRef Workbook As Excel.Workbook) As Long
' ==========================================================================
' Description : Determine the index (position) for a given Worksheet
'
' Parameters  : Name        The name that identifies the worksheet.
'               UseCodeName If True, use the CodeName instead of Name.
'               Workbook    Workbook to look in. Defaults to ThisWorkbook.
'
' Returns     : Long        The Worksheet index
' ==========================================================================

    Const sPROC     As String = "GetWorksheetIndex"

    Dim lIdx        As Long
    Dim lRtn        As Long

    Dim wkb         As Excel.Workbook
    Dim wks         As Excel.Worksheet

    Dim udtProps    As TApplicationProperties


    On Error GoTo PROC_ERR
    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)

    ' Use ThisWorkbook by default
    ' ---------------------------
    If (Workbook Is Nothing) Then
        Set wkb = ThisWorkbook
        Call Trace(tlVerbose, msMODULE, sPROC, Name)
    Else
        Set wkb = Workbook
        Call Trace(tlVerbose, msMODULE, sPROC, wkb.Name & "." & Name)
    End If

    If UseCodeName Then
        ' Search by CodeName
        ' ------------------
        For Each wks In wkb.Worksheets
            ' Increment the index
            ' -------------------
            lIdx = lIdx + 1

            ' Do a textual comparison
            ' -----------------------
            If (StrComp(wks.CodeName, Name, vbTextCompare) = 0) Then
                lRtn = lIdx
                Exit For
            End If
        Next wks

    Else
        ' Search by display (Tab) name
        ' ----------------------------
        For Each wks In wkb.Worksheets
            ' Increment the index
            ' -------------------
            lIdx = lIdx + 1

            ' Do a textual comparison
            ' -----------------------
            If (StrComp(wks.Name, Name, vbTextCompare) = 0) Then
                lRtn = lIdx
                Exit For
            End If
        Next wks

    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetWorksheetIndex = lRtn

    Set wkb = Nothing
    Set wks = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_EXIT)
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

Public Sub HideWorksheet(ByVal Name As String, _
                Optional ByVal UseCodeName As Boolean = True, _
                Optional ByVal Visibility _
                               As XlSheetVisibility = xlVeryHidden, _
                Optional ByRef Workbook As Excel.Workbook)
' ==========================================================================
' Description : Makes a Worksheet hidden if it is visible
'
' Parameters  : Name        The worksheet to be modified.
'               UseCodeName If True, use the CodeName instead of Name.
'               Visibility  The visibility to apply.
' ==========================================================================

    Const sPROC As String = "HideWorksheet"

    Dim wks     As Excel.Worksheet


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Name)

    ' ----------------------------------------------------------------------
    ' Get the worksheet to modify
    ' ---------------------------
    Set wks = GetWorksheet(Name, UseCodeName, Workbook)

    ' Set the visibility
    ' ------------------
    If (Not wks Is Nothing) Then
        wks.Visible = Visibility
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set wks = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, Name)
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

Public Sub ListWorksheets(Optional ByRef wkb As Excel.Workbook)
' ==========================================================================
' Description : List all worksheets in the workbook
'
' Parameters  : wkb   The workbook to use
' ==========================================================================

    Const sPROC As String = "ListWorksheets"

    Dim bDelete As Boolean
    Dim lIdx    As Long
    Dim sLine   As String
    Dim wks     As Excel.Worksheet


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (wkb Is Nothing) Then
        Set wkb = ThisWorkbook
        bDelete = True
    End If

    ' ----------------------------------------------------------------------

    Debug.Print String$(glLIST_LINELEN, gsLIST_LINECHAR)

    For Each wks In wkb.Worksheets
        lIdx = lIdx + 1

        sLine = wks.CodeName & " (" & wks.Name & ")"
        If (wks.Visible = xlSheetHidden) Then
            sLine = sLine & " (Hidden)"
        ElseIf (wks.Visible = xlSheetVeryHidden) Then
            sLine = sLine & " (Very Hidden)"
        End If
        Debug.Print sLine
    Next wks

    Debug.Print String$(glLIST_LINELEN, gsLIST_LINECHAR)
    Debug.Print CStr(lIdx) & " Sheets in " & wkb.Name
    Debug.Print String$(glLIST_LINELEN, gsLIST_LINECHAR)


    ' ----------------------------------------------------------------------

PROC_EXIT:

    If bDelete Then
        Set wkb = Nothing
    End If

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

Public Sub ResetWorksheet(Optional ByRef Sheet As Excel.Worksheet, _
                          Optional ByVal NewName As String, _
                          Optional ByVal ClearPageSetup As Boolean = True)
' ==========================================================================
' Description : Return worksheet to new condition.
'               Removes all values, formulas, formatting, shapes, etc.
'
' Parameters  : Sheet       The worksheet to reset
'               NewName     The new name for the sheet. CodeName is defualt.
' ==========================================================================

    Const sPROC     As String = "ResetWorksheet"

    Dim bDefault    As Boolean

    Dim lRC         As Long
    Dim lCC         As Long

    Dim oCP         As Excel.CustomProperty
    Dim oLO         As Excel.ListObject
    Dim oName       As Excel.Name
    Dim oPT         As Excel.PivotTable
    Dim oQT         As Excel.QueryTable
    Dim oShape      As Excel.Shape

    Dim Cell        As Excel.Range

    Dim wksActive   As Excel.Worksheet
    Dim eVisible    As XlSheetVisibility

    Dim udtProps    As TApplicationProperties


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    With Application
        .EnableEvents = False
        .ScreenUpdating = gbDEBUG_MODE
    End With

    ' Use the ActiveSheet if one is not provided
    ' ------------------------------------------
    If (Sheet Is Nothing) Then
        Set Sheet = ActiveSheet
        bDefault = True
    End If

    Call Trace(tlMaximum, msMODULE, sPROC, Sheet.CodeName)

    ' Set starting points
    ' -------------------
    Set wksActive = ActiveSheet
    eVisible = Sheet.Visible

    ' Show the sheet if it is hidden
    ' ------------------------------
    Sheet.Visible = xlSheetVisible

    ' Make the sheet active (required for some resets)
    ' ------------------------------------------------
    Sheet.Activate

    Call DebugAssert((ActiveSheet.CodeName = Sheet.CodeName), _
                     msMODULE, _
                     sPROC, _
                     "Unable to activate the correct worksheet.")

    With ActiveWindow

        ' Remove panes and splits
        ' -----------------------
        If .FreezePanes Then
            .FreezePanes = False
        End If

        If .Split Then
            .Split = False
        End If

        ' Reset magnification
        ' -------------------
        .Zoom = 100

        .DisplayFormulas = False
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayWorkbookTabs = True

    End With

    With Sheet

        ' Reset the tab
        ' -------------
        .Tab.Color = False

        ' Reset the name if provided
        ' --------------------------
        If (Len(NewName) > 0) Then
            .Name = NewName
        Else
            .Name = .CodeName
        End If

        ' Show any hidden rows or columns
        ' -------------------------------
        .Rows.Hidden = False
        .Columns.Hidden = False

        ' Un-filter the data
        ' ------------------
        If .FilterMode Then
            .ShowAllData
        End If

        ' Clear the AutoFilter if set
        ' ---------------------------
        If .AutoFilterMode Then
            .AutoFilterMode = False
        End If

        ' Separate the cells
        ' ------------------
        On Error Resume Next
        If .UsedRange.MergeCells Then
            .UsedRange.UnMerge

        ElseIf IsNull(.UsedRange.MergeCells) Then
            For Each Cell In GetMergedCells(.UsedRange)
                Cell.UnMerge
            Next Cell
        End If

        On Error GoTo PROC_ERR

        ' Reset height and width
        ' ----------------------
        With .UsedRange
            lRC = .Rows.Count
            lCC = .Columns.Count
            .ClearComments
            .ClearContents
            .ClearOutline
            .ClearNotes
            #If VBA7 Then
                .ClearHyperlinks
            #End If
            .ClearFormats
            .Clear
        End With

        .Rows.EntireRow.AutoFit
        .Rows.RowHeight = .StandardHeight
        .Columns.EntireColumn.AutoFit
        .Columns.ColumnWidth = .StandardWidth

        ' Clear the formatting
        ' --------------------
        With .UsedRange.Cells
            .Style = gsEXCEL_STYLE_NORMAL
            .NumberFormat = gsEXCEL_NUMFMT_GENERAL
        End With

        ' Reset PageSetup
        ' ---------------

        If ClearPageSetup Then
            Call ResetPageSetup(Sheet)
        End If

        ' Clear all of the collections
        ' ----------------------------
        For Each oCP In .CustomProperties
            oCP.Delete
        Next oCP

        For Each oLO In .ListObjects
            oLO.Delete
        Next oLO

        For Each oName In .Names
            oName.Delete
        Next oName

        For Each oName In ThisWorkbook.Names
            If (Left$(oName.RefersTo, _
                      Len(.Name) + 4) = "='" & .Name & "'!") Then
                oName.Delete
            End If
        Next oName

        For Each oPT In .PivotTables
            oPT.TableRange2.Clear
        Next oPT

        For Each oQT In .QueryTables
            oQT.Delete
        Next oQT

        For Each oShape In .Shapes
            oShape.Delete
        Next oShape
    End With

    ' ----------------------------------------------------------------------

    With Sheet
        ' Set the focus back to the home cell
        ' -----------------------------------
        Application.GoTo Reference:=.Range("A1"), _
                         Scroll:=True

        ' ------------------------------------------------------------------
        ' Make sure the UsedRange starts at the first cell
        ' ------------------------------------------------
        .Cells(1) = "Excel"

        ' Delete the data and
        ' reset the UsedRange
        ' -------------------
        .UsedRange.Delete
        lRC = .UsedRange.Rows.Count
        lCC = .UsedRange.Columns.Count

        Call ResetUsedRange(Sheet)

        Call DebugAssert((Sheet.UsedRange.Cells.Count = 1), _
                         msMODULE, _
                         sPROC, _
                         "The UsedRange for " _
                         & .CodeName & " was not cleared.")

        ' ------------------------------------------------------------------
        ' Restore previous visibility
        ' ---------------------------

        Sheet.Visible = eVisible

    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set Cell = Nothing

    Set oCP = Nothing
    Set oLO = Nothing
    Set oName = Nothing
    Set oPT = Nothing
    Set oQT = Nothing
    Set oShape = Nothing

    wksActive.Activate
    Set wksActive = Nothing

    If bDefault Then
        Set Sheet = Nothing
    End If

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

    Call Trace(tlMaximum, msMODULE, sPROC, NewName)
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

Public Sub ResetWorksheets()
' ==========================================================================
' Description : Run the reset procedure for each sheet in the workbook.
' ==========================================================================

    Const sPROC As String = "ResetWorksheets"

    Dim wks     As Excel.Worksheet


    On Error GoTo PROC_ERR
    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_ENTER)


    ' Reset all of the sheets (even if hidden)
    ' ----------------------------------------
    For Each wks In Application.Worksheets
        Call ResetWorksheet(wks)
    Next wks

    ' Make the first tab active
    ' -------------------------
    Call ActivateTab(1)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set wks = Nothing

    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_EXIT)
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

Public Sub SetCodeName(ByRef Sheet As Excel.Worksheet, _
                       ByVal CodeName As String)
' ==========================================================================
' Description : Change the CodeName for a worksheet
'
' Parameters  : Sheet       The worksheet to modify.
'               CodeName    The new CodeName for the worksheet.
'
' Comments    : This procedure should NOT be called from production code,
'               and should only be used during a build process.
'               This procedure requires the setting
'               "Trust access to the VBA project object model" to be
'               enabled in the Macro Settings area of the Trust Center.
' ==========================================================================

    Const sPROC As String = "SetCodeName"

    Dim oVBProject As Object
    Dim oVBComponent As Object

    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Sheet.Name)

    ' ----------------------------------------------------------------------

    Set oVBProject = Sheet.Parent.VBProject
    For Each oVBComponent In oVBProject.VBComponents
        If (oVBComponent.Name = Sheet.CodeName) Then
            oVBComponent.Name = CodeName
        End If
    Next oVBComponent

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set oVBProject = Nothing
    Set oVBComponent = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, CodeName)
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

Public Sub ShowWorksheet(ByVal Name As String, _
                Optional ByVal UseCodeName As Boolean = True, _
                Optional ByRef Workbook As Excel.Workbook)
' ==========================================================================
' Description : Make a Worksheet visible if it is hidden
'
' Parameters  : Name        The name that identifies the worksheet.
'               UseCodeName If True, use the CodeName instead of Name.
'               Workbook    Workbook to look in. Defaults to ThisWorkbook.
' ==========================================================================

    Const sPROC     As String = "ShowWorksheet"

    Dim wkb         As Excel.Workbook
    Dim wks         As Excel.Worksheet
    Dim udtProps    As TApplicationProperties


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)

    ' Turn off screen redrawing
    ' -------------------------
    Application.ScreenUpdating = gbDEBUG_MODE

    ' Use ThisWorkbook by default
    ' ---------------------------
    If (Workbook Is Nothing) Then
        Set wkb = ThisWorkbook
        Call Trace(tlVerbose, msMODULE, sPROC, Name)
    Else
        Set wkb = Workbook
        Call Trace(tlVerbose, msMODULE, sPROC, wkb.Name & "." & Name)
    End If

    If UseCodeName Then
        ' Search by CodeName
        ' ------------------
        For Each wks In wkb.Worksheets
            If (StrComp(wks.CodeName, Name, vbTextCompare) = 0) Then
                wks.Visible = xlSheetVisible
                Exit For
            End If
        Next wks

    Else
        ' Search by display (Tab) name
        ' ----------------------------
        For Each wks In wkb.Worksheets
            If (StrComp(wks.Name, Name, vbTextCompare) = 0) Then
                wks.Visible = xlSheetVisible
                Exit For
            End If
        Next wks

    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set wkb = Nothing
    Set wks = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

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

Public Sub ShowWorksheets()
' ==========================================================================
' Description : Make all Worksheets visible
' ==========================================================================

    Const sPROC     As String = "ShowWorksheets"

    Dim wks         As Excel.Worksheet
    Dim wksActive   As Excel.Worksheet
    Dim udtProps    As TApplicationProperties

    On Error GoTo PROC_ERR

    ' Get the current settings
    ' ------------------------
    Call GetApplicationProperties(udtProps)

    ' Turn off screen redraw
    ' ----------------------
    Application.ScreenUpdating = gbDEBUG_MODE

    ' Get the current sheet
    ' ---------------------
    Set wksActive = ActiveSheet

    ' Enumerate the Worksheets and make them visible
    ' ----------------------------------------------
    For Each wks In Application.Worksheets
        wks.Visible = xlSheetVisible
    Next wks

    ' Make sure the same sheet is active
    ' ----------------------------------
    wksActive.Activate

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set wks = Nothing
    Set wksActive = Nothing

    ' Restore the settings
    ' --------------------
    Call SetApplicationProperties(udtProps)

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

Public Function WorksheetExists(ByVal SheetName As String, _
                       Optional ByVal UseCodeName As Boolean, _
                       Optional ByRef Workbook As Excel.Workbook) As Boolean
' ==========================================================================
' Description : Determines if a given worksheet exists in the workbook
'
' Parameters  : SheetName     The name of the worksheet to find
'               UseCodeName   If True, use the CodeName for the search
'               Workbook      The workbook the sheet is located in.
'                             If not provided, ThisWorkbook will be used.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "WorksheetExists"

    Dim bRtn    As Boolean
    Dim wkb     As Excel.Workbook
    Dim wks     As Excel.Worksheet


    Dim udtProps As TApplicationProperties

    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    On Error Resume Next

    If (Workbook Is Nothing) Then
        Set wkb = ThisWorkbook
    Else
        Set wkb = Workbook
    End If

    If UseCodeName Then
        GoTo USE_CODE_NAME
    End If

    ' ----------------------------------------------------------------------
    ' Try the direct approach
    ' -----------------------
    Set wks = wkb.Worksheets(SheetName)
    bRtn = (Not (wks Is Nothing))
    If bRtn Then
        GoTo PROC_EXIT
    End If

    ' Enumerate the sheets
    ' --------------------
    For Each wks In wkb.Worksheets
        If (LCase$(wks.Name) = LCase$(SheetName)) Then
            bRtn = True
            Exit For
        End If
    Next wks

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

USE_CODE_NAME:

    For Each wks In wkb.Worksheets
        If LCase$(wks.CodeName) = LCase$(SheetName) Then
            bRtn = True
            Exit For
        End If
    Next wks

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WorksheetExists = bRtn
    Set wks = Nothing
    Set wkb = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

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

Public Function WorksheetIsVisible(ByVal Name As String, _
                          Optional ByVal UseCodeName As Boolean = True, _
                          Optional ByRef Workbook As Excel.Workbook) _
       As Boolean
' ==========================================================================
' Description : Determines if the sheet is currently visible.
'
' Parameters  : Name        The name that identifies the worksheet.
'               UseCodeName If True, use the CodeName instead of Name.
'               Workbook    Workbook to look in. Defaults to ThisWorkbook.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "WorksheetIsVisible"

    Dim bRtn    As Boolean

    On Error GoTo PROC_ERR
    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_ENTER)

    bRtn = (GetWorksheet(Name, _
                         UseCodeName, _
                         Workbook).Visible = xlSheetVisible)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    WorksheetIsVisible = bRtn

    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_EXIT)
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
