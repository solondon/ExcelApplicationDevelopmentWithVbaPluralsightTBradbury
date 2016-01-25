Attribute VB_Name = "MMSExcelRowsColumns"
' ==========================================================================
' Module      : MMSExcelRowsColumns
' Type        : Module
' Description : Support for working with rows and columns in Excel
' --------------------------------------------------------------------------
' Procedures  : ColumnLetter            String
'               ColumnNumber            Long
'               DeleteHiddenRows
'               FindColumn              Long
'               FindRow                 Long
'               FreezeHeaderRow
'               LastUsedColumn          Long
'               LastUsedRow             Long
'               ShowColumns
'               ShowRows
'               UnFreezeHeaderRow
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

' Row/Column limits
' -----------------
Public Const glEXCEL_MAXROWS_2003   As Long = 65536   '  &H10000
Public Const glEXCEL_MAXROWS_2007   As Long = 1048576    ' &H100000
Public Const glEXCEL_MAXCOLS_2003   As Long = 256     '    &H100
Public Const glEXCEL_MAXCOLS_2007   As Long = 16384   '   &H4000


' Column widths
' -------------
Public Const gdblCOLWIDTH_DEFAULT   As Double = 8.43   ' 64 pixels

'Public Const gdblCOLWIDTH_15PX      As Double = 1.43   ' 15 pixels
Public Const gdblCOLWIDTH_20PX      As Double = 2.14   ' 20 pixels
Public Const gdblCOLWIDTH_25PX      As Double = 2.86   ' 25 pixels
Public Const gdblCOLWIDTH_30PX      As Double = 3.57   ' 30 pixels
Public Const gdblCOLWIDTH_40PX      As Double = 5#     ' 40 pixels
Public Const gdblCOLWIDTH_50PX      As Double = 6.43   ' 50 pixels
Public Const gdblCOLWIDTH_60PX      As Double = 7.86   ' 60 pixels
Public Const gdblCOLWIDTH_75PX      As Double = 10#    ' 75 pixels
Public Const gdblCOLWIDTH_90PX      As Double = 12.14  ' 90 pixels

Public Const gdblCOLWIDTH_100PX     As Double = 13.57  '100 pixels
'Public Const gdblCOLWIDTH_110PX     As Double = 15#    '110 pixels
Public Const gdblCOLWIDTH_120PX     As Double = 16.43  '120 pixels
Public Const gdblCOLWIDTH_125PX     As Double = 17.14  '125 pixels
Public Const gdblCOLWIDTH_145PX     As Double = 20#    '145 pixels
Public Const gdblCOLWIDTH_150PX     As Double = 20.71  '150 pixels
Public Const gdblCOLWIDTH_175PX     As Double = 24.29  '175 pixels

Public Const gdblCOLWIDTH_200PX     As Double = 27.86  '200 pixels
Public Const gdblCOLWIDTH_215PX     As Double = 30#    '215 pixels
'Public Const gdblCOLWIDTH_225PX     As Double = 31.43  '225 pixels
'Public Const gdblCOLWIDTH_250PX     As Double = 35#    '250 pixels
Public Const gdblCOLWIDTH_275PX     As Double = 38.57  '275 pixels
'Public Const gdblCOLWIDTH_285PX     As Double = 40#    '285 pixels

Public Const gdblCOLWIDTH_300PX     As Double = 42.14  '300 pixels
Public Const gdblCOLWIDTH_325PX     As Double = 45.71  '325 pixels
Public Const gdblCOLWIDTH_350PX     As Double = 49.29  '350 pixels
'Public Const gdblCOLWIDTH_355PX     As Double = 50#    '350 pixels
'Public Const gdblCOLWIDTH_375PX     As Double = 52.86  '375 pixels

Public Const gdblCOLWIDTH_400PX     As Double = 56.43  '400 pixels
'Public Const gdblCOLWIDTH_425PX     As Double = 60#    '425 pixels
'Public Const gdblCOLWIDTH_450PX     As Double = 63.57  '450 pixels
'Public Const gdblCOLWIDTH_475PX     As Double = 67.14  '475 pixels
'Public Const gdblCOLWIDTH_495PX     As Double = 70#    '495 pixels

'Public Const gdblCOLWIDTH_500PX     As Double = 70.71  '500 pixels

Public Const gsngAUTOFILTER_SPACER  As Single = 3#

' ----------------
' Module Level
' ----------------

Private Const msMODULE              As String = "MMSExcelRowsColumns"

Public Function ColumnLetter(ByVal Column As Long) As String
' ==========================================================================
' Description : Return the letter for a column number.
'
' Parameters  : The column number to convert to letters.
'
' Returns     : String      The column letter(s).
' ==========================================================================

    Dim sRtn    As String

    While (Column > 0)
        Column = Column - 1
        sRtn = Chr(65 + (Column Mod 26)) + sRtn
        Column = Column \ 26
    Wend

    ColumnLetter = sRtn

End Function

Public Function ColumnNumber(ByVal Column As String) As Long
' ==========================================================================
' Description : Return the number for a given column letter.
'
' Parameters  : Column      The column letter(s) to convert.
'
' Returns     : Long
' ==========================================================================

    Const sPROC     As String = "ColumnNumber"

    Const lASCII_D  As Long = 68
    Const lASCII_F  As Long = 70
    Const lASCII_I  As Long = 73
    Const lASCII_V  As Long = 86
    Const lASCII_X  As Long = 88

    Dim lLen        As Long
    Dim lRtn        As Long
    Dim sSource     As String
    Dim sDescr      As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Check for upper case
    ' --------------------
    Column = Trim(UCase(Column))

    sSource = Concat(".", msMODULE, sPROC)
    lLen = Len(Column)

    ' No value provided
    ' -----------------
    If (lLen = 0) Then
        sDescr = "No column letter provided."
        Call Err.Raise(ERR_ARGUMENT_NOT_OPTIONAL, sSource, sDescr)
    End If

    ' ----------------------------------------------------------------------

    If (CInt(Application.Version) > OfficeVersion2003) Then

        ' Excel 2007 allows up to 16,384 columns (XFD)
        ' --------------------------------------------
        If (lLen > 3) Then
            sDescr = "Invalid column letter ('" & Column & "')."
            Call Err.Raise(9, sSource, sDescr)
        ElseIf (lLen = 3) Then
            If (Asc(Left$(Column, 1)) > lASCII_X) _
            Or ((Asc(Left$(Column, 1)) = lASCII_X) _
            And (Asc(Mid$(Column, 2, 1)) > lASCII_F)) _
            Or ((Asc(Left$(Column, 1)) = lASCII_X) _
            And (Asc(Mid$(Column, 2, 1)) = lASCII_F) _
            And (Asc(Right$(Column, 1)) > lASCII_D)) Then
                sDescr = "Invalid column letter ('" & Column & "')."
                Call Err.Raise(ERR_SUBSCRIPT_OUT_OF_RANGE, sSource, sDescr)
            End If
        End If

    ' ----------------------------------------------------------------------

    Else

        ' Excel 2003 allows up to 256 columns (IV)
        ' ----------------------------------------
        If (lLen > 2) Then
            sDescr = "Invalid column letter ('" & Column & "')."
            Call Err.Raise(9, sSource, sDescr)
        ElseIf (lLen = 2) Then
            If (Asc(Left(Column, 1)) > lASCII_I) _
            Or (lLen = 2) _
            And (Asc(Left(Column, 1)) = lASCII_I) _
            And (Asc(Right(Column, 1)) > lASCII_V) Then
                sDescr = "Invalid column letter ('" & Column & "')."
                Call Err.Raise(ERR_SUBSCRIPT_OUT_OF_RANGE, sSource, sDescr)
            End If
        End If

    End If

    ' ----------------------------------------------------------------------

    lRtn = ActiveSheet.Range(Column & "1").Column

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ColumnNumber = lRtn

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

Public Sub DeleteHiddenRows(ByRef Sheet As Excel.Worksheet, _
                            ByVal ShowProgress As Boolean, _
                   Optional ByRef PB As IProgressBar)
' ==========================================================================
' Description : Delete hidden rows from the UsedRange of a worksheet.
'
' Parameters  : Sheet         The worksheet to inspect
'               ShowProgress  Show a ProgressBar during operation
'               PB            The ProgressBar to use
' ==========================================================================

    Const sPROC     As String = "DeleteHiddenRows"

    Dim bCreated    As Boolean

    Dim lProgress   As Long
    Dim lRowCnt     As Long
    Dim lRowLast    As Long
    Dim lRowFirst   As Long
    Dim lRow        As Long

    Dim udtProps As TApplicationProperties

    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE

    ' ----------------------------------------------------------------------
    ' Create the ProgressBar if needed
    ' --------------------------------

    If (ShowProgress And (PB Is Nothing)) Then
        Set PB = New FProgressBar
        bCreated = True
    End If

    With Sheet
        If ShowProgress Then
            PB.Caption = "Deleting hidden rows"
            PB.Max = .Rows.Count + 1
            PB.Show
        End If

        lRowFirst = .UsedRange.Rows(1).Row
        lRowLast = (lRowFirst + .UsedRange.Rows.Count - 1)

        For lRow = lRowLast To lRowFirst Step -1

            ' Update the ProgressBar
            ' ----------------------
            If ShowProgress Then
                PB.Increment
            End If

            ' Delete the row
            ' --------------
            If (.Rows(lRow).Hidden Or (.Rows(lRow).Height = 0)) Then
                .Rows(lRow).Delete
            End If

        Next lRow
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    If bCreated Then
        PB.Hide
        Set PB = Nothing
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

Public Function FindColumn(ByRef Sheet As Excel.Worksheet, _
                           ByVal SearchVal As String, _
                  Optional ByVal SearchRow As Long = 1, _
                  Optional ByVal StartCol As Long = 1) As Long
' ==========================================================================
' Description : Search for a value and return the column number
'
' Parameters  : Sheet       The worksheet to search
'               SearchVal   The value to search for
'               SearchRow   The row to look in
'               StartCol    The starting column of the search
'
' Returns     : Long
' ==========================================================================

    Const sPROC As String = "FindCol"

    Dim lColCt  As Long

    Dim lRtn    As Long
    Dim rng     As Excel.Range
    Dim Cell    As Excel.Range


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, SearchVal)

    ' ----------------------------------------------------------------------

    If (SearchVal = vbNullString) Then
        GoTo PROC_EXIT
    End If

TRY_AGAIN:

    lColCt = Sheet.UsedRange.Columns.Count

    With Sheet
        With .Range(.Cells(SearchRow, StartCol), .Cells(SearchRow, lColCt))
            Set rng = .Find(What:=SearchVal, _
                            After:=.Cells(1, lColCt), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchDirection:=xlNext, _
                            SearchOrder:=xlByColumns, _
                            MatchCase:=False)

            If (rng Is Nothing) Then
                For Each Cell In .Cells
                    If (UCase(Cell) = UCase(SearchVal)) Then
                        lRtn = Cell.Column
                        Exit For
                    End If
                Next Cell
            End If
        End With
    End With

    If (Not (rng Is Nothing)) Then
        lRtn = rng.Column
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    FindColumn = lRtn

    Set Cell = Nothing
    Set rng = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, lRtn)
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

Public Function FindRow(ByRef Sheet As Excel.Worksheet, _
                        ByVal SearchVal As String, _
               Optional ByVal SearchCol As Long = 1, _
               Optional ByVal StartRow As Long = 1) As Long
' ==========================================================================
' Description : Search for a value and return the row number
'
' Parameters  : Sheet       The worksheet to search
'               SearchVal   The value to search for
'               SearchCol   The column to look in
'               StartRow    The starting row of the search
'
' Returns     : Long
' ==========================================================================

    Const sPROC     As String = "FindRow"

    Dim bSaved      As Boolean: bSaved = ThisWorkbook.Saved
    Dim bHidden     As Boolean

    Dim lRowCt      As Long

    Dim lRtn        As Long
    Dim rng         As Excel.Range

    Dim udtProps    As TApplicationProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, SearchVal)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE


    If (SearchVal = vbNullString) Then
        GoTo PROC_EXIT
    End If

    With Sheet.Columns(SearchCol).EntireColumn
        If .Hidden Then
            .Hidden = False
            bHidden = True
        End If
    End With

    lRowCt = LastUsedRow(Sheet)

    With Sheet
        With .Range(.Cells(StartRow, SearchCol), .Cells(lRowCt, SearchCol))
            Set rng = .Find(What:=SearchVal, _
                            After:=.Cells(1, 1), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
        End With
        If (Not rng Is Nothing) Then
            lRtn = rng.Row
        End If
    End With

    If bHidden Then
        Sheet.Columns(SearchCol).EntireColumn.Hidden = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    FindRow = lRtn

    Set rng = Nothing

    Call WorkbookSavedHasChanged(bSaved)

    Call Trace(tlMaximum, msMODULE, sPROC, lRtn)
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

Public Sub FreezeHeaderRow(ByRef Sheet As Excel.Worksheet, _
                  Optional ByVal HeaderRow As Long = 1, _
                  Optional ByVal FirstDataCol As Long = 1)
' ==========================================================================
' Description : Freeze the header row for the sheet
'
' Parameters  : Sheet         The worksheet to modify
'               HeaderRow     The last row before the split
'               FirstDataCol  The first column after the split
' ==========================================================================

    Const sPROC     As String = "FreezeHeaderRow"

    Dim sAddress    As String

    Dim eVisible    As XlSheetVisibility
    Dim wksActive   As Excel.Worksheet
    Dim udtProps    As TApplicationProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Sheet.CodeName)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE

    Set wksActive = ActiveSheet

    With Sheet
        ' Store the current visibility
        ' ----------------------------
        eVisible = Sheet.Visible

        ' Make it visible if hidden
        ' -------------------------
        .Visible = xlSheetVisible

        ' Make it active (required)
        ' -------------------------
        .Activate

        ' Locate where the new split goes
        ' -------------------------------
        sAddress = ColumnLetter(FirstDataCol) & CStr(HeaderRow + 1)
        .Cells(HeaderRow + 1, FirstDataCol).Select

        With ActiveWindow
            ' Clear old splits
            ' ----------------
            .FreezePanes = False
            .Split = False

            ' Add the new split
            ' -----------------
            .FreezePanes = True
        End With

        ' Restore previous visibility
        ' ---------------------------
        .Visible = eVisible
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ' Restore the active window
    ' -------------------------
    wksActive.Activate
    Set wksActive = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

    Call Trace(tlMaximum, msMODULE, sPROC, sAddress)
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

Public Function LastUsedColumn(ByRef Sheet As Excel.Worksheet, _
                      Optional ByVal Row As Long) As Long
' ==========================================================================
' Description : Find the last used row on a worksheet
'
' Parameters  : Sheet   The worksheet to inspect
'               Row     Limit search to a specific row
'
' Returns     : Long
' ==========================================================================

    Const sPROC As String = "LastUsedColumn"

    Dim lRtn    As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Column version
    ' --------------

    If (Row > 0) Then
        With Sheet
            lRtn = .Cells(Row, 1).SpecialCells(xlCellTypeLastCell).Column

        End With

        ' Regular version
        ' ---------------
    Else
        With Sheet
            lRtn = .UsedRange.Find(What:="*", _
                                   After:=.UsedRange.Cells(1), _
                                   LookAt:=xlPart, _
                                   LookIn:=xlFormulas, _
                                   SearchOrder:=xlByColumns, _
                                   SearchDirection:=xlPrevious, _
                                   MatchCase:=False).Column

        End With
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    LastUsedColumn = lRtn

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

Public Function LastUsedRow(ByRef Sheet As Excel.Worksheet, _
                   Optional ByVal Column As Long) As Long
' ==========================================================================
' Description : Find the last used row on a worksheet
'
' Parameters  : Sheet     The worksheet to inspect
'               Column    Limit search to a specific column
'
' Returns     : Long
' ==========================================================================

    Const sPROC As String = "LastUsedRow"

    Dim lRtn    As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Column version
    ' --------------

    If (Column > 0) Then
        With Sheet
            lRtn = .Cells(.UsedRange.Rows.Count, Column).End(xlUp).Row
        End With

        ' Regular version
        ' ---------------
    Else
        With Sheet.UsedRange
            lRtn = .Find(What:="*", _
                         After:=.Cells(1), _
                         LookAt:=xlPart, _
                         LookIn:=xlFormulas, _
                         SearchOrder:=xlByRows, _
                         SearchDirection:=xlPrevious, _
                         MatchCase:=False).Row

        End With
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    LastUsedRow = lRtn

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

Public Sub ShowColumns(Optional ByRef Sheet As Excel.Worksheet)
' ==========================================================================
' Description : Un-hide all columns on the sheet
'
' Parameters  : Sheet   The worksheet to modify
' ==========================================================================

    Const sPROC As String = "ShowColumns"

    Dim bReset  As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (Sheet Is Nothing) Then
        bReset = True
        Set Sheet = ActiveSheet
    End If

    Sheet.Columns.Hidden = False

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, Sheet.CodeName)

    If bReset Then
        Set Sheet = Nothing
    End If

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

Public Sub ShowRows(Optional ByRef Sheet As Excel.Worksheet)
' ==========================================================================
' Description : Un-hide all rows on the sheet
'
' Parameters  : Sheet   The worksheet to modify
' ==========================================================================

    Const sPROC As String = "ShowRows"

    Dim bReset  As Boolean

    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (Sheet Is Nothing) Then
        bReset = True
        Set Sheet = ActiveSheet
    End If

    Sheet.Rows.Hidden = False

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, Sheet.CodeName)

    If bReset Then
        Set Sheet = Nothing
    End If

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

Public Sub UnFreezeHeaderRow(ByRef Sheet As Excel.Worksheet)
' ==========================================================================
' Description : Clear splits and frozen rows from a sheet
'
' Parameters  : Sheet   The worksheet to modify
' ==========================================================================

    Const sPROC     As String = "UnFreezeHeaderRow"

    Dim bReset      As Boolean
    Dim wks         As Excel.Worksheet
    Dim eVis        As XlSheetVisibility

    Dim udtProps    As TApplicationProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE

    Set wks = ActiveSheet

    If (Sheet Is Nothing) Then
        bReset = True
        Set Sheet = ActiveSheet
    End If

    eVis = Sheet.Visible
    Sheet.Visible = xlSheetVisible
    Sheet.Activate

    With ActiveWindow
        If .FreezePanes Then
            .FreezePanes = False
        End If

        If .Split Then
            .Split = False
        End If
    End With

    wks.Activate
    Sheet.Visible = eVis

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

    Call Trace(tlMaximum, msMODULE, sPROC, Sheet.CodeName)

    If bReset Then
        Set Sheet = Nothing
    End If

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
