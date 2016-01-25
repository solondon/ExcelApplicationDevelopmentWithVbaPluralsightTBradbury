Attribute VB_Name = "MMSExcelRanges"
' ==========================================================================
' Module      : MMSExcelRanges
' Type        : Module
' Description : Support for working with Excel Ranges
' --------------------------------------------------------------------------
' Procedures  : ArrayToRange
'               ClearErrors
'               ClearValues
'               GetMergedCells          Excel.Range
'               GetRange                Excel.Range
'               RangesMatch             Boolean
'               RangeToArray            Variant
'               RangeToString           String
'               ResetUsedRange
'               UniqueItemsInRange      Variant
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

Private Const msMODULE As String = "MMSExcelRanges"

Public Sub ArrayToRange(ByRef Arr As Variant, ByRef Target As Excel.Range)
' ==========================================================================
' Description : Spill the contents of an array onto a worksheet
'
' Parameters  : Arr       The source array to display
'               Target    The location to place the data
' ==========================================================================

    Const sPROC As String = "ArrayToRange"

    Dim lRow    As Long    ' Destination row
    Dim lCol    As Long    ' Destination column

    Dim lIdxX   As Long    ' Array index dimensions
    Dim lIdxY   As Long    ' Array index elements

    Dim lLbx    As Long    ' Lower bounds dimensions
    Dim lUBX    As Long    ' Upper bounds dimensions

    Dim lLBY    As Long    ' Lower bounds elements
    Dim lUBY    As Long    ' Upper bounds elements


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    lLbx = 1
    lUBX = NumberOfDimensions(Arr)
    If (lUBX > 1) Then
        lUBX = UBound(Arr, 1)
        lUBY = UBound(Arr, 2)
    End If

    lCol = Target.Cells(1).Column

    If IsObject(Arr) Then
        If (TypeName(Arr) = gsTYPENAME_EXCEL_RANGE) Then
            lLBY = 1
            lUBY = Arr.Rows.Count
        End If
    End If

    ' For each Dimension
    ' ------------------
    For lIdxX = lLbx To lUBX
        lRow = Target.Cells(1).Row

        ' For each element
        ' ----------------
        For lIdxY = lLBY To lUBY
            Target.Worksheet.Cells(lRow, lCol) = Arr(lIdxX, lIdxY)
            lRow = lRow + 1
        Next lIdxY

        lCol = lCol + 1
    Next lIdxX

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

Public Sub ClearErrors(ByRef Range As Excel.Range, _
              Optional ByVal ErrorType As XlErrorChecks = xlNumberAsText, _
              Optional ByVal ConvertToNumber As Boolean = False, _
              Optional ByVal Redraw As Boolean = False, _
              Optional ByRef PB As IProgressBar, _
              Optional ByVal Message As String)
' ==========================================================================
' Description : Removes the error warnings from cells.
'
' Parameters  : Range           The range of cells to clear errors from
'               ErrorType       The type of error to clear
'               ConvertToNumber Convert numeric string values to numbers
'               Message         String to pass to the progress bar
'               Redraw          Flag to redraw during processing
' ==========================================================================

    Const sPROC     As String = "ClearErrors"

    Dim bAutoSelect As Boolean  'For debugging
    Dim bUpdatePB   As Boolean

    Dim lRow        As Long
    Dim lCol        As Long

    Dim lCellCt     As Long
    Dim lCellIdx    As Long

    Dim Cell        As Excel.Range
    Dim udtProps    As TApplicationProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Range.Worksheet.CodeName)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE

    ' Turn off options
    ' ----------------
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    ' Set the ProgressBar if needed
    ' -----------------------------
    lCellCt = Range.Cells.Count

    bUpdatePB = ((Len(Message) > 0) And (Not IsMissing(PB)))

    If bUpdatePB And (Not (PB Is Nothing)) Then
        PB.Min = 1
        PB.Max = lCellCt
        PB.Caption = Message
    End If

    ' Loop through the range
    ' ----------------------
    For Each Cell In Range.Cells

        ' Update the ProgressBar
        ' ----------------------
        lCellIdx = lCellIdx + 1

        If bUpdatePB And (Not (PB Is Nothing)) Then
            PB.Value = lCellIdx
        End If

        ' Make the address visible in the locals window
        ' ---------------------------------------------
        lRow = Cell.Row
        lCol = Cell.Column

        ' Show which cell is being processed
        ' ----------------------------------
        If bAutoSelect Then
            Cell.Select
        End If

        ' Clear the error
        ' ---------------
        If Cell.Errors.Item(ErrorType).Value = True Then
            If ConvertToNumber Then
                Cell.Copy
                Cell.TextToColumns Destination:=Range(Cell.Address)
            Else
                Cell.Errors(ErrorType).Ignore = True
            End If
        End If

    Next Cell

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

    Call Trace(tlMaximum, msMODULE, sPROC, Range.Address)
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

Public Sub ClearValues(ByRef ValueRange As Excel.Range, _
              Optional ByVal ValueTypes As XlSpecialCellsValue _
                                         = xlNumbers _
                                         + xlTextValues _
                                         + xlLogical _
                                         + xlErrors)
' ==========================================================================
' Description : Remove values from cells but leave formatting and formulas.
'
' Parameters  : ValueRange  The Range of Cells to clear.
'               ValueTypes  The types of values to clear.
'                           The default is to clear everything.
' ==========================================================================

    Const sPROC As String = "ClearValues"

    Dim rng     As Excel.Range


    On Error Resume Next
    Call Trace(tlMaximum, msMODULE, sPROC, ValueRange.Worksheet)

    ' ----------------------------------------------------------------------
    ' Get a reference to the range to clear
    ' -------------------------------------
    Set rng = ValueRange.SpecialCells(xlCellTypeConstants, ValueTypes)

    If (Err.Number = 0) Then
        rng.ClearContents
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    If (rng Is Nothing) Then
        Call Trace(tlMaximum, msMODULE, sPROC, "Error finding range")
    Else
        Call Trace(tlMaximum, msMODULE, sPROC, rng.Address)
    End If

    Set rng = Nothing

    On Error GoTo 0

End Sub

Public Function GetMergedCells(ByRef SearchRange As Variant) As Excel.Range
' ==========================================================================
' Description : Search a range for merged cells
'
' Parameters  : SearchRange   A range or address to search
'
' Returns     : Excel.Range
' ==========================================================================

    Const sPROC     As String = "GetMergedCells"

    Dim sAddress    As String

    Dim rngCell     As Excel.Range
    Dim rngRtn      As Excel.Range
    Dim rngSearch   As Excel.Range


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the range to search
    ' -----------------------

    If (TypeName(SearchRange) = gsTYPENAME_STRING) Then
        Set rngSearch = Range(SearchRange)
    Else
        Set rngSearch = SearchRange
    End If

    ' ----------------------------------------------------------------------
    ' Set the type of find to use
    ' ---------------------------

    Application.FindFormat.MergeCells = True

    ' Find the first merged cell
    ' --------------------------
    Set rngCell = rngSearch.Find(vbNullString, _
                                 LookAt:=xlPart, _
                                 SearchFormat:=True)

    ' If a merge is found look for more
    ' ---------------------------------
    If Not (rngCell Is Nothing) Then
        sAddress = rngCell.Address

        Do

            ' Combine found cells
            ' -------------------
            If (rngRtn Is Nothing) Then
                Set rngRtn = rngCell
            Else
                Set rngRtn = Union(rngRtn, rngCell)
            End If

            ' Search for the next one
            ' -----------------------
            Set rngCell = rngSearch.Find(vbNullString, _
                                         After:=rngCell, _
                                         LookAt:=xlPart, _
                                         SearchFormat:=True)
            If (rngCell Is Nothing) Then
                Exit Do
            End If
        Loop While ((rngCell.Address <> sAddress) _
                    And (Not (rngCell Is Nothing)))
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set GetMergedCells = rngRtn

    Application.FindFormat.Clear

    Set rngCell = Nothing
    Set rngRtn = Nothing
    Set rngSearch = Nothing

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

Public Function GetRange(ByRef Sheet As Excel.Worksheet, _
                Optional ByVal SearchValue As Variant = vbNullString, _
                Optional ByVal SearchColumn As Variant = 1, _
                Optional ByVal HeaderRowsToSkip As Long = 1) As Excel.Range
' ==========================================================================
' Description : Return the 'real' UsedRange from a worksheet.
'
' Parameters  : Sheet               The worksheet to use
'               SearchValue         The value in SearchColumn to look for
'               SearchColumn        The column number or letter to search
'               HeaderRowsToSkip    The number of header rows to ignore
'
' Returns     : Excel.Range
'
' Notes       : GetRange without any parameters will return the UsedRange
'               without any header row. If a subset is needed, supply a
'               column and a search value to restrict the range to rows that
'               match the search criteria. These rows must be contiguous.
' ==========================================================================

    Const sPROC     As String = "GetRange"


    Dim bFound      As Boolean
    Dim bReHide     As Boolean

    Dim lSR         As Long  ' Start row
    Dim lER         As Long  ' End row
    Dim lRC         As Long  ' Row count

    Dim lSC         As Long  ' Start column
    Dim lEC         As Long  ' End column
    Dim lCC         As Long  ' Column count

    Dim sSrchRng    As String

    Dim eVisible    As XlSheetVisibility

    Dim rngFound    As Excel.Range
    Dim rngSearch   As Excel.Range
    Dim rngLastCell As Excel.Range

    Dim udtProps    As TApplicationProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, SearchValue)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtProps)
    Application.ScreenUpdating = gbDEBUG_MODE

    ' Quit if a sheet was not provided
    ' --------------------------------
    If (Sheet Is Nothing) Then
        GoTo PROC_EXIT
    End If

    With Sheet
        eVisible = .Visible
        .Visible = xlSheetVisible

        lRC = .UsedRange.Rows.Count

        ' If the number of rows is less
        ' than the skipped rows, then quit
        ' --------------------------------
        If (lRC <= HeaderRowsToSkip) Then
            GoTo PROC_EXIT
        End If

        ' Find the start and end row/colum
        ' --------------------------------
        lSR = .UsedRange.Rows(1).Row
        lER = .UsedRange.Rows(lRC).Row

        lCC = .UsedRange.Columns.Count
        lSC = .UsedRange.Columns(1).Column
        lEC = .UsedRange.Columns(lCC).Column

        ' Adjust for skipped rows
        ' -----------------------
        If (HeaderRowsToSkip > 0) Then
            lSR = lSR + HeaderRowsToSkip
        End If

        ' If no search value use all of it
        ' --------------------------------
        If (SearchValue = vbNullString) Then
            GoTo BUILD_RANGE
        End If

        'Find the start row for the position
        '-----------------------------------
        sSrchRng = ColumnLetter(SearchColumn) & lSR & _
                   ":" & _
                   ColumnLetter(SearchColumn) & lER
        Set rngSearch = .Range(sSrchRng)
        Set rngLastCell = rngSearch.Cells(rngSearch.Cells.Count)

        ' A search cannot be done on a hidden cell
        ' ----------------------------------------
        If rngLastCell.EntireColumn.Hidden Then
            bReHide = True
            rngLastCell.EntireColumn.Hidden = False
        End If

        ' Do the search
        ' -------------
        Set rngFound = rngSearch.Find(What:=SearchValue, _
                                      After:=rngLastCell, _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole, _
                                      SearchOrder:=xlByRows, _
                                      MatchCase:=False)
        ' Quit if not found
        ' -----------------
        If (rngFound Is Nothing) Then
            GoTo PROC_EXIT
        End If

        bFound = True
        lSR = rngFound.Row
        lER = rngFound.Row

        ' Gather all of the subsequent matches
        ' ------------------------------------
        Do While bFound
            bFound = False

            Set rngFound = rngFound.Offset(RowOffset:=1)

            If (rngFound.Value = rngFound.Offset(RowOffset:=-1).Value) Then
                bFound = True
                lER = rngFound.Row
            End If
        Loop

BUILD_RANGE:

        'Return the range
        '----------------
        Set GetRange = .Range(.Cells(lSR, lSC), .Cells(lER, lEC))
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    If (Not (Sheet Is Nothing)) Then
        Sheet.Visible = eVisible
    End If

    If (Not (rngLastCell Is Nothing)) And bReHide Then
        rngLastCell.EntireColumn.Hidden = True
    End If

    Set rngSearch = Nothing
    Set rngFound = Nothing
    Set rngLastCell = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtProps)

    If (GetRange Is Nothing) Then
        Call Trace(tlMaximum, msMODULE, sPROC, "RANGE NOT FOUND")
    Else
        Call Trace(tlMaximum, msMODULE, sPROC, _
                   GetRange.Worksheet.CodeName _
                   & "(" & GetRange.Address & ")")
    End If

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

Public Function RangesMatch(ByRef Range1 As Excel.Range, _
                            ByRef Range2 As Excel.Range, _
                   Optional ByVal ProcessAll As Boolean, _
                   Optional ByVal HighlightDifferences As Boolean) _
       As Boolean
' ==========================================================================
' Description : Test if the contents of two ranges is the same
'
' Parameters  : Range1                  The first range
'               Range2                  The second range
'               ProcessAll              If True, examine all of the cells,
'                                       otherwise stop after first mismatch
'               HighlightDifferences    If True, mark cells with differences
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "RangesMatch"

    Dim bRtn    As Boolean: bRtn = True
    Dim lRow    As Long
    Dim lCell   As Long

    Dim Row     As Excel.Range
    Dim Cell    As Excel.Range


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' First check the RowCount
    ' ------------------------

    If (Range1.Rows.Count <> Range2.Rows.Count) Then
        bRtn = False
        If (Range1.Rows.Count > Range2.Rows.Count) Then
            GoTo SHORTROWS_TEST
        End If
    End If

NORMAL_TEST:

    For Each Row In Range1.Rows
        lRow = lRow + 1
        lCell = 0

        If (lRow > Range2.Rows.Count) Then
            GoTo PROC_EXIT
        End If

        If (Range1.Rows(lRow).Cells.Count _
         <> Range2.Rows(lRow).Cells.Count) Then
            bRtn = False
            If HighlightDifferences Then
                Range1.Rows(lRow).Cells.Style = gsEXCEL_STYLE_BAD
                Range2.Rows(lRow).Cells.Style = gsEXCEL_STYLE_BAD
            End If

            If (Not ProcessAll) Then
                GoTo PROC_EXIT
            End If

        Else
            For Each Cell In Row.Cells
                lCell = lCell + 1
                If (Range1.Rows(lRow).Cells(lCell) _
                 <> Range2.Rows(lRow).Cells(lCell)) Then
                    bRtn = False

                    If HighlightDifferences Then
                        Range1.Rows(lRow).Cells(lCell).Style _
                            = gsEXCEL_STYLE_BAD
                        Range2.Rows(lRow).Cells(lCell).Style _
                            = gsEXCEL_STYLE_BAD
                    End If

                    If (Not ProcessAll) Then
                        GoTo PROC_EXIT
                    End If

                End If
            Next Cell
        End If
    Next Row

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SHORTROWS_TEST:

    For Each Row In Range2.Rows
        lRow = lRow + 1
        lCell = 0

        If (lRow > Range1.Rows.Count) Then
            GoTo PROC_EXIT
        End If

        If (Range1.Rows(lRow).Cells.Count _
         <> Range2.Rows(lRow).Cells.Count) Then
            bRtn = False
            If HighlightDifferences Then
                Range1.Rows(lRow).Cells.Style _
                    = gsEXCEL_STYLE_BAD
                Range2.Rows(lRow).Cells.Style _
                    = gsEXCEL_STYLE_BAD
            End If

            If (Not ProcessAll) Then
                GoTo PROC_EXIT
            End If

        Else
            For Each Cell In Row.Cells
                lCell = lCell + 1
                If (Range1.Rows(lRow).Cells(lCell) _
                 <> Range2.Rows(lRow).Cells(lCell)) Then
                    bRtn = False

                    If HighlightDifferences Then
                        Range1.Rows(lRow).Cells(lCell).Style _
                            = gsEXCEL_STYLE_BAD
                        Range2.Rows(lRow).Cells(lCell).Style _
                            = gsEXCEL_STYLE_BAD
                    End If

                    If (Not ProcessAll) Then
                        GoTo PROC_EXIT
                    End If

                End If
            Next Cell
        End If
    Next Row


    ' ----------------------------------------------------------------------

PROC_EXIT:

    RangesMatch = bRtn

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

Public Function RangeToArray(ByRef rng As Excel.Range) As Variant
' ==========================================================================
' Description : Convert an Excel Range to a one-dimensional array
'
' Parameters  : Rng   The range to convert
'
' Returns     : Variant
' ==========================================================================

    Const sPROC As String = "RangeToArray"

    Dim lIdx    As Long: lIdx = -1
    Dim vRtn    As Variant
    Dim Cell    As Excel.Range


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    vRtn = Array()
    ReDim vRtn(0 To rng.Cells.Count - 1)
    lIdx = LBound(vRtn) - 1

    For Each Cell In rng.Cells
        lIdx = lIdx + 1
        vRtn(lIdx) = Cell
    Next Cell

    ' ----------------------------------------------------------------------

PROC_EXIT:

    RangeToArray = vRtn

    Set Cell = Nothing

    On Error Resume Next
    Erase vRtn
    vRtn = Empty

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

Public Function RangeToString(ByRef rng As Excel.Range, _
                     Optional ByVal Delimiter As String = " ") As String
' ==========================================================================
' Description : Convert a 1-dimensional array to a string
'
' Parameters  : Rng         The range to convert
'               Delimiter   The delimiter to place between elements
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "RangeToString"

    Dim sRtn    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    sRtn = Concat(Delimiter, rng)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    RangeToString = sRtn

    Call Trace(tlMaximum, msMODULE, sPROC, sRtn)
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

Public Sub ResetUsedRange(ByRef Sheet As Excel.Worksheet)
' ==========================================================================
' Description : Reset the UsedRange for a worksheet
'
' Parameters  : Sheet   The worksheet to reset
' ==========================================================================

    Const sPROC As String = "ResetUsedRange"

    Dim lRow    As Long
    Dim lCol    As Long

    Dim vRC     As Variant
    Dim vCC     As Variant

    Dim rng     As Excel.Range


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    With Sheet
        lRow = 0
        lCol = 0
        Set rng = .UsedRange

        On Error Resume Next
        lRow = .Cells.Find("*", _
                           After:=.Cells(1), _
                           LookIn:=xlFormulas, _
                           LookAt:=xlWhole, _
                           SearchDirection:=xlPrevious, _
                           SearchOrder:=xlByRows).Row
        lCol = .Cells.Find("*", _
                           After:=.Cells(1), _
                           LookIn:=xlFormulas, _
                           LookAt:=xlWhole, _
                           SearchDirection:=xlPrevious, _
                           SearchOrder:=xlByColumns).Column

        On Error GoTo PROC_ERR

        If (lRow * lCol = 0) Then
            .Columns.Delete
        Else
            .Range(.Cells(lRow + 1, 1), _
                   .Cells(.Rows.Count, 1)).EntireRow.Delete
            .Range(.Cells(1, lCol + 1), _
                   .Cells(1, .Columns.Count)).EntireColumn.Delete
        End If
        vRC = .UsedRange.Rows.Count
        vCC = .UsedRange.Columns.Count
    End With

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

Public Function UniqueItemsInRange(ByRef Source As Excel.Range, _
                          Optional ByVal Count As Boolean) As Variant
' ==========================================================================
' Purpose   : Returns the unique items within a range
'
' Arguments : Source  The range to return items from
'
'           : Count   If True, return the count of items.
'                     If false (default) or is missing,
'                     return an array of unique items.
'
' Returns   : Variant
' ==========================================================================

    Const sPROC     As String = "UniqueItemsInRange"

    Dim bMatched    As Boolean

    Dim lItemCt     As Long
    Dim lIdx        As Long
    Dim lUB         As Long

    Dim vItems()    As Variant  ' Array of items

    Dim Cell        As Excel.Range


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Loop through the source data array
    ' ----------------------------------
    For Each Cell In Source.Cells

        ' Reset the flag
        ' --------------
        bMatched = False

        ' Has the item been added?
        ' ------------------------
        If IsAllocated(vItems) Then
            For lIdx = LBound(vItems) To UBound(vItems)

                If (Cell.Value = vItems(lIdx)) Then
                    bMatched = True
                    Exit For
                End If

            Next lIdx
        End If

        ' If not in list, add the item
        ' ----------------------------
        If (Not bMatched) And (Not IsEmpty(Cell.Value)) Then
            lItemCt = lItemCt + 1
            lUB = lUB + 1
            ReDim Preserve vItems(1 To lUB)
            vItems(lUB) = Cell.Value
        End If

    Next Cell

    ' ----------------------------------------------------------------------

PROC_EXIT:

    If Count Then
        UniqueItemsInRange = CVar(lItemCt)
    Else
        UniqueItemsInRange = vItems
    End If

    ' Release the allocated memory
    ' ----------------------------
    Erase vItems

    Call Trace(tlMaximum, msMODULE, sPROC, lItemCt)
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
