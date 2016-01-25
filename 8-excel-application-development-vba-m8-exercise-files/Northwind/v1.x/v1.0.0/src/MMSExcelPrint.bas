Attribute VB_Name = "MMSExcelPrint"
' ==========================================================================
' Module      : MMSExcelPrint
' Type        : Module
' Description : Printer-related functions
' --------------------------------------------------------------------------
' Procedures  : CopyPageSetup
'               PrinterIsAvailable    Boolean
'               ResetPageSetup
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

Private Const msMODULE As String = "MMSExcelPrint"

Public Sub CopyPageSetup(ByRef Src As Excel.Worksheet, _
                         ByRef Dst As Excel.Worksheet)
' ==========================================================================
' Description : Copy the PageSetup settings from one worksheet to another
'
' Parameters  : Src   The source worksheet to copy from
'               Dst   The destination worksheet to copy to
'
' Comments    : The PrinterIsAvailable function should be called prior to
'               using this function to ensure this wil run without error.
' ==========================================================================

    Const sPROC As String = "CopyPageSetup"


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    With Dst.PageSetup
        .PrintArea = Src.PageSetup.PrintArea

        .PrintTitleRows = Src.PageSetup.PrintTitleRows
        .PrintTitleColumns = Src.PageSetup.PrintTitleColumns

        .LeftHeader = Src.PageSetup.LeftHeader
        .CenterHeader = Src.PageSetup.CenterHeader
        .RightHeader = Src.PageSetup.RightHeader

        .LeftFooter = Src.PageSetup.LeftFooter
        .CenterFooter = Src.PageSetup.CenterFooter
        .RightFooter = Src.PageSetup.RightFooter

        .LeftMargin = Src.PageSetup.LeftMargin
        .RightMargin = Src.PageSetup.RightMargin
        .TopMargin = Src.PageSetup.TopMargin
        .BottomMargin = Src.PageSetup.BottomMargin
        .HeaderMargin = Src.PageSetup.HeaderMargin
        .FooterMargin = Src.PageSetup.FooterMargin

        .PrintHeadings = Src.PageSetup.PrintHeadings
        .PrintGridlines = Src.PageSetup.PrintGridlines
        .PrintComments = Src.PageSetup.PrintComments

        If CInt(.PrintQuality(1)) = 0 Then
            .PrintQuality = Array(600, 600)    ' (H, V)
        Else
            On Error Resume Next
            .PrintQuality = Src.PageSetup.PrintQuality
            On Error GoTo PROC_ERR
        End If

        .CenterHorizontally = Src.PageSetup.CenterHorizontally
        .CenterVertically = Src.PageSetup.CenterVertically

        .Orientation = Src.PageSetup.Orientation
        .Draft = Src.PageSetup.Draft
        .PaperSize = Src.PageSetup.PaperSize
        .FirstPageNumber = Src.PageSetup.FirstPageNumber
        .Order = Src.PageSetup.Order
        .BlackAndWhite = Src.PageSetup.BlackAndWhite
        .Zoom = Src.PageSetup.Zoom
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

Public Function PrinterIsAvailable() As Boolean
' ==========================================================================
' Description : Determines if a printer is available
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "PrinterIsAvailable"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = CBool(Len(Application.ActivePrinter))

    ' ----------------------------------------------------------------------

PROC_EXIT:

    PrinterIsAvailable = bRtn

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

Public Sub ResetPageSetup(ByRef Sheet As Excel.Worksheet)
' ==========================================================================
' Description : Return the PageSetup options to normal
'
' Parameters  : Sheet   The worksheet to modify
' ==========================================================================

    Const sPROC As String = "ResetPageSetup"


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    With Sheet.PageSetup
        .BlackAndWhite = False
        .Draft = False
        .PrintArea = vbNullString
        .PrintNotes = False
        .Order = xlDownThenOver

        .Orientation = xlPortrait

        .Zoom = 100
        .FitToPagesTall = 1
        .FitToPagesWide = 1

        .PaperSize = xlPaperLetter

        .PrintTitleColumns = vbNullString
        .PrintTitleRows = vbNullString

        .CenterHorizontally = False
        .CenterVertically = False

        .LeftMargin = 50.4
        .RightMargin = 50.4
        .TopMargin = 54
        .BottomMargin = 54

        .DifferentFirstPageHeaderFooter = False
        .OddAndEvenPagesHeaderFooter = False

        .LeftHeader = vbNullString
        .CenterHeader = vbNullString
        .RightHeader = vbNullString

        .LeftFooter = vbNullString
        .CenterFooter = vbNullString
        .RightFooter = vbNullString
        With .CenterFooterPicture
            .FileName = vbNullString
        End With
    End With

    ' Remove any Page Breaks
    ' ----------------------
    With Sheet
        .ResetAllPageBreaks
        .DisplayPageBreaks = False
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
