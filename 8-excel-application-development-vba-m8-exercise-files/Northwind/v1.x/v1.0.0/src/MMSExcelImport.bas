Attribute VB_Name = "MMSExcelImport"
' ==========================================================================
' Module      : MMSExcelImport
' Type        : Module
' Description : Import functions
' ------------------------------------------------------------------------
' Procedures  : ImportTextFile
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

Private Const msMODULE As String = "MMSExcelImport"

Public Function ImportTextFile(ByRef Sheet As Excel.Worksheet, _
                               ByVal FileName As String, _
                      Optional ByVal Delimiter As String = vbTab, _
                      Optional ByVal Destination As String = "$A$1", _
                      Optional ByVal StartRow As Long = 1, _
                      Optional ByVal HasHeaders As Boolean = False, _
                      Optional ByVal Formats As Variant, _
                      Optional ByVal ColumnWidths As Variant, _
                      Optional ByVal TextQualifier _
                                  As XlTextQualifier _
                                   = xlTextQualifierNone, _
                      Optional ByVal RefreshStyle _
                                  As XlCellInsertionMode _
                                   = xlInsertDeleteCells, _
                      Optional ByVal DeleteAutoName _
                                  As Boolean = True) As Boolean
' ==========================================================================
' Description : Load a text file into a worksheet
'
' Parameters  : Sheet         The worksheet to load the data into
'               FileName      The name of the file to load
'               Delimiter     A string character to separate columns
'               Destination   The starting cell address for the data
'               StartRow      The starting row of data in the text file
'               HasHeaders    Indicates if the text file has headers
'               Formats       An array of variants indicating the cell formats
'               ColumnWidths  A variant array containing the column widths
'               TextQualifier The text qualifier specifies that
'                             the data enclosed within the qualifier
'                             is in text format
'               RefreshStyle  Returns or sets the way rows on the specified
'                             worksheet are added or deleted to accommodate
'                             the number of rows in a recordset
'                             returned by a query.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC         As String = "ImportTextFile"
    Const sTAB          As String = vbTab
    Const sCOMMA        As String = ","
    Const sSEMICOLON    As String = ";"
    Const sSPACE        As String = " "

    Dim bRtn            As Boolean

    Dim sConn           As String
    Dim sName           As String
    Dim sPath           As String
    Dim sAutoName       As String

    Dim QT              As Excel.QueryTable

    Dim udtAppProps     As TApplicationProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, FileName)

    ' ----------------------------------------------------------------------
    ' Save application state
    ' ----------------------
    Call GetApplicationProperties(udtAppProps)
    Application.ScreenUpdating = gbDEBUG_MODE

    ' Assume success
    ' --------------
    bRtn = True

    ' Initialize locals
    ' -----------------
    sPath = ParsePath(FileName, pppFullPath)
    sName = ParsePath(FileName, pppFileOnly)

    sConn = "TEXT;" & FileName

    Set QT = Sheet.QueryTables.Add(Connection:=sConn, _
                                   Destination:=Sheet.Range(Destination))
    With QT
        ' Set general properties
        ' ----------------------
        .Name = sName
        .FieldNames = HasHeaders
        .RefreshStyle = RefreshStyle
        .TextFileConsecutiveDelimiter = False
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlWindows
        .TextFileStartRow = StartRow
        .TextFileTextQualifier = TextQualifier
        .TextFileTrailingMinusNumbers = True

        ' Use column formats if provided
        ' ------------------------------
        If Not IsMissing(Formats) Then
            .TextFileColumnDataTypes = Formats
        End If

        ' Settings for fixed-width file
        ' -----------------------------
        If Not IsMissing(ColumnWidths) Then
            .TextFileParseType = xlFixedWidth
            .TextFileFixedColumnWidths = ColumnWidths

        ' Settings for delimited file
        ' ---------------------------
        Else
            .TextFileParseType = xlDelimited

            Select Case Delimiter
            Case sTAB
                .TextFileTabDelimiter = True

            Case sCOMMA
                .TextFileCommaDelimiter = True

            Case sSEMICOLON
                .TextFileSemicolonDelimiter = True

            Case sSPACE
                .TextFileSpaceDelimiter = True

            Case Else
                .TextFileOtherDelimiter = Delimiter
            End Select

        End If  ' Fixed or delimited
    End With    ' QT properties

    ' Load the file
    ' -------------
    Call Trace(tlVerbose, msMODULE, sPROC, FileName)
    QT.Refresh BackgroundQuery:=False

    ' Optionally delete the Name that was automatically created
    ' ---------------------------------------------------------
    If DeleteAutoName Then

        ' Needs to be wrapped with a single quote
        ' if there are any spaces in the sheet name
        ' -----------------------------------------
        If InStr(1, Sheet.Name, " ", vbTextCompare) > 0 Then
            sAutoName = "'" & Sheet.Name & "'!" & QT.Name
        Else
            sAutoName = Sheet.Name & "!" & QT.Name
        End If

        ' Don't give a warning if the name doesn't exist
        ' ----------------------------------------------
        On Error Resume Next

        ThisWorkbook.Names(sAutoName).Delete
    End If
    ' ----------------------------------------------------------------------

PROC_EXIT:

    ImportTextFile = bRtn

    QT.Delete
    Set QT = Nothing

    ' Restore the application state
    ' -----------------------------
    Call SetApplicationProperties(udtAppProps)

    Call Trace(tlMaximum, msMODULE, sPROC, bRtn)
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
