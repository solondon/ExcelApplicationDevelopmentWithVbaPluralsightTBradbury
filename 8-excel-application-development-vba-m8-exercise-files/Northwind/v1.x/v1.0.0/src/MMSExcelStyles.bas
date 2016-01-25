Attribute VB_Name = "MMSExcelStyles"
' ==========================================================================
' Module      : MMSExcelStyles
' Type        : Module
' Description : Support for Excel styles
' --------------------------------------------------------------------------
' Procedures  : CreateTable
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

Public Const gsEXCEL_STYLE_NORMAL   As String = "Normal"
Public Const gsEXCEL_STYLE_BAD      As String = "Bad"
Public Const gsEXCEL_STYLE_GOOD     As String = "Good"
Public Const gsEXCEL_STYLE_NEUTRAL  As String = "Neutral"

Public Const gsEXCEL_STYLE_TABLT01  As String = "TableStyleLight1"
Public Const gsEXCEL_STYLE_TABLT02  As String = "TableStyleLight2"
Public Const gsEXCEL_STYLE_TABLT03  As String = "TableStyleLight3"
Public Const gsEXCEL_STYLE_TABLT04  As String = "TableStyleLight4"
Public Const gsEXCEL_STYLE_TABLT05  As String = "TableStyleLight5"
Public Const gsEXCEL_STYLE_TABLT06  As String = "TableStyleLight6"
Public Const gsEXCEL_STYLE_TABLT07  As String = "TableStyleLight7"
Public Const gsEXCEL_STYLE_TABLT08  As String = "TableStyleLight8"
Public Const gsEXCEL_STYLE_TABLT09  As String = "TableStyleLight9"
Public Const gsEXCEL_STYLE_TABLT10  As String = "TableStyleLight10"
Public Const gsEXCEL_STYLE_TABLT11  As String = "TableStyleLight11"
Public Const gsEXCEL_STYLE_TABLT12  As String = "TableStyleLight12"
Public Const gsEXCEL_STYLE_TABLT13  As String = "TableStyleLight13"
Public Const gsEXCEL_STYLE_TABLT14  As String = "TableStyleLight14"
Public Const gsEXCEL_STYLE_TABLT15  As String = "TableStyleLight15"
Public Const gsEXCEL_STYLE_TABLT16  As String = "TableStyleLight16"
Public Const gsEXCEL_STYLE_TABLT17  As String = "TableStyleLight17"
Public Const gsEXCEL_STYLE_TABLT18  As String = "TableStyleLight18"
Public Const gsEXCEL_STYLE_TABLT19  As String = "TableStyleLight19"
Public Const gsEXCEL_STYLE_TABLT20  As String = "TableStyleLight20"
Public Const gsEXCEL_STYLE_TABLT21  As String = "TableStyleLight21"

Public Const gsEXCEL_STYLE_TABMD01  As String = "TableStyleMedium1"
Public Const gsEXCEL_STYLE_TABMD02  As String = "TableStyleMedium2"
Public Const gsEXCEL_STYLE_TABMD03  As String = "TableStyleMedium3"
Public Const gsEXCEL_STYLE_TABMD04  As String = "TableStyleMedium4"
Public Const gsEXCEL_STYLE_TABMD05  As String = "TableStyleMedium5"
Public Const gsEXCEL_STYLE_TABMD06  As String = "TableStyleMedium6"
Public Const gsEXCEL_STYLE_TABMD07  As String = "TableStyleMedium7"
Public Const gsEXCEL_STYLE_TABMD08  As String = "TableStyleMedium8"
Public Const gsEXCEL_STYLE_TABMD09  As String = "TableStyleMedium9"
Public Const gsEXCEL_STYLE_TABMD10  As String = "TableStyleMedium10"
Public Const gsEXCEL_STYLE_TABMD11  As String = "TableStyleMedium11"
Public Const gsEXCEL_STYLE_TABMD12  As String = "TableStyleMedium12"
Public Const gsEXCEL_STYLE_TABMD13  As String = "TableStyleMedium13"
Public Const gsEXCEL_STYLE_TABMD14  As String = "TableStyleMedium14"
Public Const gsEXCEL_STYLE_TABMD15  As String = "TableStyleMedium15"
Public Const gsEXCEL_STYLE_TABMD16  As String = "TableStyleMedium16"
Public Const gsEXCEL_STYLE_TABMD17  As String = "TableStyleMedium17"
Public Const gsEXCEL_STYLE_TABMD18  As String = "TableStyleMedium18"
Public Const gsEXCEL_STYLE_TABMD19  As String = "TableStyleMedium19"
Public Const gsEXCEL_STYLE_TABMD20  As String = "TableStyleMedium20"
Public Const gsEXCEL_STYLE_TABMD21  As String = "TableStyleMedium21"
Public Const gsEXCEL_STYLE_TABMD22  As String = "TableStyleMedium22"
Public Const gsEXCEL_STYLE_TABMD23  As String = "TableStyleMedium23"
Public Const gsEXCEL_STYLE_TABMD24  As String = "TableStyleMedium24"
Public Const gsEXCEL_STYLE_TABMD25  As String = "TableStyleMedium25"
Public Const gsEXCEL_STYLE_TABMD26  As String = "TableStyleMedium26"
Public Const gsEXCEL_STYLE_TABMD27  As String = "TableStyleMedium27"
Public Const gsEXCEL_STYLE_TABMD28  As String = "TableStyleMedium28"

Public Const gsEXCEL_STYLE_TABDK01  As String = "TableStyleDark1"
Public Const gsEXCEL_STYLE_TABDK02  As String = "TableStyleDark2"
Public Const gsEXCEL_STYLE_TABDK03  As String = "TableStyleDark3"
Public Const gsEXCEL_STYLE_TABDK04  As String = "TableStyleDark4"
Public Const gsEXCEL_STYLE_TABDK05  As String = "TableStyleDark5"
Public Const gsEXCEL_STYLE_TABDK06  As String = "TableStyleDark6"
Public Const gsEXCEL_STYLE_TABDK07  As String = "TableStyleDark7"
Public Const gsEXCEL_STYLE_TABDK08  As String = "TableStyleDark8"
Public Const gsEXCEL_STYLE_TABDK09  As String = "TableStyleDark9"
Public Const gsEXCEL_STYLE_TABDK10  As String = "TableStyleDark10"
Public Const gsEXCEL_STYLE_TABDK11  As String = "TableStyleDark11"

Public Const gsDEFAULT_TABLE_STYLE  As String = gsEXCEL_STYLE_TABMD02

' ----------------
' Module Level
' ----------------

Private Const msMODULE              As String = "MMSExcelStyles"

Public Sub CreateTable(ByRef Sheet As Excel.Worksheet, _
              Optional ByVal TableName As String, _
              Optional ByVal TableStyle As String = gsDEFAULT_TABLE_STYLE, _
              Optional ByVal AutoRevert As Boolean = True, _
              Optional ByVal RowStripeSize1 As Long)
' ==========================================================================
' Description : Convert the UsedRange on a data sheet to a table
'               so the user can select different color themes.
'
' Parameters  : Sheet       The worksheet to add the table to
'               TableName   The name of the new table
'               TableStyle  The name of the style to apply
'               AutoRevert  If True, automatically revert to a range
' ==========================================================================

    Const sPROC As String = "CreateTable"
    Const sTSNM As String = "_TEMPSTYLE"

    Dim bDelTS  As Boolean

    Dim oLO     As Excel.ListObject
    Dim oTS     As Excel.TableStyle


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Sheet.CodeName)

    ' ----------------------------------------------------------------------
    ' Create a copy of the TableStyle so changes can be made
    ' ------------------------------------------------------
    If (RowStripeSize1 > 0) Then

        Set oTS = ThisWorkbook.TableStyles(TableStyle).Duplicate(sTSNM)
        bDelTS = True

        ' Force AutoRevert if the StripeSize is different
        ' -----------------------------------------------
        With oTS.TableStyleElements(xlRowStripe1)
            If (.StripeSize <> RowStripeSize1) Then
                AutoRevert = True
            End If
        End With

        oTS.TableStyleElements(xlRowStripe1).StripeSize = RowStripeSize1
        oTS.TableStyleElements(xlRowStripe2).StripeSize = RowStripeSize1

    Else
        Set oTS = ThisWorkbook.TableStyles(TableStyle)
    End If

    ' Use the default table name
    ' --------------------------
    If (TableName = vbNullString) Then
        TableName = Sheet.Name
    End If

    With Sheet

        ' Make sure there isn't already an
        ' existing table with the same name
        ' ---------------------------------
        On Error Resume Next
        .ListObjects(TableName).Delete

        ' Restart error handling
        ' ----------------------
        On Error GoTo PROC_ERR

        ' Remove previous formatting
        ' --------------------------
        .UsedRange.Cells.Style = gsEXCEL_STYLE_NORMAL

        ' Add the table
        ' -------------
        Set oLO = .ListObjects.Add(xlSrcRange, Sheet.UsedRange, , xlYes)

        With oLO
            .Name = TableName
            .ShowAutoFilter = False

            .TableStyle = vbNullString
            .TableStyle = oTS

            If AutoRevert Then
                .Unlist
            End If
        End With

        .UsedRange.Rows(1).HorizontalAlignment = xlCenter

    End With

    ' Delete the temporary style
    ' --------------------------
    If bDelTS Then
        ThisWorkbook.TableStyles(sTSNM).Delete
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set oLO = Nothing
    Set oTS = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, TableName)
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
