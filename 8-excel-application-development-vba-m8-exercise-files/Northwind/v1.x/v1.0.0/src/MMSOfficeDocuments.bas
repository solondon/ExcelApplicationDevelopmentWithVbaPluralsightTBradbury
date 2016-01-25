Attribute VB_Name = "MMSOfficeDocuments"
' ==========================================================================
' Module      : MMSOfficeDocuments
' Type        : Module
' Description : Support for working with Office documents
' --------------------------------------------------------------------------
' Procedures  : GetDocument                         Object
'               GetDocumentName                     String
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

Private Const msMODULE  As String = "MMSOfficeDocuments"

Public Function GetDocument(Optional ByVal ThisInsteadOfActive _
                                        As Boolean = True) As Object
' ==========================================================================
' Description : Return the document for the application
'
' Parameters  : ThisInsteadOfActive     If the application is
'                                       Microsoft Excel, this process will
'                                       Use ThisWorkbook by default instead
'                                       of ActiveWorkbook.
'
' Returns     : Object                  A document object
' ==========================================================================

    Const sPROC As String = "GetDocument"

    Dim objApp  As Object
    Dim objRtn  As Object


    On Error GoTo PROC_ERR
    ' ----------------------------------------------------------------------

    Set objApp = Application

    Select Case objApp.Name
    Case gsOFFICE_APPNAME_EXCEL
        If ThisInsteadOfActive Then
            Set objRtn = objApp.ThisWorkbook
        Else
            Set objRtn = objApp.ActiveWorkbook
        End If

    Case gsOFFICE_APPNAME_WORD
        Set objRtn = objApp.ActiveDocument

    Case gsOFFICE_APPNAME_POWERPOINT
        Set objRtn = objApp.ActivePresentation
    
    Case gsOFFICE_APPNAME_PROJECT
        Set objRtn = objApp.ActiveProject
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set GetDocument = objRtn

    Set objRtn = Nothing
    Set objApp = Nothing

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

Public Function GetDocumentName(Optional ByVal FullName As Boolean, _
                                Optional ByVal ThisInsteadOfActive _
                                            As Boolean = True) _
       As String
' ==========================================================================
' Description : Returns the file name for the active document
'
' Parameters  : FullName                Indicates if the full name
'                                       (with path) should be returned
'               ThisInsteadOfActive     If running Excel, use ThisWorkbook
'                                       instead of ActiveWorkbook
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "GetDocumentName"

    Dim sRtn    As String

    Dim objDoc  As Object


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Get the document
    ' ----------------
    Set objDoc = GetDocument(ThisInsteadOfActive)

    If FullName Then
        sRtn = objDoc.FullName
    Else
        sRtn = objDoc.Name
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetDocumentName = sRtn

    Set objDoc = Nothing

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
