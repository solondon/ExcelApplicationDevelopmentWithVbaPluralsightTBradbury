Attribute VB_Name = "MVBAFile"
' ==========================================================================
' Module      : MVBAFile
' Type        : Module
' Description : Wrapper functions for the FileSystemObject
'               and methods for working with file names
' --------------------------------------------------------------------------
' Procedures  : FileDelete              Boolean
'               FileExists              Boolean
'               FileNameAppend          String
'               FileNameValidates       Boolean
'               FileToCollection        Long
'               FolderExists            Boolean
'               ParsePath               String
' --------------------------------------------------------------------------
' Dependencies: MVBADebug
'               MVBAError
' --------------------------------------------------------------------------
' References  : Microsoft Scripting Runtime
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

Private Const msMODULE As String = "MVBAFile"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuParsePathPart
    pppFullPath
    pppDriveOnly
    pppPathOnly
    pppFullName
    pppFileOnly
    pppFileOnlyNoExt
    pppFileExtOnly
    pppUNCServerOnly
    pppUNCShareOnly
    pppUNCPathOnly
    pppUNCFullPath
End Enum

Public Function FileDelete(ByVal FileName As String, _
                  Optional ByVal Force As Boolean) As Boolean
' ==========================================================================
' Description : Delete a disk file
'
' Parameters  : FileName    The qualified name of the file to delete.
'               Force       Force deletes for read-only files.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "FileDelete"

    Dim bRtn    As Boolean
    Dim sPath   As String
    Dim sFile   As String

    Dim FSO     As Scripting.FileSystemObject


    On Error GoTo PROC_ERR

    If ((Not BitIsSet(glTRACE_OUTPUT, toLogFile)) _
        Or (FileName <> GetLogFileName(lftTrace))) Then
        Call Trace(tlMaximum, msMODULE, sPROC, FileName)
    End If

    ' ----------------------------------------------------------------------

    Set FSO = New Scripting.FileSystemObject

    sPath = ParsePath(FileName, pppFullPath)
    sFile = ParsePath(FileName, pppFileOnly)

    If (Len(Trim(sPath)) = 0) Then
        If (CBool(Len(ThisWorkbook.Path))) Then
            sPath = ThisWorkbook.Path & "\"
        Else
            sPath = CurDir & "\"
        End If
    End If

    sFile = sPath & sFile

    Call FSO.DeleteFile(sFile, Force)

    bRtn = (Not FileExists(sFile))

    ' ----------------------------------------------------------------------

PROC_EXIT:

    FileDelete = bRtn

    If ((Not BitIsSet(glTRACE_OUTPUT, toLogFile)) _
        Or (FileName <> GetLogFileName(lftTrace))) Then
        Call Trace(tlMaximum, msMODULE, sPROC, bRtn)
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

Public Function FileExists(ByVal FileName As String) As Boolean
' ==========================================================================
' Description : Determine if a specified file exists
'
' Parameters  : FileName    The name of the file to test for
'
' Returns     : Boolean
'
' Notes       : This routine does not use the ParsePath function
'               to prevent a stack fault when tracing is logged.
' ==========================================================================

    Const sPROC As String = "FileExists"

    Dim bRtn    As Boolean
    Dim lPos    As Long
    Dim sPath   As String
    Dim sFile   As String

    Dim FSO     As Scripting.FileSystemObject


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Can trace if sending ONLY to Immediate Window
    ' ---------------------------------------------
    If (glTRACE_OUTPUT = enuTraceOutput.toImmediate) Then
        Call Trace(tlMaximum, msMODULE, sPROC, FileName)
    End If

    Set FSO = New Scripting.FileSystemObject

    ' Parse it out if there is a path
    ' -------------------------------
    lPos = InStrRev(FileName, "\")
    If (lPos > 0) Then
        sPath = Left$(FileName, lPos)
        sFile = Mid$(FileName, lPos + 1)
    Else
        sFile = FileName
    End If

    ' Look in the workbook folder
    ' then in the current folder if needed
    ' ------------------------------------
    If (Len(Trim(sPath)) = 0) Then
        If (CBool(Len(ThisWorkbook.Path))) Then
            sPath = ThisWorkbook.Path & "\"
        Else
            sPath = CurDir & "\"
        End If
    End If

    bRtn = FSO.FileExists(sPath & sFile)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    FileExists = bRtn

    Set FSO = Nothing

    On Error GoTo 0

    If (glTRACE_OUTPUT = enuTraceOutput.toImmediate) Then
        Call Trace(tlMaximum, msMODULE, sPROC, bRtn)
    End If

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

Public Function FileNameAppend(ByVal FileName As String, _
                               ByVal Text As String) As String
' ==========================================================================
' Description : Append text to the end of a filename before the extension.
'
' Parameters  : FileName    The name of the file
'               Text        The text to append
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "FileNameAppend"

    Dim sRtn    As String
    Dim sBase   As String
    Dim sFile   As String
    Dim sPath   As String
    Dim sExt    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, FileName)

    ' ----------------------------------------------------------------------
    ' Split the path, if present
    ' --------------------------

    sPath = ParsePath(FileName, pppFullPath)
    sBase = ParsePath(FileName, pppFileOnlyNoExt)
    sExt = ParsePath(FileName, pppFileExtOnly)

    ' Build the name
    ' --------------
    If (Len(sExt) = 0) Then
        sRtn = sBase & Text
    Else
        sRtn = sBase & Text & "." & sExt
    End If

    ' Add the path if it existed
    ' --------------------------
    If (Len(sPath) > 0) Then
        sRtn = sPath & sRtn
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    FileNameAppend = sRtn

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

Public Function FileNameValidates(ByRef FileName As String, _
                                  ByVal DefaultName As String) As Boolean
' ==========================================================================
' Description : Ensures a valid filename and that it exists.
'
' Parameters  : FileName    The file name being checked
'
'               DefaultName If no file name is passed, use the default.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "FileNameValidates"

    Dim bRtn    As Boolean

    Dim sFile   As String
    Dim sPath   As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, FileName)

    ' ----------------------------------------------------------------------
    ' Assume success
    ' --------------
    bRtn = True

    ' Separate the name and path
    ' --------------------------
    sPath = ParsePath(FileName, pppFullPath)
    sFile = ParsePath(FileName, pppFileOnly)

    ' If a path was not passed
    ' then use the default path
    ' -------------------------
    If Len(Trim(sPath)) = 0 Then
        sPath = ThisWorkbook.Path & "\"
    End If

    ' If the FileName was not passed
    ' then use the default FileName
    ' ------------------------------
    If Len(Trim(sFile)) = 0 Then
        sFile = DefaultName
    End If

    ' Build the real filename
    ' -----------------------
    FileName = sPath & sFile

    ' Make sure it exists
    ' -------------------
    bRtn = FileExists(FileName)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    FileNameValidates = bRtn

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

Public Function FileToCollection(ByVal FileName As String, _
                                 ByRef Col As VBA.Collection, _
                        Optional ByVal SkipEmptyLines As Boolean) As Long
' ==========================================================================
' Description : Loads a text file line-by-line into a collection
'
' Parameters  : FileName    The name of the file to import
'               Col         The collection to load the line into
'
' Returns     : Long        The number of lines loaded
' ==========================================================================

    Const sPROC     As String = "FileToCollection"

    Dim lRtn        As Long

    Dim iFileNum    As Integer
    Dim lIdx        As Long
    Dim sLine       As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Throw an exception if not found
    ' -------------------------------
    If (Not FileExists(FileName)) Then
        Call Err.Raise(ERR_FILE_NOT_FOUND, Concat(".", msMODULE, sPROC))
        GoTo PROC_EXIT
    End If

    ' First clear the collection
    ' --------------------------
    Do While Col.Count > 0
        Call Col.Remove(1)
    Loop

    ' Read the file into the collection
    ' ---------------------------------
    iFileNum = FreeFile()

    Open FileName For Input Access Read As #iFileNum

    Do While Not EOF(iFileNum)
        Line Input #iFileNum, sLine
        If (Not (SkipEmptyLines And (Trim$(sLine) = vbNullString))) Then
            Call Col.Add(sLine)
        End If
    Loop

    Close #iFileNum

    lRtn = Col.Count

    ' ----------------------------------------------------------------------

PROC_EXIT:

    FileToCollection = lRtn

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

Public Function FolderExists(ByVal Path As String) As Boolean
' ==========================================================================
' Description : Determine if a specified folder exists
'
' Parameters  : Path    The fully-qualified folder path
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "FolderExists"

    Dim bRtn    As Boolean
    Dim FSO     As Scripting.FileSystemObject


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Path)


    Set FSO = New Scripting.FileSystemObject

    bRtn = FSO.FolderExists(Path)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    FolderExists = bRtn

    Set FSO = Nothing

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

Public Function ParsePath(ByRef Path As String, _
                          ByVal PathPart As enuParsePathPart) As String
' ==========================================================================
' Description : Parse out a piece of a file name or path.
'
' Parameters  : Path        The path or file name to parse.
'               PathPart    Identifies to part to return.
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "ParsePath"

    Dim bIncludesDrive As Boolean
    Dim bIncludesPath As Boolean
    Dim bIsUNC  As Boolean

    Dim lLen    As Long
    Dim lPos    As Long

    Dim sRtn    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Path)

    ' ----------------------------------------------------------------------
    ' Quit if the entire
    ' part is requested
    ' ------------------

    If (PathPart = pppFullName) Then
        sRtn = Path
        GoTo PROC_EXIT
    End If

    ' Get the length of the path
    ' --------------------------
    lPos = InStrRev(Path, "\")

    ' Set flags
    ' ---------
    bIncludesPath = (lPos > 0)
    bIsUNC = (Left$(Path, 2) = "\\")
    bIncludesDrive = ((Not bIsUNC) And (Mid$(Path, 2, 1) = ":"))

    Select Case PathPart

    Case pppFullPath
        sRtn = Left$(Path, lPos)

    Case pppDriveOnly
        If bIncludesDrive Then
            sRtn = Left$(Path, 3)
        End If

    Case pppPathOnly
        sRtn = Left$(Path, lPos)
        If bIsUNC Then
            lPos = InStr(4, sRtn, "\")
            sRtn = Mid$(sRtn, lPos + 1)
            lPos = InStr(1, sRtn, "\")
            sRtn = Mid$(sRtn, lPos)
        Else
            lPos = InStr(1, Path, "\")
            sRtn = Mid$(sRtn, lPos)
        End If

    Case pppFullName
        sRtn = Path

    Case pppFileOnly
        If bIncludesPath Then
            sRtn = Mid$(Path, lPos + 1)
        Else
            sRtn = Path
        End If

    Case pppFileOnlyNoExt
        If bIncludesPath Then
            sRtn = Mid$(Path, lPos + 1)
            lPos = InStrRev(sRtn, ".")
            If CBool(lPos) Then
                sRtn = Left$(sRtn, lPos - 1)
            End If
        Else
            lPos = InStr(Path, ".")
            If CBool(lPos) Then
                sRtn = Left$(Path, lPos - 1)
            End If
        End If

    Case pppFileExtOnly
        If bIncludesPath Then
            sRtn = Mid$(Path, lPos + 1)
            lPos = InStrRev(sRtn, ".")
            sRtn = Mid$(sRtn, lPos + 1)
        Else
            lPos = InStrRev(Path, ".")
            sRtn = Mid$(Path, lPos + 1)
        End If

    Case pppUNCServerOnly
        If bIsUNC Then
            lPos = InStr(4, Path, "\")
            sRtn = Mid$(Path, 3, lPos - 3)  ' Don't count \\ and \
        End If

    Case pppUNCShareOnly
        If bIsUNC Then
            sRtn = Left$(Path, lPos - 1)
            lPos = InStr(4, Path, "\")
            sRtn = Mid$(sRtn, lPos + 1)
            lLen = InStr(1, sRtn, "\") - 1
            sRtn = Left$(sRtn, lLen)
        End If

    Case pppUNCPathOnly
        If bIsUNC Then
            sRtn = Left$(Path, lPos)
            lPos = InStr(4, sRtn, "\")
            sRtn = Mid$(sRtn, lPos + 1)
            lPos = InStr(1, sRtn, "\")
            sRtn = Mid$(sRtn, lPos)
        End If

    Case pppUNCFullPath
        If bIsUNC Then
            sRtn = Left$(Path, lPos)
        End If

    Case Else
        Stop
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ParsePath = sRtn

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
