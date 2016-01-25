Attribute VB_Name = "MWinAPIShell32"
' ==========================================================================
' Module      : MWinAPIShell32
' Type        : Module
' Description : Support for working with the Windows Shell
' --------------------------------------------------------------------------
' Procedures  : ShellBrowseForFolder        String
'               ShellExecute                Boolean
' --------------------------------------------------------------------------
' Dependencies: MWinAPI
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

' ----------------
' Module Level
' ----------------

Private Const msMODULE                          As String = "MWinAPIShell32"

' SHBrowseForFolder messages are described on MSDN at
' http://msdn.microsoft.com/en-us/library/bb762598(v=vs.85).aspx
' -----------------------------------
' Message from browser
' -----------------------------------
Private Const BFFM_INITIALIZED                  As Long = 1
Private Const BFFM_SELCHANGED                   As Long = 2
Private Const BFFM_VALIDATEFAILEDA              As Long = 3    ' lParam:szPath ret:1(cont),0(EndDialog)
Private Const BFFM_VALIDATEFAILEDW              As Long = 4    ' lParam:wzPath ret:1(cont),0(EndDialog)
Private Const BFFM_IUNKNOWN                     As Long = 5    ' provides IUnknown to client. lParam: IUnknown*

' Message to browser
' -----------------------------------
Private Const BFFM_SETSTATUSTEXTA               As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK                     As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA                As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW                As Long = (WM_USER + 103)
Private Const BFFM_SETSTATUSTEXTW               As Long = (WM_USER + 104)
Private Const BFFM_SETOKTEXT                    As Long = (WM_USER + 105)   ' Unicode only
Private Const BFFM_SETEXPANDED                  As Long = (WM_USER + 106)   ' Unicode only

' BROWSEINFO Flags are described on MSDN at
' http://msdn.microsoft.com/en-us/library/bb773205(v=vs.85).aspx
' -----------------------------------
Private Const BIF_RETURNONLYFSDIRS              As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN             As Long = &H2
Private Const BIF_STATUSTEXT                    As Long = &H4
Private Const BIF_RETURNFSANCESTORS             As Long = &H8
Private Const BIF_EDITBOX                       As Long = &H10
Private Const BIF_VALIDATE                      As Long = &H20
Private Const BIF_NEWDIALOGSTYLE                As Long = &H40
Private Const BIF_BROWSEINCLUDEURLS             As Long = &H80
Private Const BIF_USENEWUI                      As Long = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
Private Const BIF_UAHINT                        As Long = &H100
Private Const BIF_NONEWFOLDERBUTTON             As Long = &H200
Private Const BIF_NOTRANSLATETARGETS            As Long = &H400
Private Const BIF_BROWSEFORCOMPUTER             As Long = &H1000
Private Const BIF_BROWSEFORPRINTER              As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES            As Long = &H4000
Private Const BIF_SHAREABLE                     As Long = &H8000
Private Const BIF_BROWSEFILEJUNCTIONS           As Long = &H10000

' Local Memory Flags are described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366723(v=vs.85).aspx
' -----------------------------------
Private Const LMEM_FIXED                        As Long = &H0
Private Const LMEM_MOVEABLE                     As Long = &H2
Private Const LMEM_NOCOMPACT                    As Long = &H10
Private Const LMEM_NODISCARD                    As Long = &H20
Private Const LMEM_ZEROINIT                     As Long = &H40
Private Const LMEM_MODIFY                       As Long = &H80
Private Const LMEM_DISCARDABLE                  As Long = &HF00
Private Const LMEM_VALID_FLAGS                  As Long = &HF72
Private Const LMEM_INVALID_HANDLE               As Long = &H8000

Private Const LHND                              As Long = (LMEM_MOVEABLE Or LMEM_ZEROINIT)
Private Const LPTR                              As Long = (LMEM_FIXED Or LMEM_ZEROINIT)

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

' ----------------
' Module Level
' ----------------

Private Enum enuBrowseInfoFlags
    bifReturnOnlyFSDirs = BIF_RETURNONLYFSDIRS
    bifDontGoBelowDomain = BIF_DONTGOBELOWDOMAIN
    bifStatusText = BIF_STATUSTEXT
    bifReturnFSAncestors = BIF_RETURNFSANCESTORS
    bifEditBox = BIF_EDITBOX
    bifValidate = BIF_VALIDATE
    bifNewDialogStyle = BIF_NEWDIALOGSTYLE
    bifBrowseIncludeURLs = BIF_BROWSEINCLUDEURLS
    bifUseNewUI = BIF_USENEWUI
    bifUAHInt = BIF_UAHINT
    bifNoNewFolderButton = BIF_NONEWFOLDERBUTTON
    bifNoTranslateTargets = BIF_NOTRANSLATETARGETS
    bifBrowseForComputer = BIF_BROWSEFORCOMPUTER
    bifBrowseForPrinter = BIF_BROWSEFORPRINTER
    bifBrowseIncludeFiles = BIF_BROWSEINCLUDEFILES
    bifShareable = BIF_SHAREABLE
    bifBrowseFileJunctions = BIF_BROWSEFILEJUNCTIONS
End Enum

' -----------------------------------
' Type declarations
' -----------------------------------
' Global Level
' ----------------

' The BROWSEINFO structure is described on MSDN at
' http://msdn.microsoft.com/en-us/library/bb773205(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
Public Type TBROWSEINFO
    hWndOwner                                   As LongPtr
    pidlRoot                                    As Long
    pszDisplayName                              As String
    lpszTitle                                   As String
    ulFlags                                     As enuBrowseInfoFlags
    lpfn                                        As LongPtr
    lParam                                      As LongPtr
    iImage                                      As Long
End Type
#Else
Public Type TBROWSEINFO
    hWndOwner                                   As Long
    pidlRoot                                    As Long
    pszDisplayName                              As String
    lpszTitle                                   As String
    ulFlags                                     As enuBrowseInfoFlags
    lpfn                                        As Long
    lParam                                      As Long
    iImage                                      As Long
End Type
#End If

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The CopyMemory function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366535(VS.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Sub CopyMemory _
            Lib "Kernel32" _
            Alias "RtlMoveMemory" (ByRef pDest As Any, _
                                   ByRef pSource As Any, _
                                   ByVal dwLength As Long)
#Else
    Private Declare _
            Sub CopyMemory _
            Lib "Kernel32" _
            Alias "RtlMoveMemory" (ByRef pDest As Any, _
                                   ByRef pSource As Any, _
                                   ByVal dwLength As Long)
#End If

' The CoTaskMemFree function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms680722(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Sub CoTaskMemFree _
            Lib "ole32" (ByVal pv As LongPtr)
#Else
    Private Declare _
            Sub CoTaskMemFree _
            Lib "ole32" (ByVal pv As Long)
#End If

' The LocalAlloc function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366723(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function LocalAlloc _
            Lib "Kernel32" (ByVal uFlags As Long, _
                            ByVal uBytes As Long) _
            As LongPtr
#Else
    Private Declare _
            Function LocalAlloc _
            Lib "Kernel32" (ByVal uFlags As Long, _
                            ByVal uBytes As Long) _
            As Long
#End If


' The LocalFree function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366730(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function LocalFree _
            Lib "Kernel32" (ByVal hMem As LongPtr) _
            As LongPtr
#Else
    Private Declare _
            Function LocalFree _
            Lib "Kernel32" (ByVal hMem As Long) _
            As Long
#End If

' The GetPIDLFromPath function is undocumented.
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetPIDLFromPath _
            Lib "shell32" _
            Alias "#162" (ByVal szPath As String) _
            As Long
#Else
    Private Declare _
            Function GetPIDLFromPath _
            Lib "shell32" _
            Alias "#162" (ByVal szPath As String) _
            As Long
#End If

' The SHBrowseForFolder function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/bb762115(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SHBrowseForFolder _
            Lib "shell32" _
            Alias "SHBrowseForFolderA" (ByRef lpBrowseInfo As TBROWSEINFO) _
            As Long
#Else
    Private Declare _
            Function SHBrowseForFolder _
            Lib "shell32" _
            Alias "SHBrowseForFolderA" (ByRef lpBrowseInfo As TBROWSEINFO) _
            As Long
#End If

' The SHGetPathFromIDList function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/bb762194(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SHGetPathFromIDList _
            Lib "shell32" _
            Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
                                          ByVal pszPath As String) _
            As Long
#Else
    Private Declare _
            Function SHGetPathFromIDList _
            Lib "shell32" _
            Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
                                          ByVal pszPath As String) _
            As Long
#End If

' The ShellExecute function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/bb762153(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function ShellExecuteANSI _
            Lib "shell32" _
            Alias "ShellExecuteA" (ByVal hWnd As LongPtr, _
                                   ByVal lpOperation As String, _
                                   ByVal lpFile As String, _
                                   ByVal lpParameters As String, _
                                   ByVal lpDirectory As String, _
                                   ByVal nShowCommand As enuShowWindowCommand) _
            As LongPtr
#Else
    Private Declare _
            Function ShellExecuteANSI _
            Lib "shell32" _
            Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                   ByVal lpOperation As String, _
                                   ByVal lpFile As String, _
                                   ByVal lpParameters As String, _
                                   ByVal lpDirectory As String, _
                                   ByVal nShowCommand As enuShowWindowCommand) _
            As Long
#End If

' The SendMessage function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms644950(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SendMessage _
            Lib "User32" _
            Alias "SendMessageA" (ByVal hWnd As LongPtr, _
                                  ByVal Msg As Long, _
                                  ByVal wParam As Long, _
                                  ByRef lParam As Any) _
            As Long
#Else
    Private Declare _
            Function SendMessage _
            Lib "User32" _
            Alias "SendMessageA" (ByVal hWnd As Long, _
                                  ByVal Msg As Long, _
                                  ByVal wParam As Long, _
                                  ByRef lParam As Any) _
            As Long
#End If

#If VBA7 Then
Private Function BrowseCallbackProc(ByVal hWnd As LongPtr, _
                                    ByVal uMsg As Long, _
                                    ByVal lParam As Long, _
                                    ByVal lpData As Long) As Long
#Else
Private Function BrowseCallbackProc(ByVal hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal lParam As Long, _
                                    ByVal lpData As Long) As Long
#End If
' ==========================================================================
' Description : During initialization of the SHBrowseForFolder
'               function, this sets the pre-selected folder
'               using the pidl set as the udtBI.lParam,
'               and passed back to the callback as lpData param.
'
' Parameters  : hWnd    A handle to the window to receive the message
'               uMsg    The message to be sent
'               lParam  Additional message-specific information.
'
' Returns     : Long
'
' Comments    : This function is an implementation of a BFFCALLBACK
'             http://msdn.microsoft.com/en-us/library/bb762598(v=vs.85).aspx
' ==========================================================================

    Select Case uMsg
    Case BFFM_INITIALIZED
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, 0&, ByVal lpData)
    End Select

End Function

#If VBA7 Then
Private Function BrowseCallbackProcStr(ByVal hWnd As LongPtr, _
                                       ByVal uMsg As Long, _
                                       ByVal lParam As Long, _
                                       ByVal lpData As Long) As Long
#Else
Private Function BrowseCallbackProcStr(ByVal hWnd As Long, _
                                       ByVal uMsg As Long, _
                                       ByVal lParam As Long, _
                                       ByVal lpData As Long) As Long
#End If
' ==========================================================================
' Description : During initialization of the SHBrowseForFolder
'               function, this sets the pre-selected folder
'               using the pointer set as the udtBI.lParam,
'               and passed back to the callback as lpData param.
'
' Parameters  :
'
' Returns     : Long
'
' Comments    : This function is an implementation of a BFFCALLBACK
'             http://msdn.microsoft.com/en-us/library/bb762598(v=vs.85).aspx
' ==========================================================================

    Select Case uMsg
    Case BFFM_INITIALIZED
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, 1&, ByVal lpData)
    End Select

End Function

#If VBA7 Then
Private Function FARPROC(ByRef FuncPtr As LongPtr) As LongPtr
#Else
Private Function FARPROC(ByRef FuncPtr As Long) As Long
#End If
' ==========================================================================
' Description : FARPROC is a workaround for assigning the result of the
'               AddressOf operator to a member of a user-defined type.
'
' Parameters  : FuncPtr     A pointer to a function (AddressOf(Func))
'
' Returns     : LongPtr
' ==========================================================================

    FARPROC = FuncPtr

End Function

Public Function _
       ShellBrowseForFolder(Optional ByVal Caption As String, _
                            Optional ByVal StartFolder As String, _
                            Optional ByVal ShowNewFolderButton As Boolean) _
       As String
' ==========================================================================
' Description : Browse for folders
'
' Parameters  : Caption             The text to display
'               StartFolder         The starting location
'               ShowNewFolderButton Allow users to create new folders
'
' Returns     : String
' ==========================================================================

    Const sPROC         As String = "ShellBrowseForFolder"

    Dim bStartProvided  As Boolean: bStartProvided = Len(StartFolder)

    Dim lLen            As Long: lLen = Len(StartFolder) + 1
    Dim lRtn            As Long
    Dim lPIDL           As Long
    Dim lPos            As Long

    Dim lpPath          As Long

    Dim sRtn            As String: sRtn = Space(MAX_PATH)
    Dim udtBI           As TBROWSEINFO


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Initialize the BrowseInfo structure
    ' -----------------------------------

    With udtBI
        .hWndOwner = Application.hWnd
        .pidlRoot = 0&
        .lpszTitle = Caption
        .ulFlags = bifReturnOnlyFSDirs Or bifNewDialogStyle
        If (Not ShowNewFolderButton) Then
            .ulFlags = .ulFlags Or bifNoNewFolderButton
        End If

        If bStartProvided Then
            .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
            lpPath = LocalAlloc(LPTR, lLen)
            Call CopyMemory(ByVal lpPath, ByVal StartFolder, lLen)
            .lParam = lpPath
        End If
    End With

    ' Display the dialog
    ' ------------------
    lPIDL = SHBrowseForFolder(udtBI)

    ' Quit if canceled
    ' ----------------
    If (Not CBool(lPIDL)) Then
        GoTo PROC_EXIT
    End If

    ' Locate the selected folder
    ' --------------------------
    lRtn = SHGetPathFromIDList(ByVal lPIDL, ByVal sRtn)

    ' Trim the path
    ' -------------
    If lRtn Then
        sRtn = TrimToNull(sRtn)
    End If

    ' Make sure there is a trailing delimiter
    ' ---------------------------------------
    If ((Len(sRtn) > 3) And (Right$(sRtn, 1) <> "\")) Then
        sRtn = sRtn & "\"
    End If

    ' Release allocated memory
    ' ------------------------
    If bStartProvided Then
        Call CoTaskMemFree(lPIDL)
        Call LocalFree(lpPath)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ShellBrowseForFolder = sRtn

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

Public Function ShellExecute(ByVal FileName As String, _
                    Optional ByVal Parameters As String, _
                    Optional ByVal WorkingDir As String, _
                    Optional ByVal Verb As String, _
                    Optional ByVal Show _
                                As enuShowWindowCommand = SW_NORMAL) _
       As Boolean
' ==========================================================================
' Description : Open a file in the associated application
'
' Parameters  : FileName    The name of the file to open
'               Parameters  The parameters to be passed to the application.
'                           The format of this string is determined by
'                           the verb that is to be invoked.
'               WorkingDir  The default (working) directory for the action.
'               Verb        The verb that specifies the action to be
'                           performed. The set of available verbs
'                           depends on the particular file or folder.
'                           Generally, the actions available from an
'                           object's shortcut menu are available verbs.
'               Show        Specifies how an application is to be
'                           displayed when it is opened
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "ShellExecute"

    #If VBA7 Then
        Dim hWnd    As LongPtr
        Dim lRtn    As LongPtr
    #Else
        Dim hWnd    As Long
        Dim lRtn    As Long
    #End If

    Dim bRtn        As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (WorkingDir = vbNullString) Then
        WorkingDir = ParsePath(FileName, pppFullPath)
        FileName = ParsePath(FileName, pppFileOnly)
    End If

    lRtn = ShellExecuteANSI(hWnd, _
                            Verb, _
                            FileName, _
                            Parameters, _
                            WorkingDir, _
                            Show)

    bRtn = (lRtn > 32)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ShellExecute = bRtn

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
