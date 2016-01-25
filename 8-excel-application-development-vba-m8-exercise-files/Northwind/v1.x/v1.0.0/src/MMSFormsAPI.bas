Attribute VB_Name = "MMSFormsAPI"
' ==========================================================================
' Module      : MMSFormsAPI
' Type        : Module
' Description : Support for working with MSForms APIs
' --------------------------------------------------------------------------
' Procedures  : GetUserFormHandle               LongPtr
'               SetUserFormCloseButtonState     Boolean
'               SetUserFormParent               Boolean
'               UserFormIsMaximized             Boolean
'               UserFormIsMinimized             Boolean
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

' Window class names
' ------------------
Public Const gsCLASSNAME_VBA5_USERFORM  As String = "ThunderXFrame"
Public Const gsCLASSNAME_VBA6_USERFORM  As String = "ThunderDFrame"
Public Const gsCLASSNAME_VBA7_USERFORM  As String = "ThunderDFrame"

' ----------------
' Module Level
' ----------------

Private Const msMODULE As String = "MMSFormsAPI"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

' Used for setting the parent window
' ----------------------------------
Public Enum enuParentWindowType
    pwtNone = 0
    pwtApplication = 1
    pwtActiveWindow = 2
    pwtHandle = 3
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The FindWindow function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633499(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function FindWindow _
            Lib "user32.dll" _
            Alias "FindWindowA" (ByVal lpClassName As String, _
                                 ByVal lpClassName As String) _
            As LongPtr
#Else
    Private Declare _
            Function FindWindow _
            Lib "user32.dll" _
            Alias "FindWindowA" (ByVal lpClassName As String, _
                                 ByVal lpClassName As String) _
            As Long
#End If


' The FindWindowEx function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633500(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function FindWindowEx _
            Lib "user32.dll" _
            Alias "FindWindowExA" (ByVal hWndParent As LongPtr, _
                                   ByVal hWndChildAfter As LongPtr, _
                                   ByVal lpszClass As String, _
                                   ByVal lpszWindow As String) _
            As LongPtr
#Else
    Private Declare _
            Function FindWindowEx _
            Lib "user32.dll" _
            Alias "FindWindowExA" (ByVal hWndParent As Long, _
                                   ByVal hWndChildAfter As Long, _
                                   ByVal lpszClass As String, _
                                   ByVal lpszWindow As String) _
            As Long
#End If

' The EnableMenuItem function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms647636(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function EnableMenuItem _
            Lib "user32.dll" (ByVal hMenu As LongPtr, _
                          ByVal uIDEnableItem As Long, _
                          ByVal uEnable As enuMenuFlag) _
            As Boolean
#Else
    Private Declare _
            Function EnableMenuItem _
            Lib "user32.dll" (ByVal hMenu As Long, _
                          ByVal uIDEnableItem As Long, _
                          ByVal uEnable As enuMenuFlag) _
            As Boolean
#End If

' The GetMenuItemCount function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms647978(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetMenuItemCount _
            Lib "user32.dll" (ByVal hMenu As LongPtr) _
            As Long
#Else
    Private Declare _
            Function GetMenuItemCount _
            Lib "user32.dll" (ByVal hMenu As Long) _
            As Long
#End If

' The GetSystemMenu function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms647985(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetSystemMenu _
            Lib "user32.dll" (ByVal hWnd As LongPtr, _
                          ByVal bRevert As Long) _
            As LongPtr
#Else
    Private Declare _
            Function GetSystemMenu _
            Lib "user32.dll" (ByVal hWnd As Long, _
                          ByVal bRevert As Long) _
            As Long
#End If

' The SetParent function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms633541(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SetParent _
            Lib "user32.dll" (ByVal hWndChild As LongPtr, _
                          ByVal hWndNewParent As LongPtr) _
            As LongPtr
#Else
    Private Declare _
            Function SetParent _
            Lib "user32.dll" (ByVal hWndChild As Long, _
                          ByVal hWndNewParent As Long) _
            As Long
#End If

#If VBA7 Then
Public Function GetUserFormHandle(ByRef UF As Object) As LongPtr
#Else
Public Function GetUserFormHandle(ByRef UF As Object) As Long
#End If
' ==========================================================================
' Description : Return the handle to the UserForm
'
' Parameters  : Form        The form to find the handle for
'
' Returns     : LongPtr
' ==========================================================================

    Const sPROC     As String = "GetUserFormHandle"

    #If VBA7 Then
        Dim hWnd    As LongPtr
        Dim hWndP   As LongPtr  ' Parent
    #Else
        Dim hWnd    As Long
        Dim hWndP   As Long
    #End If

    Dim sCaption    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, UF.Caption)

    ' ----------------------------------------------------------------------
    ' Get the form caption
    ' -----------------------------
    sCaption = UF.Caption

    ' Try to find the form in the top-level windows
    ' ---------------------------------------------

    hWnd = FindWindow(gsCLASSNAME_VBA6_USERFORM, sCaption)

    If (hWnd <> 0) Then
        GoTo PROC_EXIT
    End If

    ' ----------------------------------------------------------------------
    ' Not a top level window.
    ' Search the child windows of the application
    ' -------------------------------------------
    hWndP = Application.hWnd
    hWnd = FindWindowEx(hWndP, 0&, gsCLASSNAME_VBA6_USERFORM, sCaption)

    If (hWnd <> 0) Then
        GoTo PROC_EXIT
    End If

    ' ----------------------------------------------------------------------
    ' Not a child of the application.
    ' Search for child of the ActiveWindow
    ' (Excel's ActiveWindow, not Window's ActiveWindow).
    ' --------------------------------------------------
    If (Application.ActiveWindow Is Nothing) Then
        GoTo PROC_EXIT
    End If

    hWndP = WindowHandle(Application.ActiveWindow)
    hWnd = FindWindowEx(hWndP, 0&, gsCLASSNAME_VBA6_USERFORM, sCaption)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetUserFormHandle = hWnd

    Call Trace(tlMaximum, msMODULE, sPROC, hWnd)
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

Public Function SetUserFormCloseButtonState(ByRef UF As Object, _
                                            ByVal Enabled As Boolean) _
       As Boolean
' ==========================================================================
' Description : Set the enabled state of the close button
'
' Parameters  : Enabled     Indicates the state for the close button
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "SetUserFormCloseButtonState"

    Dim bRtn        As Boolean
    Dim bRevert     As Boolean

    Dim lItemCnt    As Long
    Dim lItemID     As Long

    Dim eEnabled    As enuMenuFlag
    
    #If VBA7 Then
        Dim lhMenu  As LongPtr
        Dim lhWnd   As LongPtr
    #Else
        Dim lhMenu  As Long
        Dim lhWnd   As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the window handle
    ' ---------------------
    lhWnd = GetUserFormHandle(UF)
    If (lhWnd = 0) Then
        GoTo PROC_EXIT
    End If
    
    ' Get the menu handle
    ' -------------------
    lhMenu = GetSystemMenu(lhWnd, bRevert)

    If (lhMenu = 0) Then
        GoTo PROC_EXIT
    End If

    ' Locate the close item
    ' ---------------------
    lItemCnt = GetMenuItemCount(lhMenu)
    lItemID = lItemCnt - 1

    If Enabled Then
        eEnabled = (MF_ENABLED Or MF_BYPOSITION)
    Else
        eEnabled = (MF_DISABLED Or MF_BYPOSITION)
    End If

    bRtn = EnableMenuItem(lhMenu, lItemID, eEnabled)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    SetUserFormCloseButtonState = bRtn

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

#If VBA7 Then
    Public Function SetUserFormParent(ByRef UF As Object, _
                                      ByVal WindowType _
                                         As enuParentWindowType, _
                             Optional ByVal hWnd As LongPtr) As Boolean
#Else
    Public Function SetUserFormParent(ByRef UF As Object, _
                                      ByVal WindowType _
                                         As enuParentWindowType, _
                             Optional ByVal hWnd As Long) As Boolean
#End If
' ==========================================================================
' Description : Set the parent for the UserForm.
'               This can be the application, the ActiveWindow or no parent.
'
' Parameters  : UF              The UserForm to re-parent
'               WindowType      Identifies the new parent
'               hWnd            The handle of the new parent
'                               (valid only if WindowType = pwtHandle)
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "SetUserFormParent"

    Dim bRtn    As Boolean

    #If VBA7 Then
        Dim lhWnd   As LongPtr
        Dim lhApp   As LongPtr
        Dim lhWind  As LongPtr
        Dim lhUF    As LongPtr
    #Else
        Dim lhWnd   As Long
        Dim lhApp   As Long
        Dim lhWind  As Long
        Dim lhUF    As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Quit if there is no handle
    ' --------------------------
    lhUF = GetUserFormHandle(UF)
    If (lhUF = 0) Then
        GoTo PROC_EXIT
    End If

    ' Set the parent based on type
    ' ----------------------------
    Select Case WindowType
    Case pwtNone
        lhWnd = SetParent(lhUF, 0&)

    Case pwtApplication
        lhApp = Application.hWnd
        lhWnd = SetParent(lhUF, lhApp)
    
    Case pwtActiveWindow
        If (Application.ActiveWindow Is Nothing) Then
            GoTo PROC_EXIT
        End If
        
        lhWind = WindowHandle(ActiveWindow)
        If (lhWind = 0) Then
            GoTo PROC_EXIT
        End If
        
        lhWnd = SetParent(lhUF, lhWind)

    Case pwtHandle
        If (hWnd = 0) Then
            GoTo PROC_EXIT
        End If

        lhWnd = SetParent(lhUF, hWnd)

    End Select

    ' The function should return the
    ' handle of the previous parent
    ' ------------------------------
    bRtn = (lhWnd <> 0)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    SetUserFormParent = bRtn

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

Public Function UserFormIsMaximized(ByRef UF As Object) As Boolean
' ==========================================================================
' Description : Determines if a UserForm is maximized
'
' Parameters  : UF      The UserForm to inspect
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "UserFormIsMaximized"

    Dim bRtn        As Boolean

    #If VBA7 Then
        Dim lhWnd   As LongPtr
    #Else
        Dim lhWnd   As Long
    #End If

    
    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the handle
    ' --------------
    lhWnd = GetUserFormHandle(UF)

    If (lhWnd = 0) Then
        GoTo PROC_EXIT
    End If
    
    bRtn = WindowIsMaximized(lhWnd)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    UserFormIsMaximized = bRtn

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

Public Function UserFormIsMinimized(ByRef UF As Object) As Boolean
' ==========================================================================
' Description : Determines if a UserForm is minimized
'
' Parameters  : UF      The UserForm to inspect
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "UserFormIsMinimized"

    Dim bRtn        As Boolean

    #If VBA7 Then
        Dim lhWnd   As LongPtr
    #Else
        Dim lhWnd   As Long
    #End If

    
    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the handle
    ' --------------
    lhWnd = GetUserFormHandle(UF)

    If (lhWnd = 0) Then
        GoTo PROC_EXIT
    End If
    
    bRtn = WindowIsMinimized(lhWnd)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    UserFormIsMinimized = bRtn

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
