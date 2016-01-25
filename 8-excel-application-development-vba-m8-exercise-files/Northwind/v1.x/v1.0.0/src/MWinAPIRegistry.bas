Attribute VB_Name = "MWinAPIRegistry"
' ==========================================================================
' Module      : MWinAPIRegistry
' Type        : Module
' Description : Support for the Windows Registry
' --------------------------------------------------------------------------
' Procedures  : RegDeleteKeyValue
'               RegGetKeyValue          Variant
'               RegKeyExists            Boolean
'               RegSetKeyValue
' --------------------------------------------------------------------------
' Dependencies: MWinAPISecurity
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

Public Const ERR_REG_SUCCESS            As Long = &H0
Public Const ERR_REG_BADDB              As Long = &H1
Public Const ERR_REG_BADKEY             As Long = &H2
Public Const ERR_REG_CANTOPEN           As Long = &H3
Public Const ERR_REG_CANTREAD           As Long = &H4
Public Const ERR_REG_CANTWRITE          As Long = &H5
Public Const ERR_REG_OUTOFMEMORY        As Long = &H6
Public Const ERR_REG_INVALIDPARAMETER   As Long = &H7
Public Const ERR_REG_ACCESSDENIED       As Long = &H8
Public Const ERR_REG_INVALIDPARAMETERS  As Long = &H57
Public Const ERR_REG_NOMOREITEMS        As Long = &H103

' ----------------
' Module Level
' ----------------

Private Const msMODULE                  As String = "MWinAPIRegistry"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuRegKeySecurity
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_CREATE_LINK = &H20
    KEY_WOW64_32KEY = &H200
    KEY_WOW64_64KEY = &H100
    KEY_WOW64_RES = &H300

    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
End Enum

' Open/Create options are defined in WinNT.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724844(v=vs.85).aspx
' -----------------------------------
Public Enum enuRegOption
    REG_OPTION_NON_VOLATILE = 0
    REG_OPTION_VOLATILE = 1
    REG_OPTION_CREATE_LINK = 2
    REG_OPTION_BACKUP = 4
End Enum

' Registry keys are described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724836(v=vs.85).aspx
' -----------------------------------
Public Enum enuRegRootKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    HKCR = HKEY_CLASSES_ROOT
    HKCU = HKEY_CURRENT_USER
    HKLM = HKEY_LOCAL_MACHINE
    HKU = HKEY_USERS
End Enum

' Registry Routine Flags are described on MSDN at
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724868(v=vs.85).aspx
' -----------------------------------
Public Enum enuRegRoutineFlags
    RRF_RT_REG_NONE = &H1               ' Restrict type to REG_NONE
    RRF_RT_REG_SZ = &H2                 ' Restrict type to REG_SZ
    RRF_RT_REG_EXPAND_SZ = &H4          ' Restrict type to REG_EXPAND_SZ
    RRF_RT_REG_BINARY = &H8             ' Restrict type to REG_BINARY
    RRF_RT_REG_DWORD = &H10             ' Restrict type to REG_DWORD
    RRF_RT_REG_MULTI_SZ = &H20          ' Restrict type to REG_MULTI_SZ
    RRF_RT_REG_QWORD = &H40             ' Restrict type to REG_QWORD

    RRF_RT_DWORD = &H18                 ' Restrict type to 32-bit
                                        ' RRF_RT_REG_BINARY | RRF_RT_REG_DWORD
    RRF_RT_QWORD = &H48                 ' Restrict type to 64-bit
                                        ' RRF_RT_REG_BINARY | RRF_RT_REG_DWORD
    RRF_RT_ANY = &HFFFF                 ' No type restriction
    
    RRF_NOEXPAND = &H10000000           ' Do not automatically expand
                                        ' environment strings if the value
                                        ' is of type REG_EXPAND_SZ
    RRF_ZEROONFAILURE = &H20000000      ' If pvData is not NULL, set the
                                        ' contents of the buffer to
                                        ' zeroes on failure
End Enum

' Registry value types are defined in WinNT.h and described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724884(v=vs.85).aspx
' -----------------------------------
Public Enum enuRegValueType
    REG_NONE = 0                        ' No value type
    REG_SZ = 1                          ' Unicode null-terminated string
    REG_EXPAND_SZ = 2                   ' Unicode null-terminated string
    REG_BINARY = 3                      ' Free form binary
    REG_DWORD = 4                       ' 32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4         ' 32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = 5            ' 32-bit number
    REG_LINK = 6                        ' Symbolic Link (unicode)
    REG_MULTI_SZ = 7                    ' Multiple Unicode strings
    REG_RESOURCE_LIST = 8               ' Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = 9    ' Resource list in the hardware description
    REG_RESOURCE_REQUIREMENTS_LIST = 10
    REG_QWORD = 11                      ' 64-bit number
    REG_QWORD_LITTLE_ENDIAN = REG_QWORD ' 64-bit number (same as REG_QWORD)
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The RegCloseKey function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724837(v=vs.85)
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegCloseKey _
            Lib "advapi32" (ByVal hKey As LongPtr) _
            As LongPtr
#Else
    Private Declare _
            Function RegCloseKey _
            Lib "advapi32" (ByVal hKey As Long) _
            As Long
#End If

' The RegCreateKeyEx function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724844(v=vs.85)
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegCreateKeyEx _
            Lib "advapi32" _
            Alias "RegCreateKeyExA" (ByVal hKey As enuRegRootKey, _
                                     ByVal lpSubKey As String, _
                                     ByVal Reserved As LongPtr, _
                                     ByVal lpClass As String, _
                                     ByVal dwOptions As Long, _
                                     ByVal samDesired As enuRegKeySecurity, _
                                     ByRef lpSecurityAttributes _
                                        As TSECURITY_ATTRIBUTES, _
                                     ByRef phkResult As LongPtr, _
                                     ByRef lpdwDisposition As Long) _
            As LongPtr
#Else
    Private Declare _
            Function RegCreateKeyEx _
            Lib "advapi32" _
            Alias "RegCreateKeyExA" (ByVal hKey As enuRegRootKey, _
                                     ByVal lpSubKey As String, _
                                     ByVal Reserved As Long, _
                                     ByVal lpClass As String, _
                                     ByVal dwOptions As Long, _
                                     ByVal samDesired As enuRegKeySecurity, _
                                     ByRef lpSecurityAttributes _
                                        As TSECURITY_ATTRIBUTES, _
                                     ByRef phkResult As Long, _
                                     ByRef lpdwDisposition As Long) _
            As Long
#End If

' The RegDeleteValue function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724851(v=vs.85)
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegDeleteValue _
            Lib "advapi32" _
            Alias "RegDeleteValueA" (ByVal hKey As LongPtr, _
                                     ByVal lpValueName As String) _
            As LongPtr
#Else
    Private Declare _
            Function RegDeleteValue _
            Lib "advapi32" _
            Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                     ByVal lpValueName As String) _
            As Long
#End If

' The RegOpenKeyEx function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724897(VS.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegOpenKeyEx _
            Lib "advapi32" _
            Alias "RegOpenKeyExA" (ByVal hKey As enuRegRootKey, _
                                   ByVal lpSubKey As String, _
                                   ByVal ulOptions As LongPtr, _
                                   ByVal samDesired As LongPtr, _
                                   ByRef phkResult As LongPtr) _
            As LongPtr
#Else
    Private Declare _
            Function RegOpenKeyEx _
            Lib "advapi32" _
            Alias "RegOpenKeyExA" (ByVal hKey As enuRegRootKey, _
                                   ByVal lpSubKey As String, _
                                   ByVal ulOptions As Long, _
                                   ByVal samDesired As Long, _
                                   ByRef phkResult As Long) _
            As Long
#End If

' The RegQueryValueEx function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724911(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegQueryValueExNull _
            Lib "advapi32" _
            Alias "RegQueryValueExA" (ByVal hKey As LongPtr, _
                                      ByVal lpValueName As String, _
                                      ByVal lpReserved As LongPtr, _
                                      ByRef lpType As enuRegValueType, _
                                      ByVal lpData As LongPtr, _
                                      ByRef lpcbData As Long) _
            As LongPtr
#Else
    Private Declare _
            Function RegQueryValueExNull _
            Lib "advapi32" _
            Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                      ByVal lpValueName As String, _
                                      ByVal lpReserved As Long, _
                                      ByRef lpType As enuRegValueType, _
                                      ByVal lpData As Long, _
                                      ByRef lpcbData As Long) _
            As Long
#End If

' The RegQueryValueEx function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724911(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegQueryValueExLong _
            Lib "advapi32" _
            Alias "RegQueryValueExA" (ByVal hKey As LongPtr, _
                                      ByVal lpValueName As String, _
                                      ByVal lpReserved As LongPtr, _
                                      ByRef lpType As enuRegValueType, _
                                      ByRef lpData As LongPtr, _
                                      ByRef lpcbData As Long) _
            As LongPtr
#Else
    Private Declare _
            Function RegQueryValueExLong _
            Lib "advapi32" _
            Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                      ByVal lpValueName As String, _
                                      ByVal lpReserved As Long, _
                                      ByRef lpType As enuRegValueType, _
                                      ByRef lpData As Long, _
                                      ByRef lpcbData As Long) _
            As Long
#End If


' The RegQueryValueEx function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724911(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegQueryValueExString _
            Lib "advapi32" _
            Alias "RegQueryValueExA" (ByVal hKey As LongPtr, _
                                      ByVal lpValueName As String, _
                                      ByVal lpReserved As LongPtr, _
                                      ByRef lpType As LongPtr, _
                                      ByVal lpData As String, _
                                      ByRef lpcbData As Long) _
            As LongPtr
#Else
    Private Declare _
            Function RegQueryValueExString _
            Lib "advapi32" _
            Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                      ByVal lpValueName As String, _
                                      ByVal lpReserved As Long, _
                                      ByRef lpType As Long, _
                                      ByVal lpData As String, _
                                      ByRef lpcbData As Long) _
            As Long
#End If

' The RegSetValueEx function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724923(v=vs.85)
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegSetValueExLong _
            Lib "advapi32" _
            Alias "RegSetValueExA" (ByVal lngHKey As LongPtr, _
                                    ByVal lpValueName As String, _
                                    ByVal Reserved As LongPtr, _
                                    ByVal dwType As LongPtr, _
                                    ByRef lpValue As LongPtr, _
                                    ByVal cbData As Long) _
            As LongPtr
#Else
    Private Declare _
            Function RegSetValueExLong _
            Lib "advapi32" _
            Alias "RegSetValueExA" (ByVal lngHKey As Long, _
                                    ByVal lpValueName As String, _
                                    ByVal Reserved As Long, _
                                    ByVal dwType As Long, _
                                    ByRef lpValue As Long, _
                                    ByVal cbData As Long) _
            As Long
#End If

' The RegSetValueEx function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms724923(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function RegSetValueExString _
            Lib "advapi32" _
            Alias "RegSetValueExA" (ByVal lngHKey As LongPtr, _
                                    ByVal lpValueName As String, _
                                    ByVal Reserved As LongPtr, _
                                    ByVal dwType As LongPtr, _
                                    ByVal lpValue As String, _
                                    ByVal cbData As Long) _
            As LongPtr
#Else
    Private Declare _
            Function RegSetValueExString _
            Lib "advapi32" _
            Alias "RegSetValueExA" (ByVal lngHKey As Long, _
                                    ByVal lpValueName As String, _
                                    ByVal Reserved As Long, _
                                    ByVal dwType As Long, _
                                    ByVal lpValue As String, _
                                    ByVal cbData As Long) _
            As Long
#End If

Public Sub RegDeleteKeyValue(ByVal Key As enuRegRootKey, _
                             ByVal SubKey As String, _
                             ByVal ValueName As String)
' ==========================================================================
' Description : Delete a value from a Registry key
'
' Parameters  : Key         The root key
'               SubKey      The SubKey
'               ValueName   The name of the value to delete
' ==========================================================================

    Const sPROC As String = "RegDeleteKeyValue"

    #If VBA7 Then
        Dim lRtn As LongPtr
        Dim hKey As LongPtr
    #Else
        Dim lRtn As Long
        Dim hKey As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, SubKey)

    ' ----------------------------------------------------------------------

    lRtn = RegOpenKeyEx(Key, SubKey, 0&, KEY_WRITE, hKey)

    If (lRtn <> ERR_REG_SUCCESS) Then
        GoTo PROC_EXIT
    End If

    lRtn = RegDeleteValue(hKey, ValueName)
    lRtn = RegCloseKey(hKey)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, ValueName)
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

Public Function RegGetKeyValue(ByVal Key As enuRegRootKey, _
                               ByVal SubKey As String, _
                               ByVal ValueName As String) As Variant
' ==========================================================================
' Description : Read a value from the Windows Registry
'
' Parameters  : Key         The top-level Registry key
'               SubKey      The SubKey to look in
'               valueName   The name of the value to return
'
' Returns     : Variant
' ==========================================================================

    Const sPROC As String = "RegGetKeyValue"

    Dim vRtn    As Variant

    #If VBA7 Then
        Dim hKey As LongPtr
        Dim lRtn As LongPtr
    #Else
        Dim hKey As Long
        Dim lRtn As Long
    #End If

    Dim lSize   As Long
    Dim lType   As Long
    Dim lVAL    As Long
    Dim sVal    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, ValueName)

    ' ----------------------------------------------------------------------

    lRtn = RegOpenKeyEx(Key, SubKey, 0&, KEY_READ, hKey)

    If (lRtn <> ERR_REG_SUCCESS) Then
        GoTo PROC_EXIT
    End If

    ' First determine the datatype
    ' ----------------------------
    lRtn = RegQueryValueExNull(hKey, ValueName, 0&, lType, 0&, lSize)

    If (lRtn <> ERR_REG_SUCCESS) Then
        RegCloseKey (hKey)
        GoTo PROC_EXIT
    End If

    Select Case lType
    Case REG_SZ
        If (lSize > 0) Then
            sVal = String$(lSize, 0)
            lRtn = RegQueryValueExString(hKey, _
                                         ValueName, _
                                         0&, _
                                         lType, _
                                         sVal, _
                                         lSize)

            If (lRtn = ERR_REG_SUCCESS) Then
                vRtn = Left$(sVal, lSize)
            End If
        End If

    Case REG_DWORD
        lRtn = RegQueryValueExLong(hKey, ValueName, 0&, lType, lVAL, lSize)

        If (lRtn = ERR_REG_SUCCESS) Then
            vRtn = lVAL
        End If
    End Select

    lRtn = RegCloseKey(hKey)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    RegGetKeyValue = vRtn

    Call Trace(tlMaximum, msMODULE, sPROC, vRtn)
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

Public Function RegKeyExists(ByVal RootKey As enuRegRootKey, _
                             ByVal SubKey As String) As Boolean
' ==========================================================================
' Description : Determines if a key exists in the Windows Registry
'
' Parameters  : RootKey   The root key
'               SubKey    The SubKey
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "RegKeyExists"

    Dim bRtn    As Boolean

    #If VBA7 Then
        Dim hKey As LongPtr
        Dim lRtn As LongPtr
    #Else
        Dim hKey As Long
        Dim lRtn As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, SubKey)

    ' ----------------------------------------------------------------------

    lRtn = RegOpenKeyEx(RootKey, SubKey, 0&, KEY_READ, hKey)

    If (lRtn = ERR_REG_SUCCESS) Then
        bRtn = True
        RegCloseKey (hKey)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    RegKeyExists = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, CStr(bRtn))
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

Public Function RegValueExists(ByVal RootKey As enuRegRootKey, _
                               ByVal SubKey As String, _
                      Optional ByVal ValName As String) As Boolean
' ==========================================================================
' Description : Determines if a named value exists in the Windows Registry
'
' Parameters  : RootKey     The root key to use.
'               SubKey      The SubKey to look in.
'               ValName     The name of the value to look for.
'                           If not provided, the default value will be used.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "RegValueExists"

    Dim bRtn        As Boolean

    Dim lValType    As enuRegValueType
    Dim lSize       As Long

    #If VBA7 Then
        Dim hKey    As LongPtr
        Dim lRtn    As LongPtr
    #Else
        Dim hKey    As Long
        Dim lRtn    As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, SubKey)

    ' ----------------------------------------------------------------------

    lRtn = RegOpenKeyEx(RootKey, SubKey, 0&, KEY_READ, hKey)

    If (lRtn = ERR_REG_SUCCESS) Then
        lRtn = RegQueryValueExNull(hKey, ValName, 0&, lValType, 0&, lSize)
        If (lValType <> REG_NONE) Then
            bRtn = True
        End If
        RegCloseKey (hKey)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    RegValueExists = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, CStr(bRtn))
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

Public Function RegSetKeyValue(ByVal RootKey As enuRegRootKey, _
                               ByVal SubKey As String, _
                               ByVal Name As String, _
                               ByRef Value As Variant, _
                               ByVal ValueType As enuRegValueType) _
       As Boolean
' ==========================================================================
' Description : Write a value to the Windows Registry
'
' Parameters  : RootKey   The root key
'               SubKey    The SubKey
'               Name      The name of the value being changed
'               Value      The new value
'               ValueType  The type of Value being written
' ==========================================================================

    Const sPROC As String = "RegSetKeyValue"

    #If VBA7 Then
        Dim hKey As LongPtr
        Dim lRtn As LongPtr
    #Else
        Dim hKey As Long
        Dim lRtn As Long
    #End If

    Dim bRtn        As Boolean
    Dim lValue      As Long
    Dim sValue      As String

    Dim udtSec      As TSECURITY_ATTRIBUTES

    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Name)

    ' ----------------------------------------------------------------------
    ' Initialize the structure
    ' ------------------------

    With udtSec
        .bInheritHandle = True
        .nLength = Len(udtSec)
    End With

    ' Open or create the key
    ' ----------------------
    lRtn = RegCreateKeyEx(RootKey, _
                          SubKey, _
                          0&, _
                          vbNullString, _
                          REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, _
                          udtSec, _
                          hKey, _
                          0&)

    If ((lRtn <> ERROR_SUCCESS) Or (lRtn = ERROR_ACCESS_DENIED)) Then
        GoTo PROC_EXIT
    End If

    ' Set the value
    ' -------------
    Select Case ValueType
    Case REG_SZ
        sValue = Value & vbNullChar
        lRtn = RegSetValueExString(hKey, _
                                   Name, _
                                   0&, _
                                   ValueType, _
                                   sValue, _
                                   Len(sValue))
    Case REG_DWORD
        lValue = Value
        lRtn = RegSetValueExLong(hKey, _
                                 Name, _
                                 0&, _
                                 ValueType, _
                                 lValue, _
                                 Len(lValue))
    End Select

    If (lRtn = ERROR_SUCCESS) Then
        bRtn = True
    End If

    lRtn = RegCloseKey(hKey)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    RegSetKeyValue = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, CStr(Value))
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
