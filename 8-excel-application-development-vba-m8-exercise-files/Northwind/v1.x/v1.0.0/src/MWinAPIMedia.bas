Attribute VB_Name = "MWinAPIMedia"
' ==========================================================================
' Module      : MWinAPIMedia
' Type        : Module
' Description : Support for multimedia functions
' --------------------------------------------------------------------------
' Procedures  : CanPlayWaveData             Boolean
'               PlayEventSound              Boolean
'               PlaySound                   Boolean
'               PlaySoundAliasToString      String
'               StringToPlaySoundAlias      enuPlaySoundAlias
'               ToggleMute
' --------------------------------------------------------------------------
' Dependencies: MWinAPIRegistry
'               MWinAPIUser32Keyboard
' --------------------------------------------------------------------------
' Comments    : A list of Multimedia Functions can be found at
'             http://msdn.microsoft.com/en-us/library/dd743586(v=vs.85).aspx
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

Public Const WAVERR_BASE                    As Long = 32
Public Const MIDIERR_BASE                   As Long = 64
Public Const TIMERR_BASE                    As Long = 96
Public Const JOYERR_BASE                    As Long = 160
Public Const MCIERR_BASE                    As Long = 256
Public Const MIXERR_BASE                    As Long = 1024

Public Const MCI_STRING_OFFSET              As Long = 512
Public Const MCI_VD_OFFSET                  As Long = 1024
Public Const MCI_CD_OFFSET                  As Long = 1088
Public Const MCI_WAVE_OFFSET                As Long = 1152
Public Const MCI_SEQ_OFFSET                 As Long = 1216


' General error return values
' ---------------------------
Public Const MMSYSERR_BASE                  As Long = 0
Public Const MMSYSERR_NOERROR               As Long = 0                     ' no error
Public Const MMSYSERR_ERROR                 As Long = (MMSYSERR_BASE + 1)   ' unspecified error
Public Const MMSYSERR_BADDEVICEID           As Long = (MMSYSERR_BASE + 2)   ' device ID out of range
Public Const MMSYSERR_NOTENABLED            As Long = (MMSYSERR_BASE + 3)   ' driver failed enable
Public Const MMSYSERR_ALLOCATED             As Long = (MMSYSERR_BASE + 4)   ' device already allocated
Public Const MMSYSERR_INVALHANDLE           As Long = (MMSYSERR_BASE + 5)   ' device handle is invalid
Public Const MMSYSERR_NODRIVER              As Long = (MMSYSERR_BASE + 6)   ' no device driver present
Public Const MMSYSERR_NOMEM                 As Long = (MMSYSERR_BASE + 7)   ' memory allocation error
Public Const MMSYSERR_NOTSUPPORTED          As Long = (MMSYSERR_BASE + 8)   ' function isn't supported
Public Const MMSYSERR_BADERRNUM             As Long = (MMSYSERR_BASE + 9)   ' error value out of range
Public Const MMSYSERR_INVALFLAG             As Long = (MMSYSERR_BASE + 10)  ' invalid flag passed
Public Const MMSYSERR_INVALPARAM            As Long = (MMSYSERR_BASE + 11)  ' invalid parameter passed
Public Const MMSYSERR_HANDLEBUSY            As Long = (MMSYSERR_BASE + 12)  ' handle being used simultaneously on another thread (eg callback)
Public Const MMSYSERR_INVALIDALIAS          As Long = (MMSYSERR_BASE + 13)  ' specified alias not found
Public Const MMSYSERR_BADDB                 As Long = (MMSYSERR_BASE + 14)  ' bad registry database
Public Const MMSYSERR_KEYNOTFOUND           As Long = (MMSYSERR_BASE + 15)  ' registry key not found
Public Const MMSYSERR_READERROR             As Long = (MMSYSERR_BASE + 16)  ' registry read error
Public Const MMSYSERR_WRITEERROR            As Long = (MMSYSERR_BASE + 17)  ' registry write error
Public Const MMSYSERR_DELETEERROR           As Long = (MMSYSERR_BASE + 18)  ' registry delete error
Public Const MMSYSERR_VALNOTFOUND           As Long = (MMSYSERR_BASE + 19)  ' registry value not found
Public Const MMSYSERR_NODRIVERCB            As Long = (MMSYSERR_BASE + 20)  ' driver does not call DriverCallback
Public Const MMSYSERR_MOREDATA              As Long = (MMSYSERR_BASE + 21)  ' more data to be returned
Public Const MMSYSERR_LASTERROR             As Long = (MMSYSERR_BASE + 21)  ' last error in range

' ----------------
' Module Level
' ----------------

Private Const msMODULE                      As String = "MWinAPIMedia"

Private Const msSND_ALIAS_SYSTEMASTERISK    As String = "System Asterisk"
Private Const msSND_ALIAS_SYSTEMQUESTION    As String = "System Question"
Private Const msSND_ALIAS_SYSTEMHAND        As String = "System Hand"
Private Const msSND_ALIAS_SYSTEMEXIT        As String = "System Exit"
Private Const msSND_ALIAS_SYSTEMSTART       As String = "System Start"
Private Const msSND_ALIAS_SYSTEMWELCOME     As String = "System Welcome"
Private Const msSND_ALIAS_SYSTEMEXCLAMATION As String = "System Exclamation"
Private Const msSND_ALIAS_SYSTEMDEFAULT     As String = "System Default"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

' PlaySound flags are described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd743680(v=vs.85).aspx
' -----------------------------------
Public Enum enuPlaySoundFlag
    SND_SYNC = &H0                      ' Play synchronously (default)
    SND_ASYNC = &H1                     ' Play asynchronously
    SND_NODEFAULT = &H2                 ' Silence (!default) if sound not found
    SND_MEMORY = &H4                    ' pszSound points to a memory file
    SND_LOOP = &H8                      ' Loop the sound until next sndPlaySound
    SND_NOSTOP = &H10                   ' Don't stop any currently playing sound

    SND_PURGE = &H40                    ' Purge non-static events for task
    SND_APPLICATION = &H80              ' Look for application specific association

    SND_NOWAIT = &H2000                 ' Don't wait if the driver is busy
    SND_ALIAS = &H10000                 ' Name is a registry alias
    SND_ALIAS_ID = &H110000             ' Alias is a predefined ID
    SND_FILENAME = &H20000              ' Name is file name
    SND_RESOURCE = &H40004              ' Name is resource name or atom
End Enum

' In the Windows API, the sndAlias macro is defined as
' #define sndAlias(ch0, ch1)
' (SND_ALIAS_START + (DWORD)(BYTE)(ch0) | ((DWORD)(BYTE)(ch1) << 8))
'
' Internally, this macro takes the ASCII value of 2 characters,
' shifts the bits from the second character to the left by 8 bits,
' then uses the OR operator to arrive at the final value.
'
' PlaySound aliases are described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd743680(v=vs.85).aspx
' -----------------------------------
Public Enum enuPlaySoundAlias
    SND_ALIAS_START = 0                 ' Alias base
    SND_ALIAS_SYSTEMASTERISK = 10835    ' sndAlias('S', '*') 83 | 10752 (00101010<<8)
    SND_ALIAS_SYSTEMQUESTION = 16211    ' sndAlias('S', '?') 83 | 16128 (00111111<<8)
    SND_ALIAS_SYSTEMHAND = 18515        ' sndAlias('S', 'H') 83 | 18432 (01001000<<8)
    SND_ALIAS_SYSTEMEXIT = 17747        ' sndAlias('S', 'E') 83 | 17664 (01000101<<8)
    SND_ALIAS_SYSTEMSTART = 21331       ' sndAlias('S', 'S') 83 | 21248 (01010011<<8)
    SND_ALIAS_SYSTEMWELCOME = 22355     ' sndAlias('S', 'W') 83 | 22272 (01010111<<8)
    SND_ALIAS_SYSTEMEXCLAMATION = 8531  ' sndAlias('S', '!') 83 |  8448 (00100001<<8)
    SND_ALIAS_SYSTEMDEFAULT = 17491     ' sndAlias('S', 'D') 83 | 17408 (01000100<<8)
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The PlaySound function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd743680(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function PlayWaveFile _
            Lib "Winmm.dll" _
            Alias "PlaySoundA" (ByVal pszSound As String, _
                                ByVal hMod As LongPtr, _
                                ByVal fdwSound As enuPlaySoundFlag) _
            As Boolean
#Else
    Private Declare _
            Function PlayWaveFile _
            Lib "Winmm.dll" _
            Alias "PlaySoundA" (ByVal pszSound As String, _
                                ByVal hMod As Long, _
                                ByVal fdwSound As enuPlaySoundFlag) _
            As Boolean
#End If

' The waveOutGetNumDevs function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd743860(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function waveOutGetNumDevs _
            Lib "Winmm.dll" () _
            As Long
#Else
    Private Declare _
            Function waveOutGetNumDevs _
            Lib "Winmm.dll" () _
            As Long
#End If

' The waveOutGetVolume function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd743864(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function waveOutGetVolume _
            Lib "Winmm.dll" (ByVal uDeviceID As Long, _
                             ByRef lpdwVolume As Long) _
            As Long
#Else
    Private Declare _
            Function waveOutGetVolume _
            Lib "Winmm.dll" (ByVal uDeviceID As Long, _
                             ByRef lpdwVolume As Long) _
            As Long
#End If

' The waveOutSetVolume function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd743874(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function waveOutSetVolume _
            Lib "Winmm.dll" (ByVal uDeviceID As Long, _
                             ByVal dwVolume As Long) _
            As Long
#Else
    Private Declare _
            Function waveOutSetVolume _
            Lib "Winmm.dll" (ByVal uDeviceID As Long, _
                             ByVal dwVolume As Long) _
            As Long
#End If

Public Function CanPlayWaveData() As Boolean
' ==========================================================================
' Description : Determines if there is a device
'               capable of playing wave files.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "CanPlayWaveData"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    bRtn = (waveOutGetNumDevs() > 0)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    CanPlayWaveData = bRtn

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

Public Function PlayEventSound(ByVal EventType As enuPlaySoundAlias) _
       As Boolean
' ==========================================================================
' Description : Play a system event sound
'
' Parameters  : EventType   The pre-defined event type
' ==========================================================================

    Const sPROC     As String = "PlayEventSound"

    Dim bRtn        As Boolean

    Dim sFileName   As String
    Dim sRegKey     As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Set the base key
    ' ----------------
    sRegKey = "AppEvents\Schemes\Apps\.Default\"

    Select Case EventType
    Case SND_ALIAS_SYSTEMDEFAULT
        sRegKey = sRegKey & ".Default\.Current"

    Case SND_ALIAS_SYSTEMASTERISK
        sRegKey = sRegKey & "SystemAsterisk\.Current"

    Case SND_ALIAS_SYSTEMEXCLAMATION
        sRegKey = sRegKey & "SystemExclamation\.Current"

    Case SND_ALIAS_SYSTEMEXIT
        sRegKey = sRegKey & "SystemExit\.Current"

    Case SND_ALIAS_SYSTEMHAND
        sRegKey = sRegKey & "SystemHand\.Current"

    Case SND_ALIAS_SYSTEMQUESTION
        sRegKey = sRegKey & "SystemQuestion\.Current"

    Case SND_ALIAS_SYSTEMSTART
        sRegKey = sRegKey & "WindowsUAC\.Current"

    Case SND_ALIAS_SYSTEMWELCOME
        sRegKey = sRegKey & "WindowsLogon\.Current"

    Case Else
        sRegKey = sRegKey & ".Default\.Current"

    End Select

    sFileName = RegGetKeyValue(HKEY_CURRENT_USER, sRegKey, vbNullString)

    bRtn = PlaySound(sFileName, SND_ASYNC)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    PlayEventSound = bRtn

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

Public Function PlaySound(ByVal FileName As String, _
                 Optional ByVal Flags As enuPlaySoundFlag) As Boolean
' ==========================================================================
' Description : Play a wave file
'
' Parameters  : FileName    The name of the wave file to play
'               Flags       Optional playback modifiers
' ==========================================================================

    Const sPROC As String = "PlaySound"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, FileName)

    ' ----------------------------------------------------------------------

    bRtn = PlayWaveFile(FileName, 0&, Flags)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    PlaySound = bRtn

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

Public Function PlaySoundAliasToString(ByVal SoundAlias _
                                          As enuPlaySoundAlias) As String
' ==========================================================================
' Description : Convert an enumeration to a string
'
' Parameters  : SoundAlias  The enumeration to convert
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "PlaySoundAliasToString"

    Dim sRtn    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Select Case SoundAlias
    Case SND_ALIAS_SYSTEMASTERISK
        sRtn = msSND_ALIAS_SYSTEMASTERISK

    Case SND_ALIAS_SYSTEMQUESTION
        sRtn = msSND_ALIAS_SYSTEMQUESTION

    Case SND_ALIAS_SYSTEMHAND
        sRtn = msSND_ALIAS_SYSTEMHAND

    Case SND_ALIAS_SYSTEMEXIT
        sRtn = msSND_ALIAS_SYSTEMEXIT

    Case SND_ALIAS_SYSTEMSTART
        sRtn = msSND_ALIAS_SYSTEMSTART

    Case SND_ALIAS_SYSTEMWELCOME
        sRtn = msSND_ALIAS_SYSTEMWELCOME

    Case SND_ALIAS_SYSTEMEXCLAMATION
        sRtn = msSND_ALIAS_SYSTEMEXCLAMATION

    Case Else
        sRtn = msSND_ALIAS_SYSTEMDEFAULT

    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    PlaySoundAliasToString = sRtn

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

Public Function StringToPlaySoundAlias(ByVal SoundAlias As String) _
       As enuPlaySoundAlias
' ==========================================================================
' Description : Convert a string to an enumeration
'
' Parameters  : SoundAlias  The string to convert
'
' Returns     : enuPlaySoundAlias
' ==========================================================================

    Const sPROC As String = "StringToPlaySoundAlias"

    Dim eRtn    As enuPlaySoundAlias


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Select Case SoundAlias
    Case msSND_ALIAS_SYSTEMASTERISK
        eRtn = SND_ALIAS_SYSTEMASTERISK

    Case msSND_ALIAS_SYSTEMQUESTION
        eRtn = SND_ALIAS_SYSTEMQUESTION

    Case msSND_ALIAS_SYSTEMHAND
        eRtn = SND_ALIAS_SYSTEMHAND

    Case msSND_ALIAS_SYSTEMEXIT
        eRtn = SND_ALIAS_SYSTEMEXIT

    Case msSND_ALIAS_SYSTEMSTART
        eRtn = SND_ALIAS_SYSTEMSTART

    Case msSND_ALIAS_SYSTEMWELCOME
        eRtn = SND_ALIAS_SYSTEMWELCOME

    Case msSND_ALIAS_SYSTEMEXCLAMATION
        eRtn = SND_ALIAS_SYSTEMEXCLAMATION

    Case msSND_ALIAS_SYSTEMDEFAULT
        eRtn = SND_ALIAS_SYSTEMDEFAULT

    Case Else
        eRtn = SND_ALIAS_SYSTEMDEFAULT

    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    StringToPlaySoundAlias = eRtn

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

Public Sub ToggleMute()
' ==========================================================================
' Description : Toggle the state of the system mute
' ==========================================================================

    Const sPROC As String = "ToggleMute"

    Dim bRtn    As Boolean


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Call SendVirtualKey(VK_VOLUME_MUTE)

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
