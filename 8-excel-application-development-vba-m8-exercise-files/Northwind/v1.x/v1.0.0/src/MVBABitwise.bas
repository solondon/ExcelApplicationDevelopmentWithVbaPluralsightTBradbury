Attribute VB_Name = "MVBABitwise"
' ==========================================================================
' Module      : MVBABitwise
' Type        : Module
' Description : A collection of bitwise operations
' --------------------------------------------------------------------------
' Procedures  : BitIsSet            Boolean
'               HiByte              Byte
'               HiWord              Integer
'               LoByte              Byte
'               LoWord              Integer
'               MakeDword           Long
'               MakeWord            Integer
'               SetBit              Long
'               ShiftLeft           Long
'               ShiftRight          Long
'               SingleBitIsSet      Boolean
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

Private Const msMODULE As String = "MVBABitwise"

#If Win64 Then
    Public Function BitIsSet(ByRef Word As LongLong, _
                             ByRef Bit As LongLong) As Boolean
#Else
    Public Function BitIsSet(ByRef Word As Long, _
                             ByRef Bit As Long) As Boolean
#End If
' ==========================================================================
' Description : Test if a set of bits are set in a long
'
' Parameters  : Word        The long to test
'               Bit         The bits to compare
'
' Returns     : Boolean     Returns true if all the set bits
'                           in Bit are also set in Word.
' ==========================================================================

    Dim bRtn        As Boolean

    bRtn = ((Word And Bit) = Bit)

    BitIsSet = bRtn

End Function

Public Function HiByte(ByVal Word As Integer) As Byte
' ==========================================================================
' Description : Returns the upper byte from a Word (Integer)
'
' Parameters  : Word    The integer to parse
'
' Returns     : Byte    The most significant byte
' ==========================================================================

    Dim bytRtn  As Byte

    bytRtn = (Word And &HFF00&) \ &H100

    HiByte = bytRtn

End Function

Public Function HiWord(ByVal Dword As Long) As Integer
' ==========================================================================
' Description : Returns the upper Word (Integer) from a Dword (Long)
'
' Parameters  : Dword     The Dword to parse
'
' Returns     : Integer   The most significant 2 bytes
' ==========================================================================

    Dim lRtn    As Integer

    lRtn = (Dword And &HFFFF0000) \ &H10000

    HiWord = lRtn

End Function

Public Function LoByte(ByVal Word As Integer) As Byte
' ==========================================================================
' Description : Returns the lower byte from a Word (Integer)
'
' Parameters  : Word    The integer to parse
'
' Returns     : Byte    The least significant byte
' ==========================================================================

    Dim bytRtn  As Byte

    bytRtn = Word And &HFF

    LoByte = bytRtn

End Function

Public Function LoWord(ByVal Dword As Long) As Integer
' ==========================================================================
' Description : Returns the lower Word (Integer) from a Dword (Long)
'
' Parameters  : Dword     The Dword to parse
'
' Returns     : Integer   The least significant 2 bytes
' ==========================================================================

    Dim lRtn    As Integer

    If (Dword And &H8000&) Then
        lRtn = Dword Or &HFFFF0000
    Else
        lRtn = Dword And &HFFFF&
    End If

    LoWord = lRtn

End Function

Public Function MakeDword(ByRef HiWord As Integer, _
                          ByRef LoWord As Integer) As Long
' ==========================================================================
' Description : Combine two 2-byte Words (Integers)
'               to make a 4-byte Dword (Long)
'
' Parameters  : HiWord  The most-significant word
'               LoWord  The least-significant word
'
' Returns     : Integer
' ==========================================================================

    Dim lRtn    As Long

    lRtn = (HiWord * &H10000) Or (LoWord And &HFFFF&)

    MakeDword = lRtn

End Function

Public Function MakeWord(ByRef HiByte As Byte, _
                         ByRef LoByte As Byte) As Integer
' ==========================================================================
' Description : Combine two bytes to make a Word (Integer)
'
' Parameters  : HiByte      The most-significant byte
'               LoByte      The least-significant byte
'
'
' Returns     : Integer
' ==========================================================================

    Dim iRtn    As Integer

    If (HiByte And &H80) Then
        iRtn = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
    Else
        iRtn = (HiByte * &H100) Or LoByte
    End If

    MakeWord = iRtn

End Function

Public Function SetBit(ByVal Word As Long, _
                       ByVal Bit As Long, _
                       Optional ByVal SetFlag As Boolean = True) As Long
' ==========================================================================
' Description : Combine or remove bits from two longs
'
' Parameters  : Word        The first long to work with
'               Bit         The bits to add or remove from Word
'               SetFlag     If True (default), use logical OR.
'                           If False, use logical AND NOT.
'
' Returns     : Long
' ==========================================================================

    Const sPROC As String = "SetBit"

    Dim lRtn    As Long

    If SetFlag Then
        lRtn = (Word Or Bit)
    Else
        lRtn = (Word And Not Bit)
    End If

    SetBit = lRtn

End Function

Public Function ShiftLeft(ByVal Dword As Long, ByVal Shift As Byte) As Long
' ==========================================================================
' Description : Perform a bitwise shift left operation
'
' Parameters  : Dword   The value to modify
'               Shift   The number of bits to shift
'
' Returns     : Long
'
' Notes       : Left shifting is equal to multiplying
'               Dword by 2 to the power of Shift.
'               This routine uses a trick to avoid an overflow error.
' ==========================================================================

    Const sPROC As String = "ShiftLeft"

    Dim bytIdx  As Byte

    Dim lSave   As Long
    Dim lRtn    As Long
    Dim lWork   As Long

    ' ----------------------------------------------------------------------

    lWork = Dword

    If (Shift = 0) Then
        lRtn = Dword
        GoTo PROC_EXIT
    End If

    For bytIdx = 1 To Shift
        lSave = lWork And &H40000000    ' Save 30th bit
        lWork = lWork And &H3FFFFFFF    ' Clear 30th and 31st bits
        lWork = lWork * 2               ' Multiply by 2

        If (lSave <> 0) Then
            lWork = lWork Or &H80000000    ' Set 31st bit
        End If
    Next bytIdx

    lRtn = lWork

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ShiftLeft = lRtn

End Function

Public Function ShiftRight(ByVal Dword As Long, _
                           ByVal Shift As Byte) As Long
' ==========================================================================
' Description : Perform a bitwise shift right operation
'
' Parameters  : Dword   The value to modify
'               Shift   The number of bits to shift
'
' Returns     : Long
' ==========================================================================

    Const sPROC As String = "ShiftRight"

    Dim lRtn    As Long
    Dim lWork   As Long

    ' ----------------------------------------------------------------------

    lWork = Dword

    If (Shift = 0) Then
        lRtn = Dword
        GoTo PROC_EXIT
    End If

    lRtn = Int(lWork / (2 ^ Shift))

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ShiftRight = lRtn

End Function

#If Win64 Then
    Public Function SingleBitIsSet(ByRef Word As LongLong, _
                                   ByVal Bit As Byte) As Boolean
#Else
    Public Function SingleBitIsSet(ByRef Word As Long, _
                                   ByVal Bit As Byte) As Boolean
#End If
' ==========================================================================
' Description : Determines if a single bit is set in a long
'
' Parameters  : Word       The value to inspect
'               Bit         The zero-based bit to check.
'                           The number of valid bits
'                           depends on the platform.
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "SingleBitIsSet"

    #If Win64 Then
        Const BIT_LIMIT As Byte = 63
    #Else
        Const BIT_LIMIT As Byte = 31
    #End If

    Dim bRtn        As Boolean
    Dim lBits       As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (Bit > BIT_LIMIT) Then
        Call Err.Raise(ERR_INVALID_PROCEDURE_CALL, _
                       Concat(".", msMODULE, sPROC), _
                       "Invalid bit position.")
        GoTo PROC_EXIT
    End If

    lBits = 2 ^ Bit

    bRtn = BitIsSet(Word, lBits)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    SingleBitIsSet = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, bRtn)
    On Error GoTo 0

    Exit Function

' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume Next
    Else
        Resume PROC_EXIT
    End If

End Function
