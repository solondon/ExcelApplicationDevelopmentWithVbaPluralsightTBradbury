Attribute VB_Name = "MVBAVarTypes"
' ==========================================================================
' Module      : MVBAVarTypes
' Type        : Module
' Description : Procedures for working with VBA variable types.
' --------------------------------------------------------------------------
' Procedures  : StringToVbVarType       VbVarType
'               VbVarTypeToString       String
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

Private Const msMODULE                  As String = "MVBAVarTypes"
'                                                                             VbVarType Value
' --------------------------------------------------------------------------------------------
Public Const gsVBVARTYPE_ARRAY          As String = "Array"             ' 00100000 00000000 (8192)
Public Const gsVBVARTYPE_BOOLEAN        As String = "Boolean"           ' 00000000 00001011 (11)
Public Const gsVBVARTYPE_BYTE           As String = "Byte"              ' 00000000 00010001 (17)
Public Const gsVBVARTYPE_CURRENCY       As String = "Currency"          ' 00000000 00000110 (6)
Public Const gsVBVARTYPE_DATAOBJECT     As String = "DataObject"        ' 00000000 00001101 (13)
Public Const gsVBVARTYPE_DATE           As String = "Date"              ' 00000000 00000111 (7)
Public Const gsVBVARTYPE_DECIMAL        As String = "Decimal"           ' 00000000 00001110 (14)
Public Const gsVBVARTYPE_DOUBLE         As String = "Double"            ' 00000000 00000101 (5)
Public Const gsVBVARTYPE_EMPTY          As String = "Empty"             ' 00000000 00000000 (0)
Public Const gsVBVARTYPE_ERROR          As String = "Error"             ' 00000000 00001010 (10)
Public Const gsVBVARTYPE_INTEGER        As String = "Integer"           ' 00000000 00000010 (2)
Public Const gsVBVARTYPE_LONG           As String = "Long"              ' 00000000 00000011 (3)
Public Const gsVBVARTYPE_NULL           As String = "Null"              ' 00000000 00000001 (1)
Public Const gsVBVARTYPE_OBJECT         As String = "Object"            ' 00000000 00001001 (9)
Public Const gsVBVARTYPE_SINGLE         As String = "Single"            ' 00000000 00000100 (4)
Public Const gsVBVARTYPE_STRING         As String = "String"            ' 00000000 00001000 (8)
Public Const gsVBVARTYPE_UDT            As String = "UserDefinedType"   ' 00000000 00100100 (36)
Public Const gsVBVARTYPE_VARIANT        As String = "Variant"           ' 00000000 00001100 (12)

Public Const gsVBVARTYPE_UNKNOWN        As String = "Unknown"
Public Const gsVBVARTYPE_NOTHING        As String = "Nothing"

' Array types
' -----------
Public Const gsVBVARTYPE_BOOLEAN_ARRAY  As String = "Boolean()"
Public Const gsVBVARTYPE_BYTE_ARRAY     As String = "Byte()"
Public Const gsVBVARTYPE_CURRENCY_ARRAY As String = "Currency()"
Public Const gsVBVARTYPE_DATE_ARRAY     As String = "Date()"
Public Const gsVBVARTYPE_DECIMAL_ARRAY  As String = "Decimal()"
Public Const gsVBVARTYPE_DOUBLE_ARRAY   As String = "Double()"
Public Const gsVBVARTYPE_EMPTY_ARRAY    As String = "Empty()"
Public Const gsVBVARTYPE_INTEGER_ARRAY  As String = "Integer()"
Public Const gsVBVARTYPE_LONG_ARRAY     As String = "Long()"
Public Const gsVBVARTYPE_NULL_ARRAY     As String = "Null()"
Public Const gsVBVARTYPE_OBJECT_ARRAY   As String = "Object()"
Public Const gsVBVARTYPE_SINGLE_ARRAY   As String = "Single()"
Public Const gsVBVARTYPE_STRING_ARRAY   As String = "String()"

Public Function StringToVbVarType(ByVal VarType As String) As VbVarType
' ==========================================================================
' Description : Convert a string to an enumerated type
'
' Parameters  : VarType   The string to convert
'
' Returns     : VbVarType
' ==========================================================================

    Const sPROC As String = "StringToVbVarType"

    Dim eRtn    As VbVarType


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, VarType)

    ' ----------------------------------------------------------------------

    If (LCase$(Left$(VarType, 2)) = "vb") Then
        GoTo LONG_NAME
    End If

    ' ----------------------------------------------------------------------

SHORT_NAME:

    Select Case UCase$(VarType)
    Case UCase$(gsVBVARTYPE_ARRAY)
        eRtn = vbArray

    Case UCase$(gsVBVARTYPE_BOOLEAN)
        eRtn = vbBoolean

    Case UCase$(gsVBVARTYPE_BYTE)
        eRtn = vbByte

    Case UCase$(gsVBVARTYPE_CURRENCY)
        eRtn = vbCurrency

    Case UCase$(gsVBVARTYPE_DATAOBJECT)
        eRtn = vbDataObject

    Case UCase$(gsVBVARTYPE_DATE)
        eRtn = vbDate

    Case UCase$(gsVBVARTYPE_DECIMAL)
        eRtn = vbDecimal

    Case UCase$(gsVBVARTYPE_DOUBLE)
        eRtn = vbDouble

    Case UCase$(gsVBVARTYPE_EMPTY)
        eRtn = vbEmpty

    Case UCase$(gsVBVARTYPE_ERROR)
        eRtn = vbError

    Case UCase$(gsVBVARTYPE_INTEGER)
        eRtn = vbInteger

    Case UCase$(gsVBVARTYPE_LONG)
        eRtn = vbLong

    Case UCase$(gsVBVARTYPE_NULL)
        eRtn = vbNull

    Case UCase$(gsVBVARTYPE_OBJECT)
        eRtn = vbObject

    Case UCase$(gsVBVARTYPE_SINGLE)
        eRtn = vbSingle

    Case UCase$(gsVBVARTYPE_STRING)
        eRtn = vbString

    Case UCase$(gsVBVARTYPE_UDT)
        eRtn = vbUserDefinedType

    Case UCase$(gsVBVARTYPE_VARIANT)
        eRtn = vbVariant

    End Select

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

LONG_NAME:

    Select Case UCase$(VarType)
    Case "VB" & UCase$(gsVBVARTYPE_ARRAY)
        eRtn = vbArray

    Case "VB" & UCase$(gsVBVARTYPE_BOOLEAN)
        eRtn = vbBoolean

    Case "VB" & UCase$(gsVBVARTYPE_BYTE)
        eRtn = vbByte

    Case "VB" & UCase$(gsVBVARTYPE_CURRENCY)
        eRtn = vbCurrency

    Case "VB" & UCase$(gsVBVARTYPE_DATAOBJECT)
        eRtn = vbDataObject

    Case "VB" & UCase$(gsVBVARTYPE_DATE)
        eRtn = vbDate

    Case "VB" & UCase$(gsVBVARTYPE_DECIMAL)
        eRtn = vbDecimal

    Case "VB" & UCase$(gsVBVARTYPE_DOUBLE)
        eRtn = vbDouble

    Case "VB" & UCase$(gsVBVARTYPE_EMPTY)
        eRtn = vbEmpty

    Case "VB" & UCase$(gsVBVARTYPE_ERROR)
        eRtn = vbError

    Case "VB" & UCase$(gsVBVARTYPE_INTEGER)
        eRtn = vbInteger

    Case "VB" & UCase$(gsVBVARTYPE_LONG)
        eRtn = vbLong

    Case "VB" & UCase$(gsVBVARTYPE_NULL)
        eRtn = vbNull

    Case "VB" & UCase$(gsVBVARTYPE_OBJECT)
        eRtn = vbObject

    Case "VB" & UCase$(gsVBVARTYPE_SINGLE)
        eRtn = vbSingle

    Case "VB" & UCase$(gsVBVARTYPE_STRING)
        eRtn = vbString

    Case "VB" & UCase$(gsVBVARTYPE_UDT)
        eRtn = vbUserDefinedType

    Case "VB" & UCase$(gsVBVARTYPE_VARIANT)
        eRtn = vbVariant

    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    StringToVbVarType = eRtn

    Call Trace(tlMaximum, msMODULE, sPROC, eRtn)
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

Public Function VbVarTypeToString(ByVal VarType As VbVarType, _
                         Optional ByVal DropPrefix As Boolean) As String
' ==========================================================================
' Description : Convert an enumerated type to a string
'
' Parameters  : VarType       The enumeration to convert
'               DropPrefix    If True, remove the 'vb' prefix
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "VbVarTypeToString"

    Dim sRtn    As String


    On Error GoTo PROC_ERR
    '  Call Trace(tlMaximum, msMODULE, sPROC, VarType)

    ' ----------------------------------------------------------------------

    Select Case VarType
    Case vbArray
        sRtn = gsVBVARTYPE_ARRAY

    Case vbArray
        sRtn = gsVBVARTYPE_ARRAY

    Case vbBoolean
        sRtn = gsVBVARTYPE_BOOLEAN

    Case vbByte
        sRtn = gsVBVARTYPE_BYTE

    Case vbCurrency
        sRtn = gsVBVARTYPE_CURRENCY

    Case vbDataObject
        sRtn = gsVBVARTYPE_DATAOBJECT

    Case vbDate
        sRtn = gsVBVARTYPE_DATE

    Case vbDecimal
        sRtn = gsVBVARTYPE_DECIMAL

    Case vbDouble
        sRtn = gsVBVARTYPE_DOUBLE

    Case vbEmpty
        sRtn = gsVBVARTYPE_EMPTY

    Case vbError
        sRtn = gsVBVARTYPE_ERROR

    Case vbInteger
        sRtn = gsVBVARTYPE_INTEGER

    Case vbLong
        sRtn = gsVBVARTYPE_LONG

    Case vbNull
        sRtn = gsVBVARTYPE_NULL

    Case vbObject
        sRtn = gsVBVARTYPE_OBJECT

    Case vbSingle
        sRtn = gsVBVARTYPE_SINGLE

    Case vbString
        sRtn = gsVBVARTYPE_STRING

    Case vbUserDefinedType
        sRtn = gsVBVARTYPE_UDT

    Case vbVariant
        sRtn = gsVBVARTYPE_VARIANT

    End Select

    If DropPrefix Then
        GoTo PROC_EXIT
    End If

    sRtn = "vb" & sRtn

    ' ----------------------------------------------------------------------

PROC_EXIT:

    VbVarTypeToString = sRtn

    '  Call Trace(tlMaximum, msMODULE, sPROC, sRtn)
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
