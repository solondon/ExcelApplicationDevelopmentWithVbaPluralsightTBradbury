Attribute VB_Name = "MVBAStrings"
' ==========================================================================
' Module      : MVBAStrings
' Type        : Module
' Description : String manipulation functions.
' --------------------------------------------------------------------------
' Procedures  : Concat                  String
'               ExactWordInString       Boolean
'               MaxLen                  Long
'               PadL                    String
'               PadM                    String
'               PadR                    String
'               SingleSpace             String
'               StringRepeat            String
'               TrimToNull              String
' --------------------------------------------------------------------------
' Dependencies: MVBAError
'               MVBAMath
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

' The multi-line separator replaces
' embedded vbNewLine characters where
' a new line would interrupt the flow
' of the text, such as a log file.
' -----------------------------------
Public Const gsMULTILINE_SEP    As String = "¬" ' Chr(172)

' ----------------
' Module Level
' ----------------

Private Const msMODULE          As String = "MVBAStrings"

Public Function Concat(ByVal Delimiter As String, _
                       ParamArray Params() As Variant) As String
Attribute Concat.VB_Description = "Build a string from an array of items or multiple arguments."
' ==========================================================================
' Description : Build a string from an array of items or multiple arguments.
'
' Parameters  : Delimiter   A character to insert between each item.
'                           If no delimiter is needed, pass vbNullString.
'
' Returns     : String
'
' Notes       : Blank entries are ignored
' ==========================================================================

    Const sPROC     As String = "Concat"

    Dim bStarted    As Boolean

    Dim sRtn        As String
    Dim vParam      As Variant
    Dim vElement    As Variant


    On Error Resume Next

    ' ----------------------------------------------------------------------
    ' Build the string
    ' ----------------

    For Each vParam In Params

        ' Parse the array
        ' ---------------
        If IsArray(vParam) Then

            For Each vElement In vParam

                ' Add delimited item
                ' ------------------
                If ((Len(sRtn) > 0) Or bStarted) _
                And (Len(Delimiter) > 0) Then
                    sRtn = sRtn & Delimiter & CStr(vElement)
                    bStarted = True

                    ' First item
                    ' ----------
                ElseIf Len(sRtn) = 0 Then
                    sRtn = CStr(vElement)
                    bStarted = True

                    ' No delimiter
                    ' ------------
                Else
                    sRtn = sRtn & CStr(vElement)
                    bStarted = True
                End If
            Next vElement

            ' Add singleton item
            ' ------------------
        Else

            ' Add delimited item
            ' ------------------
            If (Len(sRtn) > 0) _
            And (Len(Delimiter) > 0) _
            And (Len(Trim$(CStr(vParam))) > 0) Then
                sRtn = sRtn & Delimiter & CStr(vParam)

            ' First item
            ' ----------
            ElseIf (Len(sRtn) = 0) Then
                sRtn = Trim$(CStr(vParam))

            ' No delimiter
            ' ------------
            Else
                sRtn = sRtn & Trim$(CStr(vParam))
            End If

        End If
    Next vParam

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Concat = sRtn

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

Public Function ExactWordInString(ByRef Text As String, _
                                  ByRef Word As String) As Boolean
Attribute ExactWordInString.VB_Description = "Determines if an exact string exists in another string. Case-sensitive."
' ==========================================================================
' Description : Test if an exact word exists in a string
'
' Parameters  : Text    The string to search
'               Word    The string to search for
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "ExactWordInString"

    Dim bRtn    As Boolean

  
    bRtn = " " & UCase(Text) & " " Like "*[!A-Z]" & UCase(Word) & "[!A-Z]*"
    
    ExactWordInString = bRtn

End Function

Public Function MaxLen(ParamArray Params() As Variant) As Long
Attribute MaxLen.VB_Description = "Determine the longest string from an array of items or multiple arguments."
' ==========================================================================
' Description : Determine the longest string from an
'               array of items or multiple arguments.
'
' Parameters  : Params    Multiple values to test. Each item can be
'                         a string or an array of strings to test.
'
' Returns     : Long
' ==========================================================================

    Const sPROC     As String = "MaxLen"

    Dim lRtn        As Long
    Dim vParam      As Variant
    Dim vElement    As Variant


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Test the parameters
    ' -------------------

    For Each vParam In Params

        ' Enumerate the array
        ' -------------------
        If IsArray(vParam) Then

            For Each vElement In vParam
                lRtn = MaxVal(lRtn, Len(vElement))
            Next vElement

        Else
            ' Test the element
            ' ----------------
            lRtn = MaxVal(lRtn, Len(vParam))
        End If
    Next vParam

    ' ----------------------------------------------------------------------

PROC_EXIT:

    MaxLen = lRtn

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

Public Function PadL(ByVal Str As String, _
                     ByVal Length As Long, _
            Optional ByVal PadChr As String = " ") As String
' ==========================================================================
' Description : Pad the left side of a string with a
'               specified character to a specific size.
'
' Parameters  : Str       The source string to pad
'
'               Length    The size of the padded string
'
'               PadChr    The character to pad the string with
'                         The default value is a space
'
' Returns     : String
' ==========================================================================

    Dim sRtn As String

    sRtn = Right$(String$(Length, PadChr) & Str, Length)

    PadL = sRtn

End Function

Public Function PadM(ByVal Str As String, _
                     ByVal Length As Long, _
            Optional ByVal PadChr As String = " ") As String
' ==========================================================================
' Description : Pad both sides of a string with a
'               specified character to a specific size.
'
' Parameters  : Str       The source string to pad
'
'               Length    The size of the padded string
'
'               PadChr    The character to pad the string with
'                         The default value is a space
'
' Returns     : String
' ==========================================================================

    Dim lLen    As Long
    Dim lFill   As Long

    Dim sRtn    As String

    'Determine how many characters should be on each side
    '----------------------------------------------------
    lLen = Len(Str)
    lFill = (Length - lLen) / 2

    'Build the return string
    'An extra character is added
    'in case Length is an odd number
    '-------------------------------
    sRtn = Left$(String$(lFill, PadChr) & Str & _
                 String$(lFill, PadChr) & PadChr, Length)

    PadM = sRtn

End Function

Public Function PadR(ByVal Str As String, _
                     ByVal Length As Long, _
            Optional ByVal PadChr As String = " ") As String
' ==========================================================================
' Description : Pad the right side of a string with a
'               specified character to a specific size.
'
' Parameters  : Str       The source string to pad
'
'               Length    The size of the padded string
'
'               PadChr    The character to pad the string with
'                         The default value is a space
'
' Returns     : String
' ==========================================================================

    Dim sRtn As String

    sRtn = Left$(Str & String$(Length, PadChr), Length)

    PadR = sRtn

End Function

Public Function SingleSpace(ByVal Text As String) As String
Attribute SingleSpace.VB_Description = "Convert all instances of multiple spaces within a string to single spaces."
' ==========================================================================
' Description : Convert all instances of multiple spaces within a string
'               to single spaces
'
' Parameters  : Text    The text to convert
'
' Returns     : String
' ==========================================================================

    Dim lPos    As Long

    ' Get the fist position
    ' ---------------------
    lPos = InStr(1, Text, Space(2), vbBinaryCompare)

    ' Loop and remove multiple spaces
    ' -------------------------------
    Do While (lPos > 0)
        Text = Replace(Text, Space(2), Space(1))
        lPos = InStr(1, Text, Space(2), vbBinaryCompare)
    Loop

    ' ----------------------------------------------------------------------

    SingleSpace = Text

End Function

Public Function StringRepeat(ByVal StringVal As String, _
                    Optional ByVal Repeat As Long = 2, _
                    Optional ByVal Delimiter As String = vbNullString) _
       As String
' ==========================================================================
' Description : Build a repeated string
'
' Parameters  : StringVal   The string to repeat
'               Repeat      The number of times to repeat
'               Delimiter   A delimiter to place between strings
'
' Returns     : String
' ==========================================================================

    Dim lIdx    As Long
    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    For lIdx = 1 To Repeat
        sRtn = sRtn & StringVal
        If (lIdx < Repeat) Then
            sRtn = sRtn & Delimiter
        End If
    Next lIdx

    ' ----------------------------------------------------------------------

    StringRepeat = sRtn

End Function

Public Function TrimToNull(ByRef Text As String) As String
Attribute TrimToNull.VB_Description = "Trim a string to the first detected null."
' ==========================================================================
' Description : Trim a string to the first detected null
'
' Parameters  : Text        The string to check
'
' Returns     : String
' ==========================================================================

    Dim lPos    As Long
    Dim sRtn    As String

    ' Locate the null
    ' ---------------
    lPos = InStr(1, Text, vbNullChar)

    ' If found then trim
    ' ------------------
    If lPos Then
        sRtn = Mid$(Text, 1, lPos - 1)
    Else
        sRtn = Trim$(Text)
    End If

    ' ----------------------------------------------------------------------

    TrimToNull = sRtn

End Function
