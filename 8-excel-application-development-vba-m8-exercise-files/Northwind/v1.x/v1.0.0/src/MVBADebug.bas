Attribute VB_Name = "MVBADebug"
' ==========================================================================
' Module      : MVBADebug
' Type        : Module
' Description : Support for application debugging
' --------------------------------------------------------------------------
' Procedures  : DebugAssert
'               DebugPrint
'               Trace
'               TraceParams
'               Trap
' --------------------------------------------------------------------------
' Dependencies: MVBABitwise
'               MVBAError
'               MVBALogFile
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuAssertOutput
    aoNull = 0
    aoImmediate = 16
    aoLogFile = 32
    aoMsgBox = 128
End Enum

Public Enum enuTraceLevel
    tlOff = 0
    tlMinimum = 1
    tlNormal = 2
    tlVerbose = 4
    tlMaximum = 8
End Enum

Public Enum enuTraceOutput
    toNull = 0
    toImmediate = 16
    toLogFile = 32
End Enum

Public Enum enuTraceStackOperation
    tsoUnknown = 0
    tsoPush
    tsoPop
End Enum

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

Public Const gbDEBUG_MODE       As Boolean = False

Public Const glASSERT_OUTPUT    As Long = enuAssertOutput.aoImmediate _
                                       Or enuAssertOutput.aoMsgBox

Public Const glTRACE_LEVEL      As Long = enuTraceLevel.tlMaximum
Public Const glTRACE_OUTPUT     As Long = enuTraceOutput.toNull
Public Const gbTRACE_STACK      As Boolean = False

Public Const gsPROC_ENTER       As String = "PROC_ENTER"
Public Const gsPROC_EXIT        As String = "PROC_EXIT"

Public Sub DebugAssert(ByVal Expression As Boolean, _
                       ByVal Module As String, _
                       ByVal Procedure As String, _
                       ByVal Message As String)
' ==========================================================================
' Description : Enhanced assertion routine.
'
' Parameters  : Expression  A boolean expression to test.
'               Module      The name of the calling module.
'               Source      The name of the calling procedure.
'               Message     The message to display if Expression is false.
' ==========================================================================

    Dim sTitle      As String: sTitle = gsAPP_NAME
    Dim sPrompt     As String
    Dim eButtons    As VbMsgBoxStyle: eButtons = vbCritical Or vbOKOnly
    Dim eMBR        As VbMsgBoxResult

    ' ----------------------------------------------------------------------
    ' Quit if not debugging
    ' ---------------------
    If (Not gbDEBUG_MODE) Then
        GoTo PROC_EXIT
    End If

    ' Quit if the assertion is True
    ' -----------------------------
    If Expression Then
        GoTo PROC_EXIT
    End If

    ' To Immediate Window
    ' -------------------
    If BitIsSet(glASSERT_OUTPUT, aoImmediate) Then
        sPrompt = "ASSERTION FAILURE (" & Now() & ")"

        Call DebugPrint(Module, Procedure, sPrompt, True)
        Call DebugPrint(Module, Procedure, Message)
    End If

    ' To Error log
    ' ------------
    If BitIsSet(glASSERT_OUTPUT, aoLogFile) Then
        sPrompt = "ASSERTION FAILED" & vbNewLine _
                  & Message
        Call LogFileEntry(lftError, _
                          Module, _
                          Procedure, _
                          sPrompt, _
                          ERR_INVALID_ASSERTION)
    End If

    ' To MsgBox
    ' ---------
    If BitIsSet(glASSERT_OUTPUT, aoMsgBox) Then
        sPrompt = "Assertion failure in " & vbNewLine _
                & Concat(".", Module, Procedure) & "." & vbNewLine _
                & Message
        eMBR = MsgBox(sPrompt, eButtons, sTitle)
    End If

    ' Stop code
    ' ---------
    Debug.Assert (Expression)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Exit Sub

End Sub

Public Static Sub DebugPrint(ByVal Module As String, _
                             ByVal Procedure As String, _
                             ByVal Message As String, _
                    Optional ByVal Reset As Boolean)
' ==========================================================================
' Description : Provide enhanced printing to the Immediate window.
'
' Parameters  : Module      The name of the calling module.
'               Procedure   The name of the calling procedure.
'               Message     The message to display.
'               Reset       Force the static variables to be reset
' ==========================================================================

    Dim ssMod   As String
    Dim ssPrc   As String

    ' Only do this if there is
    ' something new to display
    ' ------------------------
    If ((Module <> ssMod) Or (Procedure <> ssPrc) Or Reset) Then

        ' Store the new values
        ' --------------------
        ssMod = Module
        ssPrc = Procedure

        ' Add a blank line
        ' ----------------
        Debug.Print

        ' Display the new source
        ' ----------------------
        If (Len(Trim$(ssPrc)) > 0) Then
            Debug.Print "Src: " & Concat(".", ssMod, ssPrc)
        End If
    End If

    ' Display the message
    ' -------------------
    If (Len(Trim(Message)) > 0) Then
        Debug.Print "Msg: " & Message
    End If

End Sub

Public Sub Trace(ByVal TraceLevel As enuTraceLevel, _
                 ByVal Module As String, _
                 ByVal Procedure As String, _
        Optional ByVal Message As Variant = vbNullString, _
        Optional ByVal StackOperation As enuTraceStackOperation)
' ==========================================================================
' Description : Provide code-tracing capability.
'
' Parameters  : TraceLevel      Indicates the trace level to use.
'               Module          The name of the calling module.
'               Procedure       The name of the calling procedure.
'               Message         The message to display or log.
'               StackOperation  Simulate a stack construct.
'                               If used, the calling procedure must include
'                               a push at the beginning of the call, and a
'                               pop at the end to balance the stack depth.
'                               This stack depth can be used with the
'                               resulting trace log to show a hierarchy of
'                               nested procedural calls to aid in debugging.
'
' Notes       : This procedure does not use nor implement a
'               true stack, as the members are not maintained.
'               This only simulates a stack for the purposes of maintaining
'               the call depth of the procedure(s) being traced.
'
'               The Trace method only supports output
'               to the Immediate Window or a log file.
' ==========================================================================

    Static slCallDepth As Long

'    Stop

    ' Quit if tracing is turned off or the trace
    ' mode setting is higher then the global value
    ' --------------------------------------------
    If (TraceLevel > glTRACE_LEVEL) Then
        GoTo PROC_EXIT
    End If

    If (StackOperation = tsoPush) Then
        slCallDepth = slCallDepth + 1
    End If

    If BitIsSet(glTRACE_OUTPUT, toImmediate) Then
        Call DebugPrint(Module, Procedure, CStr(Message))
    End If

    If BitIsSet(glTRACE_OUTPUT, toLogFile) Then
        Call LogFileEntry(lftTrace, _
                          Module, _
                          Procedure, _
                          CStr(Message), _
                          slCallDepth)
    End If

    If (StackOperation = tsoPop) Then
        slCallDepth = slCallDepth - 1
    End If

    ' ------------------------------------------------------------------------

PROC_EXIT:

    Exit Sub

End Sub

Public Sub TraceParams(ByVal TraceLevel As enuTraceLevel, _
                       ByVal Module As String, _
                       ByVal Procedure As String, _
                       ParamArray Params() As Variant)
' ==========================================================================
' Description : Provide tracing support for parameter arrays.
'
' Parameters  : TraceLevel  Indicates the trace level to use.
'               Module      The name of the calling module.
'               Procedure   The name of the calling procedure.
'               ParamArray  The parameters to trace.
' ==========================================================================

    Dim lIdx    As Long
    Dim lLB     As Long
    Dim lUB     As Long

    Dim vParam  As Variant

    ' Determine the extent of the parameters
    ' --------------------------------------
    lLB = LBound(Params)
    lUB = UBound(Params)

    ' Call Trace for each parameter
    ' -----------------------------
    For lIdx = lLB To lUB

        If IsArray(Params(lIdx)) Then

            For Each vParam In Params(lIdx)
                Call Trace(TraceLevel, Module, Procedure, vParam)
            Next vParam

        Else
            vParam = Params(lIdx)
            Call Trace(TraceLevel, Module, Procedure, vParam)
        End If

    Next lIdx

    On Error Resume Next
    Erase vParam
    vParam = Empty

End Sub

Public Sub Trap(ByVal Module As String, _
                ByVal Procedure As String, _
                ByVal Message As String)
' ==========================================================================
' Description : Provide Trap support for testing code coverage.
'
' Parameters  : Module      The name of the calling module.
'               Procedure   The name of the calling procedure.
'               Message     The message to display.
' ==========================================================================

    If gbDEBUG_MODE Then
        Call DebugPrint(Module, Procedure, Message)
        Stop
    End If

End Sub
