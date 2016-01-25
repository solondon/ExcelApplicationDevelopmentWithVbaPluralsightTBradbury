Attribute VB_Name = "MVBAMsgBox"
' ==========================================================================
' Module      : MVBAMsgBox
' Type        : Module
' Description : Support for working with message boxes
' --------------------------------------------------------------------------
' Procedures  : ParseVbMsgBoxStyle
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

Public Enum enuVBMsgBoxStylePart
    mbspButtons
    mbspIcon
    mbspDefaultButton
    mbspModality
    mbspHelpButton
    mbspForeground
    mbspRight
    mbspRTLReading
End Enum

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE As String = "MVBAMsgBox"

Public Function ParseVbMsgBoxStyle(ByVal Style As VbMsgBoxStyle, _
                                   ByVal StylePart _
                                      As enuVBMsgBoxStylePart) _
       As VbMsgBoxStyle
' ==========================================================================
' Description : Parse a component value from a composite style
'
' Parameters  : Style           The combined style elements
'               StylePart       Identifies which part to parse out
'
' Returns     : VbMsgBoxStyle
' ==========================================================================

    Const BITMASK_BUTTONS       As Long = vbOKOnly _
                                       Or vbOKCancel _
                                       Or vbAbortRetryIgnore _
                                       Or vbRetryCancel _
                                       Or vbYesNo _
                                       Or vbYesNoCancel
    Const BITMASK_DEFBTN        As Long = vbDefaultButton1 _
                                       Or vbDefaultButton2 _
                                       Or vbDefaultButton3 _
                                       Or vbDefaultButton4
    Const BITMASK_FOREGROUND    As Long = vbMsgBoxHelpButton
    Const BITMASK_HELP          As Long = vbMsgBoxHelpButton
    Const BITMASK_ICON          As Long = vbCritical _
                                       Or vbQuestion _
                                       Or vbExclamation _
                                       Or vbInformation
    Const BITMASK_MODAL         As Long = vbApplicationModal _
                                       Or vbSystemModal
    Const BITMASK_RIGHT         As Long = vbMsgBoxRight
    Const BITMASK_RTLREADING    As Long = vbMsgBoxRtlReading

    Dim eRtn                    As VbMsgBoxStyle

    ' ----------------------------------------------------------------------

    Select Case StylePart
    Case mbspButtons
        eRtn = (Style And BITMASK_BUTTONS)

    Case mbspIcon
        eRtn = (Style And BITMASK_ICON)

    Case mbspDefaultButton
        eRtn = (Style And BITMASK_DEFBTN)

    Case mbspModality
        eRtn = (Style And BITMASK_MODAL)

    Case mbspHelpButton
        eRtn = (Style And BITMASK_HELP)

    Case mbspForeground
        eRtn = (Style And BITMASK_FOREGROUND)

    Case mbspRight
        eRtn = (Style And BITMASK_RIGHT)

    Case mbspRTLReading
        eRtn = (Style And BITMASK_RTLREADING)
    End Select

    ' ----------------------------------------------------------------------

    ParseVbMsgBoxStyle = eRtn

End Function
