Attribute VB_Name = "MRibbonCmd"
' ==========================================================================
' Module      : MRibbonCmd
' Type        : Module
' Description : Support for repurposed IRibbonControl commands
' --------------------------------------------------------------------------
' Procedures  : cmd_getEnabled
'               cmd_onAction
'               tgl_getImageCmd
'               tgl_getLabelCmd
'               tgl_onActionCmd
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

Private Const msMODULE                  As String = "MRibbonCmd"

Private Const msRXCTL_TGL_ENABLED       As String = "rxtglEnabled"
Private Const msRXCTL_TGL_REPURPOSED    As String = "rxtglRepurposed"

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private mbEnabled                       As Boolean
Private mbRepurposed                    As Boolean

Public Sub cmd_getEnabled(ByRef Control As IRibbonControl, _
                          ByRef Enabled As Variant)
' ==========================================================================
' Description : Get the enabled state for an IRibbonControl
'
' Parameters  : Control     The command initiating the callback
'               Enabled     Returns the enabled state for the command
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case "FilePrintPreview"
        bRtn = mbEnabled

    Case Else
        bRtn = Enabled
    End Select

    ' ----------------------------------------------------------------------

    Enabled = bRtn

End Sub

Public Sub cmd_onAction(ByRef Control As IRibbonControl, _
                        ByRef CancelDefault As Variant)
' ==========================================================================
' Description : Repurpose the standard FilePrintPreview command
'
' Parameters  : Control         The command initiating the callback
'               CancelDefault   If True, prevents the standard action
' ==========================================================================

    Select Case Control.Id
    Case "FilePrintPreview"

        If mbRepurposed Then
            With ActiveSheet
                .PrintPreview
                .DisplayPageBreaks = False
            End With
            
            MsgBox "PrintPreview overridden."
            CancelDefault = True

        Else
            CancelDefault = False
        End If

    Case Else
        CancelDefault = False
    End Select

End Sub

Public Sub tgl_getImageCmd(ByRef Control As IRibbonControl, _
                               ByRef Image As Variant)
' ==========================================================================
' Description : Get the image for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Image       Returns the image for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case msRXCTL_TGL_REPURPOSED
        If mbRepurposed Then
            sRtn = "ReviewAcceptChange"
        Else
            sRtn = "NewOfficeDocument"
        End If

    Case msRXCTL_TGL_ENABLED
        If mbEnabled Then
            sRtn = "ReviewAcceptChange"
        Else
            sRtn = "NewOfficeDocument"
        End If
    End Select

    ' ----------------------------------------------------------------------

    Image = sRtn

End Sub

Public Sub tgl_getLabelCmd(ByRef Control As IRibbonControl, _
                           ByRef Label As Variant)
' ==========================================================================
' Description : Get the label for an IRibbonControl
'
' Parameters  : Control     The control initiating the callback
'               Label       Returns the label for the control
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case Control.Id
    Case msRXCTL_TGL_REPURPOSED
        sRtn = "Repurposed " & CStr(mbRepurposed)

    Case msRXCTL_TGL_ENABLED
        sRtn = "Enabled " & CStr(mbEnabled)
    End Select

    ' ----------------------------------------------------------------------

    Label = sRtn

End Sub

Public Sub tgl_onActionCmd(ByRef Control As IRibbonControl, _
                           ByRef Pressed As Boolean)
' ==========================================================================
' Description : Respond to an IRibbonControl action
'
' Parameters  : Control   The control initiating the callback
'               Pressed   The pressed state for the control
' ==========================================================================

    Select Case Control.Id
    Case msRXCTL_TGL_ENABLED
        mbEnabled = Pressed

    Case msRXCTL_TGL_REPURPOSED
        mbRepurposed = Pressed
    End Select

    goRibbon.Invalidate

End Sub

