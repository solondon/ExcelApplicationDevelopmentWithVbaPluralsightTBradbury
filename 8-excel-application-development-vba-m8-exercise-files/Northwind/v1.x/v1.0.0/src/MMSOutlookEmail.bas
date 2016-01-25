Attribute VB_Name = "MMSOutlookEmail"
' ==========================================================================
' Module      : MMSOutlookEmail
' Type        : Module
' Description : Procedures to create Outlook email from other applications.
' --------------------------------------------------------------------------
' Procedures  : SendEmailOutlook
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE As String = "MMSOutlookEmail"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuOutlookBodyFormat
    olbfUnspecified = 0
    olbfPlain = 1
    olbfHTML = 2
    olbfRichText = 3
End Enum

Public Sub SendEmailOutlook(ByVal SendTo As String, _
                            ByVal Subject As String, _
                            ByVal Body As Variant, _
                   Optional ByVal BodyFormat _
                               As enuOutlookBodyFormat = olbfPlain, _
                   Optional ByVal Attachment As Variant, _
                   Optional ByVal SendCC As String, _
                   Optional ByVal SendBCC As String)
' ==========================================================================
' Description : Send an email message using default Outlook settings
'
' Parameters  : SendTo      The email address of the recipient(s)
'               Subject     The subject line of the message
'               Body        The text of the message
'               BodyFormat  Select between plain, RTF and HTML
'               Attachment  The name of the file(s) to attach
'                           Use a variant array to send multiple files
'               SendCC      The address(es) of additional recipients
'               SendBCC     The address(es) of BCCs
'
' Comments    : While this method requires Outlook to be installed,
'               it does not need to be running for it to work.
' ==========================================================================

    Const sPROC     As String = "SendEmailOutlook"

    Dim sTitle      As String: sTitle = gsAPP_NAME
    Dim sPrompt     As String
    Dim eButtons    As VbMsgBoxStyle: eButtons = vbInformation _
                                              Or vbYesNo _
                                              Or vbDefaultButton2
    Dim eMBR        As VbMsgBoxResult

    Dim vItem       As Variant

    Dim objMailApp  As Object
    Dim objMailItem As Object


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Create the objects
    ' using late binding
    ' ------------------
    Set objMailApp = CreateObject("Outlook.Application")
    Set objMailItem = objMailApp.CreateItem(0)

    ' Set the properties and send
    ' ---------------------------
    With objMailItem
        .To = SendTo
        .Subject = Subject

        .BodyFormat = BodyFormat
        If (BodyFormat = olbfHTML) Then
            .HTMLBody = Body
        ElseIf (BodyFormat = olbfRichText) Then
            .RTFBody = Body
        Else
            .Body = Body
        End If

        If IsArray(Attachment) Then
            For Each vItem In Attachment
                If FileExists(vItem) Then
                    .Attachments.Add vItem
                Else
                    sPrompt = "The attachment " _
                            & Attachment & vbNewLine _
                            & "cannot be found." & vbNewLine _
                            & "Do you wish to continue?"
                    eMBR = MsgBox(sPrompt, eButtons, sTitle)
                    If (eMBR = vbNo) Then
                        GoTo PROC_EXIT
                    End If
                End If
            Next vItem

        ElseIf (Not IsMissing(Attachment)) Then
            If (Len(Attachment) > 0) Then
                If FileExists(Attachment) Then
                    .Attachments.Add Attachment
                Else
                    sPrompt = "The attachment " _
                            & Attachment & vbNewLine _
                            & "cannot be found." & vbNewLine _
                            & "Do you wish to continue?"
                    eMBR = MsgBox(sPrompt, eButtons, sTitle)
                    If (eMBR = vbNo) Then
                        GoTo PROC_EXIT
                    End If
                End If
            End If
        End If

        If (Len(SendCC) > 0) Then
            .CC = SendCC
        End If

        If (Len(SendBCC) > 0) Then
            .BCC = SendBCC
        End If

        .Send
    End With

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set objMailApp = Nothing
    Set objMailItem = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC, True) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub
