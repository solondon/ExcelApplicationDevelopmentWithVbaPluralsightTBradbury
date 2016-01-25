Attribute VB_Name = "MMSOfficeRibbonUI"
' ==========================================================================
' Module      : MMSOfficeRibbonUI
' Type        : Module
' Description : Support for working with the Ribbon
' --------------------------------------------------------------------------
' Procedures  : ActivateRibbonTab
'               InvalidateRibbon
'               InvalidateRibbonControl
'               RibbonControlTypeToString       String
'               RibbonIsMinimized               Boolean
'               RibbonIsVisible                 Boolean
'               ShowRibbon
'               StringToRibbonControlType       enuRibbonControlType
'               ToggleRibbonVisible
' --------------------------------------------------------------------------
' References  : Microsoft Office Object Library
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

Public Const gsRXCTL_DFLT_IMAGE     As String = "HappyFace"

Public Const gsRXCTL_SIZE_NORMAL    As String = "0"
Public Const gsRXCTL_SIZE_LARGE     As String = "1"
Public Const gsRXCTL_DFLT_SIZE      As String = gsRXCTL_SIZE_NORMAL

Public Const glRXITEM_HEIGHT_SM     As Long = 16
Public Const glRXITEM_WIDTH_SM      As Long = 16
Public Const glRXITEM_HEIGHT_LG     As Long = 32
Public Const glRXITEM_WIDTH_LG      As Long = 32

Public Const glRXITEM_DFLT_HEIGHT   As Long = glRXITEM_HEIGHT_LG
Public Const glRXITEM_DFLT_WIDTH    As Long = glRXITEM_WIDTH_LG

Public Const giRXITEM_MAX_COUNT     As Integer = 1000

Public Const gbXMLNS_CUSTOMUI_ONLY  As Boolean = True

Public Const gsXMLNS_BASE           As String _
                                     = "http://schemas.microsoft.com/office/"
Public Const gsXMLNS_CUSTOMUI       As String _
                                     = gsXMLNS_BASE & "2006/01/customui"
Public Const gsXMLNS_CUSTOMUI14     As String _
                                     = gsXMLNS_BASE & "2009/07/customui"

' ----------------
' Module Level
' ----------------

Private Const msMODULE              As String = "MMSOfficeRibbonUI"

Private Const msTYPENAME_RXCTL_RBN  As String = "ribbon"
Private Const msTYPENAME_RXCTL_QAT  As String = "qat"
Private Const msTYPENAME_RXCTL_TAB  As String = "tab"
Private Const msTYPENAME_RXCTL_GRP  As String = "group"

Private Const msTYPENAME_RXCTL_LBL  As String = "labelControl"

Private Const msTYPENAME_RXCTL_BTN  As String = "button"
Private Const msTYPENAME_RXCTL_SBT  As String = "splitButton"
Private Const msTYPENAME_RXCTL_TGL  As String = "toggleButton"

Private Const msTYPENAME_RXCTL_CHK  As String = "checkBox"
Private Const msTYPENAME_RXCTL_CBO  As String = "comboBox"
Private Const msTYPENAME_RXCTL_EDT  As String = "editBox"

Private Const msTYPENAME_RXCTL_DRP  As String = "dropDown"
Private Const msTYPENAME_RXCTL_GAL  As String = "gallery"

Private Const msTYPENAME_RXCTL_MNU  As String = "menu"
Private Const msTYPENAME_RXCTL_DMN  As String = "dynamicMenu"

Private Const msTYPENAME_RXCTL_BOX  As String = "box"
Private Const msTYPENAME_RXCTL_BGP  As String = "buttonGroup"
Private Const msTYPENAME_RXCTL_DBL  As String = "dialogBoxLauncher"

Private Const msTYPENAME_RXCTL_SEP  As String = "separator"
Private Const msTYPENAME_RXCTL_MSP  As String = "menuSeparator"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuRibbonControlType
    rctUnknown = 0
    rctRibbon
    rctQAT
    rctTab
    rctGroup

    rctLabelControl

    rctButton
    rctSplitButton
    rctToggleButton

    rctCheckBox
    rctComboBox
    rctEditBox

    rctDropDown
    rctGallery

    rctMenu
    rctDynamicMenu

    rctBox
    rctButtonGroup
    rctDialogBoxLauncher

    rctSeparator
    rctMenuSeparator
End Enum

Public Sub ActivateRibbonTab(ByVal TabName As String)
' ==========================================================================
' Description : Activate a Ribbon Tab (Office 2010 and above)
'               This is a version-safe operation
'
' Parameters  : TabName   The name of the Tab to activate
' ==========================================================================

'    If (Not (goApp Is Nothing)) Then
'        If (Not (goApp.Ribbon Is Nothing)) Then
'            #If VBA7 Then
'                Call goApp.Ribbon.ActivateTab(TabName)
'            #End If
'        End If
'    End If

End Sub

Public Sub InvalidateRibbon()
' ==========================================================================
' Description : This is a 'safe' invalidate (Ribbon refresh).
'               It only functions if the needed objects exist.
' ==========================================================================

'    If (Not (goApp Is Nothing)) Then
'        If (Not (goApp.Ribbon Is Nothing)) Then
'            Call goApp.Ribbon.Invalidate
'        End If
'    End If

End Sub

Public Sub InvalidateRibbonControl(ByVal ControlID As String)
' ==========================================================================
' Description : This is a 'safe' invalidate (Control refresh).
'               It only functions if the needed objects exist.
' ==========================================================================

'    If (Not (goApp Is Nothing)) Then
'        If (Not (goApp.Ribbon Is Nothing)) Then
'            Call goApp.Ribbon.InvalidateControl(ControlID)
'        End If
'    End If

End Sub

Public Function RibbonControlTypeToString(ByVal ControlType _
                                             As enuRibbonControlType) _
       As String
' ==========================================================================
' Description : Convert an enumeration to a string
'
' Parameters  : ControlType     The enumeration to convert
'
' Returns     : String
' ==========================================================================

    Dim sRtn    As String

    ' ----------------------------------------------------------------------

    Select Case ControlType
    Case enuRibbonControlType.rctRibbon
        sRtn = msTYPENAME_RXCTL_RBN
    Case enuRibbonControlType.rctQAT
        sRtn = msTYPENAME_RXCTL_QAT
    Case enuRibbonControlType.rctTab
        sRtn = msTYPENAME_RXCTL_TAB
    Case enuRibbonControlType.rctGroup
        sRtn = msTYPENAME_RXCTL_GRP

    Case enuRibbonControlType.rctLabelControl
        sRtn = msTYPENAME_RXCTL_LBL

    Case enuRibbonControlType.rctButton
        sRtn = msTYPENAME_RXCTL_BTN
    Case enuRibbonControlType.rctSplitButton
        sRtn = msTYPENAME_RXCTL_SBT
    Case enuRibbonControlType.rctToggleButton
        sRtn = msTYPENAME_RXCTL_TGL
    
    Case enuRibbonControlType.rctCheckBox
        sRtn = msTYPENAME_RXCTL_CHK
    Case enuRibbonControlType.rctComboBox
        sRtn = msTYPENAME_RXCTL_CBO
    Case enuRibbonControlType.rctEditBox
        sRtn = msTYPENAME_RXCTL_EDT

    Case enuRibbonControlType.rctDropDown
        sRtn = msTYPENAME_RXCTL_DRP
    Case enuRibbonControlType.rctGallery
        sRtn = msTYPENAME_RXCTL_GAL

    Case enuRibbonControlType.rctMenu
        sRtn = msTYPENAME_RXCTL_MNU
    Case enuRibbonControlType.rctDynamicMenu
        sRtn = msTYPENAME_RXCTL_DMN

    Case enuRibbonControlType.rctBox
        sRtn = msTYPENAME_RXCTL_BOX
    Case enuRibbonControlType.rctButtonGroup
        sRtn = msTYPENAME_RXCTL_BGP
    Case enuRibbonControlType.rctDialogBoxLauncher
        sRtn = msTYPENAME_RXCTL_DBL

    Case enuRibbonControlType.rctSeparator
        sRtn = msTYPENAME_RXCTL_SEP
    Case enuRibbonControlType.rctMenuSeparator
        sRtn = msTYPENAME_RXCTL_MSP
    End Select

    ' ----------------------------------------------------------------------

    RibbonControlTypeToString = sRtn

End Function

Public Function RibbonIsMinimized() As Boolean
' ==========================================================================
' Description : Determine if the Ribbon is minimized
'
' Returns     : Boolean
'
' Comments    : This is a temporary implemenetation until
'               a better version implements pixel conversion
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    bRtn = (Application.CommandBars("Ribbon").Height < 75)

    ' ----------------------------------------------------------------------

    RibbonIsMinimized = bRtn

End Function

Public Function RibbonIsVisible() As Boolean
' ==========================================================================
' Description : Determines if the Ribbon is minimized
'
' Returns     : Boolean
' ==========================================================================

    Dim bRtn    As Boolean

    ' ----------------------------------------------------------------------

    bRtn = Application.CommandBars("Ribbon").Visible

    ' ----------------------------------------------------------------------

    RibbonIsVisible = bRtn

End Function

Public Sub ShowRibbon(Optional ByVal Visible As Boolean = True)
' ==========================================================================
' Description : Set the visibility of the Ribbon
'
' Parameters  : Visible     True = visible
' ==========================================================================

    Dim sMacro  As String

    sMacro = "SHOW.TOOLBAR(" _
           & Chr(34) & "Ribbon" & Chr(34) & ", " _
           & CStr(Visible) _
           & ")"

    Call Application.ExecuteExcel4Macro(sMacro)

End Sub

Public Function StringToRibbonControlType(ByVal ControlType As String) _
       As enuRibbonControlType
' ==========================================================================
' Description : Convert a string to an enumeration
'
' Parameters  : The string to convert
'
' Returns     : enuRibbonControlType
' ==========================================================================

    Dim eRtn    As enuRibbonControlType

    ' ----------------------------------------------------------------------

    Select Case ControlType
    Case msTYPENAME_RXCTL_RBN
        eRtn = rctRibbon
    Case msTYPENAME_RXCTL_QAT
        eRtn = rctQAT
    Case msTYPENAME_RXCTL_TAB
        eRtn = rctTab
    Case msTYPENAME_RXCTL_GRP
        eRtn = rctGroup

    Case msTYPENAME_RXCTL_LBL
        eRtn = rctLabelControl

    Case msTYPENAME_RXCTL_BTN
        eRtn = rctButton
    Case msTYPENAME_RXCTL_SBT
        eRtn = rctSplitButton
    Case msTYPENAME_RXCTL_TGL
        eRtn = rctToggleButton

    Case msTYPENAME_RXCTL_CHK
        eRtn = rctCheckBox
    Case msTYPENAME_RXCTL_CBO
        eRtn = rctComboBox
    Case msTYPENAME_RXCTL_EDT
        eRtn = rctEditBox

    Case msTYPENAME_RXCTL_DRP
        eRtn = rctDropDown
    Case msTYPENAME_RXCTL_GAL
        eRtn = rctGallery

    Case msTYPENAME_RXCTL_MNU
        eRtn = rctMenu
    Case msTYPENAME_RXCTL_DMN
        eRtn = rctDynamicMenu

    Case msTYPENAME_RXCTL_BOX
        eRtn = rctBox
    Case msTYPENAME_RXCTL_BGP
        eRtn = rctButtonGroup
    Case msTYPENAME_RXCTL_DBL
        eRtn = rctDialogBoxLauncher

    Case msTYPENAME_RXCTL_SEP
        eRtn = rctSeparator
    Case msTYPENAME_RXCTL_MSP
        eRtn = rctMenuSeparator

    Case Else
        eRtn = rctUnknown
    End Select

    ' ----------------------------------------------------------------------

    StringToRibbonControlType = eRtn

End Function

Public Sub ToggleRibbonVisible()
' ==========================================================================
' Description : Toggle the visibility of the ribbon
' ==========================================================================

    Dim bRtn    As Boolean

    bRtn = (Not RibbonIsVisible())
    Call ShowRibbon(bRtn)

End Sub
