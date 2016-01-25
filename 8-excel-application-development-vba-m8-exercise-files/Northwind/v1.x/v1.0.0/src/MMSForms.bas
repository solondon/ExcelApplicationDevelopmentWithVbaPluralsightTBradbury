Attribute VB_Name = "MMSForms"
' ==========================================================================
' Module      : MMSForms
' Type        : Module
' Description : General-purpose support for UserForms
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

Public Const gsFORMS_PROGID_CHECKBOX        As String = "Forms.CheckBox.1"
Public Const gsFORMS_PROGID_COMBOBOX        As String = "Forms.ComboBox.1"
Public Const gsFORMS_PROGID_COMMANDBUTTON   As String = "Forms.CommandButton.1"
Public Const gsFORMS_PROGID_FRAME           As String = "Forms.Frame.1"
Public Const gsFORMS_PROGID_IMAGE           As String = "Forms.Image.1"
Public Const gsFORMS_PROGID_LABEL           As String = "Forms.Label.1"
Public Const gsFORMS_PROGID_LISTBOX         As String = "Forms.ListBox.1"
Public Const gsFORMS_PROGID_MULTIPAGE       As String = "Forms.MultiPage.1"
Public Const gsFORMS_PROGID_OPTIONBUTTON    As String = "Forms.OptionButton.1"
Public Const gsFORMS_PROGID_SCROLLBAR       As String = "Forms.ScrollBar.1"
Public Const gsFORMS_PROGID_SPINBUTTON      As String = "Forms.SpinButton.1"
Public Const gsFORMS_PROGID_TABSTRIP        As String = "Forms.TabStrip.1"
Public Const gsFORMS_PROGID_TEXTBOX         As String = "Forms.TextBox.1"
Public Const gsFORMS_PROGID_TOGGLEBUTTON    As String = "Forms.ToggleButton.1"

' Common control
' ------------------------------

Public Const CC_SCROLLBAR_WIDTH             As Single = 18

' ----------------
' Module Level
' ----------------

Private Const msMODULE                      As String = "MMSForms"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuMouseButton
    mbNoButton = 0          ' No button is pressed.
    mbPrimaryButton = 1     ' The primary button is pressed.
    mbSecondaryButton = 2   ' The secondary button is pressed.
    mbMiddleButton = 4      ' The middle button is pressed.
End Enum

Public Enum enuMouseShift
    msNoShift = 0
    msShift = 1             ' SHIFT was pressed.
    msCtrl = 2              ' CTRL was pressed.
    msAlt = 4               ' ALT was pressed.
End Enum

Public Enum enuStartUpPosition
    supManual = 0           ' No initial setting specified.
    supCenterOwner = 1      ' Center on the UserForm owner.
    supCenterScreen = 2     ' Center on the whole screen.
    supWindowsDefault = 3   ' Position in upper-left corner of screen.
End Enum

' -----------------------------------
' Type declarations
' -----------------------------------
' Global Level
' ----------------

' POINT is a reserved Windows API structure
' (in pixels) so the name has been altered.
' This is to identify a position in points.
' -----------------------------------------
Public Type TVBPoint
    X                                       As Single
    Y                                       As Single
End Type

' RECT is a reserved Windows API structure
' (in pixels) so the name has been altered.
' This is to identify the bounds of a shape.
' ------------------------------------------
Public Type TVBRect
    Left                                    As Single
    Top                                     As Single
    Width                                   As Single
    Height                                  As Single
End Type
