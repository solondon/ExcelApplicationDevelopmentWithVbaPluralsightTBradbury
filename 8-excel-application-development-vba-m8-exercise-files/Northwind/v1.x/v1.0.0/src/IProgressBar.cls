VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Define the IProgressBar interface."
' ==========================================================================
' Module      : IProgressBar
' Type        : Class Module
' Description : Define the ProgressBar interface.
' --------------------------------------------------------------------------
' Properties  : Canceled            (Get)  Boolean
'               Canceled            (Let)  Boolean
'               CancelVisible       (Get)  Boolean
'               CancelVisible       (Let)  Boolean
'               Caption             (Get)  String
'               Caption             (Let)  String
'               ChangeRate          (Get)  Double
'               ChangeRate          (Let)  Double
'               Max                 (Get)  Long
'               Max                 (Let)  Long
'               Min                 (Get)  Long
'               Min                 (Let)  Long
'               OverallCaption      (Get)  String
'               OverallCaption      (Let)  String
'               OverallChangeRate   (Get)  Double
'               OverallChangeRate   (Let)  Double
'               OverallMax          (Get)  Long
'               OverallMax          (Let)  Long
'               OverallMin          (Get)  Long
'               OverallMin          (Let)  Long
'               OverallPercent      (Get)  Double
'               OverallValue        (Get)  Long
'               OverallValue        (Let)  Long
'               OverallVisible      (Get)  Boolean
'               OverallVisible      (Let)  Boolean
'               Percent             (Get)  Double
'               Title               (Get)  String
'               Title               (Let)  String
'               Value               (Get)  Long
'               Value               (Let)  Long
' --------------------------------------------------------------------------
' Procedures  : Complete
'               Decrement
'               Hide
'               Increment
'               OverallComplete
'               OverallDecrement
'               OverallIncrement
'               OverallReset
'               Reset
'               Show
' ==========================================================================

Public Property Get Canceled() As Boolean
End Property

Public Property Let Canceled(ByVal Canceled As Boolean)
End Property

Public Property Get CancelVisible() As Boolean
End Property

Public Property Let CancelVisible(ByVal Visible As Boolean)
End Property

Public Property Get Caption() As String
End Property

Public Property Let Caption(ByVal Caption As String)
End Property

Public Property Get ChangeRate() As Double
End Property

Public Property Let ChangeRate(ByVal Rate As Double)
End Property

Public Property Get Max() As Long
End Property

Public Property Let Max(ByVal MaxValue As Long)
End Property

Public Property Get Min() As Long
End Property

Public Property Let Min(ByVal MinValue As Long)
End Property

Public Property Get OverallCaption() As String
End Property

Public Property Let OverallCaption(ByVal Caption As String)
End Property

Public Property Get OverallChangeRate() As Double
End Property

Public Property Let OverallChangeRate(ByVal Rate As Double)
End Property

Public Property Get OverallMax() As Long
End Property

Public Property Let OverallMax(ByVal MaxValue As Long)
End Property

Public Property Get OverallMin() As Long
End Property

Public Property Let OverallMin(ByVal MinValue As Long)
End Property

Public Property Get OverallPercent() As Double
End Property

Public Property Get OverallValue() As Long
End Property

Public Property Let OverallValue(ByVal Value As Long)
End Property

Public Property Get OverallVisible() As Boolean
End Property

Public Property Let OverallVisible(ByVal Visible As Boolean)
End Property

Public Property Get Percent() As Double
End Property

Public Property Get Title() As String
End Property

Public Property Let Title(ByVal Title As String)
End Property

Public Property Get Value() As Long
End Property

Public Property Let Value(ByVal Value As Long)
End Property

Public Sub Complete()
End Sub

Public Sub Decrement()
End Sub

Public Sub Hide()
End Sub

Public Sub Increment()
End Sub

Public Sub OverallComplete()
End Sub

Public Sub OverallDecrement()
End Sub

Public Sub OverallIncrement()
End Sub

Public Sub OverallReset()
End Sub

Public Sub Refresh()
End Sub

Public Sub Reset()
End Sub

Public Sub Show()
End Sub
