VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFLightbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' Module      : IFLightbox
' Type        : Class Module
' Description : Define the FLightbox interface
' --------------------------------------------------------------------------
' Properties  : FadeIn      (Get)   Boolean
'               FadeIn      (Let)   Boolean
'               FadeOut     (Get)   Boolean
'               FadeOut     (Let)   Boolean
'               FadeSpeed   (Get)   Byte
'               FadeSpeed   (Let)   Byte
'               Opacity     (Get)   Byte
'               Opacity     (Let)   Byte
'               Tag         (Get)   String
'               Tag         (Let)   String
' --------------------------------------------------------------------------
' Procedures  : Hide
'               Repaint
'               Show
' ==========================================================================

Public Property Get FadeIn() As Boolean
End Property

Public Property Let FadeIn(ByVal bFadeIn As Boolean)
' Turn the fade-in effect on or off
End Property

Public Property Get FadeOut() As Boolean
End Property

Public Property Let FadeOut(ByVal bFadeOut As Boolean)
' Turn the fade-out effect on or off
End Property

Public Property Get FadeSpeed() As Byte
End Property

Public Property Let FadeSpeed(ByVal bytFadeSpeed As Byte)
' Set the speed of the fade effect
End Property

Public Property Get Opacity() As Byte
End Property

Public Property Let Opacity(ByVal Opacity As Byte)
' Set the transparency of the effect
End Property

Public Property Get Tag() As String
End Property

Public Property Let Tag(ByVal Tag As String)
End Property

Public Sub Hide()
End Sub

Public Sub Repaint()
End Sub

Public Sub Show()
End Sub
