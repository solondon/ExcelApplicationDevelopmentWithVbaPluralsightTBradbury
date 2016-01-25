VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFErrBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' Module      : IFErrBox
' Type        : Class Module
' Description : Define the IFErrBox interface
' --------------------------------------------------------------------------
' Properties  : Caption             (Get)  String
'               Caption             (Let)  String
'               DialogResult        (Get)  VbMsgBoxResult
'               ErrorDescription    (Let)  String
'               ErrorNumber         (Let)  Long
'               Procedure           (Let)  String
'               Tag                 (Get)  String
'               Tag                 (Let)  String
'               UserComments        (Get)  String
' --------------------------------------------------------------------------
' Procedures  : Hide
'               Show
' ==========================================================================

Public Property Let MsgBoxStyle(ByVal Style As VbMsgBoxStyle)
End Property

Public Property Get Caption() As String
End Property

Public Property Let Caption(ByVal Caption As String)
End Property

Public Property Get DialogResult() As VbMsgBoxResult
End Property

Public Property Let ErrorDescription(ByVal Description As String)
End Property

Public Property Let ErrorNumber(ByVal ErrNum As Long)
End Property

Public Property Let Procedure(ByVal Proc As String)
End Property

Public Property Get SendEmail() As Boolean
End Property

Public Property Get Tag() As String
End Property

Public Property Let Tag(ByVal Tag As String)
End Property

Public Property Get UserComments() As String
End Property

Public Sub Hide()
End Sub

Public Sub Show()
End Sub
