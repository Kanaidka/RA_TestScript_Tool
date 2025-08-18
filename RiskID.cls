VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiskID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mID As String
Private mText As String

Public Property Get ID()
    ID = mID
End Property

Public Property Get Text()
    Text = mText
End Property

Public Sub Init(ByVal iID As String, ByVal iText As String)
    mID = iID
    mText = iText
End Sub
