VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRSPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mText As String
Private mUPs() As PRSUP
Private mUPIndex As Long

Public Property Get Text() As String
    Text = mText
End Property

Public Property Let Text(ByVal iValue As String)
    mText = iValue
End Property

Public Property Get UPs()
    UPs = mUPs
End Property

Public Sub AddUP(ByVal iUp As PRSUP)
    ReDim Preserve mUPs(mUPIndex)
    Set mUPs(mUPIndex) = iUp
    mUPIndex = mUPIndex + 1
End Sub
