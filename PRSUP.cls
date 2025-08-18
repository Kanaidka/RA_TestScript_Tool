VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRSUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mText As String
Private mID As String
Private mOPs() As PRSOP
Private mOPIndex As Long

Public Property Get Text() As String
    Text = mText
End Property

Public Property Let Text(ByVal iValue As String)
    mText = iValue
End Property

Public Property Get ID() As String
    ID = mID
End Property

Public Property Let ID(ByVal iValue As String)
    mID = iValue
End Property

Public Property Get OPs()
    OPs = mOPs
End Property

Public Sub AddOP(ByVal iOp As PRSOP)
    ReDim Preserve mOPs(mOPIndex)
    Set mOPs(mOPIndex) = iOp
    mOPIndex = mOPIndex + 1
End Sub

