VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRSOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mText As String
Private mID As String
Private mComment As String
Private mPHs() As PRSPH
Private mPHIndex As Long

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

Public Property Get Comment() As String
    Comment = mComment
End Property

Public Property Let Comment(ByVal iValue As String)
    mComment = iValue
End Property

Public Property Get PHs()
    PHs = mPHs
End Property

Public Sub AddPH(ByVal iPH As PRSPH)
    ReDim Preserve mPHs(mPHIndex)
    Set mPHs(mPHIndex) = iPH
    mPHIndex = mPHIndex + 1
End Sub


