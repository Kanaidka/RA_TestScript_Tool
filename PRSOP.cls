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
'クラス　PRSOP
'PRSのOPデータを管理します

Private mText As String
Private mID As String
Private mProcessSequence As String
Private mComment As String
Private mPHs() As PRSPH
Private mPHIndex As Long

'プロパティ　Text(Get)
Public Property Get text() As String
    text = mText
End Property

'プロパティ　Text(Let)
Public Property Let text(ByVal iValue As String)
    mText = iValue
End Property

'プロパティ　ID(Get)
Public Property Get id() As String
    id = mID
End Property

'プロパティ　ID(Let)
Public Property Let id(ByVal iValue As String)
    mID = iValue
End Property

'プロパティ　Comment(Get)
Public Property Get Comment() As String
    Comment = mComment
End Property

'プロパティ　Comment(Let)
Public Property Let Comment(ByVal iValue As String)
    mComment = iValue
End Property

'プロパティ　ProcessSequence(Get)
Public Property Get ProcessSequence() As String
    ProcessSequence = mProcessSequence
End Property

'プロパティ　ProcessSequence(Let)
Public Property Let ProcessSequence(ByVal iValue As String)
    mProcessSequence = iValue
End Property

'プロパティ　PHs(Get)
Public Property Get PHs()
    PHs = mPHs
End Property

'関数　AddPH
'引数　iPH　対象PH
'PH情報を追加する
Public Sub AddPH(ByVal iPH As PRSPH)
    ReDim Preserve mPHs(mPHIndex)
    Set mPHs(mPHIndex) = iPH
    mPHIndex = mPHIndex + 1
End Sub


