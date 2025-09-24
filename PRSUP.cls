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
'クラス　PRSUP
'PRSのUPデータを管理します

Private mText As String
Private mID As String
Private mOPs() As PRSOP
Private mOPIndex As Long

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

'プロパティ　OPs(Get)
Public Property Get OPs()
    OPs = mOPs
End Property

'関数　AddUP
'引数　iOP　対象OP
'OP情報を追加する
Public Sub AddOP(ByVal iOP As PRSOP)
    ReDim Preserve mOPs(mOPIndex)
    Set mOPs(mOPIndex) = iOP
    mOPIndex = mOPIndex + 1
End Sub

