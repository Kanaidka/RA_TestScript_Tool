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
'クラス　PRSPR
'PRSのPRデータを管理します

Private mText As String
Private mUPs() As PRSUP
Private mUPIndex As Long

'プロパティ　Text(Get)
Public Property Get text() As String
    text = mText
End Property

'プロパティ　Text(Let)
Public Property Let text(ByVal iValue As String)
    mText = iValue
End Property

'プロパティ　UPs(Get)
Public Property Get UPs()
    UPs = mUPs
End Property

'関数　AddUP
'引数　iUP　対象UP
'UP情報を追加する
Public Sub AddUP(ByVal iUP As PRSUP)
    ReDim Preserve mUPs(mUPIndex)
    Set mUPs(mUPIndex) = iUP
    mUPIndex = mUPIndex + 1
End Sub
