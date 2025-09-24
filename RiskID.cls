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
'クラス　RiskID
'RiskIDデータを管理します

Private mID As String
Private mText As String

'プロパティ　ID(Get)
Public Property Get id()
    id = mID
End Property

'プロパティ　Text(Get)
Public Property Get text()
    text = mText
End Property

'関数　Init
'引数　iID　ID
'引数　iText　Text
'初期化処理
Public Sub Init(ByVal iID As String, ByVal iText As String)
    mID = iID
    mText = iText
End Sub
