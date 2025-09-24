VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'クラス　MatrixRow
'Matrix行データを管理します

Private mKey As String
Private mPhase As String
Private mBundle As String
Private mVersion As String
Private mRiskIDs() As RiskID
Private mRiskIDIndex As Long  '次の配列番号

'プロパティ　key(Get)
Public Property Get key() As String
    key = mKey
End Property

'プロパティ　Phase(Get)
Public Property Get Phase() As String
    Phase = mPhase
End Property

'プロパティ　Phase(Let)
Public Property Let Phase(ByVal iValue As String)
    mPhase = iValue
    'Phaseを基にKeyを取得しておく
    mKey = Split(iValue, " ")(0)
End Property

'プロパティ　Bundle(Get)
Public Property Get Bundle() As String
    Bundle = mBundle
End Property

'プロパティ　Bundle(Let)
Public Property Let Bundle(ByVal iValue As String)
    mBundle = iValue
End Property

'プロパティ　Version(Get)
Public Property Get Version() As String
    Version = mVersion
End Property

'プロパティ　Version(Let)
Public Property Let Version(ByVal iValue As String)
    mVersion = iValue
End Property

'プロパティ　RiskIDs(Get)
Public Property Get riskIDs()
    riskIDs = mRiskIDs
End Property

'イベント　クラスの初期化
'クラスの初期化を行う
Private Sub Class_Initialize()
    '次の配列番号初期化
    mRiskIDIndex = 0
End Sub

'関数　AddRiskID
'引数　iValue　リスクID
'リスクID配列を拡張しながら追加する
Public Sub AddRiskID(ByVal iValue As RiskID)

    '配列拡張+1
    ReDim Preserve mRiskIDs(mRiskIDIndex)
    Set mRiskIDs(mRiskIDIndex) = iValue
    '次の配列番号のインクリメント
    mRiskIDIndex = mRiskIDIndex + 1

End Sub


