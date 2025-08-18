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

Private mKey As String
Private mPhase As String
Private mBundle As String
Private mVersion As String
Private mRiskIDs() As RiskID
Private mRiskIDIndex As Long

Public Property Get Key() As String
    Key = mKey
End Property

Public Property Get Phase() As String
    Phase = mPhase
End Property

Public Property Let Phase(ByVal iValue As String)
    mPhase = iValue
    mKey = Split(iValue, " ")(0)
End Property

Public Property Get Bundle() As String
    Bundle = mBundle
End Property

Public Property Let Bundle(ByVal iValue As String)
    mBundle = iValue
End Property

Public Property Get Version() As String
    Version = mVersion
End Property

Public Property Let Version(ByVal iValue As String)
    mVersion = iValue
End Property

Public Property Get RiskIDs()
    RiskIDs = mRiskIDs
End Property

Private Sub Class_Initialize()
    mRiskIDIndex = 0
End Sub

Public Sub AddRiskID(ByVal iValue As RiskID)

    ReDim Preserve mRiskIDs(mRiskIDIndex)
    Set mRiskIDs(mRiskIDIndex) = iValue
    mRiskIDIndex = mRiskIDIndex + 1

End Sub


