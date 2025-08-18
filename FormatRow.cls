VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPhase As String
Private mRequired As String
Private mPRSReference As String
Private mData_Prerequisites As String
Private mTestInstruction As String
Private mExpectedResult As String
Private mTestResult As String
Private mEvidence As String
Private mRiskID As String

Public Property Get Phase() As String
    Phase = mPhase
End Property

Public Property Let Phase(ByVal iValue As String)
    mPhase = iValue
End Property

Public Property Get Required() As String
    Required = mRequired
End Property

Public Property Let Required(ByVal iValue As String)
    mRequired = iValue
End Property

Public Property Get PRSReference() As String
    PRSReference = mPRSReference
End Property

Public Property Let PRSReference(ByVal iValue As String)
    mPRSReference = iValue
End Property

Public Property Get Data_Prerequisites() As String
    Data_Prerequisites = mData_Prerequisites
End Property

Public Property Let Data_Prerequisites(ByVal iValue As String)
    mData_Prerequisites = iValue
End Property

Public Property Get TestInstruction() As String
    TestInstruction = mTestInstruction
End Property

Public Property Let TestInstruction(ByVal iValue As String)
    mTestInstruction = iValue
End Property

Public Property Get ExpectedResult() As String
    ExpectedResult = mExpectedResult
End Property

Public Property Let ExpectedResult(ByVal iValue As String)
    mExpectedResult = iValue
End Property

Public Property Get TestResult() As String
    TestResult = mTestResult
End Property

Public Property Let TestResult(ByVal iValue As String)
    mTestResult = iValue
End Property

Public Property Get Evidence() As String
    Evidence = mEvidence
End Property

Public Property Let Evidence(ByVal iValue As String)
    mEvidence = iValue
End Property

Public Property Get RiskID() As String
    RiskID = mRiskID
End Property

Public Property Let RiskID(ByVal iValue As String)
    mRiskID = iValue
End Property
