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
'クラス　FormatRow
'Format行データを管理します

Private mPhase() As String
Private mRequired As String
Private mSort As String
Private mSortNo As Long
Private mPRSReference As String
Private mData_Prerequisites As String
Private mTestInstruction As String
Private mExpectedResult As String
Private mTestResult As String
Private mEvidence As String
Private mRiskID As String

'プロパティ　Phase(Get)
Public Property Get Phase() As String
    Phase = Join(mPhase, ",")
End Property

'プロパティ　Phase(Let)
Public Property Let Phase(ByVal iValue As String)
    mPhase = Split(Replace(iValue, vbCrLf, vbLf), vbLf)
End Property

'プロパティ　Required(Get)
Public Property Get Required() As String
    Required = mRequired
End Property

'プロパティ　Required(Let)
Public Property Let Required(ByVal iValue As String)
    mRequired = iValue
End Property

'プロパティ　Sort(Get)
Public Property Get SortNo() As Long
    SortNo = mSortNo
End Property

'プロパティ　Sort(Get)
Public Property Get Sort() As String
    mSort = mRequired
End Property

'プロパティ　Sort(Let)
Public Property Let Sort(ByVal iValue As String)
    mSort = iValue
    If IsNumeric(iValue) Then
        mSortNo = CLng(iValue)
    Else
        mSortNo = 0
    End If
End Property

'プロパティ　PRSReference(Get)
Public Property Get PRSReference() As String
    PRSReference = mPRSReference
End Property

'プロパティ　PRSReference(Let)
Public Property Let PRSReference(ByVal iValue As String)
    mPRSReference = iValue
End Property

'プロパティ　Data_Prerequisites(Get)
Public Property Get Data_Prerequisites() As String
    Data_Prerequisites = mData_Prerequisites
End Property

'プロパティ　Data_Prerequisites(Let)
Public Property Let Data_Prerequisites(ByVal iValue As String)
    mData_Prerequisites = iValue
End Property

'プロパティ　TestInstruction(Get)
Public Property Get TestInstruction() As String
    TestInstruction = mTestInstruction
End Property

'プロパティ　TestInstruction(Let)
Public Property Let TestInstruction(ByVal iValue As String)
    mTestInstruction = iValue
End Property

'プロパティ　ExpectedResult(Get)
Public Property Get ExpectedResult() As String
    ExpectedResult = mExpectedResult
End Property

'プロパティ　ExpectedResult(Let)
Public Property Let ExpectedResult(ByVal iValue As String)
    mExpectedResult = iValue
End Property

'プロパティ　TestResult(Get)
Public Property Get TestResult() As String
    TestResult = mTestResult
End Property

'プロパティ　TestResult(Let)
Public Property Let TestResult(ByVal iValue As String)
    mTestResult = iValue
End Property

'プロパティ　Evidence(Get)
Public Property Get Evidence() As String
    Evidence = mEvidence
End Property

'プロパティ　Evidence(Let)
Public Property Let Evidence(ByVal iValue As String)
    mEvidence = iValue
End Property

'プロパティ　RiskID(Get)
Public Property Get RiskID() As String
    RiskID = mRiskID
End Property

'プロパティ　RiskID(Let)
Public Property Let RiskID(ByVal iValue As String)
    mRiskID = iValue
End Property
