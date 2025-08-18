VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ROW_INDEX_COLUMN_HEADER As Long = 1
Private Const ROW_INDEX_DATA As Long = 2

Private Const COLUMN_HEADER_PHASE As String = "Phase"
Private Const COLUMN_HEADER_REQUIRED As String = "Required"
Private Const COLUMN_HEADER_PRSREFERENCE As String = "PRS Reference"
Private Const COLUMN_HEADER_DATA_PREREQUISITES As String = "Data/Prerequisites"
Private Const COLUMN_HEADER_TESTINSTRUCTION As String = "Test Instruction"
Private Const COLUMN_HEADER_EXPECTEDRESULT As String = "Expected result"
Private Const COLUMN_HEADER_RISKID As String = "Risk ID"
Private Const COLUMN_HEADER_TESTRESULT As String = "Test Result"
Private Const COLUMN_HEADER_EVIDENCE As String = "Evidence"

Private mRows() As FormatRow

Public ColumnPhase As New Column
Public ColumnRequired As New Column
Public ColumnPRSReference As New Column
Public ColumnData_Prerequisites As New Column
Public ColumnTestInstruction As New Column
Public ColumnExpectedResult As New Column
Public ColumnRiskID As New Column
Public ColumnTestResult As New Column
Public ColumnEvidence As New Column

Private Sub Class_Initialize()
    Call ColumnPhase.Init(COLUMN_HEADER_PHASE)
    Call ColumnRequired.Init(COLUMN_HEADER_REQUIRED)
    Call ColumnPRSReference.Init(COLUMN_HEADER_PRSREFERENCE)
    Call ColumnData_Prerequisites.Init(COLUMN_HEADER_DATA_PREREQUISITES)
    Call ColumnTestInstruction.Init(COLUMN_HEADER_TESTINSTRUCTION)
    Call ColumnExpectedResult.Init(COLUMN_HEADER_EXPECTEDRESULT)
    Call ColumnRiskID.Init(COLUMN_HEADER_RISKID)
    Call ColumnTestResult.Init(COLUMN_HEADER_TESTRESULT)
    Call ColumnEvidence.Init(COLUMN_HEADER_EVIDENCE)
End Sub


Public Sub Load(ByRef iWs As Worksheet)
    
    Dim i As Long
    Dim rowIndex As Long
    Dim rowIndexLast As Long
        
    Call ColumnPhase.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnRequired.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnPRSReference.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnData_Prerequisites.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnTestInstruction.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnExpectedResult.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnRiskID.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnTestResult.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnEvidence.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    
    rowIndexLast = iWs.Cells(iWs.Rows.Count, ColumnPRSReference.Index).End(xlUp).row
    
    ReDim mRows(rowIndexLast - ROW_INDEX_DATA)
    i = 0
    For rowIndex = ROW_INDEX_DATA To rowIndexLast
        Set mRows(i) = New FormatRow
        With mRows(i)
            .Phase = iWs.Cells(rowIndex, ColumnPhase.Index).value
            .Required = iWs.Cells(rowIndex, ColumnRequired.Index).value
            .PRSReference = iWs.Cells(rowIndex, ColumnPRSReference.Index).value
            .Data_Prerequisites = iWs.Cells(rowIndex, ColumnData_Prerequisites.Index).value
            .TestInstruction = iWs.Cells(rowIndex, ColumnTestInstruction.Index).value
            .ExpectedResult = iWs.Cells(rowIndex, ColumnExpectedResult.Index).value
            .RiskID = iWs.Cells(rowIndex, ColumnRiskID.Index).value
            .TestResult = iWs.Cells(rowIndex, ColumnTestResult.Index).value
            .Evidence = iWs.Cells(rowIndex, ColumnEvidence.Index).value
        End With
        i = i + 1
    Next

End Sub

Public Function GetRow(ByVal iID As String, ByVal iPhase As String) As FormatRow

    Dim i As Long
    For i = LBound(mRows) To UBound(mRows)
        If mRows(i).RiskID = iID And mRows(i).Phase = iPhase Then
            Set GetRow = mRows(i)
            Exit For
        End If
        If mRows(i).RiskID = iID Then
            Set GetRow = mRows(i)
            Exit For
        End If
    Next
    
End Function
