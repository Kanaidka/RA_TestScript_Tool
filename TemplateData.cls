VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ROW_INDEX_COLUMN_HEADER As Long = 1
Private Const ROW_INDEX_DATA As Long = 2

Private Const COLUMN_HEADER_TESTSTEP As String = "Test -Step"
Private Const COLUMN_HEADER_PRSREFERENCE As String = "PRS Reference"
Private Const COLUMN_HEADER_DATA_PREREQUISITES As String = "Data/Prerequisites"
Private Const COLUMN_HEADER_TESTINSTRUCTION As String = "Test Instruction"
Private Const COLUMN_HEADER_EXPECTEDRESULT As String = "Expected result"
Private Const COLUMN_HEADER_RISKID As String = "Risk ID"
Private Const COLUMN_HEADER_TESTRESULT As String = "Test Result"
Private Const COLUMN_HEADER_EVIDENCE As String = "Evidence"
Private Const COLUMN_HEADER_INITIALS_DATE As String = "Initials / Date"
Private Const COLUMN_HEADER_LINK As String = "ÉäÉìÉN"
Private Const COLUMN_HEADER_REMARK As String = "îıçl"

Private mSheet As Worksheet
Private mColumnIndexMax As Long

Public ColumnTestStep As New Column
Public ColumnPRSReference As New Column
Public ColumnData_Prerequisites As New Column
Public ColumnTestInstruction As New Column
Public ColumnExpectedResult As New Column
Public ColumnRiskID As New Column
Public ColumnTestResult As New Column
Public ColumnEvidence As New Column
Public ColumnInitials_Date As New Column
Public ColumnLink As New Column
Public ColumnRemark As New Column


Private Sub Class_Initialize()

    Call ColumnTestStep.Init(COLUMN_HEADER_TESTSTEP)
    Call ColumnPRSReference.Init(COLUMN_HEADER_PRSREFERENCE)
    Call ColumnData_Prerequisites.Init(COLUMN_HEADER_DATA_PREREQUISITES)
    Call ColumnTestInstruction.Init(COLUMN_HEADER_TESTINSTRUCTION)
    Call ColumnExpectedResult.Init(COLUMN_HEADER_EXPECTEDRESULT)
    Call ColumnRiskID.Init(COLUMN_HEADER_RISKID)
    Call ColumnTestResult.Init(COLUMN_HEADER_TESTRESULT)
    Call ColumnEvidence.Init(COLUMN_HEADER_EVIDENCE)
    Call ColumnInitials_Date.Init(COLUMN_HEADER_INITIALS_DATE)
    Call ColumnLink.Init(COLUMN_HEADER_LINK)
    Call ColumnRemark.Init(COLUMN_HEADER_REMARK)

End Sub

Public Property Get RowIndexColumnHeader() As Long
    RowIndexColumnHeader = ROW_INDEX_COLUMN_HEADER
End Property

Public Property Get RowIndexData() As Long
    RowIndexData = ROW_INDEX_DATA
End Property

Public Property Get ColumnIndexMax() As Long
    ColumnIndexMax = mColumnIndexMax
End Property


Public Sub Load(ByRef iWs As Worksheet)
    
    Set mSheet = iWs
    mColumnIndexMax = iWs.Cells(ROW_INDEX_COLUMN_HEADER, iWs.Columns.Count).End(xlToLeft).Column
    Call ColumnTestStep.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnPRSReference.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnData_Prerequisites.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnTestInstruction.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnExpectedResult.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnRiskID.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnTestResult.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnEvidence.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnInitials_Date.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnLink.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnRemark.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    

End Sub


Public Function Copy(ByVal iSheetName As String) As Worksheet

    Dim ws As Worksheet
    Dim i As Long
    Dim find As Boolean
    Dim dstSheetName As String
    
    dstSheetName = iSheetName
    Do
        find = False
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = dstSheetName Then
                i = i + 1
                dstSheetName = iSheetName & "_" & Format(i, "00")
                find = True
            End If
        Next
    Loop While find
    
    Call mSheet.Copy(ThisWorkbook.Sheets(1))
    
    ThisWorkbook.Sheets(1).Name = dstSheetName
    
    Set Copy = ThisWorkbook.Sheets(1)
    
End Function

