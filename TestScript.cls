VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_OP As Long = 15128749

Private mMatrix As New MatrixData
Private mFormat As New FormatData
Private mPRS As New PRSData
Private mTemplate As New TemplateData

Private mTextConverter As New TextConverter

Public Sub LoadMatrix(ByVal iWs As Worksheet)
    Call mMatrix.Load(iWs)
End Sub

Public Sub LoadFormat(ByVal iWs As Worksheet)
    Call mFormat.Load(iWs)
End Sub

Public Sub LoadPRS(ByVal iWs As Worksheet)
    Call mPRS.Load(iWs)
End Sub

Public Sub LoadTemplate(ByVal iWs As Worksheet)
    Call mTemplate.Load(iWs)
End Sub

Public Function Create(ByVal iSheetName As String) As Boolean
    
    Dim ws As Worksheet
    Set ws = mTemplate.Copy(iSheetName)
    
    Dim m As MatrixRow
    Dim f As FormatRow
    Dim PR As PRSPR
    Dim UP As PRSUP
    Dim OP As PRSOP
    Dim ph As PRSPH
    Dim prIndex As Long
    Dim upIndex As Long
    Dim opIndex As Long
    Dim phIndex As Long
    Dim riskIDIndex As Long
    Dim rowIndex As Long
    
    
    rowIndex = mTemplate.RowIndexData
    
    Set f = mFormat.GetRow(mMatrix.ColumnRiskID_G_1.Name, "")
    If Not f Is Nothing Then
        If f.Required <> "" Then
            With ws.Cells(rowIndex, mTemplate.ColumnTestStep.Index)
                .NumberFormat = """Step""0000"
                .Formula = "=COUNTA(" & ws.Cells(mTemplate.RowIndexData, mTemplate.ColumnRiskID.Index).Address & ":" & ws.Cells(rowIndex, mTemplate.ColumnRiskID.Index).Address(RowAbsolute:=False) & ")"
            End With
            ws.Cells(rowIndex, mTemplate.ColumnPRSReference.Index).value = f.PRSReference
            ws.Cells(rowIndex, mTemplate.ColumnData_Prerequisites.Index).value = f.Data_Prerequisites
            ws.Cells(rowIndex, mTemplate.ColumnTestInstruction.Index).value = f.TestInstruction
            ws.Cells(rowIndex, mTemplate.ColumnExpectedResult.Index).value = f.ExpectedResult
            ws.Cells(rowIndex, mTemplate.ColumnRiskID.Index).value = f.RiskID
            ws.Cells(rowIndex, mTemplate.ColumnTestResult.Index).value = f.TestResult
            ws.Cells(rowIndex, mTemplate.ColumnEvidence.Index).value = f.Evidence
            rowIndex = rowIndex + 1
        End If
    End If
                    
    For prIndex = LBound(mPRS.PRs) To UBound(mPRS.PRs)
        Set PR = mPRS.PRs(prIndex)
        
        For upIndex = LBound(PR.UPs) To UBound(PR.UPs)
            Set UP = PR.UPs(upIndex)
            
            For opIndex = LBound(UP.OPs) To UBound(UP.OPs)
                Set OP = UP.OPs(opIndex)
                
                ws.Cells(rowIndex, mTemplate.ColumnPRSReference.Index).value = OP.ID & vbCrLf & OP.Text
                ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, mTemplate.ColumnIndexMax)).Interior.Color = COLOR_OP
                rowIndex = rowIndex + 1
                
                For phIndex = LBound(OP.PHs) To UBound(OP.PHs)
                    Set ph = OP.PHs(phIndex)
                                        
                    Set m = mMatrix.GetRow(ph.Key)
                    If m Is Nothing Then GoTo PH_CONTINUE
                    
                    For riskIDIndex = LBound(m.RiskIDs) To UBound(m.RiskIDs)
                        Set f = mFormat.GetRow(m.RiskIDs(riskIDIndex).ID, ph.Key)
                        
                        If f Is Nothing Then GoTo RISKID_CONTINUE
                        
                        If mTextConverter.Convert(f.Required, ph) = "" Then GoTo RISKID_CONTINUE
                        
                        With ws.Cells(rowIndex, mTemplate.ColumnTestStep.Index)
                            .NumberFormat = """Step""0000"
                            .Formula = "=COUNTA(" & ws.Cells(mTemplate.RowIndexData, mTemplate.ColumnRiskID.Index).Address & ":" & ws.Cells(rowIndex, mTemplate.ColumnRiskID.Index).Address(RowAbsolute:=False) & ")"
                        End With
                        ws.Cells(rowIndex, mTemplate.ColumnPRSReference.Index).value = mTextConverter.Convert(f.PRSReference, ph)
                        ws.Cells(rowIndex, mTemplate.ColumnData_Prerequisites.Index).value = mTextConverter.Convert(f.Data_Prerequisites, ph)
                        ws.Cells(rowIndex, mTemplate.ColumnTestInstruction.Index).value = mTextConverter.Convert(f.TestInstruction, ph)
                        ws.Cells(rowIndex, mTemplate.ColumnExpectedResult.Index).value = mTextConverter.Convert(f.ExpectedResult, ph)
                        ws.Cells(rowIndex, mTemplate.ColumnRiskID.Index).value = f.RiskID
                        ws.Cells(rowIndex, mTemplate.ColumnTestResult.Index).value = f.TestResult
                        ws.Cells(rowIndex, mTemplate.ColumnEvidence.Index).value = f.Evidence
                        rowIndex = rowIndex + 1
RISKID_CONTINUE:
                    Next
                                                                               
PH_CONTINUE:
                                        
                Next
                
            Next
            
        Next
    Next
    
    Range(ws.Cells(mTemplate.RowIndexColumnHeader, 1), ws.Cells(rowIndex - 1, mTemplate.ColumnIndexMax)).Borders.LineStyle = xlContinuous
    
End Function

