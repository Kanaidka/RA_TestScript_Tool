VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ROW_INDEX_COLUMN_HEADER As Long = 3
Private Const ROW_INDEX_DATA As Long = 4

Private Const COLUMN_HEADER_PHASE As String = "Phase"
Private Const COLUMN_HEADER_BUNDLE As String = "ímå©é“"
Private Const COLUMN_HEADER_VERSION As String = "version"
Private Const COLUMN_HEADER_LISKID_A_1 As String = "A-1"
Private Const COLUMN_HEADER_LISKID_A_2 As String = "A-2"
Private Const COLUMN_HEADER_LISKID_B_1 As String = "B-1"
Private Const COLUMN_HEADER_LISKID_B_2 As String = "B-2"
Private Const COLUMN_HEADER_LISKID_B_3 As String = "B-3"
Private Const COLUMN_HEADER_LISKID_B_4 As String = "B-4"
Private Const COLUMN_HEADER_LISKID_B_5 As String = "B-5"
Private Const COLUMN_HEADER_LISKID_B_6 As String = "B-6"
Private Const COLUMN_HEADER_LISKID_C_1 As String = "C-1"
Private Const COLUMN_HEADER_LISKID_C_2 As String = "C-2"
Private Const COLUMN_HEADER_LISKID_C_3 As String = "C-3"
Private Const COLUMN_HEADER_LISKID_C_4 As String = "C-4"
Private Const COLUMN_HEADER_LISKID_C_5 As String = "C-5"
Private Const COLUMN_HEADER_LISKID_D_1 As String = "D-1"
Private Const COLUMN_HEADER_LISKID_D_2 As String = "D-2"
Private Const COLUMN_HEADER_LISKID_D_3 As String = "D-3"
Private Const COLUMN_HEADER_LISKID_D_4 As String = "D-4"
Private Const COLUMN_HEADER_LISKID_D_5 As String = "D-5"
Private Const COLUMN_HEADER_LISKID_E_1 As String = "E-1"
Private Const COLUMN_HEADER_LISKID_E_2 As String = "E-2"
Private Const COLUMN_HEADER_LISKID_E_3 As String = "E-3"
Private Const COLUMN_HEADER_LISKID_F_1 As String = "F-1"
Private Const COLUMN_HEADER_LISKID_F_2 As String = "F-2"
Private Const COLUMN_HEADER_LISKID_F_3 As String = "F-3"
Private Const COLUMN_HEADER_LISKID_F_4 As String = "F-4"
Private Const COLUMN_HEADER_LISKID_G_1 As String = "G-1"
Private Const COLUMN_HEADER_LISKID_H_1 As String = "H-1"
Private Const COLUMN_HEADER_LISKID_I_1 As String = "I-1"
Private Const COLUMN_HEADER_LISKID_I_2 As String = "I-2"
Private Const COLUMN_HEADER_LISKID_I_3 As String = "I-3"
Private Const COLUMN_HEADER_LISKID_J_1 As String = "J-1"
Private Const COLUMN_HEADER_LISKID_K_1 As String = "K-1"
Private Const COLUMN_HEADER_LISKID_L_1 As String = "L1"


Private mRows() As MatrixRow

Public ColumnPhase As New Column
Public ColumnBundle As New Column
Public ColumnVersion As New Column
Public ColumnRiskID_A_1 As New Column
Public ColumnRiskID_A_2 As New Column
Public ColumnRiskID_B_1 As New Column
Public ColumnRiskID_B_2 As New Column
Public ColumnRiskID_B_3 As New Column
Public ColumnRiskID_B_4 As New Column
Public ColumnRiskID_B_5 As New Column
Public ColumnRiskID_B_6 As New Column
Public ColumnRiskID_C_1 As New Column
Public ColumnRiskID_C_2 As New Column
Public ColumnRiskID_C_3 As New Column
Public ColumnRiskID_C_4 As New Column
Public ColumnRiskID_C_5 As New Column
Public ColumnRiskID_D_1 As New Column
Public ColumnRiskID_D_2 As New Column
Public ColumnRiskID_D_3 As New Column
Public ColumnRiskID_D_4 As New Column
Public ColumnRiskID_D_5 As New Column
Public ColumnRiskID_E_1 As New Column
Public ColumnRiskID_E_2 As New Column
Public ColumnRiskID_E_3 As New Column
Public ColumnRiskID_F_1 As New Column
Public ColumnRiskID_F_2 As New Column
Public ColumnRiskID_F_3 As New Column
Public ColumnRiskID_F_4 As New Column
Public ColumnRiskID_G_1 As New Column
Public ColumnRiskID_H_1 As New Column
Public ColumnRiskID_I_1 As New Column
Public ColumnRiskID_I_2 As New Column
Public ColumnRiskID_I_3 As New Column
Public ColumnRiskID_J_1 As New Column
Public ColumnRiskID_K_1 As New Column
Public ColumnRiskID_L_1 As New Column

Private mColumnRiskIDs() As Column

Public Property Get Rows()
    Rows = mRows
End Property

Private Sub Class_Initialize()
    Call ColumnPhase.Init(COLUMN_HEADER_PHASE)
    Call ColumnBundle.Init(COLUMN_HEADER_BUNDLE)
    Call ColumnVersion.Init(COLUMN_HEADER_VERSION)
    Call ColumnRiskID_A_1.Init(COLUMN_HEADER_LISKID_A_1)
    Call ColumnRiskID_A_2.Init(COLUMN_HEADER_LISKID_A_2)
    Call ColumnRiskID_B_1.Init(COLUMN_HEADER_LISKID_B_1)
    Call ColumnRiskID_B_2.Init(COLUMN_HEADER_LISKID_B_2)
    Call ColumnRiskID_B_3.Init(COLUMN_HEADER_LISKID_B_3)
    Call ColumnRiskID_B_4.Init(COLUMN_HEADER_LISKID_B_4)
    Call ColumnRiskID_B_5.Init(COLUMN_HEADER_LISKID_B_5)
    Call ColumnRiskID_B_6.Init(COLUMN_HEADER_LISKID_B_6)
    Call ColumnRiskID_C_1.Init(COLUMN_HEADER_LISKID_C_1)
    Call ColumnRiskID_C_2.Init(COLUMN_HEADER_LISKID_C_2)
    Call ColumnRiskID_C_3.Init(COLUMN_HEADER_LISKID_C_3)
    Call ColumnRiskID_C_4.Init(COLUMN_HEADER_LISKID_C_4)
    Call ColumnRiskID_C_5.Init(COLUMN_HEADER_LISKID_C_5)
    Call ColumnRiskID_D_1.Init(COLUMN_HEADER_LISKID_D_1)
    Call ColumnRiskID_D_2.Init(COLUMN_HEADER_LISKID_D_2)
    Call ColumnRiskID_D_3.Init(COLUMN_HEADER_LISKID_D_3)
    Call ColumnRiskID_D_4.Init(COLUMN_HEADER_LISKID_D_4)
    Call ColumnRiskID_D_5.Init(COLUMN_HEADER_LISKID_D_5)
    Call ColumnRiskID_E_1.Init(COLUMN_HEADER_LISKID_E_1)
    Call ColumnRiskID_E_2.Init(COLUMN_HEADER_LISKID_E_2)
    Call ColumnRiskID_E_3.Init(COLUMN_HEADER_LISKID_E_3)
    Call ColumnRiskID_F_1.Init(COLUMN_HEADER_LISKID_F_1)
    Call ColumnRiskID_F_2.Init(COLUMN_HEADER_LISKID_F_2)
    Call ColumnRiskID_F_3.Init(COLUMN_HEADER_LISKID_F_3)
    Call ColumnRiskID_F_4.Init(COLUMN_HEADER_LISKID_F_4)
    Call ColumnRiskID_G_1.Init(COLUMN_HEADER_LISKID_G_1)
    Call ColumnRiskID_H_1.Init(COLUMN_HEADER_LISKID_H_1)
    Call ColumnRiskID_I_1.Init(COLUMN_HEADER_LISKID_I_1)
    Call ColumnRiskID_I_2.Init(COLUMN_HEADER_LISKID_I_2)
    Call ColumnRiskID_I_3.Init(COLUMN_HEADER_LISKID_I_3)
    Call ColumnRiskID_J_1.Init(COLUMN_HEADER_LISKID_J_1)
    Call ColumnRiskID_K_1.Init(COLUMN_HEADER_LISKID_K_1)
    Call ColumnRiskID_L_1.Init(COLUMN_HEADER_LISKID_L_1)
    
    ReDim mColumnRiskIDs(32)
    Set mColumnRiskIDs(0) = ColumnRiskID_A_2
    Set mColumnRiskIDs(1) = ColumnRiskID_A_1
    Set mColumnRiskIDs(2) = ColumnRiskID_B_1
    Set mColumnRiskIDs(3) = ColumnRiskID_B_2
    Set mColumnRiskIDs(4) = ColumnRiskID_B_3
    Set mColumnRiskIDs(5) = ColumnRiskID_B_4
    Set mColumnRiskIDs(6) = ColumnRiskID_B_5
    Set mColumnRiskIDs(7) = ColumnRiskID_B_6
    Set mColumnRiskIDs(8) = ColumnRiskID_C_1
    Set mColumnRiskIDs(9) = ColumnRiskID_C_2
    Set mColumnRiskIDs(10) = ColumnRiskID_C_3
    Set mColumnRiskIDs(11) = ColumnRiskID_C_4
    Set mColumnRiskIDs(12) = ColumnRiskID_C_5
    Set mColumnRiskIDs(13) = ColumnRiskID_D_1
    Set mColumnRiskIDs(14) = ColumnRiskID_D_2
    Set mColumnRiskIDs(15) = ColumnRiskID_D_3
    Set mColumnRiskIDs(16) = ColumnRiskID_D_4
    Set mColumnRiskIDs(17) = ColumnRiskID_D_5
    Set mColumnRiskIDs(18) = ColumnRiskID_E_1
    Set mColumnRiskIDs(19) = ColumnRiskID_E_2
    Set mColumnRiskIDs(20) = ColumnRiskID_E_3
    Set mColumnRiskIDs(21) = ColumnRiskID_F_1
    Set mColumnRiskIDs(22) = ColumnRiskID_F_2
    Set mColumnRiskIDs(23) = ColumnRiskID_F_3
    Set mColumnRiskIDs(24) = ColumnRiskID_F_4
    Set mColumnRiskIDs(25) = ColumnRiskID_G_1
    Set mColumnRiskIDs(26) = ColumnRiskID_H_1
    Set mColumnRiskIDs(27) = ColumnRiskID_I_1
    Set mColumnRiskIDs(28) = ColumnRiskID_I_2
    Set mColumnRiskIDs(29) = ColumnRiskID_I_3
    Set mColumnRiskIDs(30) = ColumnRiskID_J_1
    Set mColumnRiskIDs(31) = ColumnRiskID_K_1
    Set mColumnRiskIDs(32) = ColumnRiskID_L_1
    
End Sub

Public Sub Load(ByRef iWs As Worksheet)
    
    Dim i As Long
    Dim j As Long
    Dim rowIndex As Long
    Dim rowIndexLast As Long
        
    Call ColumnPhase.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnBundle.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnVersion.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    For i = 0 To UBound(mColumnRiskIDs)
        Call mColumnRiskIDs(i).SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Next
    
    rowIndexLast = iWs.Cells(iWs.Rows.Count, ColumnPhase.Index).End(xlUp).row
    
    ReDim mRows(rowIndexLast - ROW_INDEX_DATA)
    Dim risk As RiskID
    i = 0
    For rowIndex = ROW_INDEX_DATA To rowIndexLast
        Set mRows(i) = New MatrixRow
        With mRows(i)
            .Phase = iWs.Cells(rowIndex, ColumnPhase.Index).value
            .Bundle = iWs.Cells(rowIndex, ColumnBundle.Index).value
            .Version = iWs.Cells(rowIndex, ColumnVersion.Index).value
            
            For j = 0 To UBound(mColumnRiskIDs)
                If iWs.Cells(rowIndex, mColumnRiskIDs(j).Index).value <> "" Then
                    Set risk = New RiskID
                    Call risk.Init(mColumnRiskIDs(j).Name, iWs.Cells(rowIndex, mColumnRiskIDs(j).Index).value)
                    Call .AddRiskID(risk)
                End If
            Next
            
        End With
        i = i + 1
    Next

End Sub

Public Function GetRow(ByVal iKey As String) As MatrixRow

    Dim i As Long
    For i = LBound(mRows) To UBound(mRows)
        If mRows(i).Key = iKey Then
            Set GetRow = mRows(i)
            Exit For
        End If
    Next
    
End Function
