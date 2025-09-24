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
'クラス　FormatData
'Formatデータを管理します

Private Const ROW_INDEX_COLUMN_HEADER As Long = 1 'ヘッダ行番号
Private Const ROW_INDEX_DATA As Long = 2 'データ行番号

'列名(ヘッダ)情報
Private Const COLUMN_HEADER_PHASE As String = "Phase"
Private Const COLUMN_HEADER_REQUIRED As String = "Required"
Private Const COLUMN_HEADER_SORT As String = "Sort"
Private Const COLUMN_HEADER_PRSREFERENCE As String = "PRS Reference"
Private Const COLUMN_HEADER_DATA_PREREQUISITES As String = "Data/Prerequisites"
Private Const COLUMN_HEADER_TESTINSTRUCTION As String = "Test Instruction"
Private Const COLUMN_HEADER_EXPECTEDRESULT As String = "Expected result"
Private Const COLUMN_HEADER_RISKID As String = "Risk ID"
Private Const COLUMN_HEADER_TESTRESULT As String = "Test Result"
Private Const COLUMN_HEADER_EVIDENCE As String = "Evidence"

Private mRows() As FormatRow '行データ格納

'列情報
Public ColumnPhase As New Column
Public ColumnRequired As New Column
Public ColumnSort As New Column
Public ColumnPRSReference As New Column
Public ColumnData_Prerequisites As New Column
Public ColumnTestInstruction As New Column
Public ColumnExpectedResult As New Column
Public ColumnRiskID As New Column
Public ColumnTestResult As New Column
Public ColumnEvidence As New Column

'イベント　クラスの初期化
'クラスの初期化を行う
Private Sub Class_Initialize()
    '列情報の初期化
    Call ColumnPhase.Init(COLUMN_HEADER_PHASE)
    Call ColumnRequired.Init(COLUMN_HEADER_REQUIRED)
    Call ColumnSort.Init(COLUMN_HEADER_SORT)
    Call ColumnPRSReference.Init(COLUMN_HEADER_PRSREFERENCE)
    Call ColumnData_Prerequisites.Init(COLUMN_HEADER_DATA_PREREQUISITES)
    Call ColumnTestInstruction.Init(COLUMN_HEADER_TESTINSTRUCTION)
    Call ColumnExpectedResult.Init(COLUMN_HEADER_EXPECTEDRESULT)
    Call ColumnRiskID.Init(COLUMN_HEADER_RISKID)
    Call ColumnTestResult.Init(COLUMN_HEADER_TESTRESULT)
    Call ColumnEvidence.Init(COLUMN_HEADER_EVIDENCE)
End Sub

'関数　Load
'引数　iWs　対象ワークシート
'対象ワークシートよりフォーマット情報を取得する
Public Sub Load(ByRef iWs As Worksheet)
    
    Dim i As Long
    Dim rowIndex As Long
    Dim rowIndexLast As Long
    
    '各列番号検索
    Call ColumnPhase.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnRequired.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnSort.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnPRSReference.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnData_Prerequisites.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnTestInstruction.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnExpectedResult.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnRiskID.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnTestResult.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnEvidence.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    
    
    'PRSReferenceの最終データ行を取得
    rowIndexLast = iWs.Cells(iWs.Rows.Count, ColumnPRSReference.Index).End(xlUp).row
    
    'データ数分のデータ配列確保
    ReDim mRows(rowIndexLast - ROW_INDEX_DATA)
    i = 0
    '各行の読込
    For rowIndex = ROW_INDEX_DATA To rowIndexLast
        Set mRows(i) = New FormatRow
        With mRows(i)
            .Phase = iWs.Cells(rowIndex, ColumnPhase.Index).Value
            .Required = iWs.Cells(rowIndex, ColumnRequired.Index).Value
            .Sort = iWs.Cells(rowIndex, ColumnSort.Index).Value
            .PRSReference = iWs.Cells(rowIndex, ColumnPRSReference.Index).Value
            .Data_Prerequisites = iWs.Cells(rowIndex, ColumnData_Prerequisites.Index).Value
            .TestInstruction = iWs.Cells(rowIndex, ColumnTestInstruction.Index).Value
            .ExpectedResult = iWs.Cells(rowIndex, ColumnExpectedResult.Index).Value
            .RiskID = iWs.Cells(rowIndex, ColumnRiskID.Index).Value
            .TestResult = iWs.Cells(rowIndex, ColumnTestResult.Index).Value
            .Evidence = iWs.Cells(rowIndex, ColumnEvidence.Index).Value
        End With
        i = i + 1
    Next

    Call Sort

End Sub

'関数　GetRows
'引数　iRiskID　リスクID
'引数　iPhase　フェーズ名
'戻り値　Format行配列
'リスクIDとフェーズ名より該当するFormatRowを取得する
Public Function GetRows(ByVal iRiskIDs As String, ByVal iPhase As String) As FormatRow()

    Dim i As Long
    Dim j As Long
    Dim ret() As FormatRow
    
    j = 0
    ReDim ret(j)
    Set ret(j) = Nothing
    For i = 0 To UBound(mRows)
        If InStr(iRiskIDs, mRows(i).RiskID) > 0 And Trim(mRows(i).RiskID) <> vbNullString And _
            ( _
                mRows(i).Phase = "" Or _
                (Left(mRows(i).Phase, 4) <> "NOT," And InStr(mRows(i).Phase, iPhase) > 0) Or _
                (Left(mRows(i).Phase, 4) = "NOT," And InStr(mRows(i).Phase, iPhase) = 0) _
            ) Then
            ReDim Preserve ret(j)
            Set ret(j) = mRows(i)
            j = j + 1
        End If
    Next
    
    GetRows = ret()
        
End Function

'関数　Sort
'mRowsをSort行の内容で昇順ソートする
Private Sub Sort()
    Dim swap As FormatRow
    Dim i As Integer
    Dim j As Integer
    For i = UBound(mRows) To LBound(mRows) Step -1
        For j = LBound(mRows) To i - 1
            If mRows(j).SortNo > mRows(j + 1).SortNo Then
                Set swap = mRows(j)
                Set mRows(j) = mRows(j + 1)
                Set mRows(j + 1) = swap
            End If
        Next j
    Next i
End Sub





