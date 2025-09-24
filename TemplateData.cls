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
'クラス　TemplateData
'Templateデータを管理します

Private Const ROW_INDEX_COLUMN_HEADER As Long = 1 'ヘッダ行番号
Private Const ROW_INDEX_DATA As Long = 2 'データ行番号

'列名(ヘッダ)情報
Private Const COLUMN_HEADER_TESTSTEP As String = "Test -Step"
Private Const COLUMN_HEADER_PRSREFERENCE As String = "PRS Reference"
Private Const COLUMN_HEADER_DATA_PREREQUISITES As String = "Data/Prerequisites"
Private Const COLUMN_HEADER_TESTINSTRUCTION As String = "Test Instruction"
Private Const COLUMN_HEADER_EXPECTEDRESULT As String = "Expected result"
Private Const COLUMN_HEADER_RISKID As String = "Risk ID"
Private Const COLUMN_HEADER_TESTRESULT As String = "Test Result"
Private Const COLUMN_HEADER_EVIDENCE As String = "Evidence"
Private Const COLUMN_HEADER_INITIALS_DATE As String = "Initials / Date"
Private Const COLUMN_HEADER_LINK As String = "リンク"
Private Const COLUMN_HEADER_REMARK As String = "備考"

Private mSheet As Worksheet
Private mColumnIndexMax As Long

'列情報
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

'イベント　クラスの初期化
'クラスの初期化を行う
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

'プロパティ　RowIndexColumnHeader(Get)
'列ヘッダ行番号
Public Property Get RowIndexColumnHeader() As Long
    RowIndexColumnHeader = ROW_INDEX_COLUMN_HEADER
End Property

'プロパティ　RowIndexData(Get)
'データ行番号
Public Property Get RowIndexData() As Long
    RowIndexData = ROW_INDEX_DATA
End Property

'プロパティ　ColumnIndexMax(Get)
'最大列番号（罫線を引く処理などに利用）
Public Property Get ColumnIndexMax() As Long
    ColumnIndexMax = mColumnIndexMax
End Property

'関数　Load
'引数　iWs　読込対象ワークシート
'Templateシートの読込
Public Sub Load(ByRef iWs As Worksheet)
    
    Set mSheet = iWs
    'テンプレートの最終列を保持しておく
    mColumnIndexMax = iWs.Cells(ROW_INDEX_COLUMN_HEADER, iWs.Columns.Count).End(xlToLeft).Column
    '各行番号の設定
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

'関数　Copy
'引数　iSheetName　コピー後のシート名
'戻り値　コピーしたWorksheet
'Templateシートのコピー
Public Function Copy(ByVal iSheetName As String) As Worksheet

    Dim ws As Worksheet
    Dim i As Long
    Dim find As Boolean
    Dim dstSheetName As String
    
    dstSheetName = iSheetName
    Do
        '発見フラグ初期化
        find = False
        'シート全体を検索し、同名がないかを検索
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = dstSheetName Then
                '同名が見つかった場合
                i = i + 1
                'シート名に_2桁数値を付け、再検索
                dstSheetName = iSheetName & "_" & Format(i, "00")
                '発見フラグON
                find = True
            End If
        Next
        '同名がなくなるまで検索
    Loop While find
    
    'シートコピー
    Call mSheet.Copy(ThisWorkbook.Sheets(1))
    
    'コピーシートの名称変更
    ThisWorkbook.Sheets(1).Name = dstSheetName
    
    'コピー後のシートを戻り値に設定
    Set Copy = ThisWorkbook.Sheets(1)
    
End Function

