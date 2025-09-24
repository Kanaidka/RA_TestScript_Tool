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
'クラス　TestScript
'テストスクリプトデータを管理します

Private Const COLOR_OP As Long = 15128749 'OP行の背景色

Private mMatrix As New MatrixData
Private mFormat As New FormatData
Private mPRS As New PRSData
Private mTemplate As New TemplateData

Private mWs As Worksheet
Private mRowIndex As Long
Private mTextConverter As New TextConverter

'関数　LoadMatrix
'引数　iWs　読込対象ワークシート
'Matrixデータの読み込み
Public Sub LoadMatrix(ByVal iWs As Worksheet)
    Call mMatrix.Load(iWs)
End Sub

'関数　LoadFormat
'引数　iWs　読込対象ワークシート
'Formatデータの読み込み
Public Sub LoadFormat(ByVal iWs As Worksheet)
    Call mFormat.Load(iWs)
End Sub

'関数　LoadPRS
'引数　iWs　読込対象ワークシート
'PRSデータの読み込み
Public Sub LoadPRS(ByVal iWs As Worksheet)
    Call mPRS.Load(iWs)
    Set mTextConverter.Workbook = iWs.Parent
End Sub

'関数　LoadTemplate
'引数　iWs　読込対象ワークシート
'Templateデータの読み込み
Public Sub LoadTemplate(ByVal iWs As Worksheet)
    Call mTemplate.Load(iWs)
End Sub

'関数　Create
'引数　iSheetName　出力シート名
'TestScriptシートの作成
Public Sub Create(ByVal iSheetName As String)
    
    Dim PR As PRSPR
    Dim i As Long
    
    Set mWs = mTemplate.Copy(iSheetName)
    mRowIndex = mTemplate.RowIndexData
                     
    For i = 0 To UBound(mPRS.PRs)
        Set PR = mPRS.PRs(i)
        Call CreateHeader(PR)
        Call CreatePR(PR)
    Next
    
    With Range(mWs.Cells(mTemplate.RowIndexColumnHeader, 1), mWs.Cells(mRowIndex - 1, mTemplate.ColumnIndexMax))
        .Borders.LineStyle = xlContinuous
        .WrapText = True
    End With
    
End Sub

'関数　CreateHeader
'引数　iPR　PRデータ
'TestScriptシートヘッダ情報の作成
Public Sub CreateHeader(ByRef iPR As PRSPR)
    
    Dim UP As PRSUP
    Dim OP As PRSOP
    Dim PH As PRSPH

    'まったくPHデータがない場合はヘッダ作成を行わない
    If iPR Is Nothing Then Exit Sub
    If UBound(iPR.UPs) < 0 Then Exit Sub
    Set UP = iPR.UPs(0)
    If UBound(UP.OPs) < 0 Then Exit Sub
    Set OP = UP.OPs(0)
    If UBound(OP.PHs) < 0 Then Exit Sub
    Set PH = OP.PHs(0)

    'リスクIDのG-1は必ず最初に出力
    Call CreatePH_RiskID(PH, mMatrix.ColumnRiskID_G_1.Name)

End Sub

'関数　CreatePR
'引数　iPR　PRデータ
'PRデータに基づくシート内容作成
Public Sub CreatePR(ByRef iPR As PRSPR)

    Dim i As Long
    Dim UP As PRSUP
    
    'UPデータ行の作成
    For i = 0 To UBound(iPR.UPs)
        Set UP = iPR.UPs(i)
        Call CreateUP(UP)
    Next

End Sub

'関数　CreateUP
'引数　iUP　UPデータ
'UPデータに基づくシート内容作成
Public Sub CreateUP(ByRef iUP As PRSUP)
    
    Dim i As Long
    Dim OP As PRSOP
    
    'OPデータ行の作成
    For i = 0 To UBound(iUP.OPs)
        Set OP = iUP.OPs(i)
        Call CreateOP(OP)
    Next

End Sub

'関数　CreateOP
'引数　iOP　OPデータ
'OPデータに基づくシート内容作成
Public Sub CreateOP(ByRef iOP As PRSOP)
    
    Dim i As Long
    Dim PH As PRSPH
      
    'OPデータ行の作成
    With mWs
        .Cells(mRowIndex, mTemplate.ColumnPRSReference.Index).Value = iOP.id & vbCrLf & iOP.text
        .Range(.Cells(mRowIndex, 1), .Cells(mRowIndex, mTemplate.ColumnIndexMax)).Interior.Color = COLOR_OP
        mRowIndex = mRowIndex + 1
    End With
    
    'PHデータ行の作成
    For i = 0 To UBound(iOP.PHs)
        Set PH = iOP.PHs(i)
        Call CreatePH(PH)
    Next

End Sub

'関数　CreatePH
'引数　iPH　PHデータ
'PHデータに基づくシート内容作成
Public Sub CreatePH(ByRef iPH As PRSPH)
    
    Dim m As MatrixRow
    
    Set m = mMatrix.GetRow(iPH.key)
    If m Is Nothing Then Exit Sub
    
    Dim i As Long
    Dim fList() As FormatRow
    Dim riskIDs As String
        
    riskIDs = ""
    
    For i = 0 To UBound(m.riskIDs)
        riskIDs = riskIDs & m.riskIDs(i).id & vbCrLf
    Next
    
    fList = mFormat.GetRows(riskIDs, iPH.key)
    
    For i = 0 To UBound(fList)
        Call CreatePH_Format(iPH, fList(i))
    Next
    

End Sub

'関数　CreatePH_RiskID
'引数　iPH　PHデータ
'引数　iRiskID　リスクIDデータ
'PHデータとRiskIDに基づくシート内容作成
Public Sub CreatePH_RiskID(ByRef iPH As PRSPH, ByRef iRiskID As String)

    Dim i As Long
    Dim fList() As FormatRow
        
    fList = mFormat.GetRows(iRiskID, iPH.key)
    
    For i = 0 To UBound(fList)
        Call CreatePH_Format(iPH, fList(i))
    Next
    
End Sub

'関数　CreatePH_RiskID
'引数　iPH　PHデータ
'引数　iFormat　Formatデータ
'PHデータとiFormatに基づくシート内容作成
Public Sub CreatePH_Format(ByRef iPH As PRSPH, ByRef iFormat As FormatRow)

    Dim i As Long
        
    If iFormat Is Nothing Then Exit Sub
    
    With mWs
    
        .Cells(mRowIndex, mTemplate.ColumnPRSReference.Index).Value = mTextConverter.Convert(iFormat.Required, iPH)
        
        If Replace(.Cells(mRowIndex, mTemplate.ColumnPRSReference.Index).text, vbCrLf, "") = "" Then
            .Cells(mRowIndex, mTemplate.ColumnPRSReference.Index).Value = ""
        Else
            .Cells(mRowIndex, mTemplate.ColumnTestStep.Index).NumberFormat = """Step""0000"
            .Cells(mRowIndex, mTemplate.ColumnTestStep.Index).Formula = "=COUNTA(" & .Cells(mTemplate.RowIndexData, mTemplate.ColumnRiskID.Index).Address & ":" & .Cells(mRowIndex, mTemplate.ColumnRiskID.Index).Address(RowAbsolute:=False) & ")"
            
            .Cells(mRowIndex, mTemplate.ColumnPRSReference.Index).Value = mTextConverter.Convert(iFormat.PRSReference, iPH)
            .Cells(mRowIndex, mTemplate.ColumnData_Prerequisites.Index).Value = mTextConverter.Convert(iFormat.Data_Prerequisites, iPH)
            .Cells(mRowIndex, mTemplate.ColumnTestInstruction.Index).Value = mTextConverter.Convert(iFormat.TestInstruction, iPH)
            .Cells(mRowIndex, mTemplate.ColumnExpectedResult.Index).Value = mTextConverter.Convert(iFormat.ExpectedResult, iPH)
            .Cells(mRowIndex, mTemplate.ColumnRiskID.Index).Value = iFormat.RiskID
            .Cells(mRowIndex, mTemplate.ColumnTestResult.Index).Value = iFormat.TestResult
            .Cells(mRowIndex, mTemplate.ColumnEvidence.Index).Value = iFormat.Evidence
            mRowIndex = mRowIndex + 1
        End If
    
    End With
    
End Sub
