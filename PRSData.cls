VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRSData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ROW_INDEX_COLUMN_HEADER As Long = 1
Private Const ROW_INDEX_DATA As Long = 2

Private Const COLUMN_HEADER_PR As String = "製造工程: PR"
Private Const COLUMN_HEADER_UP As String = "工程: UP"
Private Const COLUMN_HEADER_OP As String = "サブ工程: OP"
Private Const COLUMN_HEADER_ID As String = "ID"
Private Const COLUMN_HEADER_PH As String = "工程: PH"
Private Const COLUMN_HEADER_COMMENT As String = "Comment, transitions, conditions"
Private Const COLUMN_HEADER_RECIPEPARAMETER As String = "レシピパラメータ"
Private Const COLUMN_HEADER_MATERIAL As String = "マテリアル"
Private Const COLUMN_HEADER_EQUIPMENT As String = "機器"
Private Const COLUMN_HEADER_PLACE As String = "場所"
Private Const COLUMN_HEADER_GMP As String = "GMP署名"

Private mPRS() As PRSPR
Private mPRIndex As Long

Public ColumnPR As New Column
Public ColumnUP As New Column
Public ColumnOP As New Column
Public ColumnID As New Column
Public ColumnPH As New Column
Public ColumnComment As New Column
Public ColumnRecipeParameter As New Column
Public ColumnMaterial As New Column
Public ColumnEquipment As New Column
Public ColumnPlace As New Column
Public ColumnGMP As New Column

Public Property Get PRs()
    PRs = mPRS
End Property

Private Sub Class_Initialize()

    mPRIndex = 0

    Call ColumnPR.Init(COLUMN_HEADER_PR)
    Call ColumnUP.Init(COLUMN_HEADER_UP)
    Call ColumnOP.Init(COLUMN_HEADER_OP)
    Call ColumnID.Init(COLUMN_HEADER_ID)
    Call ColumnPH.Init(COLUMN_HEADER_PH)
    Call ColumnComment.Init(COLUMN_HEADER_COMMENT)
    Call ColumnRecipeParameter.Init(COLUMN_HEADER_RECIPEPARAMETER)
    Call ColumnMaterial.Init(COLUMN_HEADER_MATERIAL)
    Call ColumnEquipment.Init(COLUMN_HEADER_EQUIPMENT)
    Call ColumnPlace.Init(COLUMN_HEADER_PLACE)
    Call ColumnGMP.Init(COLUMN_HEADER_GMP)

End Sub

Public Sub Load(ByRef iWs As Worksheet)

    Dim i As Long
    Dim rowIndex As Long
    Dim rowIndexLast As Long

    i = 0

    Call ColumnPR.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnUP.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnOP.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnID.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnPH.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnComment.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnRecipeParameter.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnMaterial.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnEquipment.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnPlace.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call ColumnGMP.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
       
    Dim PR As PRSPR
    Dim UP As PRSUP
    Dim OP As PRSOP
    Dim ph As PRSPH
   
    rowIndexLast = iWs.Cells(iWs.Rows.Count, ColumnID.Index).End(xlUp).row

    Erase mPRS
    mPRIndex = 0
    For rowIndex = ROW_INDEX_DATA To rowIndexLast
        Select Case True
            Case iWs.Cells(rowIndex, ColumnPR.Index).value <> ""
                ReDim Preserve mPRS(mPRIndex)
                Set PR = New PRSPR
                Set mPRS(mPRIndex) = PR
                PR.Text = iWs.Cells(rowIndex, ColumnPR.Index).value
                mPRIndex = mPRIndex + 1
            Case iWs.Cells(rowIndex, ColumnUP.Index).value <> ""
                Set UP = New PRSUP
                UP.Text = iWs.Cells(rowIndex, ColumnUP.Index).value
                UP.ID = iWs.Cells(rowIndex, ColumnID.Index).value
                Call PR.AddUP(UP)
            Case iWs.Cells(rowIndex, ColumnOP.Index).value <> ""
                Set OP = New PRSOP
                OP.Text = iWs.Cells(rowIndex, ColumnOP.Index).value
                OP.ID = iWs.Cells(rowIndex, ColumnID.Index).value
                Call UP.AddOP(OP)
            Case iWs.Cells(rowIndex, ColumnPH.Index).value <> ""
                Set ph = New PRSPH
                With ph
                    Call .Init(PR, UP, OP)
                    .Text = iWs.Cells(rowIndex, ColumnPH.Index).value
                    .ID = iWs.Cells(rowIndex, ColumnID.Index).value
                    .Comment = iWs.Cells(rowIndex, ColumnComment.Index).value
                    .RecipeParameter = iWs.Cells(rowIndex, ColumnRecipeParameter.Index).value
                    .Material = iWs.Cells(rowIndex, ColumnMaterial.Index).value
                    .Equipment = iWs.Cells(rowIndex, ColumnEquipment.Index).value
                    .Place = iWs.Cells(rowIndex, ColumnPlace.Index).value
                    .GMP = iWs.Cells(rowIndex, ColumnGMP.Index).value
                End With
                Call OP.AddPH(ph)
        End Select
    Next
    
    
End Sub
