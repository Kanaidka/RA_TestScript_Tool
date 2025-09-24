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
'クラス　PRSData
'PRSデータを管理します

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
Private Const COLUMN_HEADER_PROCESSSEQUENCE As String = "プロセスシーケンス"

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
Public ColumnProcessSequence As New Column

'プロパティ　PRs(Get)
Public Property Get PRs()
    PRs = mPRS
End Property

'イベント　クラスの初期化
'クラスの初期化を行う
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
    Call ColumnProcessSequence.Init(COLUMN_HEADER_PROCESSSEQUENCE)

End Sub

'関数　Load
'引数　iWs　対象ワークシート
'対象ワークシートよりPRS情報を取得する
Public Sub Load(ByRef iWs As Worksheet)

    Dim rowIndex As Long
    Dim rowIndexLast As Long

    '各列番号検索
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
    Call ColumnProcessSequence.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)

    '最終行を取得
    rowIndexLast = iWs.Cells(iWs.Rows.Count, ColumnID.Index).End(xlUp).row
    
    Dim PR As PRSPR
    Dim UP As PRSUP
    Dim OP As PRSOP
    Dim PH As PRSPH

    Erase mPRS
    mPRIndex = 0
    For rowIndex = ROW_INDEX_DATA To rowIndexLast
        Select Case True
            Case PR Is Nothing And iWs.Cells(rowIndex, ColumnPR.Index).Value <> ""
                'PRは最初以降は何が入っていても無視する
                '配列を拡張し、クラスを実体化
                ReDim Preserve mPRS(mPRIndex)
                Set PR = New PRSPR
                Set mPRS(mPRIndex) = PR
                '各項目設定
                PR.text = iWs.Cells(rowIndex, ColumnPR.Index).Value
                'INDEXを追加（1度しかこのCaseに入ってこないよう修正したため、現状は無意味なコード）
                mPRIndex = mPRIndex + 1
            
            Case iWs.Cells(rowIndex, ColumnUP.Index).Value <> ""
                'UP列に文字が入っている場合
                Set UP = New PRSUP
                '各項目設定
                UP.text = iWs.Cells(rowIndex, ColumnUP.Index).Value
                UP.id = iWs.Cells(rowIndex, ColumnID.Index).Value
                '追加
                Call PR.AddUP(UP)
            
            Case iWs.Cells(rowIndex, ColumnOP.Index).Value <> ""
                'OP列に文字が入っている場合
                Set OP = New PRSOP
                '各項目設定
                OP.text = iWs.Cells(rowIndex, ColumnOP.Index).Value
                OP.id = iWs.Cells(rowIndex, ColumnID.Index).Value
                OP.Comment = iWs.Cells(rowIndex, ColumnComment.Index).Value
                'プロセスシーケンスはカンマ区切りで情報保持する
                OP.ProcessSequence = WorksheetFunction.TextJoin(",", False, _
                    iWs.Range( _
                        iWs.Cells(rowIndex, ColumnProcessSequence.Index), _
                        iWs.Cells(rowIndex, ColumnProcessSequence.Index + ColumnProcessSequence.Count - 1) _
                    ) _
                )
                    
                '親が未作成なら作成する
                If UP Is Nothing Then Set UP = New PRSUP: Call PR.AddUP(UP)
                '追加
                Call UP.AddOP(OP)
            
            Case iWs.Cells(rowIndex, ColumnPH.Index).Value <> ""
                'PH列に文字が入っている場合
                Set PH = New PRSPH
                
                With PH
                    
                    '親クラスの逆引きをできるようにする
                    Call .Init(PR, UP, OP)
                    '各項目設定
                    .text = iWs.Cells(rowIndex, ColumnPH.Index).Value
                    .id = iWs.Cells(rowIndex, ColumnID.Index).Value
                    .Comment = iWs.Cells(rowIndex, ColumnComment.Index).Value
                    .RecipeParameter = iWs.Cells(rowIndex, ColumnRecipeParameter.Index).Value
                    .Material = iWs.Cells(rowIndex, ColumnMaterial.Index).Value
                    .Equipment = iWs.Cells(rowIndex, ColumnEquipment.Index).Value
                    .Place = iWs.Cells(rowIndex, ColumnPlace.Index).Value
                    .GMP = iWs.Cells(rowIndex, ColumnGMP.Index).Value
                    'プロセスシーケンスはカンマ区切りで情報保持する
                    .ProcessSequence = WorksheetFunction.TextJoin(",", False, _
                        iWs.Range( _
                            iWs.Cells(rowIndex, ColumnProcessSequence.Index), _
                            iWs.Cells(rowIndex, ColumnProcessSequence.Index + ColumnProcessSequence.Count - 1) _
                        ) _
                    )
                
                End With
                
                '親が未作成なら作成する
                If UP Is Nothing Then Set UP = New PRSUP: Call PR.AddUP(UP)
                If OP Is Nothing Then Set OP = New PRSOP: Call UP.AddOP(OP)
                
                '追加
                Call OP.AddPH(PH)
        End Select
    Next
    
    
End Sub
