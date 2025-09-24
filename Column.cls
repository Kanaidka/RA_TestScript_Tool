VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'クラス　Column
'列情報を管理します

Private mIndex As Long
Private mName As String
Private mCount As Long

'プロパティ　Name(Get)
'列名を取得します。
Public Property Get Name() As String
    Name = mName
End Property

'プロパティ　Index(Get)
'列番号を取得します
Public Property Get Index() As Long
    Index = mIndex
End Property

'プロパティ　Count(Get)
'結合セル時の結合列数を取得します
Public Property Get Count() As Long
    Count = mCount
End Property

'関数　Init
'初期化処理
'列名を設定します。
Public Sub Init(ByVal iName As String)
    mName = iName
End Sub

'関数　SetIndexPrefix
'引数　iWs　取込対象ワークシート
'引数　iRowIndex　取込対象行
'取込対象ワークシートの取込対象行前方より列名と前方一致する列番号を設定する
'結合セルの場合は結合列数も設定する
Public Sub SetIndexPrefix(ByVal iWs As Worksheet, Optional ByVal iRowIndex As Long = 1)
    Dim colIndex As Long
    mIndex = 0
    mCount = 0
    '1列目から以降値の入っていないセルまで
    For colIndex = 1 To iWs.Cells(iRowIndex, iWs.Columns.Count).End(xlToLeft).Column
        '対象セルの値が列名と前方一致する
        If Left(iWs.Cells(iRowIndex, colIndex).Value, Len(mName)) = mName Then
            
            If iWs.Cells(iRowIndex, colIndex).MergeCells Then
                'マージセルの場合
                With iWs.Cells(iRowIndex, colIndex).MergeArea
                    mIndex = .Cells(1, 1).Column
                    mCount = .Columns.Count
                End With
            Else
                '非マージセルの場合
                mIndex = colIndex
            End If
            '検索終了
            Exit For
        End If
    Next
End Sub
