VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mName As String
Private mText As String
Private mIndex As Long

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get Index() As Long
    Index = mIndex
End Property

Public Property Get Text() As String
    Text = mText
End Property


Public Sub Init(ByVal iName As String)
    mName = iName
End Sub

Public Sub SetCellData(ByVal iText As String, iIndex As Long)
    mText = iText
    mIndex = iIndex
End Sub

Public Sub SetIndexPrefix(ByVal iWs As Worksheet, Optional ByVal iRowIndex As Long = 1)
    For colIndex = 1 To iWs.Cells(iRowIndex, iWs.Columns.Count).End(xlToLeft).Column
        If Left(iWs.Cells(iRowIndex, colIndex).value, Len(mName)) = mName Then
            mIndex = colIndex
            Exit For
        End If
    Next
End Sub
