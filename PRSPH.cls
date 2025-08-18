VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRSPH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mID As String
Private mKey As String
Private mText As String
Private mComment As String
Private mRecipeParameter As String
Private mMaterial As String
Private mEquipment As String
Private mPlace As String
Private mGMP As String

Public PR As PRSPR
Public UP As PRSUP
Public OP As PRSOP

Public Sub Init(ByRef iPr As PRSPR, ByRef iUp As PRSUP, ByRef iOp As PRSOP)
    Set PR = iPr
    Set UP = iUp
    Set OP = iOp
End Sub

Public Property Get ID() As String
    ID = mID
End Property

Public Property Let ID(ByVal iValue As String)
    mID = iValue
End Property

Public Property Get Key() As String
    Key = mKey
End Property

Public Property Get Text() As String
    Text = mText
End Property

Public Property Let Text(ByVal iValue As String)
    mText = iValue
    mKey = Split(iValue, " ")(0)
End Property

Public Property Get Comment() As String
    Comment = mComment
End Property

Public Property Let Comment(ByVal iValue As String)
    mComment = iValue
End Property

Public Property Get RecipeParameter() As String
    RecipeParameter = mRecipeParameter
End Property

Public Property Let RecipeParameter(ByVal iValue As String)
    mRecipeParameter = iValue
End Property

Public Property Get Material() As String
    Material = mMaterial
End Property

Public Property Let Material(ByVal iValue As String)
    mMaterial = iValue
End Property

Public Property Get Equipment() As String
    Equipment = mEquipment
End Property

Public Property Let Equipment(ByVal iValue As String)
    mEquipment = iValue
End Property

Public Property Get Place() As String
    Place = mPlace
End Property

Public Property Let Place(ByVal iValue As String)
    mPlace = iValue
End Property

Public Property Get GMP() As String
    GMP = mGMP
End Property

Public Property Let GMP(ByVal iValue As String)
    mGMP = iValue
End Property

