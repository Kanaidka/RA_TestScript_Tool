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
'クラス　PRSPH
'PRSのPHデータを管理します

Private mID As String
Private mKey As String
Private mText As String
Private mComment As String
Private mRecipeParameter As String
Private mMaterial As String
Private mEquipment As String
Private mPlace As String
Private mGMP As String
Private mProcessSequence As String

Public PR As PRSPR
Public UP As PRSUP
Public OP As PRSOP

'プロパティ　ID(Get)
Public Property Get id() As String
    id = mID
End Property

'プロパティ　ID(Let)
Public Property Let id(ByVal iValue As String)
    mID = iValue
End Property

'プロパティ　key(Get)
Public Property Get key() As String
    key = mKey
End Property

'プロパティ　Text(Get)
Public Property Get text() As String
    text = mText
End Property

'プロパティ　Text(Let)
Public Property Let text(ByVal iValue As String)
    mText = iValue
    mKey = Split(iValue, " ")(0) 'スペースより前をKeyとして保持しておく
End Property

'プロパティ　Comment(Get)
Public Property Get Comment() As String
    Comment = mComment
End Property

'プロパティ　Comment(Let)
Public Property Let Comment(ByVal iValue As String)
    mComment = iValue
End Property

'プロパティ　RecipeParameter(Get)
Public Property Get RecipeParameter() As String
    RecipeParameter = mRecipeParameter
End Property

'プロパティ　RecipeParameter(Let)
Public Property Let RecipeParameter(ByVal iValue As String)
    mRecipeParameter = iValue
End Property

'プロパティ　Material(Get)
Public Property Get Material() As String
    Material = mMaterial
End Property

'プロパティ　Material(Let)
Public Property Let Material(ByVal iValue As String)
    mMaterial = iValue
End Property

'プロパティ　Equipment(Get)
Public Property Get Equipment() As String
    Equipment = mEquipment
End Property

'プロパティ　Equipment(Let)
Public Property Let Equipment(ByVal iValue As String)
    mEquipment = iValue
End Property

'プロパティ　Place(Get)
Public Property Get Place() As String
    Place = mPlace
End Property

'プロパティ　Place(Let)
Public Property Let Place(ByVal iValue As String)
    mPlace = iValue
End Property

'プロパティ　GMP(Get)
Public Property Get GMP() As String
    GMP = mGMP
End Property

'プロパティ　GMP(Let)
Public Property Let GMP(ByVal iValue As String)
    mGMP = iValue
End Property

'プロパティ　ProcessSequence(Get)
Public Property Get ProcessSequence() As String
    ProcessSequence = mProcessSequence
End Property

'プロパティ　ProcessSequence(Let)
Public Property Let ProcessSequence(ByVal iValue As String)
    mProcessSequence = iValue
End Property

'関数　AddPH
'引数　iPR　対象PR
'引数　iUP　対象UP
'引数　iOP　対象OP
'初期化処理　親クラスを参照できるようにする
Public Sub Init(ByRef iPR As PRSPR, ByRef iUP As PRSUP, ByRef iOP As PRSOP)
    Set PR = iPR
    Set UP = iUP
    Set OP = iOP
End Sub

