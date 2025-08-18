VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Convert(ByVal iText As String, ByRef iPH As PRSPH) As String

    Dim items() As String
    items = Split(iText, "%")
    Dim ret As String
    Dim i As Long
    
    ret = ""
    For i = 0 To UBound(items)
        If i Mod 2 = 0 Then
            ret = ret & items(i)
        Else
            ret = ret & ConvertItem(items(i), iPH)
        End If
    Next
    Convert = ret

End Function


Public Function ConvertItem(ByVal iText As String, ByRef iPH As PRSPH) As String

    Dim items() As String
    items = Split(iText, ",")
    Dim ret As String
    ret = ""
    
    Select Case items(0)
        Case "PR"
            ret = iPH.PR.Text
        Case "UP"
            ret = iPH.UP.Text
        Case "OP"
            ret = iPH.OP.Text
        Case "ID"
            ret = iPH.ID
        Case "PH"
            ret = iPH.Text
        Case "Comment"
            ret = iPH.Comment
        Case "レシピパラメータ"
            ret = iPH.RecipeParameter
        Case "マテリアル"
            ret = iPH.Material
        Case "機器"
            ret = iPH.Equipment
        Case "場所"
            ret = iPH.Place
        Case "GMP署名"
            ret = iPH.GMP
    End Select
    
    ConvertItem = ret
 
    If UBound(items) < 1 Then Exit Function
 
    Select Case items(1)
        Case "LINE"
            If UBound(items) < 2 Then
                ret = GetLine(ret, 1)
            ElseIf Not IsNumeric(items(2)) Then
                ret = GetLine(ret, 1)
            Else
                ret = GetLine(ret, items(2))
            End If
        Case "KEY"
            If UBound(items) < 2 Then
                ret = GetKeyValue(ret, "")
            Else
                ret = GetKeyValue(ret, items(2))
            End If
        Case "ALL"
    End Select
 
    ConvertItem = ret

End Function


Public Function GetLine(ByVal iText As String, ByVal iNumber As Long)
    Dim iIndex As Long
    Dim lines() As String
    iIndex = iNumber - 1
    lines = Split(iText, vbLf)
    If UBound(lines) < iIndex Then Exit Function
    GetLine = lines(iIndex)
End Function

Public Function GetKeyValue(ByVal iText As String, ByVal iKey As String)
    Dim lines() As String
    Dim items() As String
    Dim delimiter As String
    Dim i As Long
    Dim ret As String
    
    If iKey = "" Then
        GetKeyValue = ""
        Exit Function
    End If
    
    lines = Split(iText, vbLf)
    
    For i = 0 To UBound(lines)
        delimiter = ""
        If InStr(lines(i), ":") > 0 Then
            delimiter = ":"
        ElseIf InStr(lines(i), "：") > 0 Then
            delimiter = "："
        End If
        If delimiter <> "" Then
            items = Split(iText, delimiter)
            ret = Trim(items(1))
            Exit For
        End If
    Next
    
    GetKeyValue = ret

End Function


