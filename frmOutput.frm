VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOutput 
   ClientHeight    =   7764
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13860
   OleObjectBlob   =   "frmOutput.frx":0000
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMaxValue As Long
Private mNowValue As Long

Public Property Let Header(ByVal iValue As String)
    
    lblHeader.Caption = iValue
    
     Call Refresh
     
End Property

Public Property Let MaxValue(ByVal iValue As Long)
    
    mMaxValue = iValue

End Property

Public Property Get NowValue() As Long
    
    NowValue = mNowValue
    
End Property

Public Property Let NowValue(ByVal iValue As Long)
    
    mNowValue = iValue

    lblNow.Width = Round(lblMax.Width * mNowValue / mMaxValue, 0)
    lblProgress.Caption = mNowValue & " / " & mMaxValue

    Call Refresh

End Property

Public Sub AddLog(ByVal iText As String)

    With txtLog
        Call .SetFocus
        .Text = .Text & Format(Now, "yyyy/mm/dd hh:nn:ss") & " : " & iText & vbCrLf
        .SelStart = Len(.Text)
        .CurLine = .LineCount - 1
    End With
    
    Call Refresh

End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    MaxValue = 1
    NowValue = 0
End Sub

Private Sub Refresh()
    Call Me.Repaint
    Dim b As Boolean
    b = Application.ScreenUpdating
    Application.ScreenUpdating = True
    Application.ScreenUpdating = b
End Sub
