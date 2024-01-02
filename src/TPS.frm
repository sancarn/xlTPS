VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TPS 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   OleObjectBlob   =   "TPS.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Boolean
Private Sub UserForm_Initialize()
  Me.Show
  b = True
  While b
    Dim d As Date: d = Now()
    i = 0
    While d = Now()
      i = i + 1
    Wend
    If b Then Label1.Caption = i
    DoEvents
  Wend
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  b = False
End Sub
