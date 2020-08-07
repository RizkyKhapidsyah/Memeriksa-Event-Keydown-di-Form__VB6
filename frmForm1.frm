VERSION 5.00
Begin VB.Form frmForm1 
   Caption         =   "Memeriksa Event Keydown di Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim msg As String
    frmForm1.Cls
    frmForm1.CurrentX = 100
    frmForm1.CurrentY = 100
    msg = "Mendapat event KeyDown. KeyCode = " & KeyCode
    msg = msg & "    Shift = " & Shift
    frmForm1.Print msg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim msg As String
    frmForm1.CurrentX = 100
    frmForm1.CurrentY = 500
    msg = "Mendapat event KeyPress.  KeyAscii = " & _
          KeyAscii
    frmForm1.Print msg
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim msg As String
    frmForm1.CurrentX = 100
    frmForm1.CurrentY = 900
    msg = "Mendapat event KeyUp.    KeyCode = " & _
          KeyCode
    msg = msg & "    Shift = " & Shift
    frmForm1.Print msg
End Sub

Private Sub Form_Load()
    frmForm1.FontTransparent = False
End Sub


