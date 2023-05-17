VERSION 5.00
Begin VB.Form fIMPRESSAO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Impressão no Vídeo"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   330
   Icon            =   "fIMPRESSAO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
End
Attribute VB_Name = "fIMPRESSAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   Unload Me
End If
End Sub

Private Sub Form_Load()
Top = 0
Left = 0



End Sub


