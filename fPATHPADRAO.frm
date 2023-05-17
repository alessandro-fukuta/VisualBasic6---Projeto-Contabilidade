VERSION 5.00
Begin VB.Form fPATHPADRAO 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Cria Pasta"
      Height          =   375
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   795
      Width           =   1410
   End
   Begin VB.TextBox txtPATH 
      Height          =   300
      Left            =   75
      MaxLength       =   30
      TabIndex        =   1
      Top             =   285
      Width           =   5325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informe Path Padrão"
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1515
   End
End
Attribute VB_Name = "fPATHPADRAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Text1_Change()

End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   End
ElseIf KeyAscii = 13 Then
   SendKeys "{Tab}"
End If
End Sub


Private Sub Command1_Click()
On Error Resume Next
Dim wl_File As String
wl_File = PathWindows + pb_Sistema + ".ini"
Err = 0
If Dir(txtPATH, vbDirectory) <> "" Then
   If MsgBox("A pasta informada já existe. Continua?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbNo Then
      HomeEnd
      txtPATH.SetFocus
      Exit Sub
   End If
Else
   MkDir txtPATH
   If Err <> 0 Then
      Call MsgBox("Não foi possível criar a pasta", vbCritical, "Mensagem do Sistema")
      HomeEnd
      txtPATH.SetFocus
      Exit Sub
   End If
End If
If Right(txtPATH, 1) <> "\" Then
   txtPATH.Text = txtPATH.Text + "\"
End If
Grava_Configuracoes "PREFERENCIAS", "PathPadrao", Chr(34) + txtPATH.Text + Chr(34)
PathPadrao = txtPATH.Text
Unload Me
End Sub


Private Sub Form_Activate()
If pb_Demonstracao Then
   PathPadrao = "\INFOSOFT\"
   Unload Me
End If
End Sub

