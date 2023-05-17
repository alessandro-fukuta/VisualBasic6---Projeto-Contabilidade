VERSION 5.00
Begin VB.Form fPadraoI 
   Caption         =   "Padrão da Impressão"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3495
   HelpContextID   =   300
   Icon            =   "fPadraoI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "&Vídeo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   285
      TabIndex        =   2
      Top             =   960
      Width           =   2730
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1365
      Width           =   1365
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Padrão &Jato de Tinta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   600
      Width           =   3090
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Padrão &Matricial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   285
      TabIndex        =   0
      Top             =   285
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1365
      Width           =   1365
   End
End
Attribute VB_Name = "fPadraoI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
SetPrc 0, 0
pb_CancelaImpressao = False
If Mid(pb_Impressao_Normal, 1, 5) = "Draft" Then
   pb_ImpressaoMatricial = True
   pb_UltimaLinha = 0

   Open Printer.Port For Output As #1
   
Else
   pb_ImpressaoMatricial = False
   If Option3.Value Then
      If Dir(PathWindows + "RTF", vbDirectory) = "" Then
         MkDir PathWindows + "RTF"
      End If
      If Dir(PathWindows + "RTF\REPORT.RTF") <> "" Then Kill (PathWindows + "RTF\REPORT.RTF")
      Open PathWindows + "RTF\REPORT.RTF" For Append As #1
      Print #1, "{\rtf1\ansi {\fonttbl{\f0\fmodern\fprq1\fcharset255 " + pb_Impressao_Normal + ";}}"
   End If
End If
Unload Me
End Sub

Private Sub Command2_Click()
pb_CancelaImpressao = True
Unload Me
End Sub


Private Sub Form_Activate()
Inicializa_Impressora
If pb_NaoMatricial Then Option1.Enabled = False
If pb_NaoJato Then Option2.Enabled = False
If pb_NaoVideo Then Option3.Enabled = False
If Not pb_NaoVideo Then Option3_Click
If Not pb_NaoJato Then Option2_Click
If Not pb_NaoMatricial Then Option1_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   pb_CancelaImpressao = True
   Unload Me
End If
End Sub

Private Sub Option1_Click()
If Dir(PathPadrao + "ITUVEPLAST.SYS") = "" Then
   pb_Impressao_Normal = "Draft 12cpi"
   pb_Impressao_Condensada = "Draft 20cpi"
   pb_Impressao_Expandida = "Draft 6cpi"
   pb_Impressao_Condensada_N = "Sans Serif 20cpi"
   pb_Impressao_Normal_N = "Sans Serif 12cpi"
   pb_Impressao_Expandida_N = "Sans Serif 6cpi"
Else
   pb_Impressao_Normal = "Draft 12cpi"
   pb_Impressao_Condensada = "Draft 10cpi"
   pb_Impressao_Expandida = "Draft 11cpi"
   pb_Impressao_Normal_N = "Draft 12cpi"
   pb_Impressao_Condensada_N = "Draft 10cpi"
   pb_Impressao_Expandida_N = "Draft 11cpi"
End If
Option1.Value = True
pb_PadraoVideo = False
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   Command1_Click
End If
End Sub


Private Sub Option2_Click()
pb_Impressao_Normal = "Courier New"
pb_Impressao_Condensada = "Courier New"
pb_Impressao_Expandida = "Courier New"
pb_Impressao_Condensada_N = "Courier New"
pb_Impressao_Normal_N = "Courier New"
pb_Impressao_Expandida_N = "Courier New"
Option2.Value = True
pb_PadraoVideo = False
End Sub


Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   pb_CancelaImpressao = False
   pb_PadraoVideo = False
   Unload Me
End If
End Sub

Private Sub Option3_Click()
On Error Resume Next
pb_Impressao_Normal = "Terminal"
pb_Impressao_Condensada = "Courier New"
pb_Impressao_Expandida = "Courier New"
pb_Impressao_Condensada_N = "Courier New"
pb_Impressao_Normal_N = "Courier New"
pb_Impressao_Expandida_N = "Courier New"
pb_PadraoVideo = True
Option3.Value = True
End Sub

Private Sub Option3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   pb_PadraoVideo = True
   pb_CancelaImpressao = False
   Me.Command1.SetFocus
   Exit Sub
End If
End Sub


