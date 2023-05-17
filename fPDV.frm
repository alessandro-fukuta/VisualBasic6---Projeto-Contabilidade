VERSION 5.00
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Begin VB.Form fPDV 
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3690
   ControlBox      =   0   'False
   HelpContextID   =   350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin Mascara.Máscara txtPDV 
      Height          =   300
      Left            =   1395
      TabIndex        =   1
      Top             =   1020
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "###"
      ÉValor          =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informe o Número do Terminal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3150
   End
End
Attribute VB_Name = "fPDV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtPDV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If MsgBox("Confirma o número do PDV?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbYes Then
      Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "PDV", txtPdv.Text
      Unload Me
   Else
      txtPdv.SetFocus
      HomeEnd
   End If
ElseIf KeyAscii = 27 Then
   End
End If
End Sub


