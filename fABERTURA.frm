VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form fABERTURA 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fABERTURA.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox boxREGISTRO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   4440
      ScaleHeight     =   1845
      ScaleWidth      =   5070
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   5100
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C000&
         Caption         =   "&Registra"
         Height          =   375
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1215
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C000&
         Caption         =   "&Cancela"
         Height          =   375
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1215
         Width           =   1140
      End
      Begin MSMask.MaskEdBox txtCONTRASENHA 
         Height          =   300
         Left            =   75
         TabIndex        =   0
         Top             =   795
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   16711680
         MaxLength       =   27
         Mask            =   "##### - #### - ### - ## - #"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Controle"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   1410
      End
      Begin VB.Label lblSENHA 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   75
         TabIndex        =   6
         Top             =   255
         Width           =   4335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retorno"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   5
         Top             =   570
         Width           =   570
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2850
      Top             =   3690
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   675
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   7170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde. Inicializando aplicativo ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   5580
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   7530
   End
End
Attribute VB_Name = "fABERTURA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conta As Integer
Private wp_DIGITO As Integer
Private wp_Senha As String

Private Sub boxREGISTRO_GotFocus()
txtCONTRASENHA.Text = wp_Senha

End Sub

Private Sub Command1_Click()
On Error Resume Next

If txtCONTRASENHA.Text = wp_Senha Then
   InformaaoUsuario "Sistema Registrado. Reinicie"
   If Dir("\MDB_DEMO\99999\*.*") <> "" Then
      Kill "\MDB_DEMO\99999\*.*"
   End If
   If Dir("\MDB_DEMO\SEGURANCA\*.*") <> "" Then
      Kill "\MDB_DEMO\SEGURANCA\*.*"
   End If
   If Dir("\MDB_DEMO\99999", vbDirectory) <> "" Then
      RmDir "\MDB_DEMO\99999"
   End If
   If Dir("\MDB_DEMO\SEGURANCA", vbDirectory) <> "" Then
      RmDir "\MDB_DEMO\SEGURANCA"
   End If
   If Dir("\MDB_DEMO", vbDirectory) <> "" Then
      RmDir "\MDB_DEMO"
   End If
   Grava_Configuracoes pb_Sistema, "TipoExecucao", 1, "CGS"
   pb_Demonstracao = False
   fAgradece.Show 1
   Unload Me
Else
   MsgBox "Contra-Senha Inválida"
   txtCONTRASENHA.SetFocus
End If

End Sub


Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Activate()
Dim wl_Retorno
Dim i As Integer
Dim wl_Soma As Integer
Randomize
On Error Resume Next
Label5.Caption = pb_Sistema
Label4.Caption = pb_Sistema
wl_Retorno = RetornaConfiguracao(pb_Sistema, "TipoExecucao", "CGS.INI")
If wl_Retorno = "" Then
   Grava_Configuracoes pb_Sistema, "TipoExecucao", "0", "CGS"
   pb_Demonstracao = True
ElseIf wl_Retorno = "0" Then
   pb_Demonstracao = True
End If
pausa 3
If pb_Demonstracao Then
   pausa 1
   lblSENHA = lblSENHA + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + " - "
   lblSENHA = lblSENHA + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + " - "
   lblSENHA = lblSENHA + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + " - "
   lblSENHA = lblSENHA + Trim(Str(Int(Rnd * 9))) + Trim(Str(Int(Rnd * 9))) + " - "
   lblSENHA = lblSENHA + Trim(Str(Int(Rnd * 9)))
   wp_DIGITO = Val(Mid(lblSENHA.Caption, 11, 1))
   wp_Senha = ""
   For i = 1 To Len(lblSENHA)
      If InStr("123456789", Mid(lblSENHA, i, 1)) <> 0 Then
         wl_Soma = Val(Mid(lblSENHA.Caption, i, 1)) + wp_DIGITO
         If wl_Soma <= 9 Then
            wp_Senha = wp_Senha + Mid(Format(10 - wl_Soma, "00"), 2, 1)
         Else
            wp_Senha = wp_Senha + Mid(Format(wl_Soma, "00"), 2, 1)
         End If
      Else
         wp_Senha = wp_Senha + Mid(lblSENHA, i, 1)
      End If
   Next
   
   If Confirme("Deseja registrar agora essa cópia?") Then
      Me.boxREGISTRO.Visible = True
   Else
      pausa 1
      Timer2.Enabled = True
   End If
Else
   Timer2.Enabled = True
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If pb_Demonstracao Then Exit Sub
Unload Me
End Sub

Private Sub Form_Load()
Dim wl_Retorno
If pb_Demonstracao Then
   Timer2.Enabled = False
   Label1.Caption = "Registre o aplicativo"
End If
pb_MaquinaReady = Dir("\FONTES", vbDirectory) <> ""
End Sub

Private Sub Timer2_Timer()
Label1 = "Pronto ..."
pausa 1
Unload Me
DoEvents
Exit Sub
End Sub


