VERSION 5.00
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Begin VB.Form fEmpresa_Padrao 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1875
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   260
   Icon            =   "fEmpresa_Padrao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "&Manutenção"
      Height          =   375
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1380
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Ca&ncela"
      Height          =   375
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1380
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "&Confirma"
      Height          =   375
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1380
      Width           =   1350
   End
   Begin Mascara.Máscara txtEMPRESA 
      Height          =   300
      Left            =   135
      TabIndex        =   1
      Top             =   990
      Width           =   585
      _ExtentX        =   1032
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
      Mask            =   "#####"
      Text            =   ""
      ÉValor          =   -1  'True
   End
   Begin VB.Label lblEMPRESA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   735
      TabIndex        =   6
      Top             =   990
      Width           =   4110
   End
   Begin VB.Label lblTitulo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Empresa Padrão"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3240
      TabIndex        =   5
      Top             =   255
      Width           =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4935
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   630
      Left            =   0
      Top             =   0
      Width           =   4965
   End
End
Attribute VB_Name = "fEmpresa_Padrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tbEmpresas As Recordset
Private wp_Saida As Boolean
Private wp_Entrada As Boolean
Private Sub Command1_Click()
Dim wl_Retorno
On Error Resume Next
pb_Empresa = txtEMPRESA.VALOR
pb_RAZAOSOCIAL = IIf(Not pb_Demonstracao, tbEmpresas("RAZAOSOCIAL"), "Empresa Demonstração")
pb_FONEEMPRESA = IIf(Not pb_Demonstracao, tbEmpresas("FONE1"), "XXX")
pb_InverteOperacao = IIf(Not pb_Demonstracao, tbEmpresas("INVERTEOPERACOES"), True)
pb_Endereco = IIf(Not pb_Demonstracao, tbEmpresas("ENDERECO"), "Av Dr Soares Oliveira, ###")
pb_Cidade = IIf(Not pb_Demonstracao, tbEmpresas("CIDADE"), "ITUVERAVA - SP")
pb_Estado = IIf(Not pb_Demonstracao, IIf(IsNull(tbEmpresas("ESTADO")), "SP", tbEmpresas("ESTADO")), "SP")
Grava_Configuracoes "PREFERENCIAS", "EMPRESA_PADRAO", txtEMPRESA.VALOR
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "PADRAOPLANO")
If wl_Retorno = "" And pb_Sistema = "FATO" Then
   If Estrutura_PlanoContas Then
      If Confirme("Deseja criar o plano de contas padrao da empresa?") Then
         fPadraoPlano.Show 1
      Else
         Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "PADRAOPLANO", "1"
      End If
   End If
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command3_Click()
fEMPRESAS.Show 1
txtEMPRESA.Text = ""
txtEMPRESA.SetFocus
End Sub

Private Sub Form_Activate()
Dim wl_Retorno
If Not Estrutura_Empresas Then
   Unload Me
   Exit Sub
End If
If Not Abre_Empresas(tbEmpresas) Then
   Unload Me
   Exit Sub
End If
If Not pb_Demonstracao Then
   wl_Retorno = RetornaConfiguracao("PREFERENCIAS", "EMPRESA_PADRAO")
   If wl_Retorno <> "" Then
      pbRetornoVideo = wl_Retorno
      pb_FormAtivo = Me.Name
      pb_ObjetoAtivo = Me.ActiveControl.Name
      Me.txtEMPRESA.Text = wl_Retorno
   End If
   Me.Command3.Enabled = Verifica_Privilegio(PR_EMPRESAS, "C")
   SendKeys "{ENTER}"
Else
   Me.Command3.Enabled = False
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   If wp_Saida Then
      Unload Me
   Else
      txtEMPRESA.SetFocus
   End If
End If
End Sub


Private Sub Form_Paint()
Demarca fEmpresa_Padrao
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbEmpresas.Close
End Sub


Private Sub txtEMPRESA_GotFocus()
aviso "<F1> Consulta Empresas"
If Not wp_Entrada Then
   wp_Entrada = True
Else
   txtEMPRESA.Text = ""
   lblEMPRESA.Caption = ""
End If
wp_Saida = True
If pb_Demonstracao Then txtEMPRESA.Text = "99999"
End Sub

Private Sub txtEMPRESA_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   If Not pb_Demonstracao Then
      wl_Retorno = Most_Empresas
      txtEMPRESA.SetFocus
      If wl_Retorno <> "" Then
         txtEMPRESA.Text = wl_Retorno
         SendKeys "{ENTER}"
      End If
   Else
      InformaaoUsuario "Informe a empresa 99999 para Demonstracao"
   End If
End If
End Sub


Private Sub txtEMPRESA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtEMPRESA.Text = "" Or txtEMPRESA.Text = "0" Then
      InformaaoUsuario "Informe o código da empresa"
      txtEMPRESA.SetFocus
      HomeEnd
      Exit Sub
   End If
   If txtEMPRESA.Text = "99999" Then
      Me.lblEMPRESA.Caption = "EMPRESA DEMONSTRAÇÃO"
      SendKeys "{TAB}"
      Exit Sub
   End If
   If Not Loca_Empresas(tbEmpresas, txtEMPRESA.VALOR) Then
      InformaaoUsuario "Empresa não encontrada"
      txtEMPRESA.SetFocus
      HomeEnd
      Exit Sub
   End If
   lblEMPRESA.Caption = tbEmpresas("RAZAOSOCIAL")
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtEMPRESA_LostFocus()
wp_Saida = False
End Sub


