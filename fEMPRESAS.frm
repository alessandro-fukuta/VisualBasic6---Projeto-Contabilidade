VERSION 5.00
Object = "{5756E734-2046-400A-BC65-0E105EC5876E}#1.0#0"; "CAIXATEX.OCX"
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Begin VB.Form fEMPRESAS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresas"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Mascara.Máscara txtCODIGO 
      Height          =   300
      Left            =   150
      TabIndex        =   0
      Top             =   255
      Width           =   660
      _ExtentX        =   1164
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
   Begin VB.Frame boxdados 
      Enabled         =   0   'False
      Height          =   5640
      Left            =   30
      TabIndex        =   20
      Top             =   480
      Width           =   10455
      Begin VB.OptionButton optINVERTE 
         Caption         =   "Débito = Entrada (Modo Contábil)"
         Height          =   255
         Left            =   2970
         TabIndex        =   37
         Top             =   3450
         Value           =   -1  'True
         Width           =   2730
      End
      Begin VB.OptionButton optNORMAL 
         Caption         =   "Débito = Saída"
         Height          =   255
         Left            =   1335
         TabIndex        =   36
         Top             =   3435
         Width           =   1515
      End
      Begin CaixaTexto.Caixa_Texto txtCIDADE 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   1365
         Width           =   5535
         _ExtentX        =   9763
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
         Text            =   ""
         MaxLength       =   50
      End
      Begin CaixaTexto.Caixa_Texto txtRAZAOSOCIAL 
         Height          =   300
         Left            =   1470
         TabIndex        =   2
         Top             =   360
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   529
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
      End
      Begin Mascara.Máscara txtDATA 
         Height          =   300
         Left            =   100
         TabIndex        =   1
         Top             =   345
         Width           =   1245
         _ExtentX        =   2196
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
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         ÉData           =   -1  'True
      End
      Begin VB.CommandButton cmdDeleta 
         BackColor       =   &H00C0C000&
         Caption         =   "&Deleta"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1695
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5145
         Width           =   1400
      End
      Begin VB.CommandButton cmdGrava 
         BackColor       =   &H00C0C000&
         Caption         =   "&Grava"
         Enabled         =   0   'False
         Height          =   375
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5145
         Width           =   1400
      End
      Begin VB.TextBox txtOBSERVACAO 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   3945
         Width           =   9195
      End
      Begin CaixaTexto.Caixa_Texto txtNOMEFANTASIA 
         Height          =   300
         Left            =   5895
         TabIndex        =   3
         Top             =   360
         Width           =   4380
         _ExtentX        =   7726
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
         Text            =   ""
      End
      Begin CaixaTexto.Caixa_Texto txtCGC 
         Height          =   300
         Left            =   100
         TabIndex        =   4
         Top             =   855
         Width           =   2940
         _ExtentX        =   5186
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
         Text            =   ""
      End
      Begin CaixaTexto.Caixa_Texto txtINSCRICAO 
         Height          =   300
         Left            =   3105
         TabIndex        =   5
         Top             =   855
         Width           =   2940
         _ExtentX        =   5186
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
         Text            =   ""
      End
      Begin Mascara.Máscara txtCEP 
         Height          =   300
         Left            =   6315
         TabIndex        =   8
         Top             =   1365
         Width           =   1005
         _ExtentX        =   1773
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
         Mask            =   "##.###-###"
      End
      Begin CaixaTexto.Caixa_Texto txtENDERECO 
         Height          =   300
         Left            =   100
         TabIndex        =   9
         Top             =   1905
         Width           =   4380
         _ExtentX        =   7726
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
         Text            =   ""
      End
      Begin CaixaTexto.Caixa_Texto txtBAIRRO 
         Height          =   300
         Left            =   4590
         TabIndex        =   10
         Top             =   1905
         Width           =   4380
         _ExtentX        =   7726
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
         Text            =   ""
      End
      Begin CaixaTexto.Caixa_Texto txtTELEFONE 
         Height          =   300
         Left            =   100
         TabIndex        =   11
         Top             =   2400
         Width           =   3000
         _ExtentX        =   5292
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
         Text            =   ""
      End
      Begin CaixaTexto.Caixa_Texto txtTELEFONE2 
         Height          =   300
         Left            =   3195
         TabIndex        =   12
         Top             =   2400
         Width           =   3000
         _ExtentX        =   5292
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
         Text            =   ""
      End
      Begin CaixaTexto.Caixa_Texto txtFAX 
         Height          =   300
         Left            =   6240
         TabIndex        =   13
         Top             =   2385
         Width           =   3000
         _ExtentX        =   5292
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
         Text            =   ""
      End
      Begin CaixaTexto.Caixa_Texto txtEMAIL 
         Height          =   300
         Left            =   100
         TabIndex        =   14
         Top             =   2895
         Width           =   3000
         _ExtentX        =   5292
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
         Text            =   ""
         MaxLength       =   60
         CaixaAlta       =   0   'False
      End
      Begin Mascara.Máscara txtISS 
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   3435
         Width           =   1065
         _ExtentX        =   1879
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
         Format          =   "##0.00"
         Text            =   "_____"
         ÉValor          =   -1  'True
      End
      Begin CaixaTexto.Caixa_Texto txtESTADO 
         Height          =   300
         Left            =   5730
         TabIndex        =   7
         Top             =   1365
         Width           =   465
         _ExtentX        =   820
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
         Text            =   ""
         MaxLength       =   50
      End
      Begin VB.Label Label16 
         Caption         =   "U.F."
         Height          =   210
         Left            =   5730
         TabIndex        =   38
         Top             =   1170
         Width           =   300
      End
      Begin VB.Label Label18 
         Caption         =   "ISS"
         Height          =   210
         Left            =   120
         TabIndex        =   35
         Top             =   3240
         Width           =   630
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Left            =   100
         TabIndex        =   34
         Top             =   2700
         Width           =   435
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         Height          =   195
         Left            =   6240
         TabIndex        =   33
         Top             =   2190
         Width           =   255
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Left            =   3195
         TabIndex        =   32
         Top             =   2205
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Left            =   100
         TabIndex        =   31
         Top             =   2205
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   4590
         TabIndex        =   30
         Top             =   1710
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   100
         TabIndex        =   29
         Top             =   1710
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estadual / RG"
         Height          =   195
         Left            =   3105
         TabIndex        =   28
         Top             =   660
         Width           =   1710
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "C.G.C / CPF"
         Height          =   195
         Left            =   105
         TabIndex        =   27
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia"
         Height          =   195
         Left            =   5895
         TabIndex        =   26
         Top             =   165
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nome / Razão Social"
         Height          =   195
         Left            =   1470
         TabIndex        =   25
         Top             =   165
         Width           =   1530
      End
      Begin VB.Label Label5 
         Caption         =   "&Observações"
         Height          =   210
         Left            =   105
         TabIndex        =   24
         Top             =   3735
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         Height          =   195
         Left            =   6315
         TabIndex        =   23
         Top             =   1170
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "Cidade"
         Height          =   210
         Left            =   100
         TabIndex        =   22
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Cadastro"
         Height          =   195
         Left            =   100
         TabIndex        =   21
         Top             =   150
         Width           =   1020
      End
   End
   Begin VB.Label Label1 
      Caption         =   "&Código"
      Height          =   180
      Left            =   135
      TabIndex        =   19
      Top             =   45
      Width           =   675
   End
End
Attribute VB_Name = "fEMPRESAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private wp_Cria As Boolean
Private tbEmpresas As Recordset
Private tbCidades As Recordset
Private wp_Saida As Boolean
Private wp_edicao As Boolean


Private Function VerificaInconsistencias() As Boolean
If txtRAZAOSOCIAL.Text = "" Then
   InformaaoUsuario "O campo Razão Social é obrigatório."
   txtRAZAOSOCIAL.SetFocus
   VerificaInconsistencias = False
   Exit Function
End If
If txtCIDADE.Text = "" Then
   InformaaoUsuario "O campo Cidade é obrigatório."
   txtCIDADE.SetFocus
   VerificaInconsistencias = False
   Exit Function
End If
If txtESTADO.Text = "" Then
   InformaaoUsuario "O campo Estado é obrigatório."
   txtESTADO.SetFocus
   VerificaInconsistencias = False
   Exit Function
End If
VerificaInconsistencias = True
End Function



Function del_empresas() As Boolean
If Not edit_reg(tbEmpresas) Then
   Call MsgBox("Impossível editar registro", vbCritical, "Mensagem do Sistema")
   cmdDELETA.SetFocus
   Exit Function
End If
tbEmpresas.Delete
txtCODIGO.SetFocus
End Function


Private Function grv_empresas()
If wp_Cria Then
   If Not add_reg(tbEmpresas) Then
      Call MsgBox("Impossível acrescentar registro", vbCritical, "Mensagem do Sistema")
      cmdGrava.SetFocus
      Exit Function
   End If
   tbEmpresas("CODIGO") = txtCODIGO.Text
Else
   If Not edit_reg(tbEmpresas) Then
      Call MsgBox("Impossível editar registro", vbCritical, "Mensagem do Sistema")
      cmdGrava.SetFocus
      Exit Function
   End If
End If
tbEmpresas("DATA") = FormataData(txtDATA.Text)
tbEmpresas("RAZAOSOCIAL") = txtRAZAOSOCIAL.Text
tbEmpresas("NOMEFANTASIA") = txtNOMEFANTASIA.Text
tbEmpresas("CGC") = txtCGC.Text
tbEmpresas("INSCRICAO") = txtINSCRICAO.Text
tbEmpresas("ENDERECO") = txtENDERECO.Text
tbEmpresas("BAIRRO") = txtBAIRRO.Text
tbEmpresas("CIDADE") = IIf(txtCIDADE.Text = "", 0, txtCIDADE.Text)
tbEmpresas("ESTADO") = txtESTADO.Text
tbEmpresas("CEP") = txtCEP.Text
tbEmpresas("FONE1") = txtTELEFONE.Text
tbEmpresas("FONE2") = txtTELEFONE2.Text
tbEmpresas("FAX1") = txtFAX.Text
tbEmpresas("EMAIL") = txtEMAIL.Text
tbEmpresas("ISS") = IIf(txtISS.Text = "", 0, txtISS.Text)
tbEmpresas("OBSERVACOES") = txtOBSERVACAO.Text
tbEmpresas("INVERTEOPERACOES") = Not optNORMAL.Value
If Not update_reg(tbEmpresas) Then
   Call MsgBox("Impossível atualizar tabela", vbCritical, "Mensagem do Sistema")
   cmdGrava.SetFocus
   Exit Function
End If
txtCODIGO.SetFocus
End Function


Private Function LimpaBox()
txtRAZAOSOCIAL.Text = ""
txtNOMEFANTASIA.Text = ""
txtCGC.Text = ""
txtINSCRICAO.Text = ""
txtCIDADE.Text = ""
txtESTADO.Text = ""
txtENDERECO.Text = ""
txtBAIRRO.Text = ""
txtCEP.Text = ""
txtTELEFONE.Text = ""
txtTELEFONE2.Text = ""
txtFAX.Text = ""
txtEMAIL.Text = ""
txtOBSERVACAO.Text = ""
txtDATA.Text = ""
txtISS.Text = ""
cmdGrava.Enabled = False
cmdDELETA.Enabled = False
End Function


Private Function LtoU(ByRef KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
   End If
End Function


Private Function Mon_empresas()
On Error Resume Next
Err = 0
txtRAZAOSOCIAL.Text = tbEmpresas("RAZAOSOCIAL")
txtNOMEFANTASIA.Text = tbEmpresas("NOMEFANTASIA")
txtCGC.Text = tbEmpresas("CGC")
txtINSCRICAO.Text = tbEmpresas("INSCRICAO")
txtCIDADE.Text = tbEmpresas("CIDADE")
txtESTADO.Text = tbEmpresas("ESTADO")
txtENDERECO.Text = tbEmpresas("ENDERECO")
txtBAIRRO.Text = tbEmpresas("BAIRRO")
txtCEP.Text = tbEmpresas("CEP")
txtTELEFONE.Text = tbEmpresas("FONE1")
txtTELEFONE2.Text = tbEmpresas("FONE2")
txtFAX.Text = tbEmpresas("FAX1")
txtEMAIL.Text = tbEmpresas("EMAIL")
txtISS.Text = tbEmpresas("ISS")
txtOBSERVACAO.Text = tbEmpresas("OBSERVACOES")
txtDATA.Text = tbEmpresas("DATA")
If tbEmpresas("INVERTEOPERACOES") Then
   Me.optINVERTE.Value = True
Else
   Me.optNORMAL = True
End If
End Function


Private Sub cmdDELETA_Click()
If MsgBox("Confirma a deleção?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbYes Then
   Call del_empresas
End If
End Sub

Private Sub cmdGrava_Click()
If Not VerificaInconsistencias Then
   Exit Sub
End If
If MsgBox("Confirma a gravação ?", vbQuestion + vbYesNo, "Mensagem do Sistema") Then
   Call grv_empresas
Else
   cmdGrava.SetFocus
End If
End Sub

Private Sub Form_Activate()
If Not Abre_Empresas(tbEmpresas) Then
   Me.Visible = False: Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 And Not wp_edicao Then
   KeyAscii = 0
   If wp_Saida Then
      Me.Visible = False
      Unload Me
   Else
      txtCODIGO.SetFocus
   End If
End If
End Sub

Private Sub Form_Load()
centraobj Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbEmpresas.Close
aviso
End Sub

Private Sub txtCEP_GotFocus()
HomeEnd
End Sub


Private Sub txtCEP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   KeyCode = 0
   SendKeys "+{tab}"
End If
End Sub


Private Sub txtCEP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{tab}"
End If
End Sub


Private Sub txtCIDADE_LostFocus()
Call aviso
End Sub


Private Sub txtCODIGO_GotFocus()
Label1.FontBold = True
wp_Saida = True
If Not ShowRetorno Then
   txtCODIGO.Text = ""
Else
   ShowRetorno = False
   SendKeys "{enter}"
End If
boxDados.Enabled = False
Call LimpaBox
aviso "<F1> Empresas", True
End Sub


Private Sub txtCODIGO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
   txtCODIGO.Text = Most_Empresas
   txtCODIGO.SetFocus
End If
End Sub


Private Sub txtCODIGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtCODIGO.Text <> "" Then
      If Not Loca_Empresas(tbEmpresas, txtCODIGO.Text) Then
         If Not Confirme("Empresa não encontrada. Cadastra?") Then
            HomeEnd
            Exit Sub
         End If
         wp_Cria = True
      Else
         Call Mon_empresas
         wp_Cria = False
      End If
   Else
      If tbEmpresas.RecordCount = 0 Then
         txtCODIGO.Text = 1
      Else
         tbEmpresas.MoveLast
         txtCODIGO.Text = tbEmpresas("CODIGO") + 1
      End If
      wp_Cria = True
   End If
   If wp_Cria And Not Verifica_Privilegio(PR_EMPRESAS, "I") Then
      Call MsgBox("Usuário sem privilégio para Inclusão", vbExclamation, "Mensagem do Sistema")
      txtCODIGO.Text = ""
      txtCODIGO.SetFocus
      Exit Sub
   End If
   boxDados.Enabled = True
   If wp_Cria Or Verifica_Privilegio(PR_EMPRESAS, "A") Then
      cmdGrava.Enabled = True
   End If
   If Not wp_Cria Then
      cmdDELETA.Enabled = Verifica_Privilegio(PR_EMPRESAS, "D")
   End If
   SendKeys "{TAB}"
End If
End Sub

Private Sub txtCODIGO_LostFocus()
wp_Saida = False
aviso
End Sub


Private Sub txtDATA_GotFocus()
Label2.FontBold = True
If wp_Cria Then
   txtDATA.Text = Date
End If
HomeEnd
End Sub


Private Sub txtDATA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   KeyCode = 0
   SendKeys "+{TAB}"
End If
End Sub


Private Sub txtDATA_LostFocus()
Label2.FontBold = False
End Sub


Private Sub txtOBSERVACAO_GotFocus()
wp_edicao = True
End Sub

Private Sub txtOBSERVACAO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   SendKeys "{tab}"
End If
End Sub


Private Sub txtOBSERVACAO_LostFocus()
wp_edicao = False
End Sub

Private Sub txtRAZAOSOCIAL_KeyPress(KeyAscii As Integer)
LtoU KeyAscii
End Sub



