VERSION 5.00
Object = "{5756E734-2046-400A-BC65-0E105EC5876E}#1.0#0"; "CAIXATEX.OCX"
Object = "{BA676A3D-9505-4A77-87DC-76025E082864}#1.0#0"; "ETIQUETA.OCX"
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Begin VB.Form fLANPADRAO 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamentos Padrão"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin Mascara.Máscara txtCODIGO 
      Height          =   300
      Left            =   75
      TabIndex        =   0
      Top             =   300
      Width           =   690
      _ExtentX        =   1217
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
   Begin VB.Frame boxDADOS 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   2400
      Left            =   75
      TabIndex        =   7
      Top             =   570
      Width           =   6030
      Begin Etiq.Etiqueta lblDEBITO 
         Height          =   300
         Left            =   810
         TabIndex        =   10
         Top             =   900
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   529
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483625
      End
      Begin VB.CommandButton cmdDELETA 
         BackColor       =   &H00C0C000&
         Caption         =   "&Deleta"
         Height          =   375
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1905
         Width           =   1365
      End
      Begin VB.CommandButton cmdGRAVA 
         BackColor       =   &H00C0C000&
         Caption         =   "&Grava"
         Height          =   375
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1905
         Width           =   1365
      End
      Begin CaixaTexto.Caixa_Texto txtDESCRICAO 
         Height          =   300
         Left            =   90
         TabIndex        =   1
         Top             =   390
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   529
         BackColor       =   16777215
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
      Begin Mascara.Máscara txtDEBITO 
         Height          =   300
         Left            =   90
         TabIndex        =   2
         Top             =   900
         Width           =   690
         _ExtentX        =   1217
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
      Begin Mascara.Máscara txtCREDITO 
         Height          =   300
         Left            =   90
         TabIndex        =   3
         Top             =   1410
         Width           =   690
         _ExtentX        =   1217
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
      Begin Etiq.Etiqueta lblCREDITO 
         Height          =   300
         Left            =   810
         TabIndex        =   12
         Top             =   1410
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   529
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483625
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   1215
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Débito"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   705
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         Height          =   195
         Left            =   75
         TabIndex        =   8
         Top             =   165
         Width           =   720
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   105
      Width           =   495
   End
End
Attribute VB_Name = "fLANPADRAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private wp_Saida As Boolean
Private wp_Cria As Boolean
Private tbLanPadrao As Recordset
Private tbContas As Recordset

Private Sub grava_LanPadrao()
Dim wl_documento As Long
If wp_Cria Then
   If tbLanPadrao.RecordCount = 0 Then
      wl_documento = 1
   Else
      tbLanPadrao.MoveLast
      wl_documento = tbLanPadrao("CODIGO") + 1
   End If
   If Not add_reg(tbLanPadrao) Then
      cmdGrava.SetFocus
      Exit Sub
   End If
   tbLanPadrao("CODIGO") = wl_documento
Else
   If Not edit_reg(tbLanPadrao) Then
      cmdGrava.SetFocus
      Exit Sub
   End If
End If
tbLanPadrao("DESCRICAO") = txtDESCRICAO.Text
tbLanPadrao("CREDITO") = txtCREDITO.VALOR
tbLanPadrao("DEBITO") = txtDebito.VALOR
If Not update_reg(tbLanPadrao) Then
   cmdGrava.SetFocus
   Exit Sub
End If
txtCODIGO.SetFocus
End Sub


Private Sub Mon_LanPadrao()
On Error Resume Next
txtDESCRICAO.Text = tbLanPadrao("DESCRICAO")
txtCREDITO.Text = tbLanPadrao("CREDITO")
If txtCREDITO.VALOR <> 0 Then
   If Loca_Contas(tbContas, txtCREDITO.VALOR) Then
      lblCredito.Caption = tbContas("DESCRICAO")
   End If
End If
txtDebito.Text = tbLanPadrao("DEBITO")
If txtDebito.VALOR <> 0 Then
   If Loca_Contas(tbContas, txtDebito.VALOR) Then
      lblDebito.Caption = tbContas("DESCRICAO")
   End If
End If
End Sub

Private Sub Prepara_Form()
Dim wl_txtcodigo
wl_txtcodigo = txtCODIGO.Text
wp_Saida = True
Call LimpaCaixasTexto(Me)
cmdGrava.Enabled = False
cmdDELETA.Enabled = False
boxDados.Enabled = False
If ShowRetorno Then
   ShowRetorno = False
   txtCODIGO.Text = wl_txtcodigo
   SendKeys "{ENTER}"
End If
End Sub

Private Sub cmdDELETA_Click()
If MsgBox("Confirma a deleção?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbYes Then
   If edit_reg(tbLanPadrao) Then
      tbLanPadrao.Delete
      txtCODIGO.SetFocus
      Exit Sub
   End If
End If
End Sub


Private Sub cmdDELETA_GotFocus()
aviso "Use esse botão para deletar o registro"
End Sub


Private Sub cmdGrava_Click()
If MsgBox("Confirma a gravação?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbYes Then
   Call grava_LanPadrao
   ShowRetorno = False
End If
End Sub

Private Sub cmdGRAVA_GotFocus()
aviso "Use esse botão para gravar o registro"
End Sub


Private Sub Form_Activate()
If Not Abre_LanPadrao(tbLanPadrao) Or _
   Not Abre_PlanoContas(tbContas) Then
   Me.Visible = False: Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   If wp_Saida Then
      Me.Visible = False
      Unload Me
   Else
      txtCODIGO.SetFocus
   End If
End If
End Sub

Private Sub Máscara1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub


Private Sub Form_Load()
centraobj Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
aviso
tbLanPadrao.Close
tbContas.Close
End Sub

Private Sub optDESPESA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub


Private Sub optRECEITA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtCODIGO_GotFocus()
Call Prepara_Form
aviso "<F1> Consulta Lançamentos Padrão"
End Sub

Private Sub txtCODIGO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno As Variant
If KeyCode = vbKeyF1 Then
   wl_Retorno = Most_LanPadrao
   If wl_Retorno <> "" Then txtCODIGO.Text = wl_Retorno
   HomeEnd
   txtCODIGO.SetFocus
End If
End Sub


Private Sub txtCODIGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtCODIGO.Text <> "" And txtCODIGO.Text <> "0" Then
      If Not Loca_LanPadrao(tbLanPadrao, txtCODIGO.Text) Then
         wp_Cria = True
      Else
         Call Mon_LanPadrao
         wp_Cria = False
      End If
   Else
      wp_Cria = True
   End If
   If wp_Cria Then
      If Not Verifica_Privilegio(PR_LANPADRAO, "I") Then
         Call MsgBox("Sem privilégio para inclusão de Lançamento Padrão", vbExclamation, "Mensagem do Sistema")
         txtCODIGO.SetFocus
         Exit Sub
      End If
      cmdGrava.Enabled = True
   Else
      If Verifica_Privilegio(PR_LANPADRAO, "A") Then
         cmdGrava.Enabled = True
      End If
   End If
   If Not wp_Cria Then cmdDELETA.Enabled = Verifica_Privilegio(PR_LANPADRAO, "D")
   boxDados.Enabled = True
   SendKeys "{tab}"
End If
End Sub

Private Sub txtCODIGO_LostFocus()
aviso
wp_Saida = False
End Sub



Private Sub txtCREDITO_GotFocus()
aviso "<F1> Consulta Plano de Contas"
End Sub

Private Sub txtCREDITO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtCREDITO.SetFocus
   If wl_Retorno <> "" Then
      txtCREDITO.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If
End Sub

Private Sub txtCREDITO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtCREDITO.Text <> "" And txtCREDITO.Text <> "0" Then
      If Not Loca_Contas(tbContas, txtCREDITO.Text) Then
         InformaaoUsuario "Conta não encontrada"
         lblDebito.Caption = ""
         txtCREDITO.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblCredito.Caption = tbContas("DESCRICAO")
   End If
   SendKeys "{TAB}"
End If

End Sub


Private Sub txtCREDITO_LostFocus()
aviso
End Sub

Private Sub txtDEBITO_GotFocus()
aviso "<F1> Consulta Plano de Contas"
End Sub

Private Sub txtDEBITO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtDebito.SetFocus
   If wl_Retorno <> "" Then
      txtDebito.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If
End Sub

Private Sub txtDEBITO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtDebito.Text <> "" And txtDebito.Text <> "0" Then
      If Not Loca_Contas(tbContas, txtDebito.Text) Then
         InformaaoUsuario "Conta não encontrada"
         lblDebito.Caption = ""
         txtDebito.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblDebito.Caption = tbContas("DESCRICAO")
   End If
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtDEBITO_LostFocus()
aviso
End Sub


Private Sub txtDESCRICAO_GotFocus()
aviso "A descrição do Lançamento Padrão"
End Sub






