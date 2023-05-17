VERSION 5.00
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Begin VB.Form fHISTORICO 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "fhistorico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
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
      Height          =   2580
      Left            =   75
      TabIndex        =   5
      Top             =   570
      Width           =   6030
      Begin VB.TextBox txtDESCRICAO 
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   5775
      End
      Begin VB.CommandButton cmdDELETA 
         BackColor       =   &H00C0C000&
         Caption         =   "&Deleta"
         Height          =   375
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2100
         Width           =   1365
      End
      Begin VB.CommandButton cmdGRAVA 
         BackColor       =   &H00C0C000&
         Caption         =   "&Grava"
         Height          =   375
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2100
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         Height          =   195
         Left            =   75
         TabIndex        =   6
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
      TabIndex        =   4
      Top             =   105
      Width           =   495
   End
End
Attribute VB_Name = "fHISTORICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private wp_Saida As Boolean
Private wp_Cria As Boolean
Private tbHistorico As Recordset


Private Sub grava_Historico()
Dim wl_documento As Long
If wp_Cria Then
   If Not add_reg(tbHistorico) Then
      cmdGrava.SetFocus
      Exit Sub
   End If
   tbHistorico("CODIGO") = txtCODIGO.VALOR
Else
   If Not edit_reg(tbHistorico) Then
      cmdGrava.SetFocus
      Exit Sub
   End If
End If
tbHistorico("DESCRICAO") = txtDESCRICAO.Text
If Not update_reg(tbHistorico) Then
   cmdGrava.SetFocus
   Exit Sub
End If
txtCODIGO.SetFocus
End Sub


Private Sub Mon_Historico()
txtDESCRICAO.Text = tbHistorico("DESCRICAO")
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
   If edit_reg(tbHistorico) Then
      tbHistorico.Delete
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
   Call grava_Historico
End If
End Sub

Private Sub cmdGRAVA_GotFocus()
aviso "Use esse botão para gravar o registro"
End Sub


Private Sub Form_Activate()
If Not Abre_Historico(tbHistorico) Then
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
      If UCase(Me.ActiveControl.Name) = "TXTDESCRICAO" Then
         SendKeys "{TAB}"
      Else
         txtCODIGO.SetFocus
      End If
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
tbHistorico.Close
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
aviso "<F1> Consulta Hist¢rico"
End Sub

Private Sub txtCODIGO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno As Variant
If KeyCode = vbKeyF1 Then
   wl_Retorno = Most_Historico
   If wl_Retorno <> "" Then txtCODIGO.Text = wl_Retorno
   HomeEnd
   txtCODIGO.SetFocus
End If
End Sub


Private Sub txtCODIGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtCODIGO.Text <> "" And txtCODIGO.Text <> "0" Then
      If Not Loca_Historico(tbHistorico, txtCODIGO.Text) Then
         wp_Cria = True
      Else
         Call Mon_Historico
         wp_Cria = False
      End If
   Else
      Call MsgBox("Informe o código", vbExclamation, "Mensagem do Sistema")
      txtCODIGO.SetFocus
      Exit Sub
   End If
   If wp_Cria Then
      If Not Verifica_Privilegio(PR_HISTORICO, "I") Then
         Call MsgBox("Sem privilégio para inclusão", vbExclamation, "Mensagem do Sistema")
         txtCODIGO.SetFocus
         Exit Sub
      End If
      cmdGrava.Enabled = True
   Else
      If Verifica_Privilegio(PR_HISTORICO, "A") Then
         cmdGrava.Enabled = True
      End If
   End If
   If Not wp_Cria Then cmdDELETA.Enabled = Verifica_Privilegio(PR_HISTORICO, "D")
   boxDados.Enabled = True
   SendKeys "{tab}"
End If
End Sub

Private Sub txtCODIGO_LostFocus()
aviso
wp_Saida = False
End Sub




