VERSION 5.00
Object = "{5756E734-2046-400A-BC65-0E105EC5876E}#1.0#0"; "CAIXATEX.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fPLANOPROGRAMA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatórios Programáveis de Apresentação do Plano de Contas"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "fPLANOPROGRAMA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSALDO 
      Caption         =   "M&ostra Saldo Calculado"
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   825
      Width           =   2085
   End
   Begin VB.CheckBox chkMES 
      Caption         =   "Mês de Referência"
      Height          =   210
      Left            =   75
      TabIndex        =   2
      Top             =   615
      Width           =   1830
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFFF&
      Height          =   4770
      Left            =   2640
      Pattern         =   "*.REL"
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "..."
      Height          =   285
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.Frame boxLinha 
      Height          =   1260
      Left            =   90
      TabIndex        =   5
      Top             =   4755
      Visible         =   0   'False
      Width           =   6435
      Begin CaixaTexto.Caixa_Texto txtLinha 
         Height          =   300
         Left            =   135
         TabIndex        =   6
         Top             =   345
         Width           =   6165
         _ExtentX        =   10874
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
      Begin CaixaTexto.Caixa_Texto txtDESCRICAO 
         Height          =   300
         Left            =   135
         TabIndex        =   9
         Top             =   840
         Width           =   6150
         _ExtentX        =   10848
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   645
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conta / Fórmula"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   150
         Width           =   1140
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdPLANO 
      Height          =   3705
      Left            =   75
      TabIndex        =   4
      Top             =   1110
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   6535
      _Version        =   393216
      Rows            =   15
      FixedCols       =   0
      ForeColor       =   16711680
      BackColorBkg    =   -2147483633
      BorderStyle     =   0
      FormatString    =   "<Plano de Contas Programável                                                                                                   "
   End
   Begin CaixaTexto.Caixa_Texto txtNOME 
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Width           =   6225
      _ExtentX        =   10980
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Relatório"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   1320
   End
End
Attribute VB_Name = "fPLANOPROGRAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tbPlano As Recordset
Private wp_AddLinha As Boolean
Private wp_DAT As String
Private Sub chkMES_Click()
If chkMES.Value = 0 Then
   wp_DAT = StrTran(wp_DAT, "M", "")
Else
   wp_DAT = wp_DAT + "M"
End If
grdPLANO.SetFocus
End Sub

Private Sub chkMES_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{TAB}"
End Sub


Private Sub chkSALDO_Click()
If chkSALDO.Value = 0 Then
   wp_DAT = StrTran(wp_DAT, "S", "")
Else
   wp_DAT = wp_DAT + "S"
End If
grdPLANO.SetFocus
End Sub


Private Sub chkSALDO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{TAB}"
End Sub


Private Sub Command1_Click()
If Dir(PathPadrao + "RELATORIOS", vbDirectory) = "" Then
   MkDir PathPadrao + "RELATORIOS"
End If
Me.File1.Path = PathPadrao + "RELATORIOS"
File1.Visible = True
File1.SetFocus
SendKeys "{RIGHT}"
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
Dim wl_File As String
If KeyAscii = 27 Then
   KeyAscii = 0
   txtNOME.SetFocus
ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   If InStr(File1.FileName, ".") <> 0 Then
      wl_File = Mid(File1.FileName, 1, InStr(File1.FileName, ".") - 1)
   Else
      wl_File = File1.FileName
   End If
   txtNOME.Text = wl_File
   txtNOME.SetFocus
   HomeEnd
End If
End Sub

Private Sub File1_LostFocus()
File1.Visible = False
End Sub


Private Sub Form_Activate()
If Not Abre_PlanoContas(tbPlano) Then
   Unload Me
   Exit Sub
End If
tbPlano.Index = "iCONTA"
End Sub

Private Sub MSFlexGrid1_Click()

End Sub


Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbPlano.Close
End Sub

Private Sub grdPLANO_GotFocus()
If txtNOME.Text = "" Then txtNOME.SetFocus
End Sub

Private Sub grdPlano_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_retorno
Dim i As Integer
If KeyCode = vbKeyInsert Then
   KeyCode = 0
   GoSub add_linha
ElseIf KeyCode = 13 Then
   KeyCode = 0
   If grdPLANO.TextMatrix(grdPLANO.row - 1, 0) = "" Then
      InformaaoUsuario "Não salte linhas"
      Exit Sub
   End If
   If grdPLANO.TextMatrix(grdPLANO.row, 0) = "" Then
      wp_AddLinha = True
      txtLinha.Text = ""
      txtDESCRICAO.Text = ""
   Else
      wp_AddLinha = False
      txtDESCRICAO.Text = grdPLANO.TextMatrix(grdPLANO.row, 1)
   End If
   boxLinha.Visible = True
   txtLinha.SetFocus
   Exit Sub
ElseIf KeyCode = vbKeyDelete Then
   If MsgBox("Confirma a deleção da Linha?", vbYesNo, "Mensagem do Sistema") = vbYes Then
      GoSub del_linha
   End If
End If
Exit Sub

del_linha:
For i = grdPLANO.row To grdPLANO.rows - 2
   grdPLANO.TextMatrix(i, 0) = grdPLANO.TextMatrix(i + 1, 0)
   grdPLANO.TextMatrix(i, 1) = grdPLANO.TextMatrix(i + 1, 1)
Next
grdPLANO.TextMatrix(grdPLANO.rows - 1, 0) = ""
grdPLANO.TextMatrix(grdPLANO.rows - 1, 1) = ""
Return

add_linha:
For i = grdPLANO.rows - 1 To grdPLANO.row + 1 Step -1
   grdPLANO.TextMatrix(i, 0) = grdPLANO.TextMatrix(i - 1, 0)
   grdPLANO.TextMatrix(i, 1) = grdPLANO.TextMatrix(i - 1, 1)
Next
grdPLANO.TextMatrix(grdPLANO.row, 0) = ""
grdPLANO.TextMatrix(grdPLANO.row, 1) = ""
Return
End Sub




Private Sub grdPlano_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim wl_File As String
If KeyAscii = 27 Then
   KeyAscii = 0
   If grdPLANO.TextMatrix(1, 0) <> "" Then
      If MsgBox("Grava o Relatório?", vbYesNo, "Mensagem do Sistema") = vbYes Then
         GoSub Gera_Relatorio
      End If
   End If
   txtNOME.SetFocus
End If
Exit Sub


Gera_Relatorio:
If Dir(PathPadrao + "RELATORIOS", vbDirectory) = "" Then
   MkDir (PathPadrao + "RELATORIOS")
   If Dir(PathPadrao + "RELATORIOS", vbDirectory) = "" Then
      If MsgBox("Não foi possível criar uma pasta. Posso Terminar?", vbYesNo, "Mensagem do Sistema") = vbYes Then
         Return
      End If
   End If
End If
If InStr(txtNOME.Text, ".") = 0 Then
   wl_File = txtNOME.Text + ".REL"
Else
   wl_File = Mid(txtNOME.Text, 1, InStr(txtNOME.Text, ".") - 1) + ".REL"
End If
If Dir(PathPadrao + "RELATORIOS\" + wl_File) <> "" Then
   Kill PathPadrao + "RELATORIOS\" + wl_File
End If
Open PathPadrao + "RELATORIOS\" + wl_File For Output As 99
For i = 1 To grdPLANO.rows - 1
   If grdPLANO.TextMatrix(i, 0) = "" Then Exit For
   If Mid(grdPLANO.TextMatrix(i, 0), 1, 1) = "(" Then
      Print #99, grdPLANO.TextMatrix(i, 0) + " --> " + grdPLANO.TextMatrix(i, 1)
   Else
      Print #99, grdPLANO.TextMatrix(i, 0)
   End If
Next
Close #99
Open PathPadrao + "RELATORIOS\" + txtNOME.Text + ".DAT" For Output As 99
Print #99, wp_DAT
Close #99
Return
End Sub


Private Sub txtDESCRICAO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   grdPLANO.TextMatrix(grdPLANO.row, 1) = txtDESCRICAO.Text
   grdPLANO.SetFocus
   boxLinha.Visible = False
ElseIf KeyAscii = 27 Then
   KeyAscii = 0
   grdPLANO.SetFocus
   boxLinha.Visible = False
End If
End Sub


Private Sub txtLinha_GotFocus()
aviso "<F1> Plano de Contas"
If Not wp_AddLinha Then
   If Mid(grdPLANO.TextMatrix(grdPLANO.row, 0), 1, 1) <> "(" Then
      txtLinha.Text = Trim(Mid(grdPLANO.TextMatrix(grdPLANO.row, 0), 1, InStr(grdPLANO.TextMatrix(grdPLANO.row, 0), ":") - 1))
      txtDESCRICAO.Text = ""
   Else
      txtLinha.Text = grdPLANO.TextMatrix(grdPLANO.row, 0)
      txtDESCRICAO.Text = grdPLANO.TextMatrix(grdPLANO.row, 1)
   End If
Else
   txtLinha.Text = ""
End If
End Sub

Private Sub txtLinha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
   Most_PlanodeContas
   txtLinha.SetFocus
End If
End Sub


Private Sub txtLinha_KeyPress(KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 13 Then
   KeyAscii = 0
   GoSub Acrescenta_Conta
ElseIf KeyAscii = 27 Then
   grdPLANO.SetFocus
   Me.boxLinha.Visible = False
End If
Exit Sub


Acrescenta_Conta:
If Mid(txtLinha.Text, 1, 1) <> "(" Then
   tbPlano.Seek "=", txtLinha.Text
   If tbPlano.NoMatch Then
      InformaaoUsuario "Conta não encontrada"
      txtLinha.SetFocus
      HomeEnd
      Exit Sub
   End If
End If
If wp_AddLinha Then
   i = 0
   Do While True
      i = i + 1
      If i = grdPLANO.rows - 1 Then grdPLANO.rows = grdPLANO.rows + 1
      If grdPLANO.TextMatrix(i, 0) = "" Then
         If Mid(txtLinha.Text, 1, 1) <> "(" Then
            grdPLANO.TextMatrix(i, 0) = String(Conta_Char(txtLinha.Text, "."), ".") + txtLinha.Text + " : " + tbPlano("DESCRICAO")
         Else
            grdPLANO.TextMatrix(i, 0) = txtLinha.Text
         End If
         grdPLANO.row = i
         grdPLANO.CellForeColor = AZUL
         txtDESCRICAO.Text = ""
         txtDESCRICAO.SetFocus
         Exit Sub
      End If
   Loop
Else
   grdPLANO.CellForeColor = AZUL
   If Mid(txtLinha.Text, 1, 1) <> "(" Then
      grdPLANO.TextMatrix(grdPLANO.row, 0) = String(Conta_Char(txtLinha.Text, "."), ".") + txtLinha.Text + " : " + tbPlano("DESCRICAO")
   Else
      grdPLANO.TextMatrix(grdPLANO.row, 0) = txtLinha.Text
      txtDESCRICAO.SetFocus
      Exit Sub
   End If
   grdPLANO.SetFocus
   boxLinha.Visible = False
   Exit Sub
End If
Return
End Sub


Private Sub txtLinha_LostFocus()
aviso
End Sub


Private Sub txtNOME_GotFocus()
Dim i As Integer
grdPLANO.rows = 15
For i = 1 To 14
   grdPLANO.TextMatrix(i, 0) = ""
Next
grdPLANO.Enabled = False
chkMES.Enabled = False
chkSALDO.Enabled = False
End Sub


Private Sub txtNOME_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   Command1_Click
End If
End Sub

Private Sub txtNOME_KeyPress(KeyAscii As Integer)
Dim wl_Linha As String
Dim i As Integer
If KeyAscii = 13 Then
   KeyAscii = 0
   If InStr(txtNOME.Text, ".") <> 0 Then
      InformaaoUsuario "Informe o nome sem ."
      txtNOME.SetFocus
      Exit Sub
   End If
   chkMES.Enabled = True
   chkSALDO.Enabled = True
   grdPLANO.Enabled = True
   If Dir(PathPadrao + "RELATORIOS", vbDirectory) <> "" Then
      If Dir(PathPadrao + "RELATORIOS\" + txtNOME.Text + ".REL") <> "" Then
         Open PathPadrao + "RELATORIOS\" + txtNOME.Text + ".REL" For Input As #99
         i = 0
         Do While Not EOF(99)
            Line Input #99, wl_Linha
            i = i + 1
            If InStr(wl_Linha, "-->") <> 0 Then
               grdPLANO.TextMatrix(i, 0) = Trim(Mid(wl_Linha, 1, InStr(wl_Linha, "-->") - 1))
               grdPLANO.TextMatrix(i, 1) = Trim(Mid(wl_Linha, InStr(wl_Linha, "-->") + 3))
            Else
               grdPLANO.TextMatrix(i, 0) = wl_Linha
            End If
         Loop
         Close #99
         If Dir(PathPadrao + "RELATORIOS\" + txtNOME.Text + ".DAT") <> "" Then
            Open PathPadrao + "RELATORIOS\" + txtNOME.Text + ".DAT" For Input As #99
            Line Input #99, wp_DAT
            Close #99
            If InStr(wp_DAT, "M") <> 0 Then Me.chkMES.Value = 1 Else Me.chkMES.Value = 0
            If InStr(wp_DAT, "S") <> 0 Then Me.chkSALDO.Value = 1 Else Me.chkSALDO.Value = 0
         Else
            wp_DAT = ""
         End If
      Else
         wp_DAT = ""
      End If
   End If
   grdPLANO.SetFocus
ElseIf KeyAscii = 27 Then
   KeyAscii = 0
   Unload Me
End If
End Sub


