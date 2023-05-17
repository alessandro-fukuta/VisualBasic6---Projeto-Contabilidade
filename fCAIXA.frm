VERSION 5.00
Object = "{BA676A3D-9505-4A77-87DC-76025E082864}#1.0#0"; "ETIQUETA.OCX"
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Begin VB.Form fCAIXA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamentos Contábeis"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "fCAIXA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7185
   Visible         =   0   'False
   Begin VB.ComboBox cmbMES 
      Height          =   315
      ItemData        =   "fCAIXA.frx":000C
      Left            =   75
      List            =   "fCAIXA.frx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Selecione o Mês de Trabalho"
      Top             =   180
      Width           =   1470
   End
   Begin VB.CommandButton cmdESTORNO 
      BackColor       =   &H00C0C000&
      Caption         =   "Estorno / <F4>"
      Height          =   375
      Left            =   2055
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame boxDados 
      Enabled         =   0   'False
      Height          =   4035
      Left            =   45
      TabIndex        =   14
      Top             =   1125
      Width           =   7095
      Begin VB.CommandButton cmdDELETA 
         BackColor       =   &H00C0C000&
         Caption         =   "&Deleta"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3555
         Width           =   1005
      End
      Begin VB.CommandButton cmdGRAVA 
         BackColor       =   &H00C0C000&
         Caption         =   "&Grava"
         Enabled         =   0   'False
         Height          =   375
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3555
         Width           =   1005
      End
      Begin Etiq.Etiqueta lblCredito 
         Height          =   300
         Left            =   1155
         TabIndex        =   16
         Top             =   1365
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   529
         BackColor       =   12648447
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
      Begin VB.TextBox txtHISTORICO 
         ForeColor       =   &H00FF0000&
         Height          =   915
         Left            =   1155
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2430
         Width           =   5550
      End
      Begin Mascara.Máscara txtCREDITO 
         Height          =   300
         Left            =   480
         TabIndex        =   7
         Top             =   1365
         Width           =   675
         _ExtentX        =   1191
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
         ÉValor          =   -1  'True
      End
      Begin Mascara.Máscara txtVALOR 
         Height          =   300
         Left            =   480
         TabIndex        =   8
         Top             =   1875
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,###,##0.00"
         ÉValor          =   -1  'True
      End
      Begin Etiq.Etiqueta lblDebito 
         Height          =   300
         Left            =   1155
         TabIndex        =   24
         Top             =   855
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   529
         BackColor       =   12648447
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
      Begin Mascara.Máscara txtDebito 
         Height          =   300
         Left            =   480
         TabIndex        =   6
         Top             =   855
         Width           =   675
         _ExtentX        =   1191
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
         ÉValor          =   -1  'True
      End
      Begin Etiq.Etiqueta lblLANPADRAO 
         Height          =   300
         Left            =   780
         TabIndex        =   26
         Top             =   330
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   529
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin Mascara.Máscara txtLANPADRAO 
         Height          =   300
         Left            =   105
         TabIndex        =   5
         Top             =   330
         Width           =   675
         _ExtentX        =   1191
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
         ÉValor          =   -1  'True
      End
      Begin Mascara.Máscara txtCHISTORICO 
         Height          =   300
         Left            =   480
         TabIndex        =   9
         Top             =   2430
         Width           =   675
         _ExtentX        =   1191
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
         ÉValor          =   -1  'True
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Histórico"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   465
         TabIndex        =   31
         Top             =   2220
         Width           =   615
      End
      Begin VB.Image imgCREDITO 
         Height          =   480
         Left            =   0
         Picture         =   "fCAIXA.frx":009D
         Top             =   1245
         Width           =   480
      End
      Begin VB.Image imgDEBITO 
         Height          =   480
         Left            =   15
         Picture         =   "fCAIXA.frx":04DF
         Top             =   750
         Width           =   480
      End
      Begin VB.Image imgSai 
         Height          =   480
         Left            =   5565
         Picture         =   "fCAIXA.frx":0921
         Top             =   3495
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgENTRA 
         Height          =   480
         Left            =   5085
         Picture         =   "fCAIXA.frx":0D63
         Top             =   3465
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblTIPOCREDITO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5820
         TabIndex        =   29
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label lblTIPODEBITO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5820
         TabIndex        =   28
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Lançamento Padrão"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Conta para Débito"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   465
         TabIndex        =   25
         Top             =   645
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   465
         TabIndex        =   17
         Top             =   1665
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Conta para Crédito"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   465
         TabIndex        =   15
         Top             =   1155
         Width           =   1320
      End
   End
   Begin Mascara.Máscara txtDATA 
      Height          =   300
      Left            =   2220
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
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
      Format          =   "dd/mmm/yyyy"
      Mask            =   "##/##/####"
      ÉData           =   -1  'True
   End
   Begin Mascara.Máscara txtMOVIMENTO 
      Height          =   300
      Left            =   810
      TabIndex        =   4
      Top             =   780
      Width           =   675
      _ExtentX        =   1191
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
      ÉValor          =   -1  'True
   End
   Begin Mascara.Máscara txtANO 
      Height          =   300
      Left            =   1575
      TabIndex        =   2
      ToolTipText     =   "Informe o Ano de Trabalho"
      Top             =   195
      Width           =   570
      _ExtentX        =   1005
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
      Mask            =   "####"
      Text            =   ""
      ÉValor          =   -1  'True
   End
   Begin Mascara.Máscara txtDIA 
      Height          =   300
      Left            =   90
      TabIndex        =   3
      Top             =   780
      Width           =   675
      _ExtentX        =   1191
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
      Mask            =   "##"
      Text            =   ""
      ÉValor          =   -1  'True
   End
   Begin VB.Label lblANALISE 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   390
      Left            =   3300
      TabIndex        =   30
      Top             =   30
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Label Label8 
      Caption         =   "Ano"
      Height          =   195
      Left            =   1545
      TabIndex        =   23
      Top             =   0
      Width           =   990
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Mês"
      Height          =   195
      Left            =   90
      TabIndex        =   22
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Motivo do Estorno"
      Height          =   195
      Left            =   2070
      TabIndex        =   21
      Top             =   585
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.Label lblMOTIVO 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   2070
      TabIndex        =   20
      Top             =   780
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.Label Label2 
      Caption         =   "Movimento"
      Height          =   195
      Left            =   780
      TabIndex        =   13
      Top             =   585
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dia"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   570
      Width           =   240
   End
End
Attribute VB_Name = "fCAIXA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wp_Edita As Boolean
Dim wp_Saida As Boolean
Dim wp_Cria As Boolean
Dim tbMoviCaixa As Recordset
Dim tbContas As Recordset
Dim tbSaldo As Recordset
Dim tbLanPadrao As Recordset
Dim tbRateio As Recordset
Dim tbHistorico As Recordset
Private wp_Entrada As Boolean
Private wp_TipoDebito As String * 1
Private wp_TipoCredito As String * 1
Private wp_Rateio As Boolean
Private Sub Analise_Movimento()
If pb_InverteOperacao Then
   If wp_TipoDebito = "D" Then
      lblANALISE.Caption = "Pagamento"
      lblANALISE.Visible = True
   End If
   If wp_TipoCredito = "R" Then
      lblANALISE.Caption = "Recebimento"
      lblANALISE.Visible = True
   End If
Else
   If wp_TipoDebito = "R" Then
      lblANALISE.Caption = "Recebimento"
      lblANALISE.Visible = True
   End If
   If wp_TipoCredito = "D" Then
      lblANALISE.Caption = "Pagamento"
      lblANALISE.Visible = True
   End If
End If
If wp_TipoCredito = "A" And wp_TipoDebito = "A" Then
   lblANALISE.Caption = "Transferência"
   lblANALISE.Visible = True
End If
End Sub


Private Function Estorno()
Dim wl_Movimento As Long
Dim wl_Valor As Currency
Dim wl_TIPO As String
Dim wl_Conta As Long



End Function


Private Function Grava_MoviCaixa() As Boolean
Dim wl_Movimento As Long
Dim wl_ContaPrincipal As Long
Dim wl_SomaRateio As Currency
Dim wl_DiferencaRateio As Currency
Dim wl_ValorBase As Currency
Dim wl_MovimentoPrincipal As Long
Dim wl_ContaResultado As Long
Dim wl_Conta As Long
Dim wl_Valor As Currency
If wp_Rateio Then
   wl_ContaPrincipal = IIf(wp_TipoDebito = "A", Me.txtCREDITO.VALOR, Me.txtDebito.VALOR)
   Do While Not tbRateio.EOF
      If tbRateio("CONTAPRINCIPAL") <> wl_ContaPrincipal Then Exit Do
      wl_SomaRateio = wl_SomaRateio + tbRateio("PROPORCAO")
      tbRateio.MoveNext
   Loop
   wl_DiferencaRateio = 100 - wl_SomaRateio
   wl_ValorBase = (txtVALOR.VALOR * wl_DiferencaRateio) / 100
   wl_MovimentoPrincipal = 0
   wl_ContaResultado = wl_ContaPrincipal
   wl_SomaRateio = wl_ValorBase
   GoSub Armazena_Registro
   tbRateio.Seek "=", wl_ContaPrincipal
   Do While Not tbRateio.EOF
      If tbRateio("CONTAPRINCIPAL") <> wl_ContaPrincipal Then Exit Do
      If Loca_Contas(tbContas, tbRateio("CONTARATEIO")) Then
         wl_ValorBase = Round((txtVALOR.VALOR * tbRateio("PROPORCAO")) / 100, 2)
         wl_SomaRateio = wl_SomaRateio + Round((txtVALOR.VALOR * tbRateio("PROPORCAO")) / 100, 2)
         wl_ContaResultado = tbRateio("CONTARATEIO")
         GoSub Armazena_Registro
      End If
      tbRateio.MoveNext
   Loop
   tbMoviCaixa.Seek "=", CDate(txtDATA.Pacote), wl_MovimentoPrincipal
   wl_Valor = tbMoviCaixa("VALOR")
   wl_DiferencaRateio = txtVALOR.VALOR - wl_SomaRateio
   wl_Valor = wl_Valor + wl_DiferencaRateio
   If edit_reg(tbMoviCaixa) Then
      tbMoviCaixa("VALOR") = wl_Valor
      update_reg tbMoviCaixa
   End If
   wp_Rateio = False
Else
   wl_ContaResultado = IIf(wp_TipoDebito = "A", Me.txtCREDITO.VALOR, Me.txtDebito.VALOR)
   wl_ValorBase = txtVALOR.VALOR
   GoSub Armazena_Registro
End If
Grava_MoviCaixa = True
Exit Function

Armazena_Registro:
If wp_Cria Then
   If tbMoviCaixa.RecordCount = 0 Then
      wl_Movimento = 1
   Else
      tbMoviCaixa.Seek ">=", CDate(txtDATA.Pacote) + 1, 1
      If tbMoviCaixa.NoMatch Then
         tbMoviCaixa.MoveLast
      Else
         tbMoviCaixa.MovePrevious
         If tbMoviCaixa.BOF Then
            tbMoviCaixa.MoveFirst
         End If
      End If
      If tbMoviCaixa("DATA") = CDate(txtDATA.Pacote) Then
         wl_Movimento = tbMoviCaixa("MOVIMENTO") + 1
      Else
         wl_Movimento = 1
      End If
   End If
   If Not add_reg(tbMoviCaixa) Then
      Exit Function
   End If
   tbMoviCaixa("DATA") = txtDATA.Pacote
   tbMoviCaixa("MOVIMENTO") = wl_Movimento
Else
   If Not edit_reg(tbMoviCaixa) Then
      Exit Function
   End If
End If
tbMoviCaixa("PADRAO") = txtLANPADRAO.VALOR
If wp_Rateio Then
   tbMoviCaixa("CREDITO") = IIf(wp_TipoCredito = "A", txtCREDITO.VALOR, wl_ContaResultado)
   tbMoviCaixa("DEBITO") = IIf(wp_TipoCredito = "A", wl_ContaResultado, txtDebito.VALOR)
Else
   tbMoviCaixa("CREDITO") = txtCREDITO.VALOR
   tbMoviCaixa("DEBITO") = txtDebito.VALOR
End If
tbMoviCaixa("VALOR") = wl_ValorBase
tbMoviCaixa("HISTORICO") = txtHISTORICO.Text
If wp_Rateio Then
   If wl_MovimentoPrincipal = 0 Then
      wl_MovimentoPrincipal = wl_Movimento
   End If
   tbMoviCaixa("MOVIMENTOPRINCIPAL") = wl_MovimentoPrincipal
End If
If Not update_reg(tbMoviCaixa) Then
   Exit Function
End If
If txtCREDITO.VALOR > 0 Then
   Atualiza_SaldoContas txtCREDITO.Text, , wl_ValorBase
End If
If txtDebito.VALOR > 0 Then
   Atualiza_SaldoContas txtDebito.Text, wl_ValorBase
End If
Return
End Function



Private Function LimpaBox()
txtLANPADRAO.Text = ""
lblLANPADRAO.Caption = ""
txtCREDITO.Text = ""
lblCredito.Caption = ""
txtDebito.Text = ""
lblDebito.Caption = ""
txtVALOR.Text = ""
txtHISTORICO.Text = ""
cmdGrava.Enabled = False
cmdDELETA.Enabled = False
cmdESTORNO.Visible = False
lblMOTIVO.Visible = False
Label6.Visible = False
lblTIPODEBITO.Caption = ""
lblTIPODEBITO.BackColor = BRANCO
lblTIPOCREDITO.Caption = ""
txtHISTORICO.Text = ""
lblTIPOCREDITO.BackColor = BRANCO
lblANALISE.Visible = False
End Function

Private Function Mon_MoviCaixa()
If tbMoviCaixa("PADRAO") > 0 Then
   txtLANPADRAO.Text = tbMoviCaixa("PADRAO")
   If Loca_LanPadrao(tbLanPadrao, txtLANPADRAO.VALOR) Then
      lblLANPADRAO.Caption = tbLanPadrao("DESCRICAO")
   End If
End If
If tbMoviCaixa("CREDITO") > 0 Then
   txtCREDITO.Text = tbMoviCaixa("CREDITO")
   If Loca_Contas(tbContas, tbMoviCaixa("CREDITO")) Then
      lblCredito.Caption = tbContas("DESCRICAO")
      wp_TipoCredito = tbContas("TIPO")
      If wp_TipoCredito = "D" Then
         lblTIPOCREDITO.Caption = "Despesa"
         lblTIPOCREDITO.BackColor = VERMELHO
         lblTIPOCREDITO.ForeColor = BRANCO
      ElseIf wp_TipoCredito = "R" Then
         lblTIPOCREDITO.Caption = "Receita"
         lblTIPOCREDITO.BackColor = AZUL
         lblTIPOCREDITO.ForeColor = BRANCO
      Else
         lblTIPOCREDITO.Caption = "D/R"
         lblTIPOCREDITO.BackColor = BRANCO
         lblTIPOCREDITO.ForeColor = PRETO
      End If
   End If
End If
If tbMoviCaixa("DEBITO") > 0 Then
   txtDebito.Text = tbMoviCaixa("DEBITO")
   If Loca_Contas(tbContas, tbMoviCaixa("DEBITO")) Then
      lblDebito.Caption = tbContas("DESCRICAO")
      wp_TipoDebito = tbContas("TIPO")
      If wp_TipoDebito = "D" Then
         lblTIPODEBITO.Caption = "Despesa"
         lblTIPODEBITO.BackColor = VERMELHO
      ElseIf wp_TipoDebito = "R" Then
         lblTIPODEBITO.Caption = "Receita"
         lblTIPODEBITO.BackColor = AZUL
      Else
         lblTIPODEBITO.Caption = "D/R"
         lblTIPODEBITO.BackColor = BRANCO
      End If
   End If
End If
txtVALOR.Text = tbMoviCaixa("VALOR")
txtHISTORICO.Text = tbMoviCaixa("HISTORICO")
Analise_Movimento
End Function

Private Function PreparaForm()
wp_Saida = True
LimpaCaixasTexto Me
LimpaBox
txtDATA.Text = Date
txtDATA.Pacote = Date
End Function


Private Function Verifica_Inconsistencias() As Boolean
If txtVALOR.VALOR = 0 Then
   InformaaoUsuario "Informe o valor"
   txtVALOR.SetFocus
   HomeEnd
   Exit Function
End If
If txtCREDITO.Text <> "0" Then
   If Not Loca_Contas(tbContas, txtCREDITO.Text) Then
      InformaaoUsuario "A conta de crédito não foi encontrada"
      txtCREDITO.SetFocus
      HomeEnd
      Exit Function
   End If
End If
If txtDebito.Text <> "0" Then
   If Not Loca_Contas(tbContas, txtDebito.Text) Then
      InformaaoUsuario "A conta de débito não foi encontrada"
      txtDebito.SetFocus
      HomeEnd
      Exit Function
   End If
End If
If txtDebito.VALOR = 0 And txtCREDITO.VALOR = 0 Then
   InformaaoUsuario "Impossível gerar um movimento sem contas"
   txtDebito.SetFocus
   HomeEnd
   Exit Function
End If
If txtHISTORICO.Text = "" Then
   InformaaoUsuario "Informe um Histórico para o Movimento"
   txtHISTORICO.SetFocus
   Exit Function
End If
If txtDebito.VALOR = txtCREDITO.VALOR Then
   InformaaoUsuario "Movimento Inválido ..."
   txtDebito.SetFocus
   Exit Function
End If
'If wp_TipoDebito <> "A" And wp_TipoDebito = wp_TipoCredito Then
'   InformaaoUsuario "Não é possível movimentas duas contas de " + IIf(wp_TipoDebito = "D", "'Despesa'", "'Receita'") + " ao mesmo tempo", , "Verifique"
'   txtDebito.SetFocus
'   Exit Function
'End If
'If wp_TipoDebito <> "A" And wp_TipoCredito <> "A" Then
'   InformaaoUsuario "Movimento inconsistente verifique"
'   txtDebito.SetFocus
'   Exit Function
'End If
Verifica_Inconsistencias = True
End Function








Private Sub cmbMES_GotFocus()
Dim wl_Mes
PreparaForm
wl_Mes = Array("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro")
cmbMES.Text = wl_Mes(Month(Date) - 1)
txtANO.Text = Year(Date)
wp_Saida = True
End Sub


Private Sub cmbMES_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub


Private Sub cmbMES_LostFocus()
wp_Saida = False
End Sub

Private Sub cmdDELETA_Click()
Dim wl_MovimentoPrincipal As Long
Dim wl_Index As String
wl_Index = tbMoviCaixa.Index
If Not IsNull(tbMoviCaixa("MOVIMENTOPRINCIPAL")) Then
   If Not Confirme("Esse é um movimento de rateio. Deletá-lo implicará na eliminação de outros registros relacionados. Deseja Continuar?") Then Exit Sub
   wl_MovimentoPrincipal = tbMoviCaixa("MOVIMENTOPRINCIPAL")
   tbMoviCaixa.Index = "iPRINCIPAL"
   tbMoviCaixa.Seek "=", wl_MovimentoPrincipal
   Do While Not tbMoviCaixa.NoMatch
      If edit_reg(tbMoviCaixa) Then tbMoviCaixa.Delete
      tbMoviCaixa.Seek "=", wl_MovimentoPrincipal
   Loop
   tbMoviCaixa.Index = wl_Index
Else
   If MsgBox("Confirma a deleção do movimento?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbYes Then
      If Not edit_reg(tbMoviCaixa) Then
         InformaaoUsuario "Impossível deletar o movimento"
         txtMOVIMENTO.SetFocus
         Exit Sub
      End If
      tbMoviCaixa.Delete
      pb_DeletaMovimento = True
      Open PathPadrao + "DELETA.MOV" For Output As #1
      Print #1, "x"
      Close #1
   End If
End If
txtMOVIMENTO.Text = ""
txtMOVIMENTO.SetFocus
End Sub

Private Sub cmdESTORNO_Click()
Dim wl_Motivo As String
If Not Verifica_Privilegio(PR_LANCAMENTOS, "S") Then
   If Not NovaPermissao(PR_LANCAMENTOS, "S", "Estorno Mov.Caixa") Then
      InformaaoUsuario "Sem privilégio para estornar o movimento"
      txtDATA.SetFocus
      Exit Sub
   End If
End If
If MsgBox("Confirma o estorno?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbNo Then
   txtDATA.SetFocus
   Exit Sub
End If
Do While True
   wl_Motivo = InputBox("Qual o Motivo?", "ESTORNO DE MOVIMENTO")
   If wl_Motivo <> "" Then
      Exit Do
   End If
Loop
If Not edit_reg(tbMoviCaixa) Then
   InformaaoUsuario "Não foi possível estornar o movimento"
   txtDATA.SetFocus
   Exit Sub
End If
tbMoviCaixa("ESTORNO") = True
tbMoviCaixa("MOTIVO") = wl_Motivo
If Not update_reg(tbMoviCaixa) Then
   InformaaoUsuario "Não foi possível estornar o movimento"
End If
If txtDebito.VALOR <> 0 Then
   Atualiza_SaldoContas txtDebito.Text, , txtVALOR.Text
End If
If txtCREDITO.VALOR <> 0 Then
   Atualiza_SaldoContas txtCREDITO.Text, txtVALOR.Text
End If
txtMOVIMENTO.SetFocus
End Sub

Private Sub cmdGrava_Click()
Dim wl_Outro As String
If Verifica_Inconsistencias Then
   Analise_Movimento
   If MsgBox("Confirma a Gravação?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbNo Then
      txtDebito.SetFocus
      HomeEnd
      Exit Sub
   End If
   If Grava_MoviCaixa Then
      If (lblANALISE.Caption = "Pagamento" Or lblANALISE.Caption = "Recebimento") Then
         If MsgBox("Emite Recibo?", vbYesNo, "Mensagem do Sistema") = vbYes Then GoSub Recibo
      End If
      txtMOVIMENTO.Text = ""
      txtMOVIMENTO.SetFocus
   Else
      InformaaoUsuario "O movimento não foi " + IIf(wp_Cria, "gravado", "atualizado")
   End If
End If
Exit Sub

Recibo:
If Not PadraodeImpressao Then Return
If lblANALISE.Caption = "Recebimento" Then
   Do While True
      wl_Outro = InputBox("Informe o nome do pagador", "Recebimento")
      If wl_Outro <> "" Then
         If MsgBox(Trim(wl_Outro) + ". Está correto?", vbYesNo, "Nome do Pagador") = vbYes Then Exit Do
      End If
   Loop
   Imprime 0, 0, "R E C I B O", imp_Condensado_NEGRITO
   Imprime 0, 80 - Len("[" + Format(txtVALOR.VALOR, "R$ ##,###,##0.00") + "]"), "[" + Format(txtVALOR.VALOR, "R$ ##,###,##0.00") + "]", imp_Normal_Negrito
   Imprime 1, 0, "DATA :" + CStr(txtDATA.Pacote), Imp_Normal
   Imprime 1, 80 - Len("Movimento :" + CStr(txtMOVIMENTO.VALOR)), "Movimento :" + CStr(txtMOVIMENTO.VALOR), Imp_Normal
   Imprime 4, 0, "ATRAVES DESSE DOCUMENTO DECLARO TER RECEBIDO DE " + Trim(UCase(wl_Outro)) + " A QUANTIA DE:"
   Imprime 5, 0, Format(Me.txtVALOR.VALOR, "R$ ##,###,##0.00") + " (" + UCase(Extenso(txtVALOR.VALOR)) + "). "
   Imprime 6, 0, "REFERENTE A " + UCase(Trim(Me.txtHISTORICO.Text)) + "."
   Imprime 7, 0, "POR SER VERDADE, FIRMO O PRESENTE."
   Imprime 8, 0, pb_Cidade + ", " + Format(Day(Date), "00") + " de " + NomedoMes(Month(Date)) + " de " + Str(Year(Date))
   Imprime 11, 40 - Len(pb_RAZAOSOCIAL) / 2, String(Len(pb_RAZAOSOCIAL), "_")
   Imprime 12, 40 - Len(pb_RAZAOSOCIAL) / 2, UCase(pb_RAZAOSOCIAL)

   Imprime 14, 0, String(80, "-"), imp_Normal_Negrito
   
   Imprime 16, 0, "R E C I B O", imp_Condensado_NEGRITO
   Imprime 16, 80 - Len("[" + Format(txtVALOR.VALOR, "R$ ##,###,##0.00") + "]"), "[" + Format(txtVALOR.VALOR, "R$ ##,###,##0.00") + "]", imp_Normal_Negrito
   Imprime 17, 0, "DATA :" + CStr(txtDATA.Pacote), Imp_Normal
   Imprime 17, 80 - Len("Movimento :" + CStr(txtMOVIMENTO.VALOR)), "Movimento :" + CStr(txtMOVIMENTO.VALOR), Imp_Normal
   Imprime 19, 0, "ATRAVES DESSE DOCUMENTO DECLARO TER RECEBIDO DE " + Trim(UCase(wl_Outro)) + " A QUANTIA DE:"
   Imprime 20, 0, Format(Me.txtVALOR.VALOR, "R$ ##,###,##0.00") + " (" + UCase(Extenso(txtVALOR.VALOR)) + "). "
   Imprime 21, 0, "REFERENTE A " + UCase(Trim(Me.txtHISTORICO.Text)) + "."
   Imprime 22, 0, "POR SER VERDADE, FIRMO O PRESENTE."
   Imprime 23, 0, pb_Cidade + ", " + Format(Day(Date), "00") + " de " + NomedoMes(Month(Date)) + " de " + Str(Year(Date))
   Imprime 26, 40 - Len(pb_RAZAOSOCIAL) / 2, String(Len(pb_RAZAOSOCIAL), "_")
   Imprime 27, 40 - Len(pb_RAZAOSOCIAL) / 2, UCase(pb_RAZAOSOCIAL)
Else
   Do While True
      wl_Outro = InputBox("Informe o nome do Recebedor", "Pagamento")
      If wl_Outro <> "" Then
         If MsgBox(Trim(wl_Outro) + ". Está correto?", vbYesNo, "Nome do Pagador") = vbYes Then Exit Do
      End If
   Loop
   Imprime 0, 0, "R E C I B O", imp_Condensado_NEGRITO
   Imprime 0, 80 - Len("[" + Format(txtVALOR.VALOR, "R$ ##,###,##0.00") + "]"), "[" + Format(txtVALOR.VALOR, "R$ ##,###,##0.00") + "]", imp_Normal_Negrito
   Imprime 1, 0, "DATA :" + CStr(txtDATA.Pacote), Imp_Normal
   Imprime 1, 80 - Len("Movimento :" + CStr(txtMOVIMENTO.VALOR)), "Movimento :" + CStr(txtMOVIMENTO.VALOR), Imp_Normal
   Imprime 4, 0, "ATRAVES DESSE DOCUMENTO DECLARO TER RECEBIDO DE " + UCase(pb_RAZAOSOCIAL) + " A QUANTIA DE:"
   Imprime 5, 0, Format(Me.txtVALOR.VALOR, "R$ ##,###,##0.00") + " (" + UCase(Extenso(txtVALOR.VALOR)) + "). "
   Imprime 6, 0, "REFERENTE A " + UCase(Trim(Me.txtHISTORICO.Text)) + "."
   Imprime 7, 0, "POR SER VERDADE, FIRMO O PRESENTE."
   Imprime 8, 0, pb_Cidade + ", " + Format(Day(Date), "00") + " de " + NomedoMes(Month(Date)) + " de " + Str(Year(Date))
   Imprime 11, 40 - Len(Trim(UCase(wl_Outro))) / 2, String(Len(Trim(UCase(wl_Outro))), "_")
   Imprime 12, 40 - Len(Trim(UCase(wl_Outro))) / 2, UCase(Trim(UCase(wl_Outro)))

   Imprime 14, 0, String(80, "-"), imp_Normal_Negrito

   Imprime 16, 0, "R E C I B O", imp_Condensado_NEGRITO
   Imprime 16, 80 - Len("[" + Format(txtVALOR.VALOR, "R$ ##,###,##0.00") + "]"), "[" + Format(txtVALOR.VALOR, "R$ ##,###,##0.00") + "]", imp_Normal_Negrito
   Imprime 17, 0, "DATA :" + CStr(txtDATA.Pacote), Imp_Normal
   Imprime 17, 80 - Len("Movimento :" + CStr(txtMOVIMENTO.VALOR)), "Movimento :" + CStr(txtMOVIMENTO.VALOR), Imp_Normal
   Imprime 19, 0, "ATRAVES DESSE DOCUMENTO DECLARO TER RECEBIDO DE " + UCase(pb_RAZAOSOCIAL) + " A QUANTIA DE:"
   Imprime 20, 0, Format(Me.txtVALOR.VALOR, "R$ ##,###,##0.00") + " (" + UCase(Extenso(txtVALOR.VALOR)) + "). "
   Imprime 21, 0, "REFERENTE A " + UCase(Trim(Me.txtHISTORICO.Text)) + "."
   Imprime 22, 0, "POR SER VERDADE, FIRMO O PRESENTE."
   Imprime 23, 0, pb_Cidade + ", " + Format(Day(Date), "00") + " de " + NomedoMes(Month(Date)) + " de " + Str(Year(Date))
   Imprime 26, 40 - Len(Trim(UCase(wl_Outro))) / 2, String(Len(Trim(UCase(wl_Outro))), "_")
   Imprime 27, 40 - Len(Trim(UCase(wl_Outro))) / 2, UCase(Trim(UCase(wl_Outro)))
   Salta_Pagina
End If
Finaliza_Impressao
Return
End Sub


Private Sub Form_Activate()
If Not Abre_MoviCaixa(tbMoviCaixa) Or _
   Not Abre_PlanoContas(tbContas) Or _
   Not Abre_SaldoContas(tbSaldo) Or _
   Not Abre_LanPadrao(tbLanPadrao) Or _
   Not Abre_RateioContas(tbRateio) Or _
   Not Abre_Historico(tbHistorico) Then
   Me.Visible = False
   Unload Me
   Exit Sub
End If
If Not wp_Entrada Then
   wp_Entrada = True
   txtDIA.SetFocus
   txtDIA.Text = Format(Day(Date), "00")
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 And cmdESTORNO.Visible Then
   KeyCode = 0
   cmdESTORNO.SetFocus
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   If wp_Saida Then
      Me.Visible = False
      Unload Me
   Else
      If wp_Edita Then
         SendKeys "{TAB}"
         Exit Sub
      End If
      If Me.ActiveControl.Name = "txtDIA" Or Me.ActiveControl.Name = "txtANO" Then
         cmbMES.SetFocus
      ElseIf Me.ActiveControl.Name = "txtMOVIMENTO" Then
         txtDIA.SetFocus
      Else
         txtMOVIMENTO.Text = ""
         txtMOVIMENTO.SetFocus
      End If
   End If
End If
End Sub

Private Sub Form_Load()
centraobj Me
If pb_InverteOperacao Then
   imgDEBITO.Picture = imgENTRA.Picture
   imgCREDITO.Picture = imgSai.Picture
Else
   imgDEBITO.Picture = imgSai.Picture
   imgCREDITO.Picture = imgENTRA.Picture
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbMoviCaixa.Close
tbContas.Close
tbSaldo.Close
tbLanPadrao.Close
tbHistorico.Close
End Sub






Private Sub lblTIPO_Click()

End Sub





Private Sub tmANALISE_Timer()

End Sub

Private Sub Máscara1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub Máscara1_KeyPress(KeyAscii As Integer)
End Sub


Private Sub txtCHISTORICO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If VtoP(txtCHISTORICO.Text) <> 0 Then
      If Not Loca_Historico(tbHistorico, txtCHISTORICO.VALOR) Then
         InformaaoUsuario "Histórico não encontrado"
         HomeEnd
         Exit Sub
      End If
      txtHISTORICO.Text = tbHistorico("DESCRICAO")
      txtHISTORICO.SelStart = Len(txtHISTORICO)
   End If
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtCREDITO_GotFocus()
aviso "<F1> - Consulta Contas"
lblANALISE.Visible = False
End Sub

Private Sub txtCREDITO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno As Variant
If KeyCode = vbKeyF1 Then
   wl_Retorno = Most_PlanodeContas
   If wl_Retorno <> "" Then
      ShowRetorno = False
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
         txtCREDITO.SetFocus
         HomeEnd
         Exit Sub
      End If
      wp_TipoCredito = tbContas("TIPO")
      lblCredito.Caption = tbContas("DESCRICAO")
      If wp_TipoCredito = "D" Then
'         If pb_InverteOperacao Then
'            InformaaoUsuario "Não é possível creditar uma Despesa"
'            txtCREDITO.SetFocus
'            HomeEnd
'            Exit Sub
'         End If
         lblTIPOCREDITO.Caption = "Despesa"
         lblTIPOCREDITO.BackColor = VERMELHO
         lblTIPOCREDITO.ForeColor = BRANCO
      ElseIf wp_TipoCredito = "R" Then
'         If Not pb_InverteOperacao Then
'            InformaaoUsuario "Não é possível creditar uma Receita"
'            txtCREDITO.SetFocus
'            HomeEnd
'            Exit Sub
'         End If
         lblTIPOCREDITO.Caption = "Receita"
         lblTIPOCREDITO.BackColor = AZUL
         lblTIPOCREDITO.ForeColor = BRANCO
      Else
         lblTIPOCREDITO.Caption = "D/R"
         lblTIPOCREDITO.BackColor = BRANCO
         lblTIPOCREDITO.ForeColor = PRETO
      End If
   End If
   If Not wp_Rateio Then
      tbRateio.Seek "=", txtCREDITO.VALOR
      If Not tbRateio.NoMatch Then
         If Confirme("Confirma rateio da Conta?") Then
            wp_Rateio = True
         Else
            wp_Rateio = False
         End If
      Else
         wp_Rateio = False
      End If
   End If
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtCREDITO_LostFocus()
aviso
End Sub

Private Sub txtDATA_LostFocus()
wp_Saida = False
End Sub


Private Sub txtDEBITO_GotFocus()
wp_Rateio = False
aviso "<F1> - Consulta Contas"
lblANALISE.Visible = False
End Sub

Private Sub txtDEBITO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno As Variant
If KeyCode = vbKeyF1 Then
   wl_Retorno = Most_PlanodeContas
   If wl_Retorno <> "" Then
      ShowRetorno = False
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
         InformaaoUsuario "Conta de Débito não encontrada"
         txtDebito.SetFocus
         HomeEnd
         Exit Sub
      End If
      wp_TipoDebito = tbContas("TIPO")
      lblDebito.Caption = tbContas("DESCRICAO")
      If wp_TipoDebito = "D" Then
'         If Not pb_InverteOperacao Then
'            InformaaoUsuario "Não é possível debitar uma Despesa"
'            txtDEBITO.SetFocus
'            HomeEnd
'            Exit Sub
'         End If
         lblTIPODEBITO.Caption = "Despesa"
         lblTIPODEBITO.BackColor = VERMELHO
         lblTIPODEBITO.ForeColor = BRANCO
      ElseIf wp_TipoDebito = "R" Then
'         If pb_InverteOperacao Then
'            InformaaoUsuario "Não é possível debitar uma Receita"
'            txtDEBITO.SetFocus
'            HomeEnd
'            Exit Sub
'         End If
         lblTIPODEBITO.Caption = "Receita"
         lblTIPODEBITO.BackColor = AZUL
         lblTIPODEBITO.ForeColor = BRANCO
      Else
         lblTIPODEBITO.Caption = "D/R"
         lblTIPODEBITO.BackColor = BRANCO
         lblTIPODEBITO.ForeColor = PRETO
      End If
   End If
   tbRateio.Seek "=", txtDebito.VALOR
   If Not tbRateio.NoMatch Then
      If Confirme("Confirma rateio da Conta?") Then
         wp_Rateio = True
      Else
         wp_Rateio = False
      End If
   Else
      wp_Rateio = False
   End If
   SendKeys "{TAB}"
End If
End Sub

Private Sub txtDEBITO_LostFocus()
aviso
End Sub


Private Sub txtDIA_GotFocus()
txtDIA.Text = Day(Date)
LimpaBox
txtMOVIMENTO.Text = ""
End Sub

Private Sub txtDIA_LostFocus()
Dim wl_Mes
Dim i As Integer
wl_Mes = Array("", "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro")
For i = 1 To 12
   If cmbMES.Text = wl_Mes(i) Then
      Exit For
   End If
Next
txtDATA.Text = Format(txtDIA.VALOR, "00") + "/" + Format(i, "00") + "/" + Format(txtANO.VALOR, "0000")
If IsDate(txtDATA.Pacote) Then
   If CDate(txtDATA.Pacote) > Date Then
      InformaaoUsuario "Data Incorreta"
      txtDIA.Text = ""
      cmbMES.SetFocus
      Exit Sub
   End If
End If
End Sub


Private Sub txtHISTORICO_GotFocus()
wp_Edita = True
aviso "<ESC> Abandona a edição"
End Sub


Private Sub txtHISTORICO_LostFocus()
wp_Edita = False
aviso
End Sub


Private Sub txtLANPADRAO_GotFocus()
aviso "<F1> Lançamentos Padrao"
End Sub

Private Sub txtLANPADRAO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_LanPadrao
   txtLANPADRAO.SetFocus
   If wl_Retorno <> "" Then
      txtLANPADRAO.Text = wl_Retorno
      ShowRetorno = False
      SendKeys "{ENTER}"
   End If
End If
End Sub

Private Sub txtLANPADRAO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtLANPADRAO.Text <> "" And txtLANPADRAO.Text <> "0" Then
      If Not Loca_LanPadrao(tbLanPadrao, txtLANPADRAO.Text) Then
         InformaaoUsuario "Lançamento Padrão não encontrado"
         txtLANPADRAO.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblLANPADRAO.Caption = tbLanPadrao("DESCRICAO")
      txtDebito.Text = tbLanPadrao("DEBITO")
      If txtDebito.VALOR <> 0 Then
         If Loca_Contas(tbContas, txtDebito.VALOR) Then
            lblDebito.Caption = tbContas("DESCRICAO")
         End If
      End If
      txtCREDITO.Text = tbLanPadrao("CREDITO")
      If txtCREDITO.VALOR <> 0 Then
         If Loca_Contas(tbContas, txtCREDITO.VALOR) Then
            lblCredito.Caption = tbContas("DESCRICAO")
         End If
      End If
     txtHISTORICO.Text = tbLanPadrao("DESCRICAO")
   End If
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtMOVIMENTO_GotFocus()
If Me.txtDIA.VALOR = 0 Then
   InformaaoUsuario "É necessário informar o dia da movimentação ..."
   txtDIA.SetFocus
   HomeEnd
   Exit Sub
End If
LimpaBox
wp_Rateio = False
End Sub

Private Sub txtMOVIMENTO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno As Variant
If KeyCode = vbKeyF1 Then
   wl_Retorno = most_moviCaixa
   If wl_Retorno <> "" Then
      txtMOVIMENTO.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If
End Sub


Private Sub txtMOVIMENTO_KeyPress(KeyAscii As Integer)
Dim wl_Movimento As Long
If KeyAscii = 13 Then
   KeyAscii = 0
   wp_TipoDebito = ""
   wp_TipoCredito = ""
   If txtMOVIMENTO.Text = "" Or txtMOVIMENTO.Text = "0" Then
      LimpaBox
      If tbMoviCaixa.RecordCount = 0 Then
         wl_Movimento = 1
      Else
         tbMoviCaixa.Seek ">=", CDate(txtDATA.Pacote) + 1, 1
         If tbMoviCaixa.NoMatch Then
            tbMoviCaixa.MoveLast
         Else
            tbMoviCaixa.MovePrevious
            If tbMoviCaixa.BOF Then
               tbMoviCaixa.MoveFirst
            End If
         End If
         If tbMoviCaixa("DATA") = CDate(txtDATA.Pacote) Then
            wl_Movimento = tbMoviCaixa("MOVIMENTO") + 1
         Else
            wl_Movimento = 1
         End If
      End If
      txtMOVIMENTO.Text = wl_Movimento
      wp_Cria = True
   Else
      If tbMoviCaixa.RecordCount = 0 Then
         InformaaoUsuario "Movimento não encontrado"
         txtMOVIMENTO.SetFocus
         HomeEnd
         Exit Sub
      End If
      tbMoviCaixa.Seek "=", CDate(txtDATA.Pacote), txtMOVIMENTO.Text
      If tbMoviCaixa.NoMatch Then
         InformaaoUsuario "Movimento não encontrado"
         txtMOVIMENTO.SetFocus
         HomeEnd
         Exit Sub
      End If
      LimpaBox
      Mon_MoviCaixa
      wp_Cria = False
   End If
   If wp_Cria And Not Verifica_Privilegio(PR_LANCAMENTOS, "I") Then
      InformaaoUsuario "Usuário sem privilégio para incluir movimento"
      txtDATA.SetFocus
      Exit Sub
   End If
   If wp_Cria Then
      cmdGrava.Enabled = True
   End If
   If wp_Cria Then
      boxDados.Enabled = True
      SendKeys "{TAB}"
   Else
      cmdGrava.Enabled = False
      boxDados.Enabled = Verifica_Privilegio(PR_LANCAMENTOS, "D")
      cmdDELETA.Enabled = Verifica_Privilegio(PR_LANCAMENTOS, "D")
      If Date = tbMoviCaixa("DATA") And Not tbMoviCaixa("IMPORTADO") Then
         cmdESTORNO.Visible = Not tbMoviCaixa("ESTORNO")
      End If
      If tbMoviCaixa("ESTORNO") Then
         lblMOTIVO.Caption = tbMoviCaixa("MOTIVO")
         lblMOTIVO.Visible = True
         Label6.Visible = True
         cmdDELETA.Enabled = False
      End If
      HomeEnd
   End If
End If
End Sub
Private Sub txtMOVIMENTO_LostFocus()
aviso
End Sub

