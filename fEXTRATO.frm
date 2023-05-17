VERSION 5.00
Object = "{BA676A3D-9505-4A77-87DC-76025E082864}#1.0#0"; "ETIQUETA.OCX"
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fEXTRATO 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extrato do Movimento de Contas no Vídeo"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   Begin Etiq.Etiqueta lblDebitos 
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Alignment       =   1
      BackColor       =   -2147483624
      Caption         =   "0,00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483625
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Calculadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7170
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdIMPRIME 
      BackColor       =   &H00C0C000&
      Caption         =   "&Imprime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4005
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   630
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "&Pesquisa / <F3>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2325
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   630
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid grdFLUXO 
      Height          =   4935
      Left            =   75
      TabIndex        =   4
      Top             =   1035
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   15
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   400
      BackColor       =   16777215
      BackColorFixed  =   8404992
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483638
      GridColor       =   12582912
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      FormatString    =   $"fEXTRATO.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Etiq.Etiqueta lblCONTA 
      Height          =   300
      Left            =   1065
      TabIndex        =   6
      Top             =   195
      Width           =   6060
      _ExtentX        =   10689
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
   Begin Mascara.Máscara txtCONTA 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   195
      Width           =   975
      _ExtentX        =   1720
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
   Begin Mascara.Máscara txtINICIO 
      Height          =   300
      Left            =   75
      TabIndex        =   1
      Top             =   705
      Width           =   1050
      _ExtentX        =   1852
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
      Text            =   ""
      ÉData           =   -1  'True
   End
   Begin Mascara.Máscara txtFINAL 
      Height          =   300
      Left            =   1185
      TabIndex        =   2
      Top             =   705
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
      Text            =   ""
      ÉData           =   -1  'True
   End
   Begin Etiq.Etiqueta lblCreditos 
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Alignment       =   1
      BackColor       =   -2147483624
      Caption         =   "0,00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483625
   End
   Begin Etiq.Etiqueta lblSaldoAtual 
      Height          =   375
      Left            =   9960
      TabIndex        =   17
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Alignment       =   1
      BackColor       =   -2147483624
      Caption         =   "0,00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483625
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "EXTRATO CONTÁBIL POR CONTA"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Saldo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8970
      TabIndex        =   10
      Top             =   4725
      Width           =   585
   End
   Begin VB.Label lblSALDO 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9675
      TabIndex        =   9
      Top             =   4710
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1185
      TabIndex        =   8
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   75
      TabIndex        =   7
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   45
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "fEXTRATO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tbMoviCaixa As Recordset
Private tbContas As Recordset
Private tbSaldo As Recordset
Private wp_Saida As Boolean
Private Function Desenha_Grade()
grdFluxo.Clear
grdFluxo.TextMatrix(0, 0) = "Movto."
grdFluxo.TextMatrix(0, 1) = "Data"
grdFluxo.TextMatrix(0, 2) = "Histórico"
grdFluxo.TextMatrix(0, 3) = "Débito"
grdFluxo.TextMatrix(0, 4) = "Crédito"
grdFluxo.TextMatrix(0, 5) = "Saldo"
End Function


Private Function Monta_Grade()
Dim wl_Saldo As Currency
Dim wl_Linhas As Integer
Dim wl_MontaLinha As Boolean
Dim wl_Data As Date
Dim wl_SaldoAbertura As Currency
Dim wl_CorPadrao
Dim wl_Colunas As Integer
wl_CorPadrao = AMARELO
GoSub fSaldoAnterior
grdFluxo.row = 1
grdFluxo.Col = 1
grdFluxo.CellFontBold = True
grdFluxo.CellForeColor = BRANCO
grdFluxo.TextMatrix(1, 2) = "SALDO ANTERIOR"
grdFluxo.Col = 5
grdFluxo.CellFontBold = True
grdFluxo.CellForeColor = BRANCO
grdFluxo.TextMatrix(1, 5) = Format(wl_Saldo, "##,###,##0.00;(#,###,##0.00)")
tbMoviCaixa.Seek ">=", CDate(txtinicio.Pacote), 1
wl_Linhas = 2

If Not tbMoviCaixa.NoMatch Then
   grdFluxo.Visible = False
   Do While Not tbMoviCaixa.EOF
      wl_MontaLinha = True
      If tbMoviCaixa("CREDITO") <> txtconta.VALOR And tbMoviCaixa("DEBITO") <> txtconta.VALOR Then
         wl_MontaLinha = False
      End If
      If tbMoviCaixa("credito") = 0 And tbMoviCaixa("debito") = 0 Then
         wl_MontaLinha = False
      End If
      If tbMoviCaixa("DATA") > CDate(txtFINAL.Pacote) Then
         Exit Do
      End If
      If wl_MontaLinha Then GoSub Monta_Linha
      tbMoviCaixa.MoveNext
   Loop
End If

grdFluxo.Visible = True
lblSALDO.ForeColor = IIf(wl_Saldo < 0, VERMELHO, AZUL)
lblSALDO.Caption = Format(wl_Saldo, "##,###,##0.00;(#,###,##0.00)")
grdFluxo.row = 1
grdFluxo.Col = 0
SendKeys "{UP}"
grdFluxo.SetFocus
Exit Function

fSaldoAnterior:
tbSaldo.Seek "<", txtconta.VALOR, CDate(txtinicio.Pacote)
If Loca_Contas(tbContas, txtconta.VALOR) Then
   wl_SaldoAbertura = tbContas("SALDOABERTURA")
Else
   wl_SaldoAbertura = 0
End If
If Not tbSaldo.NoMatch Then
   If tbSaldo("CONTA") = txtconta.VALOR Then
      If Not pb_InverteOperacao Then
         wl_Saldo = wl_SaldoAbertura + tbSaldo("ANTERIOR") - tbSaldo("DEBITO") + tbSaldo("CREDITO")
      Else
         wl_Saldo = wl_SaldoAbertura + tbSaldo("ANTERIOR") + tbSaldo("DEBITO") - tbSaldo("CREDITO")
      End If
   Else
      wl_Saldo = wl_SaldoAbertura
   End If
Else
   wl_Saldo = wl_SaldoAbertura
End If
Return

Monta_Linha:
wl_Linhas = wl_Linhas + 1
If wl_Data = "0:00:00" Then
   wl_Data = tbMoviCaixa("DATA")
ElseIf wl_Data <> tbMoviCaixa("DATA") Then
'   wl_Linhas = wl_Linhas + 1
   If wl_CorPadrao = AMARELO Then
      wl_CorPadrao = BRANCO
   Else
      wl_CorPadrao = AMARELO
   End If
   wl_Data = tbMoviCaixa("DATA")
End If
If wl_Linhas = grdFluxo.rows Then
   grdFluxo.rows = grdFluxo.rows + 1
ElseIf wl_Linhas > grdFluxo.rows Then
   grdFluxo.rows = grdFluxo.rows + 2
End If
grdFluxo.TextMatrix(wl_Linhas, 0) = Format(tbMoviCaixa("MOVIMENTO"), "00000")
grdFluxo.TextMatrix(wl_Linhas, 1) = tbMoviCaixa("DATA")
grdFluxo.TextMatrix(wl_Linhas, 2) = tbMoviCaixa("HISTORICO")
If tbMoviCaixa("DEBITO") = txtconta.VALOR Then
   grdFluxo.row = wl_Linhas
   grdFluxo.Col = 3
   If Not pb_InverteOperacao Then
      grdFluxo.CellForeColor = VERMELHO
      wl_Saldo = wl_Saldo - tbMoviCaixa("VALOR")
   Else
      grdFluxo.CellForeColor = AZUL
      wl_Saldo = wl_Saldo + tbMoviCaixa("VALOR")
   End If
   grdFluxo.TextMatrix(wl_Linhas, 3) = Format(tbMoviCaixa("VALOR"), "##,###,##0.00;(#,###,##0.00)")
End If
If tbMoviCaixa("CREDITO") = txtconta.VALOR Then
   grdFluxo.row = wl_Linhas
   grdFluxo.Col = 4
   If pb_InverteOperacao Then
      grdFluxo.CellForeColor = VERMELHO
      wl_Saldo = wl_Saldo - tbMoviCaixa("VALOR")
   Else
      grdFluxo.CellForeColor = AZUL
      wl_Saldo = wl_Saldo + tbMoviCaixa("VALOR")
   End If
   grdFluxo.TextMatrix(wl_Linhas, 4) = Format(tbMoviCaixa("VALOR"), "##,###,##0.00;(#,###,##0.00)")
End If
grdFluxo.row = wl_Linhas
grdFluxo.Col = 5
grdFluxo.CellForeColor = IIf(wl_Saldo >= 0, AZUL, VERMELHO)
grdFluxo.TextMatrix(wl_Linhas, 5) = Format(wl_Saldo, "##,###,##0.00;(#,###,##0.00)")
If tbMoviCaixa("ESTORNO") Then
   GoSub Cor_Linha
   wl_Linhas = wl_Linhas + 1
   If wl_Linhas = grdFluxo.rows Then
      grdFluxo.rows = grdFluxo.rows + 1
   ElseIf wl_Linhas > grdFluxo.rows Then
      grdFluxo.rows = grdFluxo.rows + 2
   End If
   grdFluxo.TextMatrix(wl_Linhas, 0) = Format(tbMoviCaixa("MOVIMENTO"), "00000")
   grdFluxo.TextMatrix(wl_Linhas, 1) = tbMoviCaixa("DATA")
   grdFluxo.TextMatrix(wl_Linhas, 2) = "ESTORNO DO MOVIMENTO ANTERIOR"
   If tbMoviCaixa("CREDITO") = txtconta.VALOR Then
      If Not pb_InverteOperacao Then
         grdFluxo.row = wl_Linhas
         grdFluxo.Col = 3
         grdFluxo.CellForeColor = VERMELHO
         wl_Saldo = wl_Saldo - tbMoviCaixa("VALOR")
      Else
         grdFluxo.row = wl_Linhas
         grdFluxo.Col = 3
         grdFluxo.CellForeColor = AZUL
         wl_Saldo = wl_Saldo + tbMoviCaixa("VALOR")
      End If
      grdFluxo.TextMatrix(wl_Linhas, 3) = Format(tbMoviCaixa("VALOR"), "##,###,##0.00;(#,###,##0.00)")
   End If
   If tbMoviCaixa("DEBITO") = txtconta.VALOR Then
      If Not pb_InverteOperacao Then
         grdFluxo.row = wl_Linhas
         grdFluxo.Col = 4
         grdFluxo.CellForeColor = AZUL
         wl_Saldo = wl_Saldo + tbMoviCaixa("VALOR")
      Else
         grdFluxo.row = wl_Linhas
         grdFluxo.Col = 4
         grdFluxo.CellForeColor = VERMELHO
         wl_Saldo = wl_Saldo - tbMoviCaixa("VALOR")
      End If
      grdFluxo.TextMatrix(wl_Linhas, 4) = Format(tbMoviCaixa("VALOR"), "##,###,##0.00;(#,###,##0.00)")
   End If
   grdFluxo.row = wl_Linhas
   grdFluxo.Col = 5
   grdFluxo.CellForeColor = IIf(wl_Saldo >= 0, AZUL, VERMELHO)
   grdFluxo.TextMatrix(wl_Linhas, 5) = Format(wl_Saldo, "##,###,##0.00;(#,###,##0.00)")
   GoSub Cor_Linha
End If
Return



Cor_Linha:
For wl_Colunas = 0 To grdFluxo.cols - 1
   grdFluxo.Col = wl_Colunas
   grdFluxo.CellBackColor = wl_CorPadrao
Next
Return
End Function

Private Sub cmdIMPRIME_Click()
Dim i As Integer
Dim wl_Credito As Currency
Dim wl_Debito As Currency
Dim wl_Saldo As Currency
Dim CABECALHO
Dim Referencia
Dim wl_Linha As Currency
Dim pp As Integer
Dim wl_Historico As String

aadd CABECALHO, Array("Data", "Histórico", "Debito", "Crédito", "Saldo")
aadd Referencia, Array("99/99/9999", "xxxxxxxxxxggggggggggxxxxxxxxxxggggggggggxxxxxxxxxx", "9,999,999.99", "9,999,999.99", "9,999,999.99")

If Not PadraodeImpressao Then Exit Sub

For i = 1 To Me.grdFluxo.rows - 1
   If grdFluxo.TextMatrix(i, 5) <> "" Then
      If wl_Linha = 0 Then GoSub CABECALHO
      If i = 1 Then
'*********************************************************
         Monta_LinhadeImpressao wl_Linha, grdFluxo.TextMatrix(i, 2), 1, , imp_Condensado_NEGRITO
         Monta_LinhadeImpressao wl_Linha, grdFluxo.TextMatrix(i, 2), 1, , imp_Condensado_NEGRITO
         wl_Saldo = grdFluxo.TextMatrix(i, 5)
         If wl_Saldo < 0 Then
            If pb_PadraoVideo Then
              fIMPRESSAO.ForeColor = VERMELHO
            Else
               Printer.ForeColor = VERMELHO
            End If
            Monta_LinhadeImpressao wl_Linha, Format(wl_Saldo, "(#,###,##0.00;(#,###,##0.00))"), 4, "D", imp_Condensado_NEGRITO
         Else
            If pb_PadraoVideo Then
               fIMPRESSAO.ForeColor = PRETO
            Else
               Printer.ForeColor = PRETO
            End If
            Monta_LinhadeImpressao wl_Linha, Format(wl_Saldo, "#,###,##0.00;(#,###,##0.00)"), 4, "D", imp_Condensado_NEGRITO
         End If
         wl_Linha = wl_Linha + 0.5
      Else
         wl_Historico = StrTran(grdFluxo.TextMatrix(i, 2), Chr(13), " ")
         wl_Historico = StrTran(wl_Historico, Chr(10), "")
         Monta_LinhadeImpressao wl_Linha, grdFluxo.TextMatrix(i, 1), 0, , imp_Condensado
         Monta_LinhadeImpressao wl_Linha, Mid(wl_Historico, 1, 50), 1, , imp_Condensado
         wl_Debito = IIf(grdFluxo.TextMatrix(i, 3) = "", 0, grdFluxo.TextMatrix(i, 3))
         wl_Credito = IIf(grdFluxo.TextMatrix(i, 4) = "", 0, grdFluxo.TextMatrix(i, 4))
         wl_Saldo = grdFluxo.TextMatrix(i, 5)
         If wl_Debito <> 0 Then
            If pb_PadraoVideo Then
               fIMPRESSAO.ForeColor = VERMELHO
            Else
               Printer.ForeColor = VERMELHO
            End If
            Monta_LinhadeImpressao wl_Linha, Format(wl_Debito, "#,###,##0.00;(#,###,##0.00)"), 2, "D", imp_Condensado
            If pb_PadraoVideo Then
               fIMPRESSAO.ForeColor = PRETO
            Else
               Printer.ForeColor = PRETO
            End If
         End If
         If wl_Credito <> 0 Then
            Monta_LinhadeImpressao wl_Linha, Format(wl_Credito, "#,###,##0.00;(#,###,##0.00)"), 3, "D", imp_Condensado
         End If
         If wl_Saldo < 0 Then
            If pb_PadraoVideo Then
               fIMPRESSAO.ForeColor = VERMELHO
            Else
               Printer.ForeColor = VERMELHO
            End If
            Monta_LinhadeImpressao wl_Linha, Format(wl_Saldo, "#,###,##0.00;(#,###,##0.00)"), 4, "D", imp_Condensado_NEGRITO
            If pb_PadraoVideo Then
               fIMPRESSAO.ForeColor = PRETO
            Else
               Printer.ForeColor = PRETO
            End If
         Else
            Monta_LinhadeImpressao wl_Linha, Format(wl_Saldo, "#,###,##0.00;(#,###,##0.00)"), 4, "D", imp_Condensado_NEGRITO
         End If
         wl_Linha = wl_Linha + 0.5
       End If
   End If
   If wl_Linha > IIf(pb_ImpressaoMatricial, 29, 26) Then
      Salta_Pagina
      wl_Linha = 0
   End If
Next
wl_Linha = wl_Linha + 0.5
Imprime wl_Linha, 0, "Saldo Atual ...........................", imp_Condensado_NEGRITO
Monta_LinhadeImpressao wl_Linha, lblSALDO.Caption, 4, "D", imp_Condensado_NEGRITO
Finaliza_Impressao
Exit Sub


CABECALHO:
pp = pp + 1
Monta_Cabecalho CABECALHO, Referencia, 4, wl_Linha, imp_Condensado, "Extrato de Conta. Emitido em " + CStr(Date) + " . Movimento Ref. ao periodo: " + CStr(Me.txtinicio.Pacote) + " a " + CStr(Me.txtFINAL.Pacote), pp
Imprime wl_Linha, 0, "Conta :" + Str(txtconta.VALOR) + " - " + Me.lblCONTA.Caption, imp_Condensado_NEGRITO
Imprime wl_Linha, 0, "Conta :" + Str(txtconta.VALOR) + " - " + Me.lblCONTA.Caption, imp_Condensado_NEGRITO
wl_Linha = wl_Linha + 0.5
Return
End Sub

Private Sub Command1_Click()
Dim wl_InicioCalculo As String
If txtconta.Text = "" Or txtconta.Text = "0" Then
   InformaaoUsuario "Informe a conta"
   txtconta.SetFocus
   Exit Sub
End If
If txtinicio.Text = "" Then
   InformaaoUsuario "É necessário informar a data inicial"
   txtinicio.SetFocus
   Exit Sub
End If
If txtFINAL.Text = "" Then
   InformaaoUsuario "É necessário informar a data final"
   txtFINAL.SetFocus
   Exit Sub
End If
' wl_InicioCalculo = "01/" + Format(Month(CDate(txtinicio.Pacote)), "00") + "/" + Format(Year(CDate(txtinicio.Pacote)), "0000")
wl_InicioCalculo = "01/01/2002"
' If MsgBox("Recalcula Saldo?", vbYesNo + vbDefaultButton2 + vbQuestion, "Mensagem do Sistema") = vbYes Then
   Do While True
'     wl_InicioCalculo = InputBox("Data Inicial", "Recalcula Saldo", wl_InicioCalculo)
     If IsDate(wl_InicioCalculo) Then Exit Do
     InformaaoUsuario "Data Inválida"
   Loop
   Recalcula_Saldo CDate(wl_InicioCalculo)
' End If
Call Monta_Grade
Call Soma_Grade
End Sub


Private Sub Soma_Grade()

Dim xCreditos As Currency
Dim xDebitos As Currency
Dim XANTERIOR As Currency
Dim XATUAL As Currency


grdFluxo.row = 1
XANTERIOR = grdFluxo.TextMatrix(grdFluxo.row, 5)

xCreditos = 0
xDebitos = 0

grdFluxo.row = 3

Do While True

    If Val(grdFluxo.TextMatrix(grdFluxo.row, 0)) = 0 Then
       Exit Do
    End If
    
    xCreditos = xCreditos + IIf(grdFluxo.TextMatrix(grdFluxo.row, 4) <> "", grdFluxo.TextMatrix(grdFluxo.row, 4), 0)
    xDebitos = xDebitos + IIf(grdFluxo.TextMatrix(grdFluxo.row, 3) <> "", grdFluxo.TextMatrix(grdFluxo.row, 3), 0)
    XATUAL = grdFluxo.TextMatrix(grdFluxo.row, 5)
    
    If grdFluxo.row = grdFluxo.rows - 1 Then Exit Do
    grdFluxo.row = grdFluxo.row + 1
    
    
Loop
    
Me.lblCreditos.Caption = Format(xCreditos, "#,##0.00")
Me.lblDebitos.Caption = Format(xDebitos, "#,##0.00")
Me.lblSaldoAtual.Caption = Format(XATUAL, "#,##0.00")


End Sub
Private Sub Command2_Click()
Shell "CALC.EXE", vbNormalFocus
End Sub

Private Sub Form_Activate()
If Not Abre_PlanoContas(tbContas) Or _
   Not Abre_MoviCaixa(tbMoviCaixa) Or _
   Not Abre_SaldoContas(tbSaldo) Then
   Unload Me
   Exit Sub
End If
centraobj Me
'Recalcula_Saldo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
   KeyCode = 0
   Call Command1_Click
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   If wp_Saida Then
      Unload Me
   Else
      txtconta.Text = ""
      txtconta.SetFocus
   End If
End If
End Sub


Private Sub Form_Load()
Call Desenha_Grade
End Sub


Private Sub Form_Unload(Cancel As Integer)
aviso
End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub Timer2_Timer()
End Sub

Private Sub txtCONTA_GotFocus()
wp_Saida = True
LimpaCaixasTexto Me
grdFluxo.Clear
Desenha_Grade
lblSALDO.Caption = ""
aviso "<F1> Consulta Contas"
End Sub

Private Sub txtconta_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno As Variant
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   If wl_Retorno <> "" Then
      txtconta.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If
End Sub


Private Sub txtconta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtconta.Text = "" Or txtconta.Text = "0" Then
      InformaaoUsuario "É necessário informar a conta"
      txtconta.SetFocus
      Exit Sub
   Else
      If Not Loca_Contas(tbContas, txtconta.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtconta.SetFocus
         Exit Sub
      End If
      lblCONTA.Caption = tbContas("DESCRICAO")
      txtinicio.Text = Date
      txtFINAL.Text = Date
      SendKeys "{TAB}"
   End If
End If
End Sub


Private Sub txtCONTA_LostFocus()
aviso
wp_Saida = False
End Sub


Private Sub txtinicio_GotFocus()
Me.txtinicio.Text = "01/" + Format(Month(Date), "00") + "/" + Format(Year(Date), "0000")
End Sub

Private Sub txtinicio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If Me.txtinicio.Pacote < CDate("01/01/2002") Then
      MsgBox "Data inicial deve ser no mínimo em 01/01/2002"
      Me.txtinicio.SetFocus
      Exit Sub
   End If
   
   SendKeys "{tab}"
   Exit Sub
End If

End Sub
