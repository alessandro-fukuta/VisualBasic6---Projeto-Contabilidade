VERSION 5.00
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fBalanceteAnalitico 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidade"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5625
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5415
      Begin VB.CommandButton cmdImpressoras 
         Caption         =   "&Impressoras"
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "132 Colunas Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "80 Colunas Condensado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton cmdimprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin Mascara.Máscara txtinicio 
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Mascara.Máscara txtfim 
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial e Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Balancete Analítico"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "fBalanceteAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tbPlano As Recordset
Dim tbSaldo As Recordset
Dim tbMoviCaixa As Recordset

Private Sub cmdImpressoras_Click()
CommonDialog1.ShowPrinter
End Sub

Private Sub cmdimprimir_Click()
Dim XMES As String
Dim XANO As String
Dim XDATAINICIAL As Date
Dim wl_Linha As Currency
Dim pCabecalho
Dim pReferencia
Dim pCampos
Dim XTRADUTOR As Long
Dim wl_Pagina As Integer
Dim XANTERIOR As Currency
Dim xdebito As Currency
Dim xcredito As Currency
Dim XATUAL As Currency
Dim XTOTDEBITO As Currency
Dim XTOTCREDITO As Currency
Dim XULTDIA As Date
Dim XLIXO As String
Dim xlixo2 As String
Dim xgrupo As String
Dim xmovimento As Double



aadd pCabecalho, Array("CODIGO DA CONTA", "TRADUTOR", "DENOMINACAO DA CONTA", "SALDO ANTERIOR", "DEBITOS", "CREDITOS", "SALDO ATUAL")
aadd pReferencia, Array("9.9.99.99999-99", "99999999", "AAAAAAAAAAAAAABBBBBBBBBCCCDDDDDDEXXXXXXXX", "999,999,999.99", "999,999,999.99", "999,999,999.99", "999,999,999.99")

XDATAINICIAL = txtinicio.Pacote
XULTDIA = txtfim.Pacote

If file(PathWindows + "mantovani.sys") Then
   MsgBox "Configurado para a Industria Mantovani"
End If


If Not PadraodeImpressao Then

   Me.txtinicio.SetFocus
   Exit Sub
   
End If


tbPlano.Index = "iCONTA"

If tbPlano.RecordCount > 0 Then

   tbPlano.MoveFirst
   
End If

wl_Linha = 0
wl_Pagina = 0

xgrupo = ""

Do While Not tbPlano.EOF


    If Len(tbPlano("conta")) > 0 Then

        If wl_Linha = 0 Or wl_Linha > 28 Then
           If wl_Linha > 28 Then
              Salta_Pagina3
           End If
           wl_Linha = 0
           GoSub CABECALHO
        End If

           xdebito = 0
           xcredito = 0
           XANTERIOR = 0
           XATUAL = 0
           
        If Len(tbPlano("conta")) = 1 Then
           Imprime wl_Linha, 0, tbPlano("conta"), Imp_Normal
           wl_Linha = wl_Linha + 1
        End If
        
        If xgrupo = "" Or xgrupo <> Mid$(tbPlano("CONTA"), 1, 8) Then
                       wl_Linha = wl_Linha + 0.5
                       xgrupo = Mid$(tbPlano("CONTA"), 1, 8)
                       If file(PathWindows + "mantovani.sys") Then
                          Imprime wl_Linha, 0, Mid$(Formata_Conta_Mantovani(tbPlano("CONTA")), 1, 8) + " - " + RTrim(tbPlano("DESCRICAO")), imp_Normal_Negrito
                         Else
                          Imprime wl_Linha, 0, tbPlano("conta") + " - " + RTrim(tbPlano("DESCRICAO")), Imp_Normal
                       End If
                       wl_Linha = wl_Linha + 0.5
                       wl_Linha = wl_Linha + 0.5
                       
        End If
                    
        If tbPlano("TRADUTOR") > 0 Then
           
           wl_SaldoAbertura = tbPlano("SALDOABERTURA")
           XTRADUTOR = tbPlano("TRADUTOR")
           
           GoSub fSaldoAnterior
           
           xdebito = 0
           xcredito = 0
           XANTERIOR = wl_Saldo
           XATUAL = 0
           
           tbMoviCaixa.Seek ">=", XDATAINICIAL, 1
           
            If Not tbMoviCaixa.NoMatch Then
               Do While Not tbMoviCaixa.EOF
                  wl_MontaLinha = True
                  If tbMoviCaixa("CREDITO") <> XTRADUTOR And tbMoviCaixa("DEBITO") <> XTRADUTOR Then
                     wl_MontaLinha = False
                  End If
                  If tbMoviCaixa("DATA") > XULTDIA Then
                     Exit Do
                  End If
                  If wl_MontaLinha Then
                     
                     If tbMoviCaixa("estorno") = False Then
                     
                     
                        xmovimento = tbMoviCaixa("MOVIMENTO")
                     
                        If tbMoviCaixa("credito") = XTRADUTOR Then
                        
                           xcredito = xcredito + tbMoviCaixa("valor")
                           
                          Else
                          
                           xdebito = xdebito + tbMoviCaixa("valor")
                           
                        End If
                     
                     End If
                        
                  End If
                    
                  tbMoviCaixa.MoveNext
               
               Loop
               
            End If
      
            
            If Not pb_InverteOperacao Then
               XATUAL = (XANTERIOR - xdebito) + xcredito
              Else
               XATUAL = (XANTERIOR - xcredito) + xdebito
            End If
            
              If xdebito <> 0 Or xcredito <> 0 Or XANTERIOR <> 0 Then
              
                If Option1.Value = True Then
                 
                If file(PathWindows + "mantovani.sys") Then
                   Monta_LinhadeImpressao wl_Linha, RTrim(Formata_Conta_Mantovani(tbPlano("CONTA"))), 0, "E", imp_Condensado
                 Else
                   Monta_LinhadeImpressao wl_Linha, RTrim(tbPlano("CONTA")), 0, "E", imp_Condensado
                End If
                
                 Monta_LinhadeImpressao wl_Linha, LTrim(Format(XTRADUTOR, "000 0")), 1, "D", imp_Condensado
                 Monta_LinhadeImpressao wl_Linha, RTrim(Mid(tbPlano("DESCRICAO"), 1, 42)), 2, "E", imp_Condensado
              '  Monta_LinhadeIm`ressao Tl_Linha, RTrim(Format(xmovimento, "00000000")), 3, "E", imp_Condensado
                 
                 Monta_LinhadeImpressao wl_Linha, LTrim(Numero_Contabil(XANTERIOR)), 3, "D", imp_Condensado
                 Monta_LinhadeImpressao wl_Linha, LTrim(Format(pValor, "#####0.00")), 4, "D", imp_Condensado
                 Monta_LinhadeImpressao wl_Linha, LTrim(Format(xcredito, "###,##0.00")), 5, "D", imp_Condensado
                 Monta_LinhadeImpressao wl_Linha, LTrim(Numero_Contabil(XATUAL)), 6, "D", imp_Condensado
                 
                Else
                
                If file(PathWindows + "mantovani.sys") Then
                   Monta_LinhadeImpressao wl_Linha, RTrim(Formata_Conta_Mantovani(tbPlano("CONTA"))), 0, "E", Imp_Normal
                 Else
                   Monta_LinhadeImpressao wl_Linha, RTrim(tbPlano("CONTA")), 0, "E", Imp_Normal
                End If
                 
                 Monta_LinhadeImpressao wl_Linha, LTrim(Format(XTRADUTOR, "00000")), 1, "D", Imp_Normal
                 Monta_LinhadeImpressao wl_Linha, RTrim(Mid(tbPlano("DESCRICAO"), 1, 42)), 2, "E", Imp_Normal
                 
                    Monta_LinhadeImpressao wl_Linha, LTrim(Numero_Contabil(XANTERIOR)), 3, "D", Imp_Normal
                    Monta_LinhadeImpressao wl_Linha, LTrim(Format(xdebito, "###,##0.00")), 4, "D", Imp_Normal
                    Monta_LinhadeImpressao wl_Linha, LTrim(Format(xcredito, "###,##0.00")), 5, "D", Imp_Normal
                    Monta_LinhadeImpressao wl_Linha, LTrim(Numero_Contabil(XATUAL)), 6, "D", Imp_Normal
                 
                End If
                
                 XTOTDEBITO = XTOTDEBITO + xdebito
                 XTOTCREDITO = XTOTCREDITO + xcredito
                 
                 wl_Linha = wl_Linha + 0.5
                
              
              End If
           
        

      End If
      
      
    End If
    
      tbPlano.MoveNext
      
      If tbPlano.EOF Or tbPlano.NoMatch Then
         Exit Do
      End If

Loop

                 wl_Linha = wl_Linha + 0.5

If Option1.Value = True Then

      Monta_LinhadeImpressao wl_Linha, Format(XTOTDEBITO, "###,##0.00"), 4, "D", imp_Condensado_NEGRITO
      Monta_LinhadeImpressao wl_Linha, Format(XTOTCREDITO, "###,##0.00"), 5, "D", imp_Condensado_NEGRITO
    Else
      Monta_LinhadeImpressao wl_Linha, Format(XTOTDEBITO, "###,##0.00"), 4, "D", imp_Normal_Negrito
      Monta_LinhadeImpressao wl_Linha, Format(XTOTCREDITO, "###,##0.00"), 5, "D", imp_Normal_Negrito
    
End If

Salta_Pagina3
Finaliza_Impressao

MsgBox ("RELATORIO CONCLUIDO COM SUCESSO !")

Exit Sub


fSaldoAnterior:
tbSaldo.Seek "<", XTRADUTOR, XDATAINICIAL
'If Loca_Contas(tbPlano, XDRADUTOR       Then
'   wl_SaldoAbertura = tbPlano("SALDOABERTURA")
'Else
'   wl_SaldoAbertura = 0
'ENd If
If Not tbSaldo.NoMatch Then
   If tbSaldo("CONTA") = XTRADUTOR Then
      If Not pb_InverteOperacao Then
         wl_Saldo = wl_SaldoAbertura + tbSaldo("ANTERIOR") - tbSaldo("DEBITO") + tbSaldo("CREDITO")
      Else
         wl_Saldo = wl_SaldoAbertura + tbSaldo("ANTERIOR") + tbSaldo("DEBITO") - tbSaldo("CREDITO")
      End If
   Else
      wl_Saldo = wl_SaldoAbertura
   End If
Else
   Go = wl_SaldoAbertura
End If
Return


CABECALHO:
wl_Pagina = wl_Pagina + 1
wl_Linha = wl_Linha + 0.5
Monta_Cabecalho pCabecalho, pReferencia, 6, wl_Linha, IIf(Option1.Value = True, imp_Condensado, Imp_Normal), "BALANCETE ANALITIBO DE " + Mes_Extenso(Month(txtinicio.Pacote)) + "/" + Format(Year(txtinicio.Pacote), "0000"), wl_Pagina, Ultimo_DiasdoMes(Format(Month(Me.txtfim.Pacote), "00"), Format(Year(Me.txtfim.Pacote), "0000"))

Return





End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   If UCase(Me.ActiveControl.Name) = "TXTINICIO" Then
      Unload Me: Exit Sub
     Else
      Me.txtinicio.SetFocus
      Exit Sub
   End If
End If
End Sub

Private Sub Form_Load()

centraobj Me

If Not Abre_PlanoContas(tbPlano) Or Not _
       Abre_MoviCaixa(tbMoviCaixa) Or Not _
       Abre_SaldoContas(tbSaldo) Then
        
        MsgBox "OS ARQUIVOS NECESSÁRIOS ESTÃO BLOQUEADOS !"
        Unload Me
        Exit Sub
        
End If

End Sub


Private Sub Option1_Click()
cmdimprimir.SetFocus
End Sub

Private Sub Option2_Click()
cmdimprimir.SetFocus
End Sub


Private Sub txtfim_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtfim.Pacote < txtinicio.Pacote Then
      X = MsgBox("O preenchimento das datas não está correto !", vbCritical + vbOKOnly, "Aviso")
      Me.txtinicio.SetFocus
      Exit Sub
   End If
   
   SendKeys "{tab}"
   
End If

End Sub

Private Sub txtinicio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtinicio.Text = "" Then
      PO
      Me.txtinicio.SetFocus
      Exit Sub
   End If
   SendKeys "{tab}"
End If

End Sub



Function Salta_Pagina3()
If Mid(pb_Impressao_Normal, 1, 5) <> "Draft" And Not pb_PadraoVideo Then
   Printer.NewPage
Else
   If Not pb_PadraoVideo Then
      Imprime pb_LinhaBuffer + 0.5, 0, " ", pb_Tamanho
   Else
      Imprime pb_LinhaBuffer + 0.5, 0, String(123, "*"), imp_Condensado, , , pb_PadraoVideo
   End If
   For i = pb_UltimaLinha To 64
       Print #1, " "
   Next
   pb_LinhaBuffer = 0
   pb_Buffer = ""
   pb_UltimaLinha = 0
End If
End Function

