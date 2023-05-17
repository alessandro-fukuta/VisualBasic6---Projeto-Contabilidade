VERSION 5.00
Object = "{BA676A3D-9505-4A77-87DC-76025E082864}#1.0#0"; "ETIQUETA.OCX"
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fRazaoAnalitico 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidade"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5655
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   5415
      Begin VB.CommandButton cmdImpressoras 
         Caption         =   "&Impressoras"
         Height          =   495
         Left            =   3960
         TabIndex        =   11
         Top             =   720
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
      Begin Etiq.Etiqueta lbldescricao 
         Height          =   300
         Left            =   720
         TabIndex        =   10
         Top             =   1200
         Width           =   4575
         _ExtentX        =   8070
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
      Begin Mascara.Máscara txtconta 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   615
         _ExtentX        =   1085
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
         Format          =   "######"
         Text            =   ""
         ÉValor          =   -1  'True
      End
      Begin VB.CommandButton cmdimprimir 
         Caption         =   "&Imprimir"
         Height          =   495
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   1335
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
         TabIndex        =   4
         Top             =   1560
         Value           =   -1  'True
         Width           =   3015
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
         TabIndex        =   5
         Top             =   1800
         Width           =   3015
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Conta F1-Vídeo"
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
         TabIndex        =   9
         Top             =   960
         Width           =   1695
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
      Caption         =   "RAZÃO ANALÍTICO"
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
      Left            =   600
      TabIndex        =   8
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "fRazaoAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tbPlano As Recordset
Dim tbSaldo As Recordset
Dim tbMoviCaixa As Recordset

Private Sub Combo1_Change()

End Sub


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
Dim XCONTA As String
Dim XCONTRA As String
Dim XHISTORICO As String
Dim wl_MontaLinha As Boolean
Dim XTRAD As String

aadd pCabecalho, Array("DATA", "C/ PARTIDA", "TRADUTOR", "HISTORICO", "MOVTO", "DEBITOS", "CREDITOS", "SALDO ATUAL")
aadd pReferencia, Array("99/99/9999", "99999999999999", "999999", "AAAAABBBBBBBBBCCCDDDDDDDDDDEEEEEEEEEXXXXXXXXXXX", "99999", "999,999.99", "999,999.99", "999,999.99")


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


        If wl_Linha = 0 Or wl_Linha > 28 Then
           If wl_Linha > 28 Then
              Salta_Pagina
           End If
           wl_Linha = 0
           GoSub CABECALHO
        End If

           xdebito = 0
           xcredito = 0
           XANTERIOR = 0
           XATUAL = 0
                
        If txtconta.VALOR > 0 Then
           Do While tbPlano("TRADUTOR") <> txtconta.VALOR
              If tbPlano.EOF Then
                 Exit Do
              End If
              
              If tbPlano("TRADUTOR") = txtconta.VALOR Then
                 Exit Do
              End If
              
              tbPlano.MoveNext
              If tbPlano.EOF Then
                 Exit Do
              End If
              
           Loop
        End If
        
        If tbPlano.EOF Then
           Exit Do
        End If
        
        If tbPlano("TRADUTOR") > 0 Then
           
           
           wl_SaldoAbertura = tbPlano("SALDOABERTURA")
           
           XTRADUTOR = tbPlano("TRADUTOR")
           XCONTA = tbPlano("CONTA")
           
           GoSub fSaldoAnterior
           
           xdebito = 0
           xcredito = 0
           XANTERIOR = wl_Saldo
           XATUAL = 0
              
            
           wl_Linha = wl_Linha + 1
                
            If file(PathWindows + "mantovani.sys") Then
               Imprime wl_Linha, 0, RTrim(Formata_Conta_Mantovani(tbPlano("CONTA"))) + "-" + RTrim(tbPlano("DESCRICAO")) + "  TRADUTOR....:" + LTrim(Format(XTRADUTOR, "00000")), imp_Condensado
              Else
               Imprime wl_Linha, 0, RTrim(tbPlano("CONTA")) + "-" + RTrim(tbPlano("DESCRICAO")) + "  TRADUTOR....:" + LTrim(Format(XTRADUTOR, "00000")), imp_Condensado
            End If
            
           Imprime wl_Linha, 96, "SALDO ANTERIOR.:" + LTrim(Numero_Contabil(XANTERIOR)), imp_Condensado
                
     '      Monta_LinhadeImpressao wl_Linha, LTrim(Numero_Contabil(XANTERIOR)), 6, "D", imp_Condensado_NEGRITO
                
           wl_Linha = wl_Linha + 1
           
           tbMoviCaixa.Index = "iDATA"
           tbMoviCaixa.Seek ">=", XDATAINICIAL, 1
           
           Do While Not tbMoviCaixa.EOF
             
                    If wl_Linha = 0 Or wl_Linha > 28 Then
                       If wl_Linha > 28 Then
                          Salta_Pagina
                       End If
                       wl_Linha = 0
                       GoSub CABECALHO
                       wl_Linha = wl_Linha + 0.5
                             
                       If file(PathWindows + "mantovani.sys") Then
                          Imprime wl_Linha, 0, RTrim(Formata_Conta_Mantovani(tbPlano("CONTA"))) + "-" + RTrim(tbPlano("DESCRICAO")) + "  TRADUTOR....:" + LTrim(Format(XTRADUTOR, "00000")) + "   SALDO ANTERIOR ( CONTINUACAO ) ...", imp_Condensado
                         Else
                          Imprime wl_Linha, 0, RTrim(tbPlano("CONTA")) + "-" + RTrim(tbPlano("DESCRICAO")) + "  TRADUTOR....:" + LTrim(Format(XTRADUTOR, "00000")) + "   SALDO ANTERIOR ( CONTINUACAO ) ...", imp_Condensado
                       End If
                         
                       ' Imprime wl_Linha, 96, "SALDO ANTERIOR.:" + LTrim(Numero_Contabil(XATUAL)), imp_Normal_Negrito
                       ' Imprime wl_Linha, 96, "SALDO ANTERIOR.:" + LTrim(Numero_Contabil(XATUAL)), imp_Normal_Negrito
                       Monta_LinhadeImpressao wl_Linha, LTrim(Numero_Contabil(XATUAL)), 7, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                       wl_Linha = wl_Linha + 1
                       
                    End If
                    
                  wl_MontaLinha = True
                  
                  If tbMoviCaixa("CREDITO") <> XTRADUTOR And tbMoviCaixa("DEBITO") <> XTRADUTOR Then
                     wl_MontaLinha = False
                  End If
                  
                  If tbMoviCaixa("DATA") < XDATAINICIAL Then
                     wl_MontaLinha = False
                  End If
                  
                  If tbMoviCaixa("DATA") > XULTDIA Then
                     wl_MontaLinha = False
                  End If
                  
                  If wl_MontaLinha = True Then
                        
                        Monta_LinhadeImpressao wl_Linha, tbMoviCaixa("DATA"), 0, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                        XCONTRA = ""
                        XTRAD = ""
                        
                        If XTRADUTOR = tbMoviCaixa("CREDITO") Then
                           tbPlano.Index = "iTRADUTOR"
                           tbPlano.Seek "=", tbMoviCaixa("DEBITO")
                           XCONTRA = tbPlano("CONTA")
                           XTRAD = tbMoviCaixa("DEBITO")
                          Else
                           tbPlano.Index = "iTRADUTOR"
                           tbPlano.Seek "=", tbMoviCaixa("CREDITO")
                           XCONTRA = tbPlano("CONTA")
                           XTRAD = tbMoviCaixa("CREDITO")
                        End If

                        XHISTORICO = RTrim(Mid(tbMoviCaixa("HISTORICO"), 1, 50))
                        XHISTORICO = StrTran(XHISTORICO, "º", ".")
                        XHISTORICO = StrTran(XHISTORICO, Asc(13), " ")
                        XHISTORICO = StrTran(XHISTORICO, Asc(10), " ")
                        XHISTORICO = StrTran(XHISTORICO, Chr$(10), " ")
                        XHISTORICO = StrTran(XHISTORICO, Chr$(13), " ")
                        XHISTORICO = RTrim(XHISTORICO)
                        
                        If file(PathWindows + "mantovani.sys") Then
                           Monta_LinhadeImpressao wl_Linha, RTrim(Formata_Conta_Mantovani(XCONTRA)), 1, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                         Else
                           Monta_LinhadeImpressao wl_Linha, RTrim(XCONTRA), 1, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                        End If
                        
                        Monta_LinhadeImpressao wl_Linha, LTrim(Format(XTRAD, "00000")), 2, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                        Monta_LinhadeImpressao wl_Linha, XHISTORICO, 3, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                        Monta_LinhadeImpressao wl_Linha, Format(tbMoviCaixa("MOVIMENTO"), "00000"), 4, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                        
                        tbPlano.Index = "iCONTA"
                        tbPlano.Seek "=", XCONTA
                        
                        If tbMoviCaixa("credito") = XTRADUTOR Then
                        
                           xcredito = xcredito + tbMoviCaixa("valor")
                           Monta_LinhadeImpressao wl_Linha, LTrim(Format(tbMoviCaixa("VALOR"), "#,##0.00")), 6, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                          
                          ElseIf tbMoviCaixa("debito") = XTRADUTOR Then
                           
                           Monta_LinhadeImpressao wl_Linha, LTrim(Format(tbMoviCaixa("VALOR"), "#,##0.00")), 5, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                           xdebito = xdebito + tbMoviCaixa("valor")
                           
                        End If
                         
                        If Not pb_InverteOperacao Then
                           XATUAL = (XANTERIOR - xdebito) + xcredito
                         Else
                           XATUAL = (XANTERIOR - xcredito) + xdebito
                        End If
                             
                        Monta_LinhadeImpressao wl_Linha, LTrim(Numero_Contabil(XATUAL)), 7, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                             
                        wl_Linha = wl_Linha + 0.5
                      
                  End If
                    
                  tbMoviCaixa.MoveNext
               
               Loop
               
       End If
           
      tbPlano.MoveNext
      
      If tbPlano.EOF Or tbPlano.NoMatch Then
         Exit Do
      End If

Loop

                 wl_Linha = wl_Linha + 0.5

                 
Salta_Pagina
Finaliza_Impressao

MsgBox ("RELATORIO CONCLUIDO COM SUCESSO !")

Exit Sub


fSaldoAnterior:
tbSaldo.Seek "<", XTRADUTOR, XDATAINICIAL
'If Loca_Contas(tbPlano, XTRADUTOR) Then
'   wl_SaldoAbertura = tbPlano("SALDOABERTURA")
'Else
'   wl_SaldoAbertura = 0
'End If

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
   
   wl_Saldo = wl_SaldoAbertura

End If


Return


CABECALHO:
wl_Pagina = wl_Pagina + 1
Monta_Cabecalho pCabecalho, pReferencia, 7, wl_Linha, IIf(Option1.Value = True, imp_Condensado, Imp_Normal), "RAZAO ANALITICO DE " + Mes_Extenso(Month(txtinicio.Pacote)) + "/" + CStr(Year(txtinicio.Pacote)), wl_Pagina, Ultimo_DiasdoMes(Format(Month(Me.txtfim.Pacote), "00"), Format(Year(Me.txtfim.Pacote), "0000"))

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



Private Sub txtconta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   txtconta.Text = Most_PlanodeContas
   If txtconta.Text <> "" Then
      SendKeys "{enter}"
   End If
End If
   
End Sub

Private Sub txtconta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtconta.VALOR = 0 Then
      lbldescricao.Caption = "RAZÃO COMPLETO"
      SendKeys "{TAB}"
      Exit Sub
   End If
   
   tbPlano.Index = "iTRADUTOR"
   tbPlano.Seek "=", txtconta.VALOR
   
   If tbPlano.NoMatch Then
      MsgBox "Conta não encontrada !"
      Me.txtconta.SetFocus
      Exit Sub
     Else
      Me.lbldescricao.Caption = tbPlano("descricao")
      SendKeys "{tab}"
   End If
   
End If
End Sub






