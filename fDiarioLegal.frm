VERSION 5.00
Object = "{5756E734-2046-400A-BC65-0E105EC5876E}#1.0#0"; "CAIXATEX.OCX"
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fDiarioLegal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilidade"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   5415
      Begin VB.CommandButton cmdImpressoras 
         Caption         =   "&Impressoras"
         Height          =   375
         Left            =   3960
         TabIndex        =   11
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   960
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton cmdimprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin CaixaTexto.Caixa_Texto txtpagina 
         Height          =   300
         Left            =   3000
         TabIndex        =   7
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
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
         Text            =   "1"
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
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Pag.Inicial"
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
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "DIÁRIO LEGAL"
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
      TabIndex        =   6
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "fDiarioLegal"
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
Dim xdia As Date
Dim XLIXO As String
Dim xlixo2 As String
Dim xgrupo As String
Dim XCONTA As String
Dim XCONTRA As String
Dim XHISTORICO As String
Dim wl_MontaLinha As Boolean
Dim XTRAD As String
Dim XNOMECONTRA As String
Dim IMPRIMIU As Boolean
Dim xtotal As Currency
Dim XTOTALGERAL As Currency
Dim xlixo1 As Long
Dim xlixo3 As Long
aadd pCabecalho, Array("DATA", "C/ PARTIDA", "TRADUTOR", "DENOMINACAO DA CONTA", "HISTORICO", "MOVTO", "VALOR MOVIMENTO")
aadd pReferencia, Array("99/99/9999", "99999999999999", "99999", "AAAAAAAAAAAAAAAAAAAAAAAA", "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", "99999", "9,999,999.99")

xlixo1 = tbMoviCaixa.RecordCount
xlixo3 = tbPlano.RecordCount

Bar1.Max = xlixo1 * xlixo3
Bar1.Value = 1

If file(PathWindows + "mantovani.sys") Then
   MsgBox "Configurado para a Industria Mantovani"
End If

If Not PadraodeImpressao Then

   Me.txtinicio.SetFocus
   Exit Sub

End If


MsgBox "Prepare a impressora e confirme !"

Me.cmdimprimir.SetFocus


tbPlano.Index = "iCONTA"

If tbPlano.RecordCount > 0 Then

   tbPlano.MoveFirst
   
End If

wl_Linha = 0
wl_Pagina = Val(txtpagina.Text) - 1

XDATAINICIAL = txtinicio.Pacote
XULTDIA = txtfim.Pacote
xdia = XDATAINICIAL

xgrupo = ""
XTOTALGERAL = 0
xtotal = 0
Do While xdia <= XULTDIA

Do While Not tbPlano.EOF


        If wl_Linha = 0 Or wl_Linha > 28 Then
           If wl_Linha > 28 Then
              Salta_Pagina
           End If
           wl_Linha = 0
           GoSub CABECALHO
           wl_Linha = wl_Linha + 1
            
        End If

           xdebito = 0
           xcredito = 0
           XANTERIOR = 0
           XATUAL = 0
    
'        If XGRUPO = "" Or XGRUPO <> Mid$(tbPlano("CONTA"), 1, 5) Then
'                       wl_Linha = wl_Linha + 0.5
'                       XGRUPO = Mid$(tbPlano("CONTA"), 1, 5)
'                       Imprime wl_Linha, 0, Mid$(tbPlano("CONTA"), 1, 5) + " - " + RTrim(tbPlano("DESCRICAO")), imp_Condensado
'                       wl_Linha = wl_Linha + 0.5
'
'        End If
                
    
        If tbPlano("TRADUTOR") > 0 Then
           
           
           wl_SaldoAbertura = tbPlano("SALDOABERTURA")
           
           XTRADUTOR = tbPlano("TRADUTOR")
           XCONTA = tbPlano("CONTA")
           
           GoSub fSaldoAnterior
           
           xdebito = 0
           xcredito = 0
           XANTERIOR = wl_Saldo
           XATUAL = 0
              
           IMPRIMIU = False
           xtotal = 0
           
           tbMoviCaixa.Index = "iDATA"
           tbMoviCaixa.Seek "=", xdia, 1
           
           Do While Not tbMoviCaixa.EOF
             
                    If tbMoviCaixa.EOF Or tbMoviCaixa.NoMatch Then
                       Exit Do
                    End If
                    
                    If wl_Linha = 0 Or wl_Linha > 28 Then
                       If wl_Linha > 28 Then
                          Salta_Pagina
                       End If
                       wl_Linha = 0
                       GoSub CABECALHO
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
                  
                  If tbMoviCaixa("DATA") <> xdia Then
                     Exit Do
                  End If
                  
                  If wl_MontaLinha = True Then
                        
                        If Not IMPRIMIU Then
                        
                        If file(PathWindows + "mantovani.sys") Then
                           wl_Linha = wl_Linha + 1
                           Imprime wl_Linha, 0, RTrim(Formata_Conta_Mantovani(tbPlano("CONTA"))) + "-" + RTrim(tbPlano("DESCRICAO")) + "  TRADUTOR:" + LTrim(Format(XTRADUTOR, "00000")), Imp_Normal
                           wl_Linha = wl_Linha + 1
                           IMPRIMIU = True
                         Else
                           wl_Linha = wl_Linha + 1
                           Imprime wl_Linha, 0, RTrim(tbPlano("CONTA")) + "-" + RTrim(tbPlano("DESCRICAO")) + "  TRADUTOR:" + LTrim(Format(XTRADUTOR, "00000")), Imp_Normal
                           wl_Linha = wl_Linha + 1
                           IMPRIMIU = True
                        End If
                        
                        End If
                        
                        Monta_LinhadeImpressao wl_Linha, tbMoviCaixa("DATA"), 0, "E", imp_Condensado
                        XCONTRA = ""
                        XTRAD = ""
                        XNOMECONTRA = ""
                        
                        If XTRADUTOR = tbMoviCaixa("CREDITO") Then
                           tbPlano.Index = "iTRADUTOR"
                           tbPlano.Seek "=", tbMoviCaixa("DEBITO")
                           XCONTRA = tbPlano("CONTA")
                           XTRAD = tbMoviCaixa("DEBITO")
                           XNOMECONTRA = tbPlano("DESCRICAO")
                          Else
                           tbPlano.Index = "iTRADUTOR"
                           tbPlano.Seek "=", tbMoviCaixa("CREDITO")
                           XCONTRA = tbPlano("CONTA")
                           XNOMECONTRA = tbPlano("DESCRICAO")
                           XTRAD = tbMoviCaixa("CREDITO")
                        End If

                        XHISTORICO = RTrim(Mid(tbMoviCaixa("HISTORICO"), 1, 40))
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
                        Monta_LinhadeImpressao wl_Linha, RTrim(Mid$(XNOMECONTRA, 1, 30)), 3, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                        Monta_LinhadeImpressao wl_Linha, XHISTORICO, 4, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                        Monta_LinhadeImpressao wl_Linha, Format(tbMoviCaixa("MOVIMENTO"), "00000"), 5, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                        
                        tbPlano.Index = "iCONTA"
                        tbPlano.Seek "=", XCONTA
                        
                        If tbMoviCaixa("credito") = XTRADUTOR Then
                        
                           xcredito = xcredito + tbMoviCaixa("valor")
                           Monta_LinhadeImpressao wl_Linha, LTrim(Format(tbMoviCaixa("VALOR"), "#,##0.00")), 6, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                           xtotal = xtotal + tbMoviCaixa("VALOR")
                           
                          ElseIf tbMoviCaixa("debito") = XTRADUTOR Then
                           
                           Monta_LinhadeImpressao wl_Linha, LTrim(Format(tbMoviCaixa("VALOR"), "#,##0.00")), 6, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
                           xdebito = xdebito + tbMoviCaixa("valor")
                           xtotal = xtotal + tbMoviCaixa("VALOR")
                           
                        End If
                         
                        If Not pb_InverteOperacao Then
                           XATUAL = (XANTERIOR - xdebito) + xcredito
                         Else
                           XATUAL = (XANTERIOR - xcredito) + xdebito
                        End If
                                          
                                         
                        wl_Linha = wl_Linha + 0.5
                                     
                        
                  End If
                    
                  If Bar1.Value + 1 > Bar1.Max Then
                     Bar1.Max = Bar1.Max + 1
                  End If
                  
                  Bar1.Value = Bar1.Value + 1
                    
                  tbMoviCaixa.MoveNext
               
               Loop
               
       End If
           
     If xtotal > 0 Then
           
       wl_Linha = wl_Linha + 0.5
       
       Monta_LinhadeImpressao wl_Linha, "**********", 0, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
       Monta_LinhadeImpressao wl_Linha, "TOTAL DA CONTA NO DIA:" + dtoc(xdia), 4, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
       Monta_LinhadeImpressao wl_Linha, LTrim(Format(xtotal, "###,##0.00")), 6, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
       wl_Linha = wl_Linha + 0.5
       wl_Linha = wl_Linha + 0.5
       XTOTALGERAL = XTOTALGERAL + xtotal
       xtotal = 0
     End If
     
     If Bar1.Value + 1 > Bar1.Max Then
        Bar1.Max = Bar1.Max + 1
     End If
       
     Bar1.Value = Bar1.Value + 1
     
      tbPlano.MoveNext
      
      If tbPlano.EOF Or tbPlano.NoMatch Then
         Exit Do
      End If


Loop


      tbPlano.MoveFirst

      xdia = xdia + 1
     
      If xdia > XULTDIA Then
         Exit Do
      End If
  
Loop


     If XTOTALGERAL > 0 Then
           
       wl_Linha = wl_Linha + 0.5
       
       Monta_LinhadeImpressao wl_Linha, "**********", 0, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
       Monta_LinhadeImpressao wl_Linha, "TOTAL DO EXERCICIO:", 4, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
       Monta_LinhadeImpressao wl_Linha, LTrim(Format(XTOTALGERAL, "###,##0.00")), 6, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
       wl_Linha = wl_Linha + 0.5
       wl_Linha = wl_Linha + 0.5
       XTOTALGERAL = XTOTALGERAL + xtotal
       xtotal = 0
     
     
     End If


wl_Linha = wl_Linha + 0.5
          
Salta_Pagina
Finaliza_Impressao

MsgBox ("RELATORIO CONCLUIDO COM SUCESSO !")

Exit Sub


fSaldoAnterior:
tbSaldo.Seek "<", XTRADUTOR, XDATAINICIAL

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
Monta_Cabecalho pCabecalho, pReferencia, 6, wl_Linha, IIf(Option1.Value = True, imp_Condensado, Imp_Normal), "D I A R I O   L E G A L   D E     " + Mes_Extenso(Month(txtinicio.Pacote)) + "/" + Format(Year(txtinicio.Pacote), "0000"), wl_Pagina, Ultimo_DiasdoMes(Format(Month(Me.txtfim.Pacote), "00"), Format(Year(Me.txtfim.Pacote), "0000"))

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

Private Sub txtpagina_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtpagina.Text = "" Then
      X = MsgBox("É Necessário a informação da página inicial")
      Me.txtpagina.SetFocus
      Exit Sub
    Else
      Me.cmdimprimir.SetFocus
      Exit Sub
   End If
End If
End Sub


