VERSION 5.00
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fBalanceteSintetico 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balancete Sintético"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   100
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   "CONTA                     | DENOMINAÇÃO DA CONTA                                                        |  SALDO APURADO"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   8295
      Begin VB.CommandButton cmdImpressoras 
         Caption         =   "&Impressoras"
         Height          =   375
         Left            =   6840
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdimprimir2 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdimprimir 
         Caption         =   "&Gerar"
         Height          =   375
         Left            =   6840
         TabIndex        =   2
         Top             =   360
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
         TabIndex        =   7
         Top             =   960
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "80 Colunas Normal"
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
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Balancete Sintético"
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
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "fBalanceteSintetico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tbPlano As Recordset
Dim tbSaldo As Recordset
Dim tbMoviCaixa As Recordset
Dim tbBalanco As Recordset



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
Dim ZZ As Double

If tbBalanco.RecordCount > 0 Then
   tbBalanco.MoveFirst
End If
Grid1.row = 1
Do While Not tbBalanco.EOF

    If edit_reg(tbBalanco) Then
       tbBalanco.Delete
    End If
    tbBalanco.MoveNext
    
Loop

For ZZ = 1 To 99 Step 1

    Grid1.TextMatrix(ZZ, 0) = ""
    Grid1.TextMatrix(ZZ, 1) = ""
    Grid1.TextMatrix(ZZ, 2) = ""
    
Next ZZ


XDATAINICIAL = txtinicio.Pacote
XULTDIA = txtfim.Pacote

tbPlano.Index = "iCONTA"

If tbPlano.RecordCount > 0 Then

   tbPlano.MoveFirst
   
End If

wl_Linha = 0
wl_Pagina = 0

xgrupo = ""

Do While Not tbPlano.EOF



           xdebito = 0
           xcredito = 0
           XANTERIOR = 0
           XATUAL = 0
    
        If xgrupo = "" Or xgrupo <> Mid$(tbPlano("CONTA"), 1, 8) Then
                       wl_Linha = wl_Linha + 0.5
                       xgrupo = Mid$(tbPlano("CONTA"), 1, 8)
                       
                       
                       If Len(xgrupo) = 1 Then
                          xgrau1 = RTrim(Mid$(xgrupo, 1, 1))
                       End If
                       
                       If Len(xgrupo) = 3 Then
                          xgrau2 = RTrim(Mid$(xgrupo, 1, 3))
                       End If
                       
                       If Len(xgrupo) = 5 Then
                          xgrau3 = RTrim(Mid$(xgrupo, 1, 5))
                       End If
                       
                       If Len(xgrupo) = 8 Then
                          XGRAU4 = RTrim(Mid$(xgrupo, 1, 8))
                       End If
                     
                       If add_reg(tbBalanco) Then
                          tbBalanco("conta") = xgrupo
                          tbBalanco("descricao") = tbPlano("descricao")
                          tbBalanco("valor") = 0
                          
                          If update_reg(tbBalanco) Then
                          End If
                       
                       End If
                                                
                        
                       Grid1.row = Grid1.row + 1
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
            
            If xgrau1 <> "" Then
               tbBalanco.Seek "=", xgrau1
               If Not tbBalanco.NoMatch Then
                  If edit_reg(tbBalanco) Then
                     tbBalanco("valor") = tbBalanco("valor") + XATUAL
                     If update_reg(tbBalanco) Then
                     End If
                     
                  End If
               End If
            End If
            
            If xgrau3 <> "" And XGRAU4 = "" Then
               tbBalanco.Seek "=", xgrau3
               If Not tbBalanco.NoMatch Then
                  If edit_reg(tbBalanco) Then
                     tbBalanco("valor") = tbBalanco("valor") + XATUAL
                     If update_reg(tbBalanco) Then
                     End If
                     
                  End If
               End If
               
               tbBalanco.Seek "=", xgrau2
               If Not tbBalanco.NoMatch Then
                  If edit_reg(tbBalanco) Then
                     tbBalanco("valor") = tbBalanco("valor") + XATUAL
                     If update_reg(tbBalanco) Then
                     End If
                     
                  End If
               End If
            
            End If
            
            If XGRAU4 <> "" Then
               
               tbBalanco.Seek "=", XGRAU4
               If Not tbBalanco.NoMatch Then
                  If edit_reg(tbBalanco) Then
                     tbBalanco("valor") = tbBalanco("valor") + XATUAL
                     If update_reg(tbBalanco) Then
                     End If
                     
                  End If
               End If
            
               tbBalanco.Seek "=", xgrau3
               If Not tbBalanco.NoMatch Then
                  If edit_reg(tbBalanco) Then
                     tbBalanco("valor") = tbBalanco("valor") + XATUAL
                     If update_reg(tbBalanco) Then
                     End If
                     
                  End If
               End If
               
               tbBalanco.Seek "=", xgrau2
               If Not tbBalanco.NoMatch Then
                  If edit_reg(tbBalanco) Then
                     tbBalanco("valor") = tbBalanco("valor") + XATUAL
                     If update_reg(tbBalanco) Then
                     End If
                     
                  End If
               End If
            
            
            End If
            
         
            
              If xdebito <> 0 Or xcredito <> 0 Or XANTERIOR <> 0 Then
              
                If Option1.Value = True Then
                
                 
                 
                Else
                
                

                End If
                
                 XTOTDEBITO = XTOTDEBITO + xdebito
                 XTOTCREDITO = XTOTCREDITO + xcredito
                 
                 wl_Linha = wl_Linha + 0.5
                
              
              End If
           
        

      End If
      
      tbPlano.MoveNext
      
      If tbPlano.EOF Or tbPlano.NoMatch Then
         Exit Do
      End If

Loop

                 wl_Linha = wl_Linha + 0.5

If tbBalanco.RecordCount > 0 Then
   tbBalanco.MoveFirst
End If
Grid1.row = 1

Do While Not tbBalanco.EOF

        Grid1.TextMatrix(Grid1.row, 0) = " " + tbBalanco("conta")
        Grid1.TextMatrix(Grid1.row, 1) = Space(Len(tbBalanco("conta"))) + "" + tbBalanco("descricao")
        Grid1.TextMatrix(Grid1.row, 2) = Format(tbBalanco("valor"), "#,##0.00")
        
        Grid1.row = Grid1.row + 1
        
        tbBalanco.MoveNext
        
        If tbBalanco.EOF Then
           Exit Do
        End If
        
        
Loop

Grid1.row = 1

cmdimprimir2.Enabled = True

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
' Monta_Cabecalho pCabecalho, pReferencia, 6, wl_Linha, IIf(Option1.Value = True, imp_Condensado, Imp_Normal), "BALANCETE ANALITICO DE " + UCase(Me.txtmes.Text) + "/" + txtano.Text, wl_Pagina

Return





End Sub

Private Sub cmdimprimir2_Click()
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
Dim xgrau As String
Dim xtotal As Currency


If tbBalanco.RecordCount <= 0 Then
   X = MsgBox("Não existem relatórios gerados no momento !")
   Me.cmdimprimir2.SetFocus
   Exit Sub
  Else
   tbBalanco.MoveFirst
End If


aadd pCabecalho, Array("CODIGO DA CONTA", "DENOMINACAO DA CONTA", "SALDO APURADO")
aadd pReferencia, Array("9.9.99.99999-99", "AAAAAAAAAAAAAAAAAAAAAAAAAAABBBBBBBBBCCCDDXXX", "9,999,999,999.99")

XDATAINICIAL = Me.txtinicio.Pacote
XULTDIA = Me.txtfim.Pacote

If Not PadraodeImpressao Then

   Me.txtinicio.SetFocus
   Exit Sub
   
End If


wl_Linha = 0
wl_Pagina = 0
   
tbBalanco.MoveFirst
xgrau = Mid$(tbBalanco("conta"), 1, 1)
xtotal = 0

Do While Not tbBalanco.EOF
         
        
        If wl_Linha = 0 Or wl_Linha > 28 Then
           If wl_Linha > 28 Then
              Salta_Pagina
           End If
           wl_Linha = 0
           GoSub CABECALHO
        End If


        If Len(tbBalanco("conta")) = 1 Then
           xtotal = tbBalanco("valor")
        End If

        Monta_LinhadeImpressao wl_Linha, tbBalanco("conta"), 0, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
        Monta_LinhadeImpressao wl_Linha, Space(Len(tbBalanco("conta"))) + "" + tbBalanco("descricao"), 1, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
        Monta_LinhadeImpressao wl_Linha, Numero_Contabil(tbBalanco("valor")), 2, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
        
        wl_Linha = wl_Linha + 0.5
        
        tbBalanco.MoveNext
        
        If tbBalanco.EOF Or tbBalanco.NoMatch Then
           Exit Do
        End If
        
        If Mid(tbBalanco("conta"), 1, 1) <> xgrau Then
           xgrau = Mid$(tbBalanco("conta"), 1, 1)
           wl_Linha = wl_Linha + 1
           Monta_LinhadeImpressao wl_Linha, "T O T A L . . . . . .", 1, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
           Monta_LinhadeImpressao wl_Linha, Numero_Contabil(xtotal), 2, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
           wl_Linha = wl_Linha + 0.5
           Salta_Pagina
           wl_Linha = 0
           GoSub CABECALHO
        End If
        

Loop

           wl_Linha = wl_Linha + 1
           Monta_LinhadeImpressao wl_Linha, "T O T A L . . . . . .", 1, "E", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
           Monta_LinhadeImpressao wl_Linha, Numero_Contabil(xtotal), 2, "D", IIf(Option1.Value = True, imp_Condensado, Imp_Normal)
           wl_Linha = wl_Linha + 0.5

Salta_Pagina
Finaliza_Impressao

MsgBox ("RELATORIO IMPRESSO COM SUCESSO !")
Exit Sub


CABECALHO:
wl_Pagina = wl_Pagina + 1
Monta_Cabecalho pCabecalho, pReferencia, 2, wl_Linha, IIf(Option1.Value = True, imp_Condensado, Imp_Normal), "BALANCETE SINTETICO DE " + Mes_Extenso(Month(txtinicio.Pacote)) + "/" + Format(Year(txtinicio.Pacote), "0000"), wl_Pagina, Ultimo_DiasdoMes(Format(Month(Me.txtfim.Pacote), "00"), Format(Year(Me.txtfim.Pacote), "0000"))
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
       Abre_SaldoContas(tbSaldo) Or Not _
       Abre_BalanceteSintetico(tbBalanco) Then
        
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





