VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm fMENU 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Infosoft Contabilidade Geral"
   ClientHeight    =   6525
   ClientLeft      =   3660
   ClientTop       =   -270
   ClientWidth     =   8880
   Icon            =   "fMENU.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "fMENU.frx":0442
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1376
      ButtonWidth     =   2249
      ButtonHeight    =   1323
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lançamentos"
            Description     =   "pagar"
            Object.ToolTipText     =   "Cadastro de Contas a Pagar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Plano de Contas"
            Description     =   "receber"
            Object.ToolTipText     =   "Lançamentos de Contas à Receber"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Históricos"
            Description     =   "contabilidade"
            Object.ToolTipText     =   "Lançamentos Contábeis"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Extratos"
            Description     =   "fluxocaixa"
            Object.ToolTipText     =   "Fluxo de Caixa"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Saída"
            Description     =   "saida"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   31
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMENU.frx":4E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMENU.frx":51B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMENU.frx":560A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMENU.frx":5A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMENU.frx":5FFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   9480
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar Progressao 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   6060
      Visible         =   0   'False
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar BarraStatus 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   6285
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9798
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "O nome do usuário atual {S} = Supervisor"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
      EndProperty
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
   Begin VB.Menu MENU00 
      Caption         =   "&Manutenção"
      Begin VB.Menu Manutencao_Empresas 
         Caption         =   "&Empresas"
         Begin VB.Menu Empresa_Manutencao 
            Caption         =   "&Manutenção"
         End
         Begin VB.Menu Empresa_Relatorio 
            Caption         =   "&Relatório"
         End
      End
      Begin VB.Menu hist0101 
         Caption         =   "Históricos"
      End
      Begin VB.Menu Manutencao_PlanoContas 
         Caption         =   "&Plano de Contas"
      End
      Begin VB.Menu Manutencao_LanPadrao 
         Caption         =   "Lançamentos Padrão"
      End
      Begin VB.Menu trc 
         Caption         =   "-"
      End
      Begin VB.Menu Movimento_Lancamento 
         Caption         =   "&Lançamentos"
         Shortcut        =   ^C
      End
      Begin VB.Menu Manutencoes_Extrato_MoviCaixa 
         Caption         =   "&Extrato de Conta"
         Shortcut        =   ^E
      End
      Begin VB.Menu vericon 
         Caption         =   "Verifica Inconsistências"
      End
   End
   Begin VB.Menu Menu_Relatorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu Movimento_Conferencia 
         Caption         =   "&Relatório de Conferência"
      End
      Begin VB.Menu rolhist 
         Caption         =   "Históricos"
      End
      Begin VB.Menu Relatorio_Plano 
         Caption         =   "&Plano de Contas"
      End
      Begin VB.Menu GRE1 
         Caption         =   "-"
      End
      Begin VB.Menu rolcont01 
         Caption         =   "Balancete Analítico"
      End
      Begin VB.Menu rolcont02 
         Caption         =   "Balancete Sintético"
      End
      Begin VB.Menu rolcont03 
         Caption         =   "Razão Analítico"
      End
      Begin VB.Menu rolcont04 
         Caption         =   "Diário Legal"
      End
   End
   Begin VB.Menu MENU01 
      Caption         =   "R&otina Diversas"
      Begin VB.Menu MENU01_01 
         Caption         =   "&Troca de Usuário"
         Shortcut        =   ^T
      End
      Begin VB.Menu MENU01_03 
         Caption         =   "&Segurança"
         Shortcut        =   ^S
      End
      Begin VB.Menu MENU01_05 
         Caption         =   "&Preferências"
      End
      Begin VB.Menu TR 
         Caption         =   "-"
      End
      Begin VB.Menu verifica 
         Caption         =   "Verifica Estruturas de Dados"
      End
      Begin VB.Menu Rotina_Repara 
         Caption         =   "Repara &Banco de Dados"
      End
      Begin VB.Menu Rotina_Menu 
         Caption         =   "&Atualiza Menu"
      End
      Begin VB.Menu Rotina_Saldo 
         Caption         =   "Recalcula &Saldo"
      End
      Begin VB.Menu Rotina_GeraLanca 
         Caption         =   "&Gera Lançamentos Contábeis"
      End
      Begin VB.Menu Rotina_PlanoPrograma 
         Caption         =   "&Plano de Contas Programado"
      End
      Begin VB.Menu Rotina_Visualiza 
         Caption         =   "&Visualiza Relatório Último Relatório no Vídeo"
         Shortcut        =   ^W
      End
      Begin VB.Menu rotinas_campos 
         Caption         =   "&Analisa campos adicionados"
      End
      Begin VB.Menu rotina_importa 
         Caption         =   "&Importa Plano de Contas"
      End
   End
   Begin VB.Menu MENU90 
      Caption         =   "&Aplicativos"
      Begin VB.Menu MENU90_01 
         Caption         =   "&Calculadora"
      End
      Begin VB.Menu MENU90_02 
         Caption         =   "&Data e Hora"
      End
      Begin VB.Menu MENU90_03 
         Caption         =   "&Configurações Regionais"
      End
      Begin VB.Menu MENU90_04 
         Caption         =   "&Monitor"
      End
   End
   Begin VB.Menu MENU99 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "fMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private wp_Entrada As Boolean
Public pb_Online As Boolean

Private Sub Configura_Menus()
'Configura Menu Comum
On Error Resume Next
End Sub

Private Sub Monta_2Plano()
If Me.Width > 11000 And Dir(PathWindows + "logo800600.jpg") <> "" Then
   Me.Picture = LoadPicture(PathWindows + "logo800600.jpg")
ElseIf Me.Width < 11000 And Dir(PathWindows + "logo640480.JPG") <> "" Then
   Me.Picture = LoadPicture(PathWindows + "logo640480.jpg")
End If
End Sub


Private Sub Empresa_Manutencao_Click()
If Not Verifica_Privilegio(PR_EMPRESAS, "C") Then
   InformaaoUsuario "Usuário sem privilégio para acessar Empresas"
   Exit Sub
End If
fEMPRESAS.Show 1
End Sub

Private Sub Empresa_Relatorio_Click()
Dim tbEmpresas As Recordset
Dim pCabecalho
Dim pReferencia
Dim pCampos
If Not Abre_Empresas(tbEmpresas) Then
  Exit Sub
End If

tbEmpresas.Index = "iRAZAO"

aadd pCabecalho, Array("Código", "Razao Social")
aadd pReferencia, Array("99999", "AAAAAAAAAABBBBBBBBBBCCCCCCCCCCDDDDDDDDDDEEEEEEEEEE")

AddColunaImpressao pCampos, "CODIGO", , "00000"
AddColunaImpressao pCampos, "RAZAOSOCIAL"

RelatorioPadrao tbEmpresas, "Relatório de Empresas", pCabecalho, pReferencia, pCampos
tbEmpresas.Close
End Sub


Private Sub hist0101_Click()
fHISTORICO.Show
End Sub


Private Sub Manutencao_LanPadrao_Click()
If Not Verifica_Privilegio(PR_LANPADRAO, "C") Then
   InformaaoUsuario "Sem privilégio para acessar Lançamentos Padrao"
   Exit Sub
End If
fLANPADRAO.Show
fLANPADRAO.SetFocus
End Sub

Private Sub Manutencao_PlanoContas_Click()
fPLANOCONTAS.Show
fPLANOCONTAS.SetFocus
End Sub

Private Sub Manutencoes_Extrato_MoviCaixa_Click()
If Not Verifica_Privilegio(PR_EXTRATO, "E") Then
   InformaaoUsuario "Usuário sem privilégio para Extrato"
   Exit Sub
End If
fEXTRATO.Show
End Sub

Private Sub MDIForm_Activate()
On Error Resume Next
Dim wl_retorno As String
If Not wp_Entrada Then
   wp_Entrada = True
  
   Configura_Menus
   If Not Permissao(PR_SISTEMA, "C", "Acesso ao Sistema") Then
      Call MsgBox("Usuário sem permissão para acesso", vbCritical, "Mensagem do Sistema")
      End
   End If
'   pic_DEMO.Left = Width - pic_DEMO.Width - 100
'   pic_DEMO.Visible = pb_Demonstracao
   If pb_Demonstracao Then
      MsgBox "Essa é uma cópia não registrada do aplicativo. Entre em contato com a Infosoft para regularização.", vbInformation, "Mensagem do Sistema"
   End If
   InformaEmpresa
   wl_retorno = RetornaConfiguracao("Preferencias_" + Format(pb_Empresa, "000"), "PDV")
   If wl_retorno = "" Then
      fPDV.Show 1
   End If
   pb_NivelPlano = Val(RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "NivelPlanodeContas"))
   wl_retorno = Val(RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "RegimeCompetencia"))
   pb_RegimeCompetencia = IIf(wl_retorno = "", True, IIf(wl_retorno = "1", True, False))
   wl_retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "Online")
   pb_Online = False 'IIf(wl_Retorno = "", False, IIf(wl_Retorno = "1", True, False))
   If pb_InverteOperacao Then
      Me.BarraStatus.Panels(3).Text = "D+"
      Me.BarraStatus.Panels(3).ToolTipText = "Débito Adiciona ao Saldo"
   Else
      Me.BarraStatus.Panels(3).Text = "D-"
      Me.BarraStatus.Panels(3).ToolTipText = "Débito Subtrai do Saldo"
   End If
'    Verifica_Estruturas
   Display_Usuario
   Inicializa_Impressora

End If
End Sub





Private Sub MENU01_01_Click()
Call Troca_Usuario
End Sub


Private Sub MENU01_03_Click()
If Verifica_Privilegio(PR_SEGURANCA, "C", "Sem privilégio para acessar SEGURANÇA") Then
   fSEGURANCA.Show 1
End If
End Sub


Private Sub MENU01_05_Click()
Dim WL_CAPTION As String
On Error Resume Next
Err = 0
WL_CAPTION = fMENU.ActiveForm.Caption
If Err = 0 Then
   MsgBox "Feche o formulário '" + WL_CAPTION + "', antes de alterar preferências", vbExclamation, "Mensagem do Sistema"
   Exit Sub
End If
If Verifica_Privilegio(PR_ROTINA, "P", "Sem privilégio para acessar preferências") Then
   fPREFERENCIAS.Show 1
End If
End Sub


Private Sub MENU90_01_Click()
Shell "CALC.EXE", vbNormalFocus
End Sub

Private Sub MENU90_02_Click()
Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End Sub

Private Sub MENU90_03_Click()
 Call Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
End Sub


Private Sub MENU90_04_Click()
Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Sub


Private Sub MENU99_Click()
If Confirme("Confirma a saída do Sistema") Then End
End
End Sub

Private Sub Movimento_Conferencia_Click()
fCONFEREMOVIMENTO.Show 1
End Sub

Private Sub Movimento_Lancamento_Click()
If Not Verifica_Privilegio(PR_LANCAMENTOS, "C") Then
   InformaaoUsuario "Usuário sem privilégio para acessar Movimento de Caixa"
   Exit Sub
End If
fCAIXA.Show
fCAIXA.SetFocus
End Sub


Private Sub Relatorio_Plano_Click()
Dim tbPlano As Recordset
Dim wl_Linha As Currency
Dim wl_Cabecalho
Dim wl_Referencia
Dim wl_Pagina As Integer
Dim wl_Tradutor As Long
aadd wl_Cabecalho, Array("CONTA", "TRADUTOR", "DESCRIÇÃO", "SALDO ABERTURA")
aadd wl_Referencia, Array("99.999.99999", "99999", "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", "9,999,999.99")


If file(PathWindows + "mantovani.sys") Then
   MsgBox "Configurado para a Industria Mantovani"
End If

If Not Abre_PlanoContas(tbPlano) Then
   Exit Sub
End If

If tbPlano.RecordCount = 0 Then tbPlano.Close: Exit Sub

tbPlano.Index = "iCONTA"
tbPlano.MoveFirst

If Not PadraodeImpressao Then Exit Sub

Do While Not tbPlano.EOF
   
   If wl_Linha = 0 Then
      wl_Pagina = wl_Pagina + 1
      Monta_Cabecalho wl_Cabecalho, wl_Referencia, 3, wl_Linha, imp_Condensado, "Plano de Contas", wl_Pagina
   End If
   
   If Len(tbPlano("conta")) <= 7 Then
      
      If Len(tbPlano("conta")) = 1 Then
         If file(PathWindows + "mantovani.sys") Then
            Monta_LinhadeImpressao wl_Linha, tbPlano("conta") + "0000", 0, "E", imp_Condensado
           Else
            Monta_LinhadeImpressao wl_Linha, tbPlano("conta"), 0, "E", imp_Condensado
         End If
      End If
             
      If Len(tbPlano("conta")) = 3 Then
         If file(PathWindows + "mantovani.sys") Then
            Monta_LinhadeImpressao wl_Linha, Mid(tbPlano("conta"), 1, 1) + Mid(tbPlano("conta"), 3, 1) + "000", 0, "E", imp_Condensado
           Else
            Monta_LinhadeImpressao wl_Linha, tbPlano("conta"), 0, "E", imp_Condensado
         End If
      End If
             
      If Len(tbPlano("conta")) = 5 Then
         If file(PathWindows + "mantovani.sys") Then
            Monta_LinhadeImpressao wl_Linha, Mid(tbPlano("conta"), 1, 1) + Mid(tbPlano("conta"), 3, 1) + Mid(tbPlano("conta"), 5, 1) + "00", 0, "E", imp_Condensado
           Else
            Monta_LinhadeImpressao wl_Linha, tbPlano("conta"), 0, "E", imp_Condensado
         End If
      End If

      If Len(tbPlano("conta")) = 7 Then
         If file(PathWindows + "mantovani.sys") Then
            Monta_LinhadeImpressao wl_Linha, tbPlano("conta") + "00000", 0, "E", imp_Condensado
          Else
            Monta_LinhadeImpressao wl_Linha, tbPlano("conta"), 0, "E", imp_Condensado
         End If
      End If
         
      Else
         
      If file(PathWindows + "mantovani.sys") Then
         Monta_LinhadeImpressao wl_Linha, Formata_Conta_Mantovani(tbPlano("conta")), 0, , imp_Condensado
        Else
         Monta_LinhadeImpressao wl_Linha, tbPlano("CONTA"), 0, , imp_Condensado
      End If
   
   End If
   
   If tbPlano("tradutor") > 0 Then
      Monta_LinhadeImpressao wl_Linha, LTrim(Format(tbPlano("TRADUTOR"), "00000")), 1, "D", imp_Condensado
   End If
   
   Monta_LinhadeImpressao wl_Linha, Space(Len(tbPlano("conta"))) + RTrim(tbPlano("DESCRICAO")), 2, "E", imp_Condensado
   
   If tbPlano("tradutor") > 0 Then
      Monta_LinhadeImpressao wl_Linha, Format(tbPlano("SALDOABERTURA"), "###,##0.00"), 3, "D", imp_Condensado
   End If
   
   wl_Linha = wl_Linha + 0.5
   If wl_Linha > IIf(pb_ImpressaoMatricial, 29, 26) Then
      wl_Linha = 0
      Salta_Pagina
   End If
   
   wl_Tradutor = tbPlano("TRADUTOR")
   tbPlano.MoveNext
   If Not tbPlano.EOF Then
      If wl_Tradutor = 0 Or tbPlano("TRADUTOR") = 0 Then
         wl_Linha = wl_Linha + 0.5
      End If
   End If

Loop
Finaliza_Impressao
tbPlano.Close
Exit Sub
      
   






End Sub


Private Sub rolcont01_Click()
fBalanceteAnalitico.Show
End Sub

Private Sub rolcont02_Click()
fBalanceteSintetico.Show
End Sub

Private Sub rolcont03_Click()
fRazaoAnalitico.Show
End Sub

Private Sub rolcont04_Click()
fDiarioLegal.Show
End Sub



Private Sub rolhist_Click()
Dim tbhistoricos As Recordset
Dim pCabecalho
Dim pReferencia
Dim pCampos
If Not Abre_Historico(tbhistoricos) Then
  Exit Sub
End If

aadd pCabecalho, Array("Código", "Descrição")
aadd pReferencia, Array("99999", "AAAAAAAAAABBBBBBBBBBCCCCCCCCCCDDDDDDDDDDEEEEEEEEEEFFFFFFFFFFFFFFFFFFFFFFFFF")

AddColunaImpressao pCampos, "CODIGO", , "00000"
AddColunaImpressao pCampos, "DESCRICAO"

RelatorioPadrao tbhistoricos, "Relatório de Historicos", pCabecalho, pReferencia, pCampos, imp_Condensado
tbhistoricos.Close
Exit Sub

End Sub

Private Sub Rotina_GeraLanca_Click()
If RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "FornecedoresDiversos") = "" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "ClientesDiversos") = "" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "JURO_C") = "" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "DESCONTO_C") = "" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "CORRECAO_C") = "" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "Taxa") = "" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "FornecedoresDiversos") = "" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "JURO_F") = "" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "DESCONTO_F") = "" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "CORRECAO_F") = "" Then
    InformaaoUsuario "Verifique as preferências de contas ..."
    Exit Sub
End If
If RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "FornecedoresDiversos") = "0" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "ClientesDiversos") = "0" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "JURO_C") = "0" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "DESCONTO_C") = "0" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "CORRECAO_C") = "0" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "Taxa") = "0" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "FornecedoresDiversos") = "0" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "JURO_F") = "0" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "DESCONTO_F") = "0" Or _
    RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "CORRECAO_F") = "0" Then
    InformaaoUsuario "Verifique as preferências de contas ..."
    Exit Sub
End If
fGERALANCA.Show 1
End Sub

Private Sub rotina_importa_Click()
' Form1.Show
'FImportaPlanoRural.Show
End Sub


Private Sub Rotina_PlanoPrograma_Click()
fPLANOPROGRAMA.Show 1
End Sub


Private Sub Rotina_Saldo_Click()
Dim wl_InicioCalculo As String
wl_InicioCalculo = "01/" + Format(Month(Date), "00") + "/" + Format(Year(Date), "0000")
Do While True
  wl_InicioCalculo = InputBox("Data Inicial", "Recalcula Saldo", wl_InicioCalculo)
  If IsDate(wl_InicioCalculo) Or wl_InicioCalculo = "" Then Exit Do
  If MsgBox("Data inválida. Deseja Abandonar?", vbYesNo, "Mensagem do Sistema") = vbYes Then
     Exit Sub
   End If
Loop
If wl_InicioCalculo = "" Then
   Recalcula_Saldo
Else
   Recalcula_Saldo CDate(wl_InicioCalculo)
End If
End Sub

Private Sub Rotina_Visualiza_Click()
If Dir(PathWindows + "WORDPAD.EXE") = "" Then
   InformaaoUsuario "Para visualizar os relatórios no vídeo localize e copie o aplicativo: WORDPAD.EXE, para a pasta " + PathWindows
   Exit Sub
End If
If Dir(PathWindows + "RTF\REPORT.RTF") = "" Then
   InformaaoUsuario "Não existe relatório impresso"
   Exit Sub
End If
Shell "WORDPAD.EXE " + PathWindows + "RTF\REPORT.RTF", vbMaximizedFocus
End Sub



Private Sub Rotina_Menu_Click()
Call Configura_Menus
End Sub

Private Sub Rotina_Repara_Click()
Dim WL_CAPTION As String
On Error Resume Next
WL_CAPTION = fMENU.ActiveForm.Caption
If Err = 0 Then
   MsgBox "Feche o formulário '" + WL_CAPTION + "', antes de reparar banco de dados", vbExclamation, "Mensagem do Sistema"
   Exit Sub
End If
If Not Verifica_Privilegio(PR_ROTINA, "R") Then
   InformaaoUsuario "Usuário sem privilégio para reparar banco de dados"
   Exit Sub
End If
If MsgBox("Essa opção executa um operação demorada. Tem certeza que deseja continuar?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbYes Then
   dbFINANCEIRO.Close
   dbPROTECAO.Close
   dbPRODUTOS.Close
   dbMOVIMENTO.Close
   dbVendas.Close
   dbFluxo.Close
   dbEMPRESAS.Close
   dbPreferencias.Close
   Call Repara_Banco
End If
End Sub

'Private Sub rotinas_campos_Click()
'Dim tbPagar As Recordset
'Dim tbClientes As Recordset
'Dim tbCidades As Recordset
'Dim tbBairros As Recordset
'Dim tbFornecedores As Recordset
'Dim tbPlano As Recordset
'If Not Abre_APagar(tbPagar) Or _
'   Not Abre_Clientes(tbClientes) Or _
'   Not Abre_Cidades(tbCidades) Or _
'   Not Abre_Bairros(tbBairros) Or _
'   Not Abre_Fornecedores(tbFornecedores) Or _
'   Not Abre_PlanoContas(tbPlano) Then
'   Exit Sub
'End If
'Informacao
'Do While Not tbPlano.EOF
'   If IsNull(tbPlano("ATIVOPASSIVO")) Then
'      If edit_reg(tbPlano) Then
'         tbPlano("ATIVOPASSIVO") = 0
'         update_reg tbPlano
'      End If
'   End If
'   tbPlano.MoveNext
'Loop
'Do While Not tbPagar.EOF
'   DisplayMensagem "Atualizando movimento " + Str(tbPagar("MOVIMENTO")) + " do Contas a Pagar."
'   If IsNull(tbPagar("PROCESSO")) Then
'      If edit_reg(tbPagar) Then
'         tbPagar("PROCESSO") = 0
'         update_reg tbPagar
'      End If
'   End If
'   tbPagar.MoveNext
'Loop
'Do While Not tbClientes.EOF
'   DisplayMensagem "Atualizando cliente " + tbClientes("NOME") + Space(50 - Len(tbClientes("NOME")))
'   If Loca_Cidades(tbCidades, tbClientes("CIDADE")) Then
'      If edit_reg(tbClientes) Then
'         tbClientes("DESCR_CIDADE") = tbCidades("DESCRICAO")
'         update_reg tbClientes
'      End If
'   End If
'   If Loca_Bairros(tbBairros, tbClientes("CIDADE"), tbClientes("BAIRRO")) Then
'      If edit_reg(tbClientes) Then
'         tbClientes("DESCR_BAIRRO") = tbBairros("NOME")
'         update_reg tbClientes
'      End If
'   End If
'   tbClientes.MoveNext
'Loop
'Do While Not tbFornecedores.EOF
'   DisplayMensagem "Atualizando Fornecedor " + tbFornecedores("RAZAOSOCIAL") + Space(50 - Len(tbFornecedores("RAZAOSOCIAL")))
'   If Not IsNull(tbFornecedores("CIDADE")) Then
'      If Loca_Cidades(tbCidades, tbFornecedores("CIDADE")) Then
'         If edit_reg(tbFornecedores) Then
'            tbFornecedores("DESCR_CIDADE") = tbCidades("DESCRICAO")
'            update_reg tbFornecedores
'         End If
'      End If
'   End If
'   tbFornecedores.MoveNext
'Loop
'tbPagar.Close
'tbClientes.Close
'tbBairros.Close
'tbCidades.Close
'tbFornecedores.Close
'Informacao
'End Sub


Private Sub Timer2_Timer()
'If Not pic_DEMO.Visible Then pic_DEMO.Visible = False
'lblDEMO1.Visible = Not lblDEMO1.Visible
'lblDEMO2.Visible = Not lblDEMO2.Visible
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Dim wl_botao As Button
Dim wl_Opcao As String
Dim wl_Tag As String
Set wl_botao = Button
wl_Opcao = Button.key
wl_Tag = Button.Tag
If Button.Description = "pagar" Then
   Movimento_Lancamento_Click
ElseIf Button.Description = "receber" Then
   Manutencao_PlanoContas_Click
ElseIf Button.Description = "contabilidade" Then
   hist0101_Click
ElseIf Button.Description = "fluxocaixa" Then
   Manutencoes_Extrato_MoviCaixa_Click
ElseIf Button.Description = "saida" Then
   MENU99_Click
End If
End Sub






'Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'Dim wl_escolha As String
'wl_escolha = UCase(ButtonMenu.Text)
'If wl_escolha = "CONTAS A RECEBER" Then Manutencoes_ContasaReceber_Manutencao_Click
'If wl_escolha = "CONTAS A PAGAR" Then Manutencoes_ContasaPagar_Manutencao_Click
'If wl_escolha = "FLUXO DE CAIXA" Then Relatorio_Fluxo_Click
'If wl_escolha = "EXTRATO DE CONTA" Then Relatorio_Extrato_Click
'If wl_escolha = "RECALCULA SALDO" Then Rotina_Saldo_Click
'If wl_escolha = "MOVIMENTO DE CONTAS" Then Movimento_Lancamento_Click
'End Sub




Private Sub XX_Click()
Dim tbMovi As Recordset
If Not Abre_MoviCaixa(tbMovi) Then Exit Sub
tbMovi.MoveFirst
Do While Not tbMovi.EOF
   If UCase(Mid(tbMovi("HISTORICO"), 1, 6)) = "REGIME" Then
      If edit_reg(tbMovi) Then tbMovi.Delete
   ElseIf AT("BAIXA", UCase(tbMovi("HISTORICO"))) <> 0 And _
      (AT("CR", UCase(tbMovi("HISTORICO"))) <> 0 Or AT("CP", UCase(tbMovi("HISTORICO"))) <> 0) Then
         If edit_reg(tbMovi) Then tbMovi.Delete
   ElseIf AT("JUROS", UCase(tbMovi("HISTORICO"))) <> 0 And _
      (AT("CR", UCase(tbMovi("HISTORICO"))) <> 0 Or AT("CP", UCase(tbMovi("HISTORICO"))) <> 0) Then
         If edit_reg(tbMovi) Then tbMovi.Delete
   ElseIf AT("CORREÇÃO", UCase(tbMovi("HISTORICO"))) <> 0 And _
      (AT("CR", UCase(tbMovi("HISTORICO"))) <> 0 Or AT("CP", UCase(tbMovi("HISTORICO"))) <> 0) Then
         If edit_reg(tbMovi) Then tbMovi.Delete
   ElseIf AT("CORRECAO", UCase(tbMovi("HISTORICO"))) <> 0 And _
      (AT("CR", UCase(tbMovi("HISTORICO"))) <> 0 Or AT("CP", UCase(tbMovi("HISTORICO"))) <> 0) Then
         If edit_reg(tbMovi) Then tbMovi.Delete
   ElseIf AT("DESCONTO", UCase(tbMovi("HISTORICO"))) <> 0 And _
      (AT("CR", UCase(tbMovi("HISTORICO"))) <> 0 Or AT("CP", UCase(tbMovi("HISTORICO"))) <> 0) Then
         If edit_reg(tbMovi) Then tbMovi.Delete
   End If
   tbMovi.MoveNext
Loop
End
End Sub


Private Sub vericon_Click()
fVerificaInconsistencias.Show
End Sub

Private Sub verifica_Click()
Verifica_Estruturas
End Sub
