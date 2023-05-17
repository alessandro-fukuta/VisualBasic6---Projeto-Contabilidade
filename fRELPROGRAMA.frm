VERSION 5.00
Object = "{5756E734-2046-400A-BC65-0E105EC5876E}#1.0#0"; "CAIXATEX.OCX"
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Begin VB.Form fRELPROGRAMA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatórios Programados"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "fRELPROGRAMA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFFF&
      Height          =   2820
      Left            =   2565
      Pattern         =   "*.REL"
      TabIndex        =   6
      Top             =   210
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.ComboBox cmbANO 
      Height          =   315
      ItemData        =   "fRELPROGRAMA.frx":000C
      Left            =   4365
      List            =   "fRELPROGRAMA.frx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1170
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ComboBox cmbMES 
      Height          =   315
      ItemData        =   "fRELPROGRAMA.frx":00D9
      Left            =   2265
      List            =   "fRELPROGRAMA.frx":0101
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1170
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "&Imprime"
      Enabled         =   0   'False
      Height          =   375
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1425
   End
   Begin VB.CheckBox chkSALDO 
      Caption         =   "Mostra Saldo Calculado"
      Enabled         =   0   'False
      Height          =   240
      Left            =   30
      TabIndex        =   8
      Top             =   855
      Value           =   1  'Checked
      Width           =   2280
   End
   Begin VB.CheckBox chkMES 
      Caption         =   "&Mês de Referência"
      Enabled         =   0   'False
      Height          =   240
      Left            =   30
      TabIndex        =   4
      Top             =   630
      Value           =   1  'Checked
      Width           =   1680
   End
   Begin CaixaTexto.Caixa_Texto txtNOME 
      Height          =   300
      Left            =   30
      TabIndex        =   0
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "..."
      Height          =   285
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   405
   End
   Begin Mascara.Máscara txtDATAI 
      Height          =   300
      Left            =   930
      TabIndex        =   1
      Top             =   1140
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
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
   Begin Mascara.Máscara txtDATAF 
      Height          =   300
      Left            =   930
      TabIndex        =   2
      Top             =   1470
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
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
   Begin VB.Label lblDATAF 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblDATAI 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   1230
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblMES 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Mês de Referência:"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   1230
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Relatório"
      Height          =   195
      Left            =   30
      TabIndex        =   7
      Top             =   45
      Width           =   1320
   End
End
Attribute VB_Name = "fRELPROGRAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub Imprime_RelPrograma()
Dim tbPlano As Recordset
Dim tbMovi As Recordset
Dim tbSaldo As Recordset
Dim tbRelPrograma As Recordset
Dim wl_DataInicial As String
Dim wl_Datafinal As String
Dim wl_dia As Integer
Dim wl_Cabecalho
Dim wl_Referencia
Dim wl_Linha As String
Dim wl_Conta As String
Dim i As Integer
Dim wl_SaldoAnterior As Currency
Dim wl_SaldoDebito As Currency
Dim wl_SaldoCredito As Currency
Dim wl_NivelAcima As String
Dim wl_LL As Currency
Dim wl_UConta As String
'Dim i As Integer
Dim wl_Sinal As String * 1
Dim wl_Resgata As Boolean
Dim wl_FormulaAnterior As Currency
Dim wl_FormulaDebito As Currency
Dim wl_FormulaCredito As Currency
Dim wl_Extensao As String
Dim wl_Tradutor As Long
Dim wl_Folha As Integer
Dim wl_Calculo As Currency
Dim wl_Descricao As String
Dim wl_Fator As Currency
If Not Abre_PlanoContas(tbPlano) Or _
   Not Abre_MoviCaixa(tbMovi) Or _
   Not Abre_SaldoContas(tbSaldo) Or _
   Not Abre_RelatorioProgramado(tbRelPrograma, True) Then
   Exit Sub
End If
If tbRelPrograma.RecordCount > 0 Then
   tbRelPrograma.MoveFirst
   Do While Not tbRelPrograma.EOF
      If tbRelPrograma.EOF Or tbRelPrograma.NoMatch Then Exit Do
      If edit_reg(tbRelPrograma) Then tbRelPrograma.Delete
      tbRelPrograma.MoveNext
   Loop
End If
tbPlano.Index = "iCONTA"
If chkMES.Value = 1 Then
   wl_DataInicial = "01/" + Mid(cmbMES.Text, 1, 2) + "/" + cmbANO.Text
   wl_dia = 31
   Do While Not IsDate(Format(wl_dia, "00") + "/" + Mid(cmbMES.Text, 1, 2) + "/" + cmbANO.Text)
      wl_dia = wl_dia - 1
   Loop
   wl_Datafinal = Format(wl_dia, "00") + "/" + Mid(cmbMES.Text, 1, 2) + "/" + cmbANO.Text
   If chkSALDO.Value = 1 Then
      aadd wl_Cabecalho, Array("Conta", "Descricao", "Saldo")
      aadd wl_Referencia, Array("9.99.99.99.99999   99999", "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", "99,999,999.99")
   Else
      aadd wl_Cabecalho, Array("Conta", "Descricao", "Anterior", "Débito", "Crédito", "Saldo")
      aadd wl_Referencia, Array("9.99.99.99.99999   99999", "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", "99,999,999.99", "99,999,999.99", "99,999,999.99", "99,999,999.99")
   End If
   Recalcula_Saldo CDate(wl_DataInicial)
End If
Open PathPadrao + "RELATORIOS\" + txtnome.Text + ".REL" For Input As #99
Do While Not EOF(99)
   Input #99, wl_Linha
   wl_Conta = ""
   If Mid(wl_Linha, 1, 1) <> "(" Then
      For i = 1 To Len(wl_Linha)
         If Mid(wl_Linha, i, 1) <> "." Then Exit For
      Next
      wl_Conta = Trim(Mid(wl_Linha, i, InStr(wl_Linha, ":") - i))
   Else
      wl_Resgata = False
      wl_Conta = ""
      For i = 1 To Len(wl_Linha)
         If Mid(wl_Linha, i, 1) = ")" Then Exit For
         If InStr("+-", Mid(wl_Linha, i, 1)) <> 0 Then
            If wl_Resgata Then
               GoSub Busca_Conta
               wl_Conta = ""
               i = i + 1
            Else
               wl_Resgata = True
               i = i + 1
            End If
         End If
         If wl_Resgata Then wl_Conta = wl_Conta + Mid(wl_Linha, i, 1)
         If Not wl_Resgata And wl_Conta <> "" Then
            GoSub Busca_Conta
            wl_Conta = ""
         End If
      Next
   End If
   GoSub Busca_Conta
Loop
Close #99
GoSub Imprime
Finaliza_Impressao
Exit Sub


Busca_Conta:
tbPlano.Seek "=", wl_Conta
If tbPlano.NoMatch Then Return
Do While Not tbPlano.EOF
   If Mid(tbPlano("CONTA"), 1, Len(wl_Conta)) <> wl_Conta Then Exit Do
   If Mid(tbPlano("CONTA"), 1, Len(wl_Conta)) = wl_Conta Then
      If Mid(tbPlano("CONTA"), Len(wl_Conta) + 1, 1) <> "." And Mid(tbPlano("CONTA"), Len(wl_Conta) + 1, 1) <> "" Then Exit Do
   End If
   tbRelPrograma.Seek "=", tbPlano("CONTA")
   If tbRelPrograma.NoMatch Then
      wl_SaldoAnterior = 0
      wl_SaldoDebito = 0
      wl_SaldoCredito = 0
      If tbPlano("TRADUTOR") <> 0 Then
         wl_SaldoAnterior = (RetornaSaldoAnterior(tbPlano("TRADUTOR"), CDate(wl_DataInicial)) + tbPlano("SALDOABERTURA"))
         wl_SaldoDebito = (RetornaSaldoAtual(tbPlano("TRADUTOR"), CDate(wl_DataInicial), CDate(wl_Datafinal), "D"))
         wl_SaldoCredito = (RetornaSaldoAtual(tbPlano("TRADUTOR"), CDate(wl_DataInicial), CDate(wl_Datafinal), "C"))
      End If
      If add_reg(tbRelPrograma) Then
         tbRelPrograma("CONTA") = tbPlano("CONTA")
         tbRelPrograma("TRADUTOR") = tbPlano("TRADUTOR")
         tbRelPrograma("ANTERIOR") = wl_SaldoAnterior
         tbRelPrograma("DEBITO") = wl_SaldoDebito
         tbRelPrograma("CREDITO") = wl_SaldoCredito
         update_reg tbRelPrograma
         If wl_Conta <> tbPlano("CONTA") Then GoSub Nivel_Acima
      End If
   End If
   tbPlano.MoveNext
Loop
Return


Nivel_Acima:
wl_NivelAcima = tbPlano("CONTA")
Do While True
   For i = Len(wl_NivelAcima) To 1 Step -1
      If Mid(wl_NivelAcima, i, 1) = "." Then Exit For
   Next
   If i <> 0 Then
      wl_NivelAcima = Mid(wl_NivelAcima, 1, i - 1)
   End If
   tbRelPrograma.Seek "=", wl_NivelAcima
   If Not tbRelPrograma.NoMatch Then
      If edit_reg(tbRelPrograma) Then
         tbRelPrograma("ANTERIOR") = tbRelPrograma("ANTERIOR") + wl_SaldoAnterior
         tbRelPrograma("DEBITO") = tbRelPrograma("DEBITO") + wl_SaldoDebito
         tbRelPrograma("CREDITO") = tbRelPrograma("CREDITO") + wl_SaldoCredito
         update_reg tbRelPrograma
      End If
   End If
   If wl_NivelAcima = wl_Conta Or Len(wl_NivelAcima) = 1 Then Exit Do
Loop
Return

Imprime:
If tbRelPrograma.RecordCount = 0 Then Return
Open PathPadrao + "RELATORIOS\" + txtnome.Text + ".REL" For Input As #99
Do While Not EOF(99)
   wl_Linha = 0
   Input #99, wl_Linha
   wl_Conta = ""
   If Mid(wl_Linha, 1, 1) <> "(" Then
      For i = 1 To Len(wl_Linha)
         If Mid(wl_Linha, i, 1) <> "." Then Exit For
      Next
      wl_Conta = Trim(Mid(wl_Linha, i, InStr(wl_Linha, ":") - i))
   Else
      wl_Resgata = False
      wl_Conta = ""
      wl_Sinal = ""
      wl_FormulaAnterior = 0
      wl_FormulaDebito = 0
      wl_FormulaCredito = 0
      For i = 1 To Len(wl_Linha)
         If Mid(wl_Linha, i, 1) = ")" Then GoSub Calcula_Formula: Exit For
         If InStr("+-", Mid(wl_Linha, i, 1)) <> 0 Then
            If wl_Sinal <> " " Then
               GoSub Calcula_Formula
            Else
               wl_Resgata = True
            End If
            wl_Sinal = Mid(wl_Linha, i, 1)
            i = i + 1
            wl_Conta = ""
         End If
         If wl_Resgata Then wl_Conta = wl_Conta + Mid(wl_Linha, i, 1)
         If Not wl_Resgata And wl_Conta <> "" Then wl_Conta = ""
      Next
   End If
   If Mid(wl_Linha, 1, 1) <> "(" Then
      tbRelPrograma.Seek "=", wl_Conta
   Else
      If tbRelPrograma.RecordCount > 0 Then tbRelPrograma.MoveFirst
   End If
   Do While Not tbRelPrograma.EOF
      If Mid(wl_Linha, 1, 1) <> "(" Then
         If Mid(tbRelPrograma("CONTA"), 1, Len(wl_Conta)) <> wl_Conta Then Exit Do
         If Mid(tbRelPrograma("CONTA"), 1, Len(wl_Conta)) = wl_Conta Then
            If Mid(tbRelPrograma("CONTA"), Len(wl_Conta) + 1, 1) <> "" And Mid(tbRelPrograma("CONTA"), Len(wl_Conta) + 1, 1) <> "." Then Exit Do
         End If
         tbPlano.Seek "=", tbRelPrograma("CONTA")
      End If
      If wl_LL = 0 Then
         wl_Folha = wl_Folha + 1
         If chkMES.Value = 1 Then
            wl_Extensao = "Referente ao mês de " + Me.cmbMES.Text + " / " + Me.cmbANO.Text
         Else
            wl_Extensao = "Referente ao período de " + CStr(Me.txtdatai.Pacote) + " a " + CStr(Me.txtdataf.Pacote)
         End If
         Monta_Cabecalho wl_Cabecalho, wl_Referencia, IIf(chkSALDO.Value = 1, 2, 5), wl_LL, imp_Condensado, Trim(txtnome.Text) + "  -- " + wl_Extensao, wl_Folha
      End If
      If Mid(wl_Linha, 1, 1) <> "(" Then
         If tbPlano("TRADUTOR") > 0 Then
            If wl_Tradutor = 0 Then wl_LL = wl_LL + 0.5
            Monta_LinhadeImpressao wl_LL, tbPlano("CONTA") + IIf(tbPlano("TRADUTOR") = 0, "", "   -  " + Format(tbPlano("TRADUTOR"), "00000")), 0, , imp_Condensado
            Monta_LinhadeImpressao wl_LL, tbPlano("DESCRICAO"), 1, , imp_Condensado
         Else
            wl_LL = wl_LL + 0.5
            Monta_LinhadeImpressao wl_LL, tbPlano("CONTA") + IIf(tbPlano("TRADUTOR") = 0, "", "   -  " + Format(tbPlano("TRADUTOR"), "00000")), 0, , imp_Condensado_NEGRITO
            Monta_LinhadeImpressao wl_LL, tbPlano("DESCRICAO"), 1, , imp_Condensado_NEGRITO
         End If
         wl_SaldoAnterior = tbRelPrograma("ANTERIOR")
         wl_SaldoDebito = tbRelPrograma("DEBITO")
         wl_SaldoCredito = tbRelPrograma("CREDITO")
      Else
         wl_LL = wl_LL + 0.5
         wl_Descricao = Trim(Mid(wl_Linha, InStr(wl_Linha, "-->") + 3))
         If pb_InverteOperacao Then
            wl_Calculo = wl_FormulaAnterior + wl_FormulaDebito - wl_FormulaCredito
            If wl_Calculo < 0 Then
               wl_Descricao = UCase(StrTran(UCase(wl_Descricao), "#", "DÉFICIT"))
            Else
               wl_Descricao = UCase(StrTran(UCase(wl_Descricao), "#", "SUPERÁVIT"))
            End If
         Else
            wl_Calculo = wl_FormulaAnterior - wl_FormulaDebito + wl_FormulaCredito
            If wl_Calculo < 0 Then
               wl_Descricao = UCase(StrTran(UCase(wl_Descricao), "#", "DÉFICIT"))
            Else
               wl_Descricao = UCase(StrTran(UCase(wl_Descricao), "#", "SUPERÁVIT"))
            End If
         End If
         Monta_LinhadeImpressao wl_LL, wl_Descricao, 0, , imp_Condensado_NEGRITO
      End If
      If chkSALDO.Value = 1 Then
         If Mid(wl_Linha, 1, 1) <> "(" Then
            If pb_InverteOperacao Then
               wl_Calculo = wl_SaldoAnterior + wl_SaldoDebito - wl_SaldoCredito
               Monta_LinhadeImpressao wl_LL, Format(wl_Calculo, "##,###,##0.00;(##,###,##0.00)"), 2, "D", imp_Condensado
            Else
               wl_Calculo = wl_SaldoAnterior - wl_SaldoDebito + wl_SaldoCredito
               Monta_LinhadeImpressao wl_LL, Format(wl_Calculo, "##,###,##0.00;(##,###,##0.00)"), 2, "D", imp_Condensado
            End If
         Else
            If pb_InverteOperacao Then
               Monta_LinhadeImpressao wl_LL, Format(wl_FormulaAnterior + wl_FormulaDebito - wl_FormulaCredito, "##,###,##0.00;(##,###,##0.00)"), 2, "D", imp_Condensado_NEGRITO
            Else
               Monta_LinhadeImpressao wl_LL, Format(wl_FormulaAnterior - wl_FormulaDebito + wl_FormulaCredito, "##,###,##0.00;(##,###,##0.00)"), 2, "D", imp_Condensado_NEGRITO
            End If
         End If
      Else
         If Mid(wl_Linha, 1, 1) <> "(" Then
            Monta_LinhadeImpressao wl_LL, Format(wl_SaldoAnterior, "##,###,##0.00;(##,###,##0.00)"), 2, "D", imp_Condensado
            Monta_LinhadeImpressao wl_LL, Format(wl_SaldoDebito, "##,###,##0.00;(##,###,##0.00)"), 3, "D", imp_Condensado
            Monta_LinhadeImpressao wl_LL, Format(wl_SaldoCredito, "##,###,##0.00;(##,###,##0.00)"), 4, "D", imp_Condensado
            If pb_InverteOperacao Then
               Monta_LinhadeImpressao wl_LL, Format(wl_SaldoAnterior + wl_SaldoDebito - wl_SaldoCredito, "##,###,##0.00"), 5, "D", imp_Condensado
            Else
               Monta_LinhadeImpressao wl_LL, Format(wl_SaldoAnterior - wl_SaldoDebito + wl_SaldoCredito, "##,###,##0.00"), 5, "D", imp_Condensado
            End If
         Else
            Monta_LinhadeImpressao wl_LL, Format(wl_FormulaAnterior, "##,###,##0.00;(##,###,##0.00)"), 2, "D", imp_Condensado_NEGRITO
            Monta_LinhadeImpressao wl_LL, Format(wl_FormulaDebito, "##,###,##0.00;(##,###,##0.00)"), 3, "D", imp_Condensado_NEGRITO
            Monta_LinhadeImpressao wl_LL, Format(wl_FormulaCredito, "##,###,##0.00;(##,###,##0.00)"), 4, "D", imp_Condensado_NEGRITO
            If pb_InverteOperacao Then
               Monta_LinhadeImpressao wl_LL, Format(wl_FormulaAnterior + wl_FormulaDebito - wl_FormulaCredito, "##,###,##0.00;(##,###,##0.00)"), 5, "D", imp_Condensado_NEGRITO
            Else
               Monta_LinhadeImpressao wl_LL, Format(wl_FormulaAnterior - wl_FormulaDebito + wl_FormulaCredito, "##,###,##0.00;(##,###,##0.00)"), 5, "D", imp_Condensado_NEGRITO
            End If
         End If
      End If
      wl_LL = wl_LL + 0.5
      If wl_LL > IIf(pb_ImpressaoMatricial, 29, 26) Then
         Salta_Pagina
         wl_LL = 0
      End If
      If Mid(wl_Linha, 1, 1) = "(" Then Exit Do
      wl_Tradutor = tbPlano("TRADUTOR")
      tbRelPrograma.MoveNext
   Loop
Loop
Close #99
Return



Calcula_Formula:
wl_SaldoAnterior = 0
wl_SaldoDebito = 0
wl_SaldoCredito = 0
tbRelPrograma.Seek "=", wl_Conta
If tbRelPrograma.NoMatch Then Return
Do While Not tbRelPrograma.EOF
   If Mid(tbRelPrograma("CONTA"), 1, Len(wl_Conta)) <> wl_Conta Then Exit Do
   If Mid(tbRelPrograma("CONTA"), 1, Len(wl_Conta)) = wl_Conta Then
      If Mid(tbRelPrograma("CONTA"), Len(wl_Conta) + 1, 1) <> "" Then Exit Do
   End If
   wl_SaldoAnterior = wl_SaldoAnterior + tbRelPrograma("ANTERIOR")
   wl_SaldoDebito = wl_SaldoDebito + tbRelPrograma("DEBITO")
   wl_SaldoCredito = wl_SaldoCredito + tbRelPrograma("CREDITO")
   tbRelPrograma.MoveNext
Loop
If wl_Sinal = "+" Then
   wl_FormulaAnterior = wl_FormulaAnterior + wl_SaldoAnterior
   wl_FormulaCredito = wl_FormulaCredito + wl_SaldoCredito
   wl_FormulaDebito = wl_FormulaDebito + wl_SaldoDebito
Else
   wl_FormulaAnterior = wl_FormulaAnterior - wl_SaldoAnterior
   wl_FormulaCredito = wl_FormulaCredito - wl_SaldoCredito
   wl_FormulaDebito = wl_FormulaDebito - wl_SaldoDebito
End If
Return
End Sub



Private Sub Command1_Click()
Me.File1.Path = PathPadrao + "RELATORIOS"
File1.Visible = True
File1.SetFocus
SendKeys "{RIGHT}"
End Sub

Private Sub Command2_Click()
If Not PadraodeImpressao Then Exit Sub
Imprime_RelPrograma
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
Dim wl_File As String
If KeyAscii = 27 Then
   KeyAscii = 0
   txtnome.SetFocus
ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   If InStr(File1.FileName, ".") <> 0 Then
      wl_File = Mid(File1.FileName, 1, InStr(File1.FileName, ".") - 1)
   Else
      wl_File = File1.FileName
   End If
   txtnome.Text = wl_File
   txtnome.SetFocus
   HomeEnd
End If
End Sub


Private Sub File1_LostFocus()
File1.Visible = False
End Sub


Private Sub txtNOME_GotFocus()
Command2.Enabled = False
End Sub

Private Sub txtNOME_KeyPress(KeyAscii As Integer)
Dim wl_Line As String
Dim i As Integer
If KeyAscii = 13 Then
   KeyAscii = 0
   Command2.Enabled = True
   If InStr(txtnome.Text, ".") <> 0 Then
      InformaaoUsuario "Informe o nome sem ."
      txtnome.SetFocus
      Exit Sub
   End If
   If Dir(PathPadrao + "RELATORIOS\" + txtnome.Text + ".DAT") = "" Then
      chkMES.Value = 0
      chkSALDO.Value = 0
   Else
      Open PathPadrao + "RELATORIOS\" + txtnome.Text + ".DAT" For Input As #99
      Line Input #99, wl_Line
      Close #99
      If InStr(wl_Line, "M") <> 0 Then
         chkMES.Value = 1
         lblMES.Visible = True
         cmbMES.Visible = True
         cmbANO.Visible = True
         lblDATAI.Visible = False
         txtdatai.Visible = False
         lblDATAF.Visible = False
         txtdataf.Visible = False
         If Month(Date) = 1 Then
            cmbMES.Text = "01 - Janeiro"
         ElseIf Month(Date) = 2 Then
            cmbMES.Text = "02 - Fevereiro"
         ElseIf Month(Date) = 3 Then
            cmbMES.Text = "03 - Março"
         ElseIf Month(Date) = 4 Then
            cmbMES.Text = "04 - Abril"
         ElseIf Month(Date) = 5 Then
            cmbMES.Text = "05 - Maio"
         ElseIf Month(Date) = 6 Then
            cmbMES.Text = "06 - Junho"
         ElseIf Month(Date) = 7 Then
            cmbMES.Text = "07 - Julho"
         ElseIf Month(Date) = 8 Then
            cmbMES.Text = "08 - Agosto"
         ElseIf Month(Date) = 9 Then
            cmbMES.Text = "09 - Setembro"
         ElseIf Month(Date) = 10 Then
            cmbMES.Text = "10 - Outubro"
         ElseIf Month(Date) = 11 Then
            cmbMES.Text = "11 - Novembro"
         ElseIf Month(Date) = 12 Then
            cmbMES.Text = "12 - Dezembro"
         End If
         cmbANO.Clear
         For i = Year(Date) - 10 To Year(Date) + 10
            cmbANO.AddItem Format(i, "0000")
         Next
         cmbANO.Text = Format(Year(Date), "0000")
         cmbMES.SetFocus
      Else
         chkMES.Value = 0
         lblMES.Visible = False
         cmbMES.Visible = False
         cmbANO.Visible = False
         lblDATAI.Visible = True
         txtdatai.Visible = True
         lblDATAF.Visible = True
         txtdataf.Visible = True
         txtdatai.SetFocus
      End If
      If InStr(wl_Line, "S") <> 0 Then chkSALDO.Value = 1 Else chkSALDO.Value = 0
   End If
End If
End Sub


