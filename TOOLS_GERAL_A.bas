Attribute VB_Name = "TOOLS_GERAL_A"
Public pb_LinhaImpressaoMatricial As Integer
Public pb_ColunaImpressaoMatricial As Integer
Public pb_MontaProgressao As Boolean

Sub PO()
Dim LIXO
    
    LIXO = MsgBox("Preenchimento Obrigatório Desta Informação !", vbOKOnly + vbCritical, "Aviso")
    
End Sub


Function DIFFHORAS(X_HORA_I As Date, X_HORA_F As Date) As Date
   Dim X_HORAS, X_VAR1, X_VAR2 As String
   If X_HORA_I >= X_HORA_F Then
      DIFFHORAS = "00:00:00"
      Exit Function
   End If
   X_VAR1 = Mid(X_HORA_I, 1, 2) + ":00:00"
   X_VAR2 = Mid(X_HORA_F, 1, 2) + ":00:00"
   X_HORAS = Format(DateDiff("h", X_VAR1, X_VAR2), "00") + ":"
   X_VAR1 = "00:" + Mid(X_HORA_I, 4, 2) + ":00"
   X_VAR2 = "00:" + Mid(X_HORA_F, 4, 2) + ":00"
   If X_VAR1 > X_VAR2 Then
      X_VAR1 = "00:" + Mid(X_HORA_I, 4, 2) + ":00"
      X_VAR2 = "01:" + Mid(X_HORA_F, 4, 2) + ":00"
      X_HORAS = Format(Val(Mid(X_HORAS, 1, 2)) - 1, "00") + ":"
   End If
   X_HORAS = X_HORAS + Format(DateDiff("n", X_VAR1, X_VAR2), "00") + ":"
   X_VAR1 = "00:00:" + Mid(X_HORA_I, 7, 2)
   X_VAR2 = "00:00:" + Mid(X_HORA_F, 7, 2)
   If X_VAR1 > X_VAR2 Then
      X_VAR1 = "00:00:" + Mid(X_HORA_I, 7, 2)
      X_VAR2 = "00:01:" + Mid(X_HORA_F, 7, 2)
      X_HORAS = Mid(X_HORAS, 1, 3) + Format(Val(Mid(X_HORAS, 4, 2)) - 1, "00") + ":"
   End If
   DIFFHORAS = X_HORAS + Format(DateDiff("s", X_VAR1, X_VAR2), "00")
End Function

Function Conta_Char(pString, pchar As String)
Dim i As Integer
Dim wl_Conta As Integer
For i = 1 To Len(pString)
   If Mid(pString, i, Len(pchar)) = pchar Then wl_Conta = wl_Conta + 1
Next
Conta_Char = wl_Conta
End Function


Function dbSeek(ByRef pTabela As Recordset, pOqueBuscar)
pTabela.Seek "=", pOqueBuscar
dbSeek = Not pTabela.NoMatch
End Function

Sub Demarca(pForm As Form)
Dim Controle As Control
For Each Controle In pForm
   If TypeOf Controle Is TextBox Or TypeOf Controle Is Caixa_Texto Then
      Controle.BorderStyle = "0"
      pForm.Line (Controle.Left - 15, Controle.Top)-(Controle.Left - 15, Controle.Top + Controle.Height)
      pForm.Line (Controle.Left - 15, Controle.Top + Controle.Height)-(Controle.Left - 15 + Controle.Width + 15, Controle.Top + Controle.Height)
   End If
Next
End Sub

Sub EmiteBoleto(pBanco As Long, pVencimento As Date, pGeracao As Date, pDocumento As String, pValor As Currency, pSacado As String, pEndereço As String, pCPF As String)
Dim wl_Valor As Currency
Dim wl_Retorno As String
Dim wl_Linha As Currency
Dim wl_Coluna As Currency
Dim wl_File As String
Dim wl_LinhaInicial As Currency
Dim wl_Diferenca As Currency
Static wl_Contaboleto As Integer
If pb_PadraoVideo Then
   Exit Sub
End If
Printer.ScaleMode = 7
wl_File = PathPadrao + "BOLETOS\" + Format(pBanco, "000") + ".INI"
If Dir(wl_File) = "" Then
   Exit Sub
End If
wl_LinhaInicial = Printer.CurrentY
If Printer.CurrentY = 0 Then
   wl_Contaboleto = 0
End If
wl_Contaboleto = wl_Contaboleto + 1
wl_Linha = RetornaConfiguracao("LOCAL DE PAGAMENTO", "Linha", wl_File)
wl_Diferenca = RetornaConfiguracao("VENCIMENTO", "Linha", wl_File)
wl_Coluna = RetornaConfiguracao("LOCAL DE PAGAMENTO", "Coluna", wl_File)
wl_Retorno = RetornaConfiguracao("LOCAL DE PAGAMENTO", "Texto", wl_File)
Imprime wl_LinhaInicial + wl_Linha, wl_Coluna, wl_Retorno, Imp_Normal
GoSub Correcao

wl_Linha = RetornaConfiguracao("VENCIMENTO", "Linha", wl_File)
wl_Diferenca = RetornaConfiguracao("DATA DO DOCUMENTO", "Linha", wl_File)
wl_Coluna = RetornaConfiguracao("VENCIMENTO", "Coluna", wl_File)
wl_Retorno = CStr(pVencimento)
Imprime wl_LinhaInicial + wl_Linha, wl_Coluna, CStr(pVencimento)
GoSub Correcao

wl_Linha = RetornaConfiguracao("DATA DO DOCUMENTO", "Linha", wl_File)
wl_Diferenca = RetornaConfiguracao("NRO. DO DOCUMENTO", "Linha", wl_File)
wl_Coluna = RetornaConfiguracao("DATA DO DOCUMENTO", "Coluna", wl_File)
wl_Retorno = CStr(pGeracao)
Imprime wl_LinhaInicial + wl_Linha, wl_Coluna, CStr(pGeracao)
GoSub Correcao

wl_Linha = RetornaConfiguracao("NRO. DO DOCUMENTO", "Linha", wl_File)
wl_Diferenca = RetornaConfiguracao("DATA DO PROCESSAMENTO", "Linha", wl_File)
wl_Coluna = RetornaConfiguracao("NRO. DO DOCUMENTO", "Coluna", wl_File)
wl_Retorno = pDocumento
Imprime wl_LinhaInicial + wl_Linha, wl_Coluna, pDocumento
GoSub Correcao

wl_Linha = RetornaConfiguracao("DATA DO PROCESSAMENTO", "Linha", wl_File)
wl_Diferenca = RetornaConfiguracao("VALOR DO DOCUMENTO", "Linha", wl_File)
wl_Coluna = RetornaConfiguracao("DATA DO PROCESSAMENTO", "Coluna", wl_File)
wl_Retorno = CStr(Date)
Imprime wl_LinhaInicial + wl_Linha, wl_Coluna, CStr(Date)
GoSub Correcao

wl_Linha = RetornaConfiguracao("VALOR DO DOCUMENTO", "Linha", wl_File)
wl_Diferenca = RetornaConfiguracao("INSTRUCAO", "Linhas", wl_File)
wl_Coluna = RetornaConfiguracao("VALOR DO DOCUMENTO", "Coluna", wl_File)
wl_Retorno = Format(pValor, "##,###,##0.00")
Imprime wl_LinhaInicial + wl_Linha, wl_Coluna, Format(pValor, "##,###,##0.00"), imp_Normal_Negrito
GoSub Correcao

wl_Linha = RetornaConfiguracao("INSTRUCAO", "Linhas", wl_File)
wl_Diferenca = wl_Linha + 0.5
wl_Retorno = RetornaConfiguracao("INSTRUCAO", "I01", wl_File)
Imprime wl_LinhaInicial + wl_Linha, 0, wl_Retorno
GoSub Correcao

wl_Linha = wl_Linha + 0.5
wl_Diferenca = wl_Linha + 0.5
wl_Retorno = RetornaConfiguracao("INSTRUCAO", "I02", wl_File)
Imprime wl_LinhaInicial + wl_Linha, 0, wl_Retorno
GoSub Correcao

wl_Linha = wl_Linha + 0.5
wl_Diferenca = wl_Linha + 0.5
wl_Retorno = RetornaConfiguracao("INSTRUCAO", "I03", wl_File)
Imprime wl_LinhaInicial + wl_Linha, 0, wl_Retorno
GoSub Correcao

wl_Linha = wl_Linha + 0.5
wl_Diferenca = RetornaConfiguracao("SACADO", "Linha", wl_File)
wl_Retorno = RetornaConfiguracao("INSTRUCAO", "I04", wl_File)
Imprime wl_LinhaInicial + wl_Linha, 0, wl_Retorno
GoSub Correcao

wl_Linha = RetornaConfiguracao("SACADO", "Linha", wl_File)
wl_Diferenca = RetornaConfiguracao("CPF OU CGC", "Linha", wl_File)
wl_Coluna = RetornaConfiguracao("SACADO", "Coluna", wl_File)
wl_Retorno = CStr(pSacado)
Imprime wl_LinhaInicial + wl_Linha, wl_Coluna, CStr(pSacado)
GoSub Correcao

wl_Linha = RetornaConfiguracao("CPF OU CGC", "Linha", wl_File)
wl_Diferenca = RetornaConfiguracao("ESPACAMENTO", "Linhas", wl_File)
wl_Coluna = RetornaConfiguracao("CPF OU CGC", "Coluna", wl_File)
wl_Retorno = CStr(pCPF)
Imprime wl_LinhaInicial + wl_Linha, wl_Coluna, CStr(pCPF)
GoSub Correcao

wl_Linha = RetornaConfiguracao("ESPACAMENTO", "Linhas", wl_File)
wl_Diferenca = 0
wl_Retorno = ""
Imprime wl_LinhaInicial + wl_Linha, 0, ""
GoSub Correcao
Exit Sub


Correcao:
wl_Contaboleto = wl_Contaboleto
If Printer.CurrentY < wl_LinhaInicial Then
   If wl_LinhaInicial + wl_Linha > Printer.ScaleHeight Then
      wl_LinhaInicial = Printer.CurrentY - wl_Linha - 0.38
   End If
End If
Return
End Sub

Function EstourodePagina(pLinha As Currency) As Boolean
If pLinha > IIf(pb_ImpressaoMatricial, 29, 26) Then EstourodePagina = True
End Function

Sub HomeEnd()
SendKeys "{Home}+{End}"
End Sub


Function Mes_Extenso(pMes As Integer) As String

    If pMes = 1 Then
       Mes_Extenso = "JANEIRO  "
ElseIf pMes = 2 Then
       Mes_Extenso = "FEVEREIRO"
ElseIf pMes = 3 Then
       Mes_Extenso = "MARÇO    "
ElseIf pMes = 4 Then
       Mes_Extenso = "ABRIL    "
ElseIf pMes = 5 Then
       Mes_Extenso = "MAIO     "
ElseIf pMes = 6 Then
       Mes_Extenso = "JUNHO    "
ElseIf pMes = 7 Then
       Mes_Extenso = "JULHO    "
ElseIf pMes = 8 Then
       Mes_Extenso = "AGOSTO   "
ElseIf pMes = 9 Then
       Mes_Extenso = "SETEMBRO "
ElseIf pMes = 10 Then
       Mes_Extenso = "OUTUBRO  "
ElseIf pMes = 11 Then
       Mes_Extenso = "NOVEMBRO "
ElseIf pMes = 12 Then
       Mes_Extenso = "DEZEMBRO "
End If


End Function

Function PathWindows() As String
Dim Temp As String
Dim Ret As Long
Const MAX_LENGTH = 145

Temp = String(MAX_LENGTH, 0)
Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
Temp = Left$(Temp, Ret)
If Temp <> "" And Right$(Temp, 1) <> "\" Then
   PathWindows = Temp & "\"
Else
   PathWindows = Temp
End If
End Function


Sub centraobj(ByVal pForm As Form)
Dim wl_HeightBar As Currency
Dim wl_BarrastatusTop As Currency
Dim wl_AreaLivre As Currency
On Error Resume Next
pForm.Top = 0
pForm.Left = 0
End Sub


Public Sub InformaaoUsuario(pMensagem As String, Optional pIcone = vbExclamation, Optional pTitulo As String = "Mensagem do Sistema")
On Error Resume Next
MsgBox pMensagem, pIcone, pTitulo
VB.Screen.ActiveForm.ActiveControl.SetFocus
End Sub

Public Sub LimpaCaixasTexto(Janela As Form, Optional pCheck As Boolean = True)
Dim Controle As Control
Dim wl_Vazio As String
Dim wl_Rows As Long
Dim wl_Cols As Long
On Error Resume Next
wl_Vazio = "X"
If IsNumeric(pbRetornoVideo) Then
   If pbRetornoVideo = 0 Then wl_Vazio = ""
Else
   If pbRetornoVideo = "" Then wl_Vazio = ""
End If
For Each Controle In Janela.Controls
   If Controle.Name <> Janela.ActiveControl.Name Or Janela.Name <> pb_FormAtivo Or wl_Vazio = "" Or _
   Janela.ActiveControl.Name <> pb_ObjetoAtivo Then
      If TypeOf Controle Is Máscara Then
         If Not Controle.ÉData Then
            Controle.Text = ""
         Else
            Controle.Text = ""
         End If
      ElseIf TypeOf Controle Is Caixa_Texto Then
         Controle.Text = ""
      ElseIf TypeOf Controle Is Etiqueta Then
         Controle.Caption = ""
      ElseIf TypeOf Controle Is CheckBox And pCheck Then
           Controle.Value = 0
      ElseIf TypeOf Controle Is TextBox Then
         Controle.Text = ""
      ElseIf TypeOf Controle Is MaskEdBox Then
         If Controle.PromptInclude Then
            Controle.PromptInclude = False
            Controle.Text = ""
            Controle.PromptInclude = True
         Else
            Controle.Text = ""
         End If
      ElseIf TypeOf Controle Is MSFlexGrid Then
         For wl_Rows = Controle.FixedRows To Controle.rows - 1
            For wl_Cols = Controle.FixedCols To Controle.cols - 1
               Controle.TextMatrix(wl_Rows, wl_Cols) = ""
            Next
         Next
      End If
   Else
      Controle.Text = pbRetornoVideo
      pbRetornoVideo = ""
   End If
Next Controle
Set Controle = Nothing
End Sub





Function PadraodeImpressao(Optional pNaoVideo As Boolean = False, Optional pNaoMatricial As Boolean = False, Optional pNaoJato As Boolean = False) As Boolean
pb_NaoVideo = pNaoVideo
pb_NaoMatricial = pNaoMatricial
pb_NaoJato = pNaoJato
fPadraoI.Show 1
PadraodeImpressao = Not pb_CancelaImpressao
End Function




Sub PreparaProximoDocumento(Optional pSaltoLinhas As Currency = 0)
Imprime (pb_LinhaImpressaoMatricial / 2) + 0.5, 0, " ", pb_Tamanho
pb_Buffer = ""
pb_LinhaBuffer = 0
pb_LinhaImpressaoMatricial = 0
pb_UltimaLinha = 0
Imprime pSaltoLinhas, 0, "", imp_Condensado
End Sub

Function Proximo_Mes(pData As Date, Optional pDia)
Dim wl_dia As Integer
Dim wl_Mes As Integer
Dim wl_Ano As Integer
Dim wl_Proximo As String
If IsMissing(pDia) Then
   pDia = Day(pData)
End If
wl_dia = pDia
wl_Mes = IIf(Month(pData) < 12, Month(pData) + 1, 1)
wl_Ano = IIf(Month(pData) < 12, Year(pData), Year(pData) + 1)

wl_Proximo = Format(wl_dia, "00") + "/" + Format(wl_Mes, "00") + "/" + Format(wl_Ano, "0000")

Do While Not IsDate(wl_Proximo)
   wl_dia = wl_dia - 1
   wl_Proximo = Format(wl_dia, "00") + "/" + Format(wl_Mes, "00") + "/" + Format(wl_Ano, "0000")
Loop
Proximo_Mes = wl_Proximo
End Function


Sub ReparaBancodeDados(pBanco)
If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\" + pBanco) <> "" Then
   DisplayMensagem "Aguarde, reparando " + pBanco + " ..."
   DBEngine.RepairDatabase PathPadrao + Format(pb_Empresa, "00000") + "\" + pBanco
   DisplayMensagem "Compactando " + pBanco + " ..."
   DBEngine.CompactDatabase PathPadrao + Format(pb_Empresa, "00000") + "\" + pBanco, PathPadrao + Format(pb_Empresa, "00000") + "\C_" + Mid(pBanco, 2)
   If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\X_" + Mid(pBanco, 2)) <> "" Then
      Kill PathPadrao + Format(pb_Empresa, "00000") + "\X_" + Mid(pBanco, 2)
   End If
   Name PathPadrao + Format(pb_Empresa, "00000") + "\" + pBanco As PathPadrao + Format(pb_Empresa, "00000") + "\X_" + Mid(pBanco, 2)
   Name PathPadrao + Format(pb_Empresa, "00000") + "\C_" + Mid(pBanco, 2) As PathPadrao + Format(pb_Empresa, "00000") + "\" + pBanco
   aviso
End If
End Sub

Function Replicate(pString As String, pvezes As Long) As String
Dim a As Long
Dim xstring As String

For a = 1 To pvezes Step 1
    xstring = xstring + pString
Next a

Replicate = xstring

End Function

Function Round(pValor, pDecimais As Integer) As Currency
Dim wl_Decimal As String
Dim wl_Retorno
Dim wl_Conversao As Currency
wl_Decimal = String(pDecimais, "0")
wl_Retorno = Format(pValor, "##############0." + wl_Decimal)
wl_Conversao = wl_Retorno
Round = wl_Conversao
End Function


Sub SetPrc(pLinha, pColuna)
pb_LinhaImpressaoMatricial = pLinha
pb_ColunaImpressaoMatricial = pColuna
End Sub



