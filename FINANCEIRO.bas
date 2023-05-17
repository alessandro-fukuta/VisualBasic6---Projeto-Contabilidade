Attribute VB_Name = "FINANCEIRO"
Option Explicit
Public pb_DeletaMovimento As Boolean 'Verifica se o movimento de conta foi deletado
Public pb_InverteOperacao As Boolean
Public pb_NivelPlano As Integer
Public pb_Nivel1 As Long
Public pb_Nivel2 As Long
Public pb_nivel3 As Long
Public pb_nivel4 As Long
Public pb_Conta As Long
Public dbFaturamento As Database
Public pb_RegimeCompetencia As Boolean


Function Most_PlanodeContas() As Long
fMOSTPLANO.Show 1
Most_PlanodeContas = IIf(pbRetornoVideo = "", 0, Val(pbRetornoVideo))
End Function


Function CalculaDataBase(ByRef txtDATA As Date, pOperacao As String)
Dim wl_Diferenca As Long
Dim wl_FimMes As Date
If Month(txtDATA) = 12 Then
   wl_FimMes = CDate("31/12/" + Format(Year(txtDATA), "0000"))
Else
   wl_FimMes = CDate("01/" + Format(Month(txtDATA) + 1, "00") + "/" + Format(Year(txtDATA), "0000")) - 1
End If
If Mid(pOperacao, 1, 3) = "DFS" Then
   wl_Diferenca = 7 - Weekday(txtDATA) + 1
ElseIf Mid(pOperacao, 1, 3) = "DFD" Then
   If Day(txtDATA) <= 10 Then
      wl_Diferenca = 10 - Day(txtDATA)
   ElseIf Day(txtDATA) <= 20 Then
      wl_Diferenca = 20 - Day(txtDATA)
   ElseIf Day(txtDATA) <= Day(wl_FimMes) Then
      wl_Diferenca = Day(wl_FimMes) - Day(txtDATA)
   End If
ElseIf Mid(pOperacao, 1, 3) = "DFQ" Then
   If Day(txtDATA) <= 15 Then
      wl_Diferenca = 15 - Day(txtDATA)
   Else
      wl_Diferenca = Day(wl_FimMes) - Day(txtDATA)
   End If
ElseIf Mid(pOperacao, 1, 3) = "DFM" Then
   wl_Diferenca = Day(wl_FimMes) - Day(txtDATA)
Else
   wl_Diferenca = 0
End If
CalculaDataBase = txtDATA + wl_Diferenca
End Function

Function Abre_MoviCaixa(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)

If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Movimento_de_Caixa", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_MoviCaixa = True
   pObjeto_Recordset.Index = "iDATA"
Else
   Abre_MoviCaixa = False
End If

End Function


Public Function LancaMoviCaixa(pConta As Long, pValor As Currency, pTipo As String, pHistorico As String, Optional pDate, Optional pContraPartida As Long = 0, Optional ByRef pLancamento As Long = 0)
Dim wl_Movimento As Long
Dim tbMoviCaixa As Recordset
If IsMissing(pDate) Then
   pDate = Date
End If
If Not Abre_MoviCaixa(tbMoviCaixa) Then
   Exit Function
End If
If tbMoviCaixa.RecordCount = 0 Then
   wl_Movimento = 1
Else
   tbMoviCaixa.Seek ">=", pDate + 1, 1
   If tbMoviCaixa.NoMatch Then
      tbMoviCaixa.MoveLast
   Else
      tbMoviCaixa.MovePrevious
      If tbMoviCaixa.BOF Then
         tbMoviCaixa.MoveFirst
      End If
   End If
   If tbMoviCaixa("DATA") = pDate Then
      wl_Movimento = tbMoviCaixa("MOVIMENTO") + 1
   Else
      wl_Movimento = 1
   End If
End If
If Not add_reg(tbMoviCaixa) Then
   MsgBox "Favor fazer lançamento manual no movimento de caixa"
   Exit Function
End If
tbMoviCaixa("DATA") = pDate
tbMoviCaixa("MOVIMENTO") = wl_Movimento
If pTipo = "C" Then
   tbMoviCaixa("CREDITO") = pConta
   tbMoviCaixa("DEBITO") = pContraPartida
Else
   tbMoviCaixa("DEBITO") = pConta
   tbMoviCaixa("CREDITO") = pContraPartida
End If
tbMoviCaixa("VALOR") = pValor
tbMoviCaixa("HISTORICO") = pHistorico
tbMoviCaixa("IMPORTADO") = True
If Not update_reg(tbMoviCaixa) Then
   MsgBox "Favor fazer lançamento manual no movimento de caixa"
   Exit Function
End If
pLancamento = wl_Movimento
If pTipo = "D" Then
   Atualiza_SaldoContas pConta, pValor
   If pContraPartida > 0 Then
      Atualiza_SaldoContas pContraPartida, , pValor
   End If
Else
   Atualiza_SaldoContas pConta, , pValor
   If pContraPartida > 0 Then
      Atualiza_SaldoContas pContraPartida, pValor
   End If
End If
tbMoviCaixa.Close
End Function


Public Function Atualiza_SaldoContas(pConta As Long, Optional pDebito As Currency = 0, Optional pCredito As Currency = 0) As Boolean
Dim tbSaldo As Recordset
Dim wl_SaldoAnterior As Currency
Dim wl_SaldoDebitoAnterior  As Currency
Dim wl_SaldoCreditoAnterior As Currency
If Not Abre_SaldoContas(tbSaldo) Then
   Exit Function
End If
If tbSaldo.RecordCount > 0 Then
   tbSaldo.Seek "=", pConta, Date
   If tbSaldo.NoMatch Then
      GoSub fSaldoAnterior
      If Not add_reg(tbSaldo) Then
         Exit Function
      End If
      tbSaldo("CONTA") = pConta
      tbSaldo("DATA") = Date
      If Not pb_InverteOperacao Then
         tbSaldo("ANTERIOR") = wl_SaldoAnterior + wl_SaldoCreditoAnterior - wl_SaldoDebitoAnterior
      Else
         tbSaldo("ANTERIOR") = wl_SaldoAnterior - wl_SaldoCreditoAnterior + wl_SaldoDebitoAnterior
      End If
      tbSaldo("DEBITO") = pDebito
      tbSaldo("CREDITO") = pCredito
      If Not update_reg(tbSaldo) Then
         Exit Function
      End If
      Atualiza_SaldoContas = True
   Else
      If Not edit_reg(tbSaldo) Then
         Exit Function
      End If
      tbSaldo("DEBITO") = tbSaldo("DEBITO") + pDebito
      tbSaldo("CREDITO") = tbSaldo("CREDITO") + pCredito
      If Not update_reg(tbSaldo) Then
         Exit Function
      End If
      Atualiza_SaldoContas = True
   End If
Else
   If Not add_reg(tbSaldo) Then
      Exit Function
   End If
   tbSaldo("CONTA") = pConta
   tbSaldo("DATA") = Date
   tbSaldo("ANTERIOR") = 0
   tbSaldo("DEBITO") = pDebito
   tbSaldo("CREDITO") = pCredito
   If Not update_reg(tbSaldo) Then
      Exit Function
   End If
   Atualiza_SaldoContas = True
End If
Exit Function


fSaldoAnterior:
tbSaldo.Seek "<", pConta, Date
If Not tbSaldo.NoMatch Then
   If tbSaldo("CONTA") = pConta Then
      wl_SaldoAnterior = tbSaldo("ANTERIOR")
      wl_SaldoDebitoAnterior = tbSaldo("DEBITO")
      wl_SaldoCreditoAnterior = tbSaldo("CREDITO")
   End If
End If
Return
End Function

Function Estrutura_Fluxo()
Dim wl_PDV As Integer
On Error Resume Next
wl_PDV = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "PDV")
ReDim campo(0)
ReDim indices(0)
If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\FLUXO" + Format(wl_PDV, "000") + ".MDB") <> "" Then
   Kill PathPadrao + Format(pb_Empresa, "00000") + "\FLUXO" + Format(wl_PDV, "000") + ".MDB"
End If
wl_PDV = RetornaConfiguracao("PREFERENCIAS", "PDV")

AddField campo, "DATA", dbDate, , True
AddField campo, "CODIGO", dbLong, , True
AddField campo, "NOME", dbText, 80
AddField campo, "MOVIMENTO", dbLong, , True
AddField campo, "aRECEBER", dbCurrency
AddField campo, "aPAGAR", dbCurrency

AddIndex indices, "iDATA", Array("DATA", "MOVIMENTO")

Estrutura_Fluxo = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FLUXO" + Format(wl_PDV, "000") + ".MDB", "Fluxo", campo, indices)

End Function

Function Loca_Plano(ByRef ptbPlano As Recordset, pNivel1 As Integer, Optional pNivel2 As Integer = 0, Optional pNivel3 As Integer = 0, Optional pNivel4 As Integer = 0, Optional pConta As Long = 0) As Boolean
Dim wl_Busca As String
Dim wl_Index As String
wl_Index = ptbPlano.Index
If ptbPlano.RecordCount = 0 Then
   Loca_Plano = False
Else
   wl_Busca = Format(pNivel1, "0")
   wl_Busca = wl_Busca + IIf(pNivel2 > 0, "." + Trim(Format(pNivel2, "##")), "")
   wl_Busca = wl_Busca + IIf(pNivel3 > 0, "." + Trim(Format(pNivel3, "##")), "")
   wl_Busca = wl_Busca + IIf(pNivel4 > 0, "." + Format(pNivel4, "00"), "")
   wl_Busca = wl_Busca + IIf(pConta > 0, "." + Format(pConta, "00000"), "")
   ptbPlano.Seek "=", wl_Busca
   If ptbPlano.NoMatch Then
      Loca_Plano = False
   Else
      Loca_Plano = True
   End If
End If
End Function

Function Loca_Contas(ByRef tbPlano As Recordset, pConta As Long) As Boolean
tbPlano.Seek "=", pConta
Loca_Contas = Not tbPlano.NoMatch
End Function


Function Loca_Bancos(ByRef tbBancos As Recordset, txtCODIGO As Long) As Boolean
If tbBancos.RecordCount = 0 Then
   Loca_Bancos = False
   Exit Function
End If
tbBancos.Seek "=", txtCODIGO
Loca_Bancos = Not tbBancos.NoMatch
End Function


Function Loca_Eventos(ByRef tbEventos As Recordset, txtgrupo As Long, txtCODIGO As Long) As Boolean
If tbEventos.RecordCount = 0 Then
   Loca_Eventos = False
   Exit Function
End If
tbEventos.Seek "=", txtgrupo, txtCODIGO
Loca_Eventos = Not tbEventos.NoMatch
End Function


Function Loca_LancaEventos(ByRef tblanca As Recordset, pANO As Integer, pMes As Integer, pEhSocio As Boolean, pCodigo As Long) As Boolean
If tblanca.RecordCount = 0 Then
   Loca_LancaEventos = False
   Exit Function
End If
tblanca.Seek "=", pANO, pMes, pEhSocio, pCodigo
Loca_LancaEventos = Not tblanca.NoMatch
End Function



Function Loca_LanPadrao(ByRef tbLanPadrao As Recordset, txtCODIGO As Long) As Boolean
If tbLanPadrao.RecordCount = 0 Then
   Loca_LanPadrao = False
   Exit Function
End If
tbLanPadrao.Seek "=", txtCODIGO
Loca_LanPadrao = Not tbLanPadrao.NoMatch
End Function

Function Loca_TipoCredor(ByVal tbTipoCredor As Recordset, txtCODIGO As String) As Boolean
If tbTipoCredor.RecordCount = 0 Then
   Loca_TipoCredor = False
   Exit Function
End If
tbTipoCredor.Seek "=", txtCODIGO
Loca_TipoCredor = Not tbTipoCredor.NoMatch
End Function

Function Loca_Historico(ByVal tbHistorico As Recordset, txtCODIGO As String) As Boolean
If tbHistorico.RecordCount = 0 Then
   Loca_Historico = False
   Exit Function
End If
tbHistorico.Seek "=", txtCODIGO
Loca_Historico = Not tbHistorico.NoMatch
End Function



Function Loca_Credores(ByVal tbCredores As Recordset, txtCODIGO As String) As Boolean
If tbCredores.RecordCount = 0 Then
   Loca_Credores = False
   Exit Function
End If
tbCredores.Seek "=", txtCODIGO
Loca_Credores = Not tbCredores.NoMatch
End Function



Function Abre_Contas(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)

If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Contas_de_Resultado", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_Contas = True
   pObjeto_Recordset.Index = "iCODIGO"
Else
   Abre_Contas = False
End If

End Function

Function Abre_RateioContas(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)
If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "RateiodeContas", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_RateioContas = True
   pObjeto_Recordset.Index = "iCONTA"
Else
   Abre_RateioContas = False
End If
End Function



Function Abre_OrdemPlano(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)

If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Ordem_Plano", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_OrdemPlano = True
Else
   Abre_OrdemPlano = False
End If

End Function


Function Abre_SaldoContas(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)

If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Saldo_Contas", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_SaldoContas = True
   pObjeto_Recordset.Index = "iCONTA"
Else
   Abre_SaldoContas = False
End If

End Function


Function Loca_SaldoPlano(ByRef ptbSaldoPlano As Recordset, pTradutor As Long, pExercicio As Long) As Boolean
If ptbSaldoPlano.RecordCount > 0 Then
   ptbSaldoPlano.Seek "=", pTradutor, pExercicio
   If ptbSaldoPlano.NoMatch Then
      Loca_SaldoPlano = False
   Else
      Loca_SaldoPlano = True
   End If
Else
   Loca_SaldoPlano = False
End If
End Function




Function Most_LanPadrao()
ReDim aCampo(0)
aadd aCampo, Array("Código", "CODIGO", 2000, "000")
aadd aCampo, Array("Descrição", "DESCRICAO", 6000, "")

Most_LanPadrao = Video("Lançamentos Padrão", aCampo, Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Lancamentos_Padrao", "iDESCRICAO", "CODIGO")
End Function

Function Most_OrdemPlano()
ReDim aCampo(0)
aadd aCampo, Array("Código", "CODIGO", 2000, "000")
aadd aCampo, Array("Descrição", "DESCRICAO", 6000, "")

Most_OrdemPlano = Video("Ordem Plano Contas", aCampo, Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Ordem_Plano", , "CODIGO")
End Function



Function Estrutura_Contas() As Boolean

ReDim campo(0)
ReDim indices(0)

Call AddField(campo, "CODIGO", dbLong, 0, True)
Call AddField(campo, "DESCRICAO", dbText, 50)
Call AddField(campo, "SALDO ANTERIOR", dbCurrency)
Call AddField(campo, "SALDO ATUAL", dbCurrency)

Call AddIndex(indices, "iCODIGO", Array("CODIGO"), True, True, True)
Call AddIndex(indices, "iDESCRICAO", Array("DESCRICAO"))

Estrutura_Contas = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Contas_de_Resultado", campo, indices)

End Function

Function Estrutura_RateioContas() As Boolean

ReDim campo(0)
ReDim indices(0)

AddField campo, "CONTAPRINCIPAL", dbLong, 0, True
AddField campo, "CONTARATEIO", dbLong
AddField campo, "PROPORCAO", dbCurrency

AddIndex indices, "iCONTA", "CONTAPRINCIPAL"
Estrutura_RateioContas = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "RateiodeContas", campo, indices)
End Function


Function Estrutura_BalanceteSintetico() As Boolean
ReDim campo(0)
ReDim indices(0)

Call AddField(campo, "CONTA", dbText, 10, True)
Call AddField(campo, "DESCRICAO", dbText, 50)
Call AddField(campo, "VALOR", dbCurrency)

Call AddIndex(indices, "iCONTA", Array("CONTA"), True, True, True)

Estrutura_BalanceteSintetico = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "BALANCETESINTETICO", campo, indices)

End Function


Function Estrutura_SaldoContas() As Boolean
ReDim campo(0)
ReDim indices(0)

AddField campo, "CONTA", dbLong, 0, True, False
AddField campo, "DATA", dbDate
AddField campo, "ANTERIOR", dbCurrency
AddField campo, "DEBITO", dbCurrency
AddField campo, "CREDITO", dbCurrency

AddIndex indices, "iCONTA", Array("CONTA", "DATA")
AddIndex indices, "iDATA", Array("DATA", "CONTA")

Estrutura_SaldoContas = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Saldo_Contas", campo, indices)

End Function





Function RetornaSaldoAnterior(pConta As Long, pDatabase As Date) As Currency
Dim wl_tbSaldo As Recordset
Dim wl_SaldoAnterior As Currency
If Not Abre_SaldoContas(wl_tbSaldo) Then
   Exit Function
End If
wl_tbSaldo.Seek "<", pConta, pDatabase
If Not wl_tbSaldo.NoMatch Then
   If wl_tbSaldo("CONTA") = pConta Then
      If Not pb_InverteOperacao Then
         wl_SaldoAnterior = wl_tbSaldo("ANTERIOR") - wl_tbSaldo("DEBITO") + wl_tbSaldo("CREDITO")
      Else
         wl_SaldoAnterior = wl_tbSaldo("ANTERIOR") + wl_tbSaldo("DEBITO") - wl_tbSaldo("CREDITO")
      End If
   End If
End If
RetornaSaldoAnterior = wl_SaldoAnterior
End Function

Function RetornaSaldoAtual(pConta As Long, pDataInicio As Date, pDataFinal As Date, Optional pTipo As String = "T")
Dim wl_SaldoAnterior As Currency
Dim wl_TotalCredito As Currency
Dim wl_TotalDebito As Currency
Dim wl_tbSaldo As Recordset
Dim wl_tbContas As Recordset
If Not Abre_SaldoContas(wl_tbSaldo) Or _
   Not Abre_PlanoContas(wl_tbContas) Then
   Exit Function
End If
If Not Loca_Contas(wl_tbContas, pConta) Then
   Exit Function
End If
wl_SaldoAnterior = RetornaSaldoAnterior(pConta, pDataInicio)
wl_tbSaldo.Seek ">=", pConta, pDataInicio
If Not wl_tbSaldo.NoMatch Then
   If wl_tbSaldo("CONTA") = pConta Then
      Do While Not wl_tbSaldo.EOF
         If wl_tbSaldo("DATA") > pDataFinal Then
            Exit Do
         End If
         If wl_tbSaldo("CONTA") <> pConta Then
            Exit Do
         End If
         wl_TotalCredito = wl_TotalCredito + wl_tbSaldo("CREDITO")
         wl_TotalDebito = wl_TotalDebito + wl_tbSaldo("DEBITO")
         wl_tbSaldo.MoveNext
      Loop
   End If
End If
If pTipo = "D" Then
   RetornaSaldoAtual = wl_TotalDebito
ElseIf pTipo = "C" Then
   RetornaSaldoAtual = wl_TotalCredito
ElseIf pTipo = "A" Then
   RetornaSaldoAtual = wl_SaldoAnterior
Else
   If Not pb_InverteOperacao Then
      RetornaSaldoAtual = wl_SaldoAnterior + wl_TotalCredito - wl_TotalDebito
   Else
      RetornaSaldoAtual = wl_SaldoAnterior - wl_TotalCredito + wl_TotalDebito
   End If
End If
End Function



Public Function Estrutura_SaldoPlano()
ReDim campo(0)
ReDim indices(0)
AddField campo, "TRADUTOR", dbLong, , True
AddField campo, "EXERCICIO", dbLong, , True
AddField campo, "SALDO ABERTURA", dbCurrency
AddField campo, "D01", dbCurrency
AddField campo, "D02", dbCurrency
AddField campo, "D03", dbCurrency
AddField campo, "D04", dbCurrency
AddField campo, "D05", dbCurrency
AddField campo, "D06", dbCurrency
AddField campo, "D07", dbCurrency
AddField campo, "D08", dbCurrency
AddField campo, "D09", dbCurrency
AddField campo, "D10", dbCurrency
AddField campo, "D11", dbCurrency
AddField campo, "D12", dbCurrency
AddField campo, "C01", dbCurrency
AddField campo, "C02", dbCurrency
AddField campo, "C03", dbCurrency
AddField campo, "C04", dbCurrency
AddField campo, "C05", dbCurrency
AddField campo, "C06", dbCurrency
AddField campo, "C07", dbCurrency
AddField campo, "C08", dbCurrency
AddField campo, "C09", dbCurrency
AddField campo, "C10", dbCurrency
AddField campo, "C11", dbCurrency
AddField campo, "C12", dbCurrency

AddIndex indices, "iTRADUTOR", Array("TRADUTOR", "EXERCICIO"), True, True, True
Estrutura_SaldoPlano = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Saldo_Plano_de_Contas", campo, indices)
End Function

Function Abre_BalanceteSintetico(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)
If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "BALANCETESINTETICO", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_BalanceteSintetico = True
   pObjeto_Recordset.Index = "iCONTA"
Else
   Abre_BalanceteSintetico = False
End If
End Function

Function Abre_Historico(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)

If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Historico", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_Historico = True
   pObjeto_Recordset.Index = "iCODIGO"
Else
   Abre_Historico = False
End If
End Function

Function Most_Historico()
ReDim aCampo(0)
aadd aCampo, Array("Código", "CODIGO", 2000, "000")
aadd aCampo, Array("Descrição", "DESCRICAO", 6000, "")
Most_Historico = Video("Histórico", aCampo, Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Historico", "iCODIGO", "CODIGO")
End Function


Function Estrutura_Historico() As Boolean
ReDim campo(0)
ReDim indices(0)

Call AddField(campo, "CODIGO", dbLong, 0, True)
Call AddField(campo, "DESCRICAO", dbMemo)

Call AddIndex(indices, "iCODIGO", Array("CODIGO"), True, True, True)

Estrutura_Historico = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Historico", campo, indices)
End Function

Function Estrutura_LanPadrao() As Boolean
ReDim campo(0)
ReDim indices(0)

AddField campo, "CODIGO", dbLong, 0, True
AddField campo, "DESCRICAO", dbText, 50
AddField campo, "CREDITO", dbLong
AddField campo, "DEBITO", dbLong

AddIndex indices, "iCODIGO", Array("CODIGO"), True, True, True
AddIndex indices, "iDESCRICAO", Array("DESCRICAO")

Estrutura_LanPadrao = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Lancamentos_Padrao", campo, indices)

End Function

Function Abre_LanPadrao(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)

If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Lancamentos_Padrao", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_LanPadrao = True
   pObjeto_Recordset.Index = "iCODIGO"
Else
   Abre_LanPadrao = False
End If

End Function


Function Estrutura_MoviCaixa() As Boolean
ReDim campo(0)
ReDim indices(0)

AddField campo, "DATA", dbDate, , True
AddField campo, "MOVIMENTO", dbLong, , True
AddField campo, "CREDITO", dbLong
AddField campo, "DEBITO", dbLong
AddField campo, "VALOR", dbCurrency
AddField campo, "HISTORICO", dbMemo
AddField campo, "ESTORNO", dbBoolean
AddField campo, "MOTIVO", dbMemo
AddField campo, "IMPORTADO", dbBoolean
AddField campo, "PADRAO", dbLong
AddField campo, "MOVIMENTOPRINCIPAL", dbLong

AddIndex indices, "iDATA", Array("DATA", "MOVIMENTO"), True, True, True
AddIndex indices, "iPRINCIPAL", "MOVIMENTOPRINCIPAL"

Estrutura_MoviCaixa = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Movimento_de_Caixa", campo, indices)
End Function

