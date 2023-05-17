Attribute VB_Name = "Comum"
Option Explicit
Public dbFINANCEIRO As Database
Public dbPROTECAO As Database
Public dbEMPRESAS As Database
Public dbVendas As Database
Public dbPRODUTOS As Database
Public dbMOVIMENTO As Database
Public dbFluxo As Database
Public dbComum As Database
Public dbPreferencias As Database
Public pb_TitularBloqueado As Boolean
Public pb_Cria As Boolean
Public pb_Codigo As Long
Public pb_MaquinaReady As Boolean
Public pb_NaoMatricial As Boolean
Public pb_NaoVideo As Boolean
Public pb_NaoJato As Boolean
Public pb_RAZAOSOCIAL As String
Public pb_FONEEMPRESA As String
Public pb_Cidade As String
Public pb_Estado As String
Public pb_Empresa As Long
Public pb_FormAtivo As String
Public pb_Endereco As String
Public pb_ObjetoAtivo As String
Public pb_Informacao As Boolean
Public pb_Modal As Integer
Public pb_Online As Boolean



Sub Main()
Dim i As Integer
Dim wl_retorno As String
fABERTURA.Show 1
' fApresentaFlash.Show 1
wl_retorno = RetornaConfiguracao("Preferencias", "PathPadrao")
If wl_retorno = "" Then
   fPATHPADRAO.Show 1
Else
   PathPadrao = wl_retorno
End If
fMENU.Show
End Sub

Function Estrutura_Preferencias() As Boolean
'Criando Tabela Bairros
ReDim campo(0)
ReDim indices(0)
aviso "Verificando estrutura da tabela Preferencias"
AddField campo, "CONTA_MASTER", dbLong
AddField campo, "CONTA_FUNCIONARIO", dbLong, , True

Estrutura_Preferencias = cria(PathPadrao + Format(pb_Empresa, "00000") + "\PREFERENCIAS.MDB", "Gerais", campo, indices)
aviso
End Function

Sub Recalcula_Saldo(Optional pData)
Dim tbMovi As Recordset
Dim tbSaldo As Recordset
Dim tbContas As Recordset
Dim wl_Data As Date
Dim wl_SaldoAnterior As Currency
Dim wl_Conta As Long
If Not Abre_MoviCaixa(tbMovi) Or _
   Not Abre_SaldoContas(tbSaldo) Or _
   Not Abre_PlanoContas(tbContas) Then
   Exit Sub
End If
tbSaldo.Index = "iDATA"
If tbMovi.RecordCount = 0 Then
   Exit Sub
End If
Informacao
DisplayMensagem "Aguarde. Limpando saldos antigos ..."
If tbSaldo.RecordCount > 0 Then
   If IsMissing(pData) Then
      DisplayMensagem "Recalculando Todos os saldos ..."
      tbSaldo.MoveFirst
   Else
      If pData = "00:00:00" Or pData = "" Then
         DisplayMensagem "Recalculando Todos os saldos ..."
         tbSaldo.MoveFirst
      Else
         DisplayMensagem "Recalculando Saldos a partir de " + CStr(pData)
         tbSaldo.Seek ">=", pData, 1
      End If
   End If
   Do While Not tbSaldo.EOF
      If edit_reg(tbSaldo) Then
         tbSaldo.Delete
      End If
      tbSaldo.MoveNext
   Loop
End If
tbSaldo.Index = "iCONTA"
If IsMissing(pData) Then
   tbMovi.MoveFirst
Else
   tbMovi.Seek ">=", pData, 1
   If tbMovi.NoMatch Then
      GoSub FechaTabelas
      Informacao
      Exit Sub
   End If
End If
Do While Not tbMovi.EOF
   If tbMovi("CREDITO") > 0 And Not tbMovi("ESTORNO") Then
      wl_Conta = tbMovi("CREDITO")
      tbSaldo.Seek "=", wl_Conta, tbMovi("DATA")
      If tbSaldo.NoMatch Then
         GoSub Anterior
         tbSaldo.AddNew
         tbSaldo("CONTA") = wl_Conta
         tbSaldo("DATA") = tbMovi("DATA")
         tbSaldo("ANTERIOR") = wl_SaldoAnterior
         tbSaldo("CREDITO") = tbMovi("VALOR")
         tbSaldo("DEBITO") = 0
         tbSaldo.Update
      Else
         tbSaldo.Edit
         tbSaldo("CREDITO") = tbSaldo("CREDITO") + tbMovi("VALOR")
         tbSaldo.Update
      End If
   End If
   If tbMovi("DEBITO") > 0 And Not tbMovi("ESTORNO") Then
      wl_Conta = tbMovi("DEBITO")
      tbSaldo.Seek "=", wl_Conta, tbMovi("DATA")
      If tbSaldo.NoMatch Then
         GoSub Anterior
         tbSaldo.AddNew
         tbSaldo("CONTA") = wl_Conta
         tbSaldo("DATA") = tbMovi("DATA")
         tbSaldo("ANTERIOR") = wl_SaldoAnterior
         tbSaldo("DEBITO") = tbMovi("VALOR")
         tbSaldo("CREDITO") = 0
         tbSaldo.Update
      Else
         tbSaldo.Edit
         tbSaldo("DEBITO") = tbSaldo("DEBITO") + tbMovi("VALOR")
         tbSaldo.Update
      End If
   End If
   tbMovi.MoveNext
Loop
pb_DeletaMovimento = False
If Dir(PathPadrao + "DELETA.MOV") <> "" Then Kill PathPadrao + "DELETA.MOV"
Informacao
Exit Sub


Anterior:
wl_SaldoAnterior = 0
tbSaldo.Seek "<", wl_Conta, tbMovi("DATA")
If Not tbSaldo.NoMatch Then
   If tbSaldo("CONTA") = wl_Conta Then
      If Not pb_InverteOperacao Then
         wl_SaldoAnterior = tbSaldo("ANTERIOR") - tbSaldo("DEBITO") + tbSaldo("CREDITO")
      Else
         wl_SaldoAnterior = tbSaldo("ANTERIOR") + tbSaldo("DEBITO") - tbSaldo("CREDITO")
      End If
   End If
End If
Return

FechaTabelas:
tbContas.Close
tbMovi.Close
tbSaldo.Close
Return
End Sub


Function aviso(Optional pMensagem As String = "", Optional pLuz As Boolean = False)
Exit Function
End Function

Sub EliminaLinhadaGrade(pGrade As Object)
Dim wl_Rows As Integer
Dim wl_Cols As Integer
Dim ii As Integer
Dim jj As Integer
wl_Rows = pGrade.rows
wl_Cols = pGrade.cols - 1
For ii = pGrade.row To pGrade.rows - 2
   For jj = 0 To wl_Cols
      pGrade.TextMatrix(ii, jj) = pGrade.TextMatrix(ii + 1, jj)
   Next
Next
For ii = 0 To wl_Cols
   pGrade.TextMatrix(pGrade.rows - 1, ii) = ""
Next
End Sub


Function Informacao()
pb_Informacao = Not pb_Informacao
If Not pb_Informacao Then
   Unload fInformacao
End If
End Function

Function DisplayMensagem(Optional pMensagem = "")
On Error Resume Next
If Not pb_Informacao Then Exit Function
fInformacao.Show , fMENU
If Err <> 0 Then Exit Function
fInformacao.lblMensagem.Caption = pMensagem
DoEvents
End Function

Function Abre_RelatorioProgramado(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)
If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Relatorio_Programado", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_RelatorioProgramado = True
   pObjeto_Recordset.Index = "iCONTA"
Else
   Abre_RelatorioProgramado = False
End If
aviso
End Function



Function Abre_Preferencias(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)
'aviso "Abrindo tabela Preferencias Gerais ..."
If aOpen(Format(pb_Empresa, "00000") + "\PREFERENCIAS.MDB", "Gerais", dbPROTECAO, pObjeto_Recordset) Then
   Abre_Preferencias = True
Else
   Abre_Preferencias = False
End If
aviso
End Function



Sub InformaEmpresa()
Dim tbEmpresas As Recordset
fEmpresa_Padrao.Show 1
If pb_Empresa = 0 Then
   InformaaoUsuario "A empresa não foi informada. O sistema será finalizado"
   End
End If
If Not Abre_Empresas(tbEmpresas) Then
   InformaaoUsuario "Não foi possível abrir a tabela empresas. O sistema será finalizado"
   End
End If
If Not Loca_Empresas(tbEmpresas, pb_Empresa) And Not pb_Demonstracao Then
   InformaaoUsuario "Empresa não encontrada"
   tbEmpresas.Close
   End
End If
If Dir(PathPadrao + Format(pb_Empresa, "00000"), vbDirectory) = "" Then
   MkDir PathPadrao + Format(pb_Empresa, "00000")
End If
fMENU.BarraStatus.Panels(1).Text = UCase(tbEmpresas("RAZAOSOCIAL"))
tbEmpresas.Close
End Sub


Function Confirme(pMensagem As String, Optional pButton As Integer = 1) As Boolean
Dim wl_retorno As Integer
Dim wl_Form As Form
On Error Resume Next
wl_retorno = MsgBox(pMensagem, vbQuestion + vbYesNo + IIf(pButton = 1, vbDefaultButton1, vbDefaultButton2), "Mensagem do Sistema")
If wl_retorno = vbYes Then
   Confirme = True
Else
   Confirme = False
End If
VB.Screen.ActiveForm.ActiveControl.SetFocus
End Function

Function Loca_Empresas(ByVal tbEmpresas As Recordset, txtCODIGO As Long) As Boolean
If pb_Demonstracao And tbEmpresas.RecordCount = 0 Then
   If add_reg(tbEmpresas) Then
      tbEmpresas("CODIGO") = 99999
      tbEmpresas("RAZAOSOCIAL") = "EMPRESA DEMONSTRACAO"
      update_reg tbEmpresas
   End If
End If
If tbEmpresas.RecordCount = 0 Then
   Loca_Empresas = False
   Exit Function
End If
tbEmpresas.Seek "=", txtCODIGO
Loca_Empresas = Not tbEmpresas.NoMatch
End Function

Function Most_Empresas()
ReDim aCampo(0)
aadd aCampo, Array("Código", "CODIGO", 800, "00000")
aadd aCampo, Array("Razão Social", "RAZAOSOCIAL", 5000, "")
aadd aCampo, Array("Telefone", "FONE1", 2000, "")
aadd aCampo, Array("Telefone", "FONE2", 2000, "")
aadd aCampo, Array("FAX", "FAX1", 2000, "")
aadd aCampo, Array("E-Mail", "EMAIL", 3000, "")

Most_Empresas = Video("Empresas", aCampo, "SEGURANCA\EMPRESAS.MDB", "Empresas", "iRAZAO", "CODIGO")
End Function




Function Estrutura_Empresas() As Boolean

ReDim campo(0)
ReDim indices(0)
aviso "Verificando estrutura da tabela Empresas"
Call AddField(campo, "CODIGO", dbLong, , True)
Call AddField(campo, "RAZAOSOCIAL", dbText, 50)
Call AddField(campo, "NOMEFANTASIA", dbText, 50)
Call AddField(campo, "CGC", dbText, 50)
Call AddField(campo, "INSCRICAO", dbText, 50)
Call AddField(campo, "CIDADE", dbText, 50)
Call AddField(campo, "ESTADO", dbText, 2)
Call AddField(campo, "CEP", dbText, 9)
Call AddField(campo, "ENDERECO", dbText, 50)
Call AddField(campo, "BAIRRO", dbText, 50)
Call AddField(campo, "EMAIL", dbText, 50)
Call AddField(campo, "OBSERVACOES", dbMemo)
Call AddField(campo, "DATA", dbDate)
Call AddField(campo, "FONE1", dbText, 50)
Call AddField(campo, "FONE2", dbText, 50)
Call AddField(campo, "FAX1", dbText, 50)
Call AddField(campo, "FAX2", dbText, 50)
Call AddField(campo, "ISS", dbCurrency)
AddField campo, "PEDIDO", dbLong
AddField campo, "INVERTEOPERACOES", dbBoolean
AddField campo, "LOTE", dbLong

Call AddIndex(indices, "iCODIGO", "CODIGO", True, True, True)
Call AddIndex(indices, "iRAZAO", "RAZAOSOCIAL")

If Dir(PathPadrao, vbDirectory) = "" Then
   MkDir PathPadrao
End If
If Dir(PathPadrao + "SEGURANCA", vbDirectory) = "" Then
   MkDir PathPadrao + "SEGURANCA"
End If
Estrutura_Empresas = cria(PathPadrao + "SEGURANCA\EMPRESAS.MDB", "Empresas", campo, indices)
aviso
End Function



Function Abre_Empresas(pObjeto_Recordset As Recordset, Optional EXCLUSIVO = False)
Abre_Empresas = aOpen("SEGURANCA\EMPRESAS.MDB", "EMPRESAS", dbEMPRESAS, pObjeto_Recordset)
aviso
End Function


Function Abre_PlanoContas(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False) As Boolean
If aOpen(Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Plano_de_Contas", dbFINANCEIRO, pObjeto_Recordset) Then
   Abre_PlanoContas = True
   pObjeto_Recordset.Index = "iTRADUTOR"
Else
   Abre_PlanoContas = False
End If
aviso
End Function


Public Function Estrutura_PlanoContas()
ReDim campo(0)
ReDim indices(0)
AddField campo, "CONTA", dbText, 30, True
AddField campo, "TRADUTOR", dbLong
AddField campo, "DESCRICAO", dbText, 50
AddField campo, "SALDOABERTURA", dbCurrency
AddField campo, "TIPO", dbText, 1
AddField campo, "ATIVOPASSIVO", dbLong

AddIndex indices, "iCONTA", "CONTA", True, True, True
AddIndex indices, "iTRADUTOR", "TRADUTOR"

Estrutura_PlanoContas = cria(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", "Plano_de_Contas", campo, indices)
End Function


