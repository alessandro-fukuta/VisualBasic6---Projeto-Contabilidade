Attribute VB_Name = "Local"
Option Explicit
Public pb_ClienteAtualizado As Boolean
Public pb_Controle As Control
Public pb_NovoVencimento As Date
Public pb_RETIRADA As Currency
Public pb_HISTORICOdaRETIRADA As String
Public pb_REFORCO As Currency
Public pb_HISTORICOdoREFORCO As String
Public pb_linha As Long
Public Const pb_Sistema = "CONTABILIDADE"
Public dollar As Currency



'Sub Main()
'Dim i As Integer
'Dim wl_retorno As String
'fABERTURA.Show 1
'If Not pb_Demonstracao Then
'   wl_retorno = RetornaConfiguracao("Preferencias", "PathPadrao")
'   If wl_retorno = "" Then
'      fPATHPADRAO.Show 1
'   Else
'      PathPadrao = wl_retorno
'   End If
'Else
'   PathPadrao = "\MDB_DEMO\"
'   If Dir("\MDB_DEMO", vbDirectory) = "" Then
'      MkDir "\MDB_DEMO"
'   End If
'End If
'fMENU.Show
'End Sub


Public Function Repara_Banco()
On Error Resume Next
fMENU.MousePointer = 11
dbPROTECAO.Close
dbFINANCEIRO.Close
dbPRODUTOS.Close
dbMOVIMENTO.Close
On Error GoTo erro
Err = 0
Informacao
If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.MDB") <> "" Then
   DisplayMensagem "Aguarde, reparando PROTECAO.MDB ..."
   DBEngine.RepairDatabase PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.MDB"
   DisplayMensagem "Compactando PROTECAO.MDB"
   DBEngine.CompactDatabase PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.MDB", PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.COP"
   If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.OLD") <> "" Then
      Kill PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.OLD"
   End If
   Name PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.MDB" As PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.OLD"
   Name PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.COP" As PathPadrao + Format(pb_Empresa, "00000") + "\PROTECAO.MDB"
End If
If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.MDB") <> "" Then
   DisplayMensagem "Aguarde, reparando PRODUTOS.MDB ..."
   DBEngine.RepairDatabase PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.MDB"
   DisplayMensagem "Compactando PRODUTOS.MDB"
   DBEngine.CompactDatabase PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.MDB", PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.COP"
   If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.OLD") <> "" Then
      Kill PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.OLD"
   End If
   Name PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.MDB" As PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.OLD"
   Name PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.COP" As PathPadrao + Format(pb_Empresa, "00000") + "\PRODUTOS.MDB"
End If
If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB") <> "" Then
   DisplayMensagem "Aguarde, reparando FINANCEIRO.MDB ..."
   DBEngine.RepairDatabase PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB"
   DisplayMensagem "Compactando FINANCEIRO.MDB"
   DBEngine.CompactDatabase PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB", PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.COP"
   If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.OLD") <> "" Then
      Kill PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.OLD"
   End If
   Name PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB" As PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.OLD"
   Name PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.COP" As PathPadrao + Format(pb_Empresa, "00000") + "\FINANCEIRO.MDB"
End If
If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.MDB") <> "" Then
   DisplayMensagem "Aguarde, reparando MOVIMENTO.MDB ..."
   DBEngine.RepairDatabase PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.MDB"
   DisplayMensagem "Compactando MOVIMENTO.MDB"
   DBEngine.CompactDatabase PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.MDB", PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.COP"
   If Dir(PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.OLD") <> "" Then
      Kill PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.OLD"
   End If
   Name PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.MDB" As PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.OLD"
   Name PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.COP" As PathPadrao + Format(pb_Empresa, "00000") + "\MOVIMENTO.MDB"
End If
Informacao
fMENU.MousePointer = 0
Exit Function


erro:
fMENU.MousePointer = 0
MsgBox Err.Description, vbCritical, "Erro nro. " + Str(Err.Number)
Exit Function
End Function

Function Verifica_Estruturas() As Boolean
Dim wl_Retorno As Boolean
pb_MontaProgressao = True
fMENU.Progressao.Min = 0
fMENU.Progressao.Max = 11
fMENU.Progressao.Value = 0
fMENU.Progressao.Visible = True
wl_Retorno = True
Informacao
DisplayMensagem "Infosoft CONTABIL Verificando estruturas, aguarde ..."
If Not Estrutura_Privilegios Or Not _
       Estrutura_Usuarios Or Not _
       Estrutura_PlanoContas Or Not _
       Estrutura_Preferencias Or Not _
       Estrutura_SaldoContas Or Not _
       Estrutura_Empresas Or Not _
       Estrutura_MoviCaixa Or Not _
       Estrutura_RateioContas Or Not _
       Estrutura_Historico Or Not _
       Estrutura_LanPadrao Then

   wl_Retorno = False
End If

'   Estrutura_LanPadrao Or Not _
'   Estrutura_RelatorioProgramado Or Not _

If Not Estrutura_BalanceteSintetico Then
   wl_Retorno = False
End If

Informacao
Verifica_Estruturas = wl_Retorno
pb_MontaProgressao = False
fMENU.Progressao.Visible = False
aviso

End Function
