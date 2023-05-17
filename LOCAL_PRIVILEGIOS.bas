Attribute VB_Name = "Privilégios"
Option Explicit
Public Const PR_SUPERVISOR = 0
Public Const PR_SISTEMA = 1
Public Const PR_SEGURANCA = 2
Public Const PR_USUARIOS = 3
Public Const PR_LANPADRAO = 4
Public Const PR_EMPRESAS = 5
Public Const PR_LANCAMENTOS = 6
Public Const PR_HISTORICO = 7
Public Const PR_PLANOCONTAS = 8
Public Const PR_BALANCOANALITICO = 9
Public Const PR_BALANCOSINTETICO = 10
Public Const PR_ROLPLANOCONTAS = 11
Public Const PR_ROLHISTORICOS = 12
Public Const PR_DIARIO = 13
Public Const PR_RAZAO = 14
Public Const PR_EXTRATO = 15
Public Const PR_ROTINA = 16

Function Carrega_Opcoes()
ReDim pb_Opcoes(0)
ReDim pb_Privilegios(0)
Dim wl_CCAA As Boolean


'Segundo Parametro da Matriz pb_opcoes
'<Vazio> ---> Todos Usuários
'<1>     ---> Usuários Supervisor
'<2>     ---> Usuário Master  -- as configurações influem em todos os usuários
'<3>     ---> Usuário Master

wl_CCAA = Dir(PathPadrao + "CCAA.SYS") <> ""

Call aadd(pb_Opcoes, Array("Supervisor", 1))
Call aadd(pb_Privilegios, Array("&Supervisor", "", "", "", "", "", "", "", "", "", "", "", "", "", ""))

Call aadd(pb_Opcoes, Array("Sistema", 1))
Call aadd(pb_Privilegios, Array("A&cesso"))

Call aadd(pb_Opcoes, Array("Segurança", 1))
Call aadd(pb_Privilegios, Array("A&cesso"))

Call aadd(pb_Opcoes, Array("Usuários", 1))
Call aadd(pb_Privilegios, Array("A&cesso", "&Altera", "&Deleta", "&Inclui"))

Call aadd(pb_Opcoes, Array("Lançamentos Padrão", 1))
Call aadd(pb_Privilegios, Array("A&cesso", "&Altera", "&Deleta", "&Inclui"))

Call aadd(pb_Opcoes, Array("Empresas", 1))
Call aadd(pb_Privilegios, Array("A&cesso", "&Altera", "&Deleta", "&Inclui"))

Call aadd(pb_Opcoes, Array("Lançamentos Contábeis"))
Call aadd(pb_Privilegios, Array("A&cesso", "&Altera", "&Deleta", "&Inclui"))

Call aadd(pb_Opcoes, Array("Histórico"))
Call aadd(pb_Privilegios, Array("A&cesso", "&Altera", "&Deleta", "&Inclui"))

Call aadd(pb_Opcoes, Array("Plano de Contas"))
Call aadd(pb_Privilegios, Array("A&cesso", "&Altera", "&Deleta", "&Inclui"))

Call aadd(pb_Opcoes, Array("Balancete Analítico"))
Call aadd(pb_Privilegios, Array("A&cesso"))

Call aadd(pb_Opcoes, Array("Balancete Sintético"))
Call aadd(pb_Privilegios, Array("A&cesso"))

Call aadd(pb_Opcoes, Array("Relatório de Plano de Contas"))
Call aadd(pb_Privilegios, Array("A&cesso"))

Call aadd(pb_Opcoes, Array("Relatório de Históricos"))
Call aadd(pb_Privilegios, Array("A&cesso"))

Call aadd(pb_Opcoes, Array("Diário Legal"))
Call aadd(pb_Privilegios, Array("A&cesso"))

Call aadd(pb_Opcoes, Array("Razão Analítico"))
Call aadd(pb_Privilegios, Array("A&cesso"))

Call aadd(pb_Opcoes, Array("Preferencias"))
Call aadd(pb_Privilegios, Array("A&cesso"))

End Function

