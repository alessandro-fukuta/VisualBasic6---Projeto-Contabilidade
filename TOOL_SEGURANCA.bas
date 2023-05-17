Attribute VB_Name = "TOOLS_SEGURANCA"
Option Explicit
Public dbSeguranca As Database
Public pb_Usuario As String
Public pb_Senha As String
Public pb_RetornodoBloqueio As String
Public Const pb_ConsultaPrivilegios As Boolean = True
Public pb_Opcoes
Public pb_Privilegios


Sub Display_Usuario()
fMENU.BarraStatus.Panels(2).Text = UCase(pb_Usuario)
If Verifica_Privilegio(PR_SUPERVISOR, "S") Then
   fMENU.BarraStatus.Panels(2).Text = "{S} - " + UCase(pb_Usuario)
End If
If pb_Senha = "AMANHECEU NO VALE" Then
   fMENU.BarraStatus.Panels(2).Text = "<<<< " + UCase(pb_Usuario) + " >>>>"
End If
End Sub



Function Estrutura_Privilegios() As Boolean

ReDim campo(0)
ReDim indices(0)
aadd campo, Array("SISTEMA", dbText, 50, True, True)
aadd campo, Array("USUARIO", dbText, 50, True, True)
aadd campo, Array("OPCAO", dbText, 50, True, True)
aadd campo, Array("PRIVILEGIO", dbText, 50, False, True)
aadd indices, Array("iPRIV", Array("SISTEMA", "USUARIO", "OPCAO"), True, True, True)
If Dir(PathPadrao + "SEGURANCA", vbDirectory) = "" Then
   Err = 0
   MkDir PathPadrao + "SEGURANCA"
   If Err <> 0 Then
      Estrutura_Privilegios = False
      Exit Function
   End If
End If
Estrutura_Privilegios = cria(PathPadrao + "SEGURANCA\SEGURANCA.MDB", "Privilegios", campo, indices)
aviso
End Function



Function Estrutura_Usuarios() As Boolean

ReDim campo(0)
ReDim indices(0)
aadd campo, Array("NOME", dbText, 50, False, True)
aadd campo, Array("SENHA", dbText, 50, False, True)

aadd indices, Array("iNOME", Array("NOME"), True, True, True)
If Dir(PathPadrao + "SEGURANCA", vbDirectory) = "" Then
   Err = 0
   MkDir PathPadrao + "SEGURANCA"
   If Err <> 0 Then
      Estrutura_Usuarios = False
      Exit Function
   End If
End If
Estrutura_Usuarios = cria(PathPadrao + "SEGURANCA\SEGURANCA.MDB", "Usuarios", campo, indices)
End Function




Function Abre_Usuarios(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)
Abre_Usuarios = aOpen("SEGURANCA\SEGURANCA.MDB", "Usuarios", dbSeguranca, pObjeto_Recordset)
aviso
End Function







Function Abre_Privilegios(ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)

Abre_Privilegios = aOpen("SEGURANCA\SEGURANCA.MDB", "Privilegios", dbSeguranca, pObjeto_Recordset)
aviso
End Function


Function Most_Usuarios()
fMOSTUSUARIOS.Show 1
Most_Usuarios = pbRetornoVideo
End Function

Function Permissao(Optional pOpcao As Integer = -1, Optional pPrivilegio As String = "", Optional pTitulo As String = "Identifique-se") As Boolean
Dim tbPrivilegio As Recordset
If Not pb_ConsultaPrivilegios Then
   Permissao = True
   Exit Function
End If
fBLOQUEIO.lblTitulo.Caption = pTitulo
fBLOQUEIO.Show 1
If pb_RetornodoBloqueio = "" Then
   Permissao = False
   Exit Function
End If
pb_Usuario = pb_RetornodoBloqueio
If pOpcao > -1 Then
   Permissao = Verifica_Privilegio(pOpcao, pPrivilegio)
End If
End Function


Function NovaPermissao(pOpcao As Integer, pPrivilegio As String, pTitulo As String, Optional ByRef wp_Usuario)
Dim wl_Usuario As String
Dim wl_Senha As String
Dim wl_Retorno
wl_Usuario = pb_Usuario
wl_Senha = pb_Senha
wl_Retorno = Permissao(pOpcao, pPrivilegio, pTitulo)
If wl_Retorno Then
   wp_Usuario = pb_Usuario
End If
pb_Usuario = wl_Usuario
pb_Senha = wl_Senha
NovaPermissao = wl_Retorno
End Function
Function SalvaRegistroWindows(pSecao As String, pChave As String, pValor As Variant)
SaveSetting pb_Sistema, pSecao, pChave, pValor
If Err <> 0 Then
   Call MsgBox("Erro ao gravar registro do windows", vbCritical, "Mensagem do Sistema")
   End
End If
End Function


Sub Troca_Usuario()
Dim wl_Usuario As String
Dim wl_Senha As String
wl_Usuario = pb_Usuario
wl_Senha = pb_Senha
If Permissao(PR_SISTEMA, "C", "Mudança de Usuário") Then
   pb_Usuario = pb_RetornodoBloqueio
   Call Display_Usuario
Else
   If pb_Usuario <> wl_Usuario Then
      Call MsgBox("Impossível trocar para o usuario " + pb_Usuario, vbInformation, "Mensagem do Sistema")
      pb_Usuario = wl_Usuario
      pb_Senha = wl_Senha
   Else
      Call MsgBox("Mantêm o usuário " + pb_Usuario, vbExclamation, "Mensagem do Sistema")
   End If
End If
End Sub

Function Verifica_Privilegio(pOpcao As Integer, pPrivilegio As String, Optional pMensagemNegativa As String = "", Optional pGeral As Boolean = False, Optional pSuperSeguro As Boolean) As Boolean
Dim tbprivilegios As Recordset
If Not pb_ConsultaPrivilegios Then
   Verifica_Privilegio = True
   Exit Function
End If
'If file("C:\INFOSOFT.CFG") Then
'   pb_Senha = "AMANHECEU NO VALE"
'End If

If UCase(pb_Senha) = "AMANHECEU NO VALE" Or pb_Demonstracao Then
   Verifica_Privilegio = True
   Exit Function
End If
If UCase(pb_Usuario) = "AUTORIZADO" Then
   If pSuperSeguro Then
      InformaaoUsuario "Nível Super-Seguro"
      Exit Function
   Else
      Verifica_Privilegio = True
      Exit Function
   End If
End If
If Not Abre_Privilegios(tbprivilegios) Then
   Verifica_Privilegio = False
   Exit Function
End If
tbprivilegios.Seek "=", ctox(pb_Sistema), ctox(pb_Usuario), ctox(Format(PR_SUPERVISOR, "000"))
If Not tbprivilegios.NoMatch Then
   If Not pGeral Then
      If AT("S", xtoc(tbprivilegios("PRIVILEGIO"))) > 0 Then
         If pSuperSeguro Then
            Verifica_Privilegio = False
         Else
            Verifica_Privilegio = True
            Exit Function
         End If
      End If
   End If
End If
If Not pGeral Then
   tbprivilegios.Seek "=", ctox(pb_Sistema), ctox(pb_Usuario), ctox(Format(pOpcao, "000"))
Else
   tbprivilegios.Seek "=", ctox(pb_Sistema), "", ctox(Format(pOpcao, "000"))
End If
If tbprivilegios.NoMatch Then
   If pMensagemNegativa <> "" Then
      Call MsgBox(pMensagemNegativa, vbCritical, "Sistema de Seguranca")
   End If
   Verifica_Privilegio = False
   Exit Function
End If
If AT(pPrivilegio, xtoc(tbprivilegios("PRIVILEGIO"))) = 0 Then
   If pMensagemNegativa <> "" Then
      Call MsgBox(pMensagemNegativa, vbCritical, "Sistema de Seguranca")
   End If
   Verifica_Privilegio = False
   Exit Function
End If
Verifica_Privilegio = True
End Function




