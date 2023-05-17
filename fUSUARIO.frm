VERSION 5.00
Begin VB.Form fUSUARIO 
   Caption         =   "Inserção de Usuários"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   310
   Icon            =   "fUSUARIO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2220
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSENHAATUAL 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   75
      MaxLength       =   10
      PasswordChar    =   "X"
      TabIndex        =   1
      Top             =   825
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtCONFIRMA 
      BackColor       =   &H00000080&
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   75
      MaxLength       =   10
      PasswordChar    =   "X"
      TabIndex        =   3
      Top             =   1815
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtSENHA 
      BackColor       =   &H00000080&
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   75
      MaxLength       =   10
      PasswordChar    =   "X"
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtUSUARIO 
      Height          =   300
      Left            =   75
      MaxLength       =   10
      TabIndex        =   0
      Top             =   300
      Width           =   5385
   End
   Begin VB.Label legenda_SenhaAtual 
      Caption         =   "Senha Atual"
      Height          =   195
      Left            =   75
      TabIndex        =   7
      Top             =   630
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label legenda_CONFIRMA 
      Caption         =   "Senha"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label legenda_SENHA 
      Caption         =   "Senha"
      Height          =   195
      Left            =   75
      TabIndex        =   5
      Top             =   1125
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Nome do Usuário"
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   105
      Width           =   1305
   End
End
Attribute VB_Name = "fUSUARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private wp_Saida As Boolean
Dim tbSeguranca As Recordset
Dim wp_Cria As Boolean


Private Function Del_Usuario()
Dim tbprivilegios As Recordset
If Not Abre_Privilegios(tbprivilegios) Then
   Exit Function
End If
tbprivilegios.Seek ">=", ctox(pb_Sistema), ctox(txtUSUARIO.Text), ctox("000")
If Not tbprivilegios.NoMatch Then
   If tbprivilegios("USUARIO") = ctox(txtUSUARIO.Text) Then
      Do While Not tbprivilegios.EOF
         If tbprivilegios("USUARIO") <> ctox(txtUSUARIO.Text) Then
            Exit Do
         End If
         If Not edit_reg(tbprivilegios) Then
            Call MsgBox("Erro ao deletar registro de privilégio", vbExclamation, "Mensagem do Sistema")
            Exit Function
         End If
         tbprivilegios.Delete
         tbprivilegios.MoveNext
      Loop
   End If
End If
tbSeguranca.Delete
tbprivilegios.Close
End Function


Private Sub Form_Activate()
If Not Abre_Usuarios(tbSeguranca) Then
   Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   If wp_Saida Then
      Unload Me
   Else
      txtUSUARIO.SetFocus
   End If
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_GotFocus()
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbSeguranca.Close
dbSeguranca.Close
aviso
End Sub

Private Sub txtCONFIRMA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   KeyCode = 0
   SendKeys "+{tab}"
End If
End Sub


Private Sub txtCONFIRMA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtCONFIRMA.Text <> txtSENHA.Text Then
      Call MsgBox("Confirmação de senha inconsistente", vbCritical, "Mensagem do Sistema")
      txtSENHA.SetFocus
      Exit Sub
   End If
   If wp_Cria Then
      If Not add_reg(tbSeguranca) Then
         txtUSUARIO.SetFocus
         Exit Sub
      End If
      tbSeguranca("NOME") = ctox(txtUSUARIO.Text)
   Else
      If Not edit_reg(tbSeguranca) Then
         txtUSUARIO.SetFocus
         Exit Sub
      End If
   End If
   tbSeguranca("SENHA") = ctox(txtSENHA.Text)
   If Not update_reg(tbSeguranca) Then
      txtUSUARIO.SetFocus
      Exit Sub
   End If
   SendKeys "{tab}"
End If
End Sub


Private Sub txtSENHA_GotFocus()
txtSENHAATUAL.Visible = False
legenda_SenhaAtual.Visible = False
txtSENHA.Text = ""
txtCONFIRMA.Text = ""
txtCONFIRMA.Visible = False
legenda_CONFIRMA.Visible = False
End Sub

Private Sub txtSENHA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   KeyCode = 0
   SendKeys "+{tab}"
End If
End Sub


Private Sub txtSENHA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtSENHA.Text = txtSENHAATUAL.Text Then
      Call MsgBox("Impossível mudar para mesma senha", vbExclamation, "Mensagem do Sistema")
      txtUSUARIO.SetFocus
      Exit Sub
   End If
   legenda_CONFIRMA.Top = 1155
   legenda_CONFIRMA.Caption = "Confirme a Senha"
   txtCONFIRMA.Top = 1350
   legenda_CONFIRMA.Visible = True
   txtCONFIRMA.Visible = True
   SendKeys "{tab}"
End If
End Sub


Private Sub txtSENHAATUAL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
   If MsgBox("Deseja deletar o usuário?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbYes Then
      If Not edit_reg(tbSeguranca) Then
         txtUSUARIO.SetFocus
         Exit Sub
      End If
      Call Del_Usuario
   End If
   txtUSUARIO.SetFocus

End If
End Sub

Private Sub txtSENHAATUAL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If ctox(txtSENHAATUAL.Text) <> tbSeguranca("SENHA") Then
      Call MsgBox("Senha Inconsistente", vbCritical, "Mensagem do Sistema")
      Unload Me
   End If
   legenda_SENHA.Top = 630
   txtSENHA.Top = 825
   legenda_SENHA.Caption = "Nova Senha"
   legenda_SENHA.Visible = True
   txtSENHA.Visible = True
   SendKeys "{tab}"
End If
End Sub


Private Sub txtUSUARIO_Change()
If UCase(txtUSUARIO.Text) = "ISSO N" Or _
   UCase(txtUSUARIO.Text) = "RAIN" Or _
   UCase(txtUSUARIO.Text) = "QUI MIN" Or _
   UCase(txtUSUARIO.Text) = "AUTORIZ" Then
      InformaaoUsuario "Nome de usuário não permitido"
      txtUSUARIO.Text = ""
End If
End Sub

Private Sub txtUSUARIO_GotFocus()
aviso "<F1> Usuários"
wp_Saida = True
If Not ShowRetorno Then
   txtUSUARIO.Text = ""
Else
   ShowRetorno = False
End If
txtSENHA.Text = ""
txtCONFIRMA.Text = ""
txtSENHAATUAL.Text = ""

txtSENHAATUAL.Visible = False
legenda_SenhaAtual.Visible = False
legenda_SENHA.Visible = False
txtSENHA.Visible = False
legenda_CONFIRMA.Visible = False
txtCONFIRMA.Visible = False
End Sub


Private Sub txtUSUARIO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno As String
If KeyCode = vbKeyF1 Then
   wl_Retorno = Most_Usuarios
   If wl_Retorno <> "" Then txtUSUARIO.Text = wl_Retorno
   txtUSUARIO.SetFocus
End If
End Sub

Private Sub txtUSUARIO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtUSUARIO.Text = "" Then
      Call MsgBox("Informe o usuário", vbExclamation, "Mensagem do Sistema")
      txtUSUARIO.SetFocus
      Exit Sub
   End If
   If tbSeguranca.RecordCount > 0 Then
      tbSeguranca.Seek "=", ctox(txtUSUARIO.Text)
      If Not tbSeguranca.NoMatch Then
         txtSENHAATUAL.Visible = True
         legenda_SenhaAtual.Visible = True
         txtSENHAATUAL.SetFocus
         wp_Cria = False
         Exit Sub
      Else
         If MsgBox("Deseja criar um novo usuário?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbNo Then
            txtUSUARIO.Text = ""
            txtUSUARIO.SetFocus
            Exit Sub
         End If
         legenda_SENHA.Top = 630
         txtSENHA.Top = 825
         legenda_SENHA.Visible = True
         txtSENHA.Visible = True
         wp_Cria = True
      End If
   Else
      legenda_SENHA.Top = 630
      txtSENHA.Top = 825
      legenda_SENHA.Visible = True
      txtSENHA.Visible = True
         wp_Cria = True

   End If
   SendKeys "{tab}"
End If
End Sub

Private Sub txtUSUARIO_LostFocus()
aviso
wp_Saida = False
End Sub


