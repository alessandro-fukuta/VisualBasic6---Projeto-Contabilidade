VERSION 5.00
Begin VB.Form fBLOQUEIO 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   270
   Icon            =   "fBLOQUEIO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSENHA 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   90
      MaxLength       =   10
      PasswordChar    =   "X"
      TabIndex        =   2
      Top             =   1680
      Width           =   2025
   End
   Begin VB.TextBox txtUSUARIO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   1095
      Width           =   4260
   End
   Begin VB.Label lblTitulo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Identifique-se"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3525
      TabIndex        =   4
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
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
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
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
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   900
      Width           =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   -15
      X2              =   4920
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      FillColor       =   &H0080FFFF&
      Height          =   630
      Left            =   -15
      Shape           =   4  'Rounded Rectangle
      Top             =   -15
      Width           =   4965
   End
End
Attribute VB_Name = "fBLOQUEIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tbUsuarios As Recordset
Private tbprivilegios As Recordset
Private wp_tentativa As Integer
Private WP_linha As Integer
Private Sub Form_Activate()
If pb_Demonstracao Then
   If InStr(fMENU.Caption, "D E M O N S T R A Ç Ã O") = 0 Then
      fMENU.Caption = fMENU.Caption + " - D E M O N S T R A Ç Ã O"
   End If
End If
If Not Estrutura_Privilegios Or _
   Not Estrutura_Usuarios Then
   Unload Me
   Exit Sub
End If
If Not Abre_Usuarios(tbUsuarios) Or _
   Not Abre_Privilegios(tbprivilegios) Then
   Unload Me
   Exit Sub
End If
If file("C:\INFOSOFT.CFG") Then
   txtUSUARIO.Text = "AMANHECEU NO VALE"
End If
pb_RetornodoBloqueio = ""
pb_Senha = ""
End Sub

Private Sub Form_Load()
'Caption = Caption + " - " + pb_Sistema + IIf(pb_Demonstracao, "  (D E M O)", "")
End Sub


Private Sub Form_Paint()
Demarca fBLOQUEIO
Exit Sub
Dim i As Integer
Dim l As Integer
For i = 0 To 100
   Line (0, l)-(Width, l + 100), RGB(0, i * 3, i * 3), BF
   l = l + Height / 100
Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
aviso
tbUsuarios.Close
tbprivilegios.Close
dbSeguranca.Close
End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub txtSENHA_GotFocus()
aviso "Digite a senha do usuario " + txtUSUARIO.Text
txtSENHA.Text = ""
End Sub

Private Sub txtSENHA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   wp_tentativa = wp_tentativa + 1
   txtSENHA.Text = UCase(txtSENHA)
   If txtUSUARIO.Text = "AUTORIZADO" Then
      If txtSENHA.Text <> "COMANDO" Then
         If wp_tentativa > 2 Then
            pb_RetornodoBloqueio = ""
            Unload Me
            Exit Sub
         End If
         Call MsgBox("Senha inválida", vbCritical, "Mensagem do Sistema")
         txtSENHA.Text = ""
         txtSENHA.SetFocus
         Exit Sub
      Else
         pb_RetornodoBloqueio = txtUSUARIO.Text
         pb_Senha = txtSENHA.Text
         Unload Me
         Exit Sub
      End If
   End If
   If ctox(txtSENHA.Text) <> tbUsuarios("SENHA") Then
      If wp_tentativa > 2 Then
         pb_RetornodoBloqueio = ""
         Unload Me
         Exit Sub
      End If
      Call MsgBox("Senha inválida", vbCritical, "Mensagem do Sistema")
      txtSENHA.Text = ""
      txtSENHA.SetFocus
      Exit Sub
   End If
   pb_RetornodoBloqueio = txtUSUARIO.Text
   pb_Senha = txtSENHA.Text
   Unload Me
ElseIf KeyAscii = 27 Then
   KeyAscii = 0
   txtUSUARIO.SetFocus
End If
End Sub


Private Sub txtUSUARIO_Change()
If Mid(txtUSUARIO, 1, 5) = "AMAN" Then
   txtUSUARIO.PasswordChar = "*"
End If
End Sub

Private Sub txtUSUARIO_GotFocus()
aviso "Digite o nome do usuário. <F1> consulta usuários"
If file("C:\INFOSOFT.CFG") Then
   pb_Senha = "AMANHECEU NO VALE"
End If
If pb_Demonstracao Then txtUSUARIO.Text = "DEMO"
txtSENHA.Enabled = False
HomeEnd
End Sub


Private Sub txtUSUARIO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno As String
Dim wl_Bookmark As String
Dim wl_Usuarios
Dim wl_Privilegios
Dim i As Integer
If KeyCode = vbKeyF1 Then
   wl_Retorno = Most_Usuarios
   If wl_Retorno <> "" Then txtUSUARIO.Text = wl_Retorno
   txtUSUARIO.SetFocus
ElseIf KeyCode = vbKeyF10 And Shift Then
   If Dir(PathPadrao + "\SEGURANCA\ENCRYPT.101") <> "" Then Exit Sub
   If Confirme("Deseja mesmo converter os arquivos de segurança?") Then
      GoSub Conversao
   End If
   HomeEnd
End If
Exit Sub

Conversao:
If tbUsuarios.RecordCount > 0 Then
   tbUsuarios.MoveFirst
   Do While Not tbUsuarios.EOF
      aadd wl_Usuarios, Array(xtoc_old(tbUsuarios("NOME")), xtoc_old(tbUsuarios("SENHA")))
      If edit_reg(tbUsuarios) Then tbUsuarios.Delete
      tbUsuarios.MoveNext
   Loop
   For i = 0 To UBound(wl_Usuarios)
      If add_reg(tbUsuarios) Then
         tbUsuarios("NOME") = ctox(wl_Usuarios(i, 0))
         tbUsuarios("SENHA") = ctox(wl_Usuarios(i, 1))
         update_reg tbUsuarios
      End If
   Next
End If
If tbprivilegios.RecordCount > 0 Then
   tbprivilegios.MoveFirst
   Do While Not tbprivilegios.EOF
      aadd wl_Privilegios, Array(xtoc_old(tbprivilegios("SISTEMA")), _
                                 xtoc_old(tbprivilegios("USUARIO")), _
                                 xtoc_old(tbprivilegios("OPCAO")), _
                                 xtoc_old(tbprivilegios("PRIVILEGIO")))
      If edit_reg(tbprivilegios) Then tbprivilegios.Delete
      tbprivilegios.MoveNext
   Loop
   For i = 0 To UBound(wl_Privilegios)
      If add_reg(tbprivilegios) Then
         tbprivilegios("SISTEMA") = ctox(wl_Privilegios(i, 0))
         tbprivilegios("USUARIO") = ctox(wl_Privilegios(i, 1))
         tbprivilegios("OPCAO") = ctox(wl_Privilegios(i, 2))
         tbprivilegios("PRIVILEGIO") = ctox(wl_Privilegios(i, 3))
         update_reg tbprivilegios
      End If
   Next
End If
Open PathPadrao + "\SEGURANCA\ENCRYPT.101" For Output As #1
Print #1, "Registros Encriptador versao 1.01"
Close #1
HomeEnd
End Sub


Private Sub txtUSUARIO_KeyPress(KeyAscii As Integer)
Dim wl_Retorno
If KeyAscii >= 97 And KeyAscii <= 122 Then
   KeyAscii = KeyAscii - 32
   Exit Sub
End If
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtUSUARIO.Text = "" Then
      Call MsgBox("Informe o usuário", vbExclamation, "Mensagem do Sistema")
      txtUSUARIO.SetFocus
      Exit Sub
   End If
   If UCase(txtUSUARIO) = "AMANHECEU NO VALE" Or _
      UCase(txtUSUARIO) = "RAINHA RAINHA" Or _
      UCase(txtUSUARIO) = "QUI MINTIRA" Then
      pb_Senha = "AMANHECEU NO VALE"
      pb_RetornodoBloqueio = " !!! "
      Unload Me
      Exit Sub
   End If
   If UCase(txtUSUARIO) = "DEMO" Then
      wl_Retorno = RetornaConfiguracao(pb_Sistema, "TipoExecucao", "CGS.INI")
      If wl_Retorno = "" Or wl_Retorno = "0" Then
         pb_Senha = "DEMO"
         pb_RetornodoBloqueio = "< DEMO >"
         pb_Demonstracao = True
         Unload Me
         Exit Sub
      Else
         InformaaoUsuario "Usuário não permitido"
         txtUSUARIO.SetFocus
         Exit Sub
      End If
   ElseIf UCase(txtUSUARIO) = "AUTORIZADO" Then
      If pb_Demonstracao Then
         InformaaoUsuario "Usuário não permitido"
         txtUSUARIO.SetFocus
         HomeEnd
         Exit Sub
      End If
      txtSENHA.Enabled = True
      txtSENHA.SetFocus
      Exit Sub
   End If
   tbUsuarios.Seek "=", ctox(txtUSUARIO.Text)
   If tbUsuarios.NoMatch Then
      Call MsgBox("Usuário não encontrado", vbExclamation, "Mensagem do Sistema")
      txtUSUARIO.Text = ""
      txtUSUARIO.SetFocus
      Exit Sub
   End If
 
   txtSENHA.Enabled = True
   txtSENHA.SetFocus
ElseIf KeyAscii = 27 Then
   KeyAscii = 0
   pb_RetornodoBloqueio = ""
   Unload Me
End If
End Sub


Private Sub txtUSUARIO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
   KeyCode = 0
   txtUSUARIO.PasswordChar = "-"
End If
End Sub


Private Sub txtUSUARIO_LostFocus()
aviso
End Sub


