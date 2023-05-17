VERSION 5.00
Object = "{5756E734-2046-400A-BC65-0E105EC5876E}#1.0#0"; "CAIXATEX.OCX"
Begin VB.Form fSEGURANCA 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Segurança"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   ControlBox      =   0   'False
   HelpContextID   =   320
   Icon            =   "fSEGURANCA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSupervisor 
      Caption         =   "&Apenas Nível 3"
      Height          =   255
      Left            =   8355
      TabIndex        =   36
      Top             =   345
      Width           =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Usuários"
      Height          =   405
      Left            =   6825
      TabIndex        =   20
      Top             =   270
      Width           =   1380
   End
   Begin VB.Frame box_Privilegios 
      BackColor       =   &H8000000A&
      Caption         =   "Privilégios"
      Height          =   4245
      Left            =   3120
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   7950
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   29
         Left            =   4005
         TabIndex        =   35
         Top             =   3855
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   28
         Left            =   4005
         TabIndex        =   34
         Top             =   3600
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   27
         Left            =   4005
         TabIndex        =   33
         Top             =   3345
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   26
         Left            =   4005
         TabIndex        =   32
         Top             =   3090
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   25
         Left            =   4005
         TabIndex        =   31
         Top             =   2835
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   24
         Left            =   4005
         TabIndex        =   30
         Top             =   2580
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   23
         Left            =   4005
         TabIndex        =   29
         Top             =   2325
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   22
         Left            =   4005
         TabIndex        =   28
         Top             =   2070
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   21
         Left            =   4005
         TabIndex        =   27
         Top             =   1815
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   20
         Left            =   4005
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   19
         Left            =   4005
         TabIndex        =   25
         Top             =   1305
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   18
         Left            =   4005
         TabIndex        =   24
         Top             =   1050
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   17
         Left            =   4005
         TabIndex        =   23
         Top             =   795
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   16
         Left            =   4005
         TabIndex        =   22
         Top             =   540
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   15
         Left            =   4005
         TabIndex        =   21
         Top             =   285
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   4
         Top             =   525
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   780
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   3
         Left            =   135
         TabIndex        =   6
         Top             =   1035
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   4
         Left            =   135
         TabIndex        =   7
         Top             =   1290
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   5
         Left            =   135
         TabIndex        =   8
         Top             =   1545
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   6
         Left            =   135
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   7
         Left            =   135
         TabIndex        =   10
         Top             =   2055
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   8
         Left            =   135
         TabIndex        =   11
         Top             =   2310
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   9
         Left            =   135
         TabIndex        =   12
         Top             =   2565
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   10
         Left            =   135
         TabIndex        =   13
         Top             =   2820
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   11
         Left            =   135
         TabIndex        =   14
         Top             =   3075
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   12
         Left            =   135
         TabIndex        =   15
         Top             =   3330
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   13
         Left            =   135
         TabIndex        =   16
         Top             =   3585
         Visible         =   0   'False
         Width           =   3550
      End
      Begin VB.CheckBox chkPRIVILEGIO 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   315
         Index           =   14
         Left            =   135
         TabIndex        =   17
         Top             =   3840
         Visible         =   0   'False
         Width           =   3550
      End
   End
   Begin VB.ListBox lstOPCOES 
      Height          =   4155
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   930
      Visible         =   0   'False
      Width           =   2820
   End
   Begin CaixaTexto.Caixa_Texto txtUSUARIO 
      Height          =   360
      Left            =   150
      TabIndex        =   1
      Top             =   285
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   635
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
   Begin VB.Label Legenda_Lista 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opções"
      Height          =   195
      Left            =   150
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   540
   End
End
Attribute VB_Name = "fSEGURANCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tbUsuarios As Recordset
Private tbprivilegios As Recordset
Private wp_opcao As Integer
Private wp_Entrada As Boolean
Private wp_Usuario As String
Private wp_Senha As String

Private Function Carrega_Privilegios(pPonteiro)
Dim i As Integer
Dim wl_Char As String
On Error Resume Next
For i = 0 To 29
   chkPRIVILEGIO(i).Caption = pb_Privilegios(pPonteiro, i)
   If Err <> 0 Or pb_Privilegios(pPonteiro, i) = "" Then
      Exit For
   End If
   chkPRIVILEGIO(i).Value = 0
   wl_Char = UCase(Mid(pb_Privilegios(pPonteiro, i), AT("&", pb_Privilegios(pPonteiro, i)) + 1, 1))
   If pb_Opcoes(pPonteiro, 1) = 2 Then
      tbprivilegios.Seek "=", ctox(pb_Sistema), "", ctox(Format(wp_opcao, "000"))
   Else
      tbprivilegios.Seek "=", ctox(pb_Sistema), ctox(txtUSUARIO.Text), ctox(Format(wp_opcao, "000"))
   End If
   If Not tbprivilegios.NoMatch Then
      If AT(wl_Char, xtoc(tbprivilegios("PRIVILEGIO"))) > 0 Then
         chkPRIVILEGIO(i).Value = 1
      End If
   End If
   chkPRIVILEGIO(i).Visible = True
Next
box_Privilegios.Visible = True
End Function



Private Function Monta_Lista()
On Error Resume Next
Dim i As Integer
Dim z
lstOPCOES.Clear
DoEvents
Call Carrega_Opcoes
For i = 0 To UBound(pb_Opcoes)
   If Not IsEmpty(pb_Opcoes(i, 1)) Then
      If pb_Opcoes(i, 1) = 1 Then
         If chkSupervisor.Value = 0 And txtUSUARIO.Text <> "" Then
            If Verifica_Privilegio(PR_SUPERVISOR, "S") Then
               lstOPCOES.AddItem UCase(pb_Opcoes(i, 0))
            End If
         End If
      ElseIf pb_Opcoes(i, 1) <> 0 Then
         If pb_Opcoes(i, 1) = 2 Or pb_Opcoes(i, 1) = 3 Then
            If pb_Senha = "AMANHECEU NO VALE" Then
               lstOPCOES.AddItem UCase(pb_Opcoes(i, 0))
            End If
         Else
            lstOPCOES.AddItem UCase(pb_Opcoes(i, 0))
         End If
      End If
   Else
      If chkSupervisor.Value = 0 And txtUSUARIO.Text <> "" Then
         lstOPCOES.AddItem UCase(pb_Opcoes(i, 0))
      End If
   End If
Next
End Function


Function Ver_Privilegio(pOpcao, pPrivilegio As String)

End Function

Private Sub chkPRIVILEGIO_Click(Index As Integer)
Dim i As Integer
Dim wl_String As String
If Not chkPRIVILEGIO(Index).Visible Then
   Exit Sub
End If
wl_String = ""
For i = 0 To 29
   If chkPRIVILEGIO(i).Visible = False Then
      Exit For
   End If
   If chkPRIVILEGIO(i).Value = 1 Then
      wl_String = wl_String + UCase(Mid(pb_Privilegios(wp_opcao, i), AT("&", pb_Privilegios(wp_opcao, i)) + 1, 1))
   End If
Next
If pb_Opcoes(wp_opcao, 1) = 2 Then
   tbprivilegios.Seek "=", ctox(pb_Sistema), "", ctox(Format(wp_opcao, "000"))
Else
   tbprivilegios.Seek "=", ctox(pb_Sistema), ctox(txtUSUARIO.Text), ctox(Format(wp_opcao, "000"))
End If
If tbprivilegios.NoMatch Then
   If Not add_reg(tbprivilegios) Then
      lstOPCOES.SetFocus
      Exit Sub
   End If
   tbprivilegios("SISTEMA") = ctox(pb_Sistema)
   tbprivilegios("USUARIO") = IIf(pb_Opcoes(wp_opcao, 1) = 2, "", ctox(txtUSUARIO.Text))
   tbprivilegios("OPCAO") = ctox(Format(wp_opcao, "000"))
Else
   If Not edit_reg(tbprivilegios) Then
      lstOPCOES.SetFocus
      Exit Sub
   End If
End If
tbprivilegios("PRIVILEGIO") = ctox(wl_String)
If Not update_reg(tbprivilegios) Then
   lstOPCOES.SetFocus
   Exit Sub
End If
End Sub

Private Sub chkPRIVILEGIO_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 27 Then
   KeyAscii = 0
   lstOPCOES.SetFocus
   For i = 0 To 29
      chkPRIVILEGIO(i).Visible = False
   Next
   box_Privilegios.Visible = False
End If
End Sub


Private Sub chkSupervisor_Click()
Monta_Lista
lstOPCOES.SetFocus
End Sub

Private Sub Command1_Click()
If Verifica_Privilegio(PR_USUARIOS, "C", "Sem privilégio para acessar usuários") Then
   fUSUARIO.Show 1
End If
txtUSUARIO.SetFocus
End Sub

Private Sub Form_Activate()
Dim wl_Usuario As String
Dim wl_Senha As String
If Not wp_Entrada Then
   wp_Usuario = pb_Usuario
   wp_Senha = pb_Senha
   wp_Entrada = True
   wl_Usuario = pb_Usuario
   wl_Senha = pb_Senha
   If Not Permissao(PR_SEGURANCA, "C", "Identifique-se") Then
      pb_Usuario = wl_Usuario
      pb_Senha = wl_Senha
      Unload Me
      Exit Sub
   End If
End If
If pb_Demonstracao Then
   InformaaoUsuario "Não é permitido um usuário DEMO acessar o nível de segurança"
   Unload Me
End If
If Not Abre_Usuarios(tbUsuarios) Or _
   Not Abre_Privilegios(tbprivilegios) Then
   Unload Me
End If
wp_opcao = 0
Command1.Enabled = Verifica_Privilegio(PR_SUPERVISOR, "S")
End Sub

Private Sub Form_Load()
wp_Entrada = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbUsuarios.Close
tbprivilegios.Close
dbSeguranca.Close
aviso
pb_Usuario = wp_Usuario
pb_Senha = wp_Senha
Display_Usuario
End Sub


Private Sub lstOPCOES_DblClick()
SendKeys "{ENTER}"
End Sub

Private Sub lstOPCOES_GotFocus()
On Error Resume Next
If lstOPCOES.ListIndex = -1 Then
   lstOPCOES.ListIndex = 0
End If
If box_Privilegios.Visible Then
   Call chkPRIVILEGIO_KeyPress(0, 27)
End If
End Sub

Private Sub lstOPCOES_KeyPress(KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 13 Then
   KeyAscii = 0
   For i = 0 To UBound(pb_Opcoes)
      If UCase(pb_Opcoes(i, 0)) = lstOPCOES.List(lstOPCOES.ListIndex) Then
         wp_opcao = i
         Call Carrega_Privilegios(i)
         If chkPRIVILEGIO(0).Visible Then
            chkPRIVILEGIO(0).SetFocus
         End If
         Exit For
      End If
   Next
ElseIf KeyAscii = 27 Then
   KeyAscii = 0
   txtUSUARIO.SetFocus
End If
End Sub


Private Sub txtUSUARIO_GotFocus()
Dim i As Integer
aviso "<F1> Usuários"
chkSupervisor.Enabled = False
If Not ShowRetorno Then
   txtUSUARIO.Text = ""
Else
   ShowRetorno = False
End If
lstOPCOES.Visible = False
Legenda_Lista.Visible = False
For i = 0 To 29
   chkPRIVILEGIO(i).Visible = False
Next
box_Privilegios.Visible = False
HomeEnd
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
      If pb_Senha = "AMANHECEU NO VALE" Then
         Call Monta_Lista
         chkSupervisor.Enabled = True
         Legenda_Lista.Visible = True
         lstOPCOES.Visible = True
         lstOPCOES.SetFocus
         Exit Sub
      Else
         InformaaoUsuario "Informe corretamente o usuário"
         txtUSUARIO.SetFocus
         Exit Sub
      End If
   End If
   tbUsuarios.Seek "=", ctox(txtUSUARIO.Text)
   If tbUsuarios.NoMatch Then
      Call MsgBox("Usuário não encontrado", vbExclamation, "Mensagem do Sistema")
      txtUSUARIO.Text = ""
      txtUSUARIO.SetFocus
      Exit Sub
   End If
   Call Monta_Lista
   chkSupervisor.Enabled = True
   Legenda_Lista.Visible = True
   lstOPCOES.Visible = True
   lstOPCOES.SetFocus
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub


Private Sub txtUSUARIO_LostFocus()
aviso
End Sub


