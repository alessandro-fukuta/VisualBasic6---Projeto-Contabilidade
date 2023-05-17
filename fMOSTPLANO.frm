VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMOSTPLANO 
   Appearance      =   0  'Flat
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Plano de Contas"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9690
   Icon            =   "fMOSTPLANO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox boxAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   1260
      ScaleHeight     =   690
      ScaleWidth      =   7185
      TabIndex        =   1
      Top             =   2760
      Width           =   7215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aguarde. Montando Plano de Contas ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   4050
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3285
      Top             =   2820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMOSTPLANO.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMOSTPLANO.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMOSTPLANO.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMOSTPLANO.frx":113E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreePlano 
      Height          =   6690
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   11800
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      ImageList       =   "ImageList1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "fMOSTPLANO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tbPlano As Recordset
Private Sub Monta_Grade()
Dim wl_Retorno
Dim wl_Aberto As Boolean
aviso "Aguarde. Montando Plano de Contas ..."
TreePlano.Nodes.Clear
TreePlano.Nodes.Add , , "PLANO", "Plano de Contas", 4
TreePlano.Nodes.Item("PLANO").Expanded = True
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "PlanodeContasAberto")
wl_Aberto = IIf(wl_Retorno = "", True, IIf(wl_Retorno = "1", True, False))
Do While Not tbPlano.EOF
   If Len(tbPlano("CONTA")) = 1 Then
      TreePlano.Nodes.Add "PLANO", tvwChild, "N" + tbPlano("CONTA"), tbPlano("CONTA") + " - " + tbPlano("DESCRICAO"), 1
      TreePlano.Nodes.Item("N" + tbPlano("CONTA")).Expanded = wl_Aberto
      TreePlano.Nodes.Item("N" + tbPlano("CONTA")).Bold = True
      Monta_SubNivel "N" + tbPlano("CONTA")
   ElseIf tbPlano("CONTA") = "" Then
      tbPlano.MoveNext
   Else
      tbPlano.MoveNext
   End If
Loop
aviso
End Sub

Private Sub Monta_SubNivel(pConta As String)
Dim wl_Nivel As String
wl_Nivel = tbPlano("CONTA")
tbPlano.MoveNext
Do While Not tbPlano.EOF
   If Len(tbPlano("CONTA")) = 1 Then
      Exit Sub
   End If
   If Mid(tbPlano("CONTA"), 1, Len(wl_Nivel)) <> wl_Nivel Or _
   Mid(tbPlano("CONTA"), Len(wl_Nivel) + 1, 1) <> "." Then
      Exit Sub
   End If
   If tbPlano("TRADUTOR") = 0 Then
      TreePlano.Nodes.Add pConta, tvwChild, "N" + tbPlano("CONTA"), tbPlano("CONTA") + " - " + tbPlano("DESCRICAO"), 2
      TreePlano.Nodes.Item("N" + tbPlano("CONTA")).Expanded = True
   Else
      TreePlano.Nodes.Add pConta, tvwChild, "N" + tbPlano("CONTA"), tbPlano("CONTA") + "  [ " + Format(tbPlano("TRADUTOR"), "00000") + " ] - " + tbPlano("DESCRICAO"), 3
      TreePlano.Nodes.Item("N" + tbPlano("CONTA")).ForeColor = AZUL
   End If
   tbPlano.MoveNext
   If Not tbPlano.EOF Then
      If (Mid(tbPlano("CONTA"), 1, Len(wl_Nivel)) = wl_Nivel And Len(tbPlano("CONTA")) > Len(wl_Nivel)) Or _
      Mid(tbPlano("CONTA"), Len(wl_Nivel) + 1, 1) <> "." Then
         tbPlano.MovePrevious
         Monta_SubNivel "N" + tbPlano("CONTA")
      End If
   End If
Loop
End Sub

Private Sub Form_Activate()
Dim wl_Linha As Currency
Dim i As Integer
Dim wl_ContaPonto As Integer
Dim wl_StringNivel As String
If Not Abre_PlanoContas(tbPlano) Then
   Unload Me
   Exit Sub
End If
tbPlano.Index = "iCONTA"
If tbPlano.RecordCount = 0 Then
   InformaaoUsuario "Sem registros"
   Unload Me
   Exit Sub
End If
Me.MousePointer = 11
DoEvents
tbPlano.MoveFirst
Monta_Grade
Me.MousePointer = 0
Me.boxAguarde.Visible = False
DoEvents
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   pbRetornoVideo = ""
   Unload Me
   Exit Sub
End If
End Sub


Private Sub Form_Load()
pb_FormAtivo = VB.Screen.ActiveForm.Name
pb_ObjetoAtivo = VB.Screen.ActiveForm.ActiveControl.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbPlano.Close
End Sub


Private Sub MSFlexGrid1_Click()

End Sub


Private Sub grdPlano_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_String As String
Dim i As Integer
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_String = InputBox("O que pesquisar?", "Pesquisa Plano de Contas")
   If wl_String = "" Then
      Exit Sub
   End If
   For i = 1 To grdPLANO.rows - 1
      If grdPLANO.TextMatrix(i, 1) <> "" Then
         If InStr(UCase(grdPLANO.TextMatrix(i, 2)), UCase(wl_String)) > 0 Then
            grdPLANO.row = i
            SendKeys "{LEFT}"
            Exit For
         End If
      End If
   Next
End If
End Sub

Private Sub grdPlano_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   pbRetornoVideo = 0
   Unload Me
End If
If KeyAscii = 13 Then
   KeyAscii = 0
   If grdPLANO.TextMatrix(grdPLANO.row, 1) = "" Then
      InformaaoUsuario "Não é possível selecionar essa conta"
      grdPLANO.SetFocus
      Exit Sub
   End If
   pbRetornoVideo = grdPLANO.TextMatrix(grdPLANO.row, 1)
   Unload Me
End If
End Sub


Private Sub Timer1_Timer()
End Sub

Private Sub TreePlano_DblClick()
InformaaoUsuario "Use <ENTER>"
End Sub

Private Sub TreePlano_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim j As Integer
Dim wl_Pesquisa As String
Dim wl_Ok As Boolean
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Pesquisa = InputBox("O que pesquisar?", "Pesquisa Plano de Contas")
   If wl_Pesquisa = "" Then Exit Sub
   For i = 1 To TreePlano.Nodes.Count
      wl_Ok = False
      For j = 1 To Len(TreePlano.Nodes.Item(i).Text)
         If UCase(wl_Pesquisa) = UCase(Mid(TreePlano.Nodes.Item(i).Text, j, Len(wl_Pesquisa))) Then
            TreePlano.Nodes.Item(i).Selected = True
            wl_Ok = True
            Exit For
         End If
      Next
      If wl_Ok Then Exit For
   Next
End If
End Sub

Private Sub TreePlano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If InStr(TreePlano.SelectedItem.Text, "[") = 0 Then
      Exit Sub
   End If
   pbRetornoVideo = Mid(TreePlano.SelectedItem.Text, InStr(TreePlano.SelectedItem.Text, "[") + 2, 5)
   Unload Me
End If
End Sub


