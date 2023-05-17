VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVideo 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6300
   ClientLeft      =   2370
   ClientTop       =   1650
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   280
   Icon            =   "Frmvideo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Frmvideo.frx":0442
   ScaleHeight     =   6300
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin TrueDBGrid50.TDBGrid TDBGrid1 
      Bindings        =   "Frmvideo.frx":0884
      Height          =   4920
      Left            =   90
      OleObjectBlob   =   "Frmvideo.frx":0898
      TabIndex        =   0
      Top             =   915
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.PictureBox box_Filtro 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   30
      ScaleHeight     =   795
      ScaleWidth      =   7455
      TabIndex        =   7
      Top             =   45
      Visible         =   0   'False
      Width           =   7485
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Exatamente"
         Height          =   210
         Left            =   5415
         TabIndex        =   15
         Top             =   480
         Width           =   1485
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Início do Campo"
         Height          =   210
         Left            =   3840
         TabIndex        =   14
         Top             =   465
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Todo o campo"
         Height          =   210
         Left            =   2370
         TabIndex        =   13
         Top             =   465
         Width           =   1440
      End
      Begin VB.CommandButton cmdFILTRO 
         Caption         =   "&Filtro"
         Height          =   285
         Left            =   6660
         TabIndex        =   12
         Top             =   150
         Width           =   675
      End
      Begin VB.TextBox txtFILTRO 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   2385
         TabIndex        =   10
         Top             =   150
         Width           =   4260
      End
      Begin VB.ComboBox cmbCampo 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   150
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro à aplicar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2385
         TabIndex        =   11
         Top             =   0
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione o Campo para filtro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   60
         TabIndex        =   9
         Top             =   0
         Width           =   1860
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4935
      Top             =   2220
   End
   Begin VB.PictureBox SubHelp 
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      Height          =   3240
      Left            =   1260
      ScaleHeight     =   3180
      ScaleWidth      =   6165
      TabIndex        =   3
      Top             =   2595
      Visible         =   0   'False
      Width           =   6225
      Begin MSFlexGridLib.MSFlexGrid GrdSubHelp 
         Height          =   3240
         Left            =   -30
         TabIndex        =   4
         Top             =   -15
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   5715
         _Version        =   393216
         Rows            =   100
         FixedCols       =   0
         BackColor       =   8454143
         BackColorFixed  =   128
         ForeColorFixed  =   16777215
         BackColorBkg    =   16777215
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         GridLinesFixed  =   1
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "Legenda                      | <Conteúdo                                                                                 "
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4815
      Top             =   4830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmvideo.frx":2277
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmvideo.frx":26CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmvideo.frx":2B1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmvideo.frx":2F73
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmvideo.frx":33C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmvideo.frx":381B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmvideo.frx":3C6F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   -1  'True
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5025
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   " <F3> Filtra Conteúdo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2985
      TabIndex        =   6
      Top             =   5985
      Width           =   2055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   " <F2> Detalhes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1545
      TabIndex        =   5
      Top             =   5985
      Width           =   1410
   End
   Begin VB.Label lblTITULO 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   75
      TabIndex        =   2
      Top             =   270
      Width           =   7350
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   " <F1> Consulta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   105
      TabIndex        =   1
      Top             =   5985
      Width           =   1410
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private wp_opcao As Integer
Private wp_Top As Long
Private wp_Left As Long
Private wp_SubHelp As Boolean

Private Function Atu_Opcao()
Shape1.Visible = True
Shape4.Visible = True
Shape3.Visible = True
Shape6.Visible = True
Shape5.Visible = True
If wp_opcao = 1 And IndicedaConsulta <> "" Then
   Shape1.Visible = False
ElseIf wp_opcao = 2 Then
   Shape4.Visible = False
ElseIf wp_opcao = 3 Then
   Shape3.Visible = False
ElseIf wp_opcao = 4 Then
   Shape6.Visible = False
ElseIf wp_opcao = 5 Then
   Shape5.Visible = False
End If
End Function


Private Function CentraForm()
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2
End Function

Private Function Efeito()
Dim i As Long
Dim X As Variant
For i = 0 To 100
   Line (X, 200)-(X + 105, 550), RGB(i * 3, i * 3, 130 - i), BF
   X = X + (Me.Width / 2) / 100
Next
X = Me.Width / 2
For i = 100 To 0 Step -1
   Line (X, 200)-(X + 105, 550), RGB(i * 3, i * 3, 130 - i), BF
   X = X + (Me.Width / 2) / 100
Next
Me.Visible = True
End Function

Private Function Inc_Abertura()
Dim i As Long
Dim l As Long
l = Label2.Left
For i = l To 1620 Step -1
   Label2.Left = i
   DoEvents
Next
End Function

Private Sub MenuLocal()
End Sub

Function Redesenha()
Dim Coluna As TrueDBGrid50.Column
Dim i As Integer
Dim tamanho As Integer
If IndicedaConsulta <> "" Then
   Data1.Recordset.Index = IndicedaConsulta
End If
Do While TDBGrid1.Columns.Count <> 0
   TDBGrid1.Columns.Remove 0
Loop
For i = 0 To UBound(a_Browse001)
   tamanho = tamanho + a_Browse001(i, 2)
   Set Coluna = TDBGrid1.Columns.Add(i)
   With Coluna
      .Visible = True
      .DataField = a_Browse001(i, 1)
      .Caption = a_Browse001(i, 0)
      .Width = a_Browse001(i, 2)
      .NumberFormat = a_Browse001(i, 3)
      .Alignment = a_Browse001(i, 4)
   End With
   Me.cmbCampo.AddItem Format(i, "000") + " - " + UCase(a_Browse001(i, 0))
Next
Me.cmbCampo.Text = "000 - " + UCase(a_Browse001(0, 0))
If tamanho > Screen.Width - 300 Then
   tamanho = Screen.Width - 300
End If
TDBGrid1.Width = tamanho - 90
Width = tamanho + 200
lblTitulo.Width = Width - 300
TDBGrid1.Visible = True
Call CentraForm

End Function


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub


Sub Refresh_SubHelp()
On Error Resume Next
Dim i As Integer
Dim wl_Coluna1 As Long
If wp_SubHelp Then GoSub Monta_SubHelp
Exit Sub

Monta_SubHelp:
For i = 0 To UBound(a_Browse001)
   If a_Browse001(i, 2) > GrdSubHelp.ColWidth(1) Then GrdSubHelp.ColWidth(1) = a_Browse001(i, 2)
   If IsNumeric(Data1.Recordset(a_Browse001(i, 1))) Then
      Me.GrdSubHelp.TextMatrix(i + 1, 0) = a_Browse001(i, 0)
      Me.GrdSubHelp.TextMatrix(i + 1, 1) = Format(Data1.Recordset(a_Browse001(i, 1)), a_Browse001(i, 3))
   ElseIf IsDate(Data1.Recordset(a_Browse001(i, 1))) Then
      Me.GrdSubHelp.TextMatrix(i + 1, 0) = a_Browse001(i, 0)
      Me.GrdSubHelp.TextMatrix(i + 1, 1) = CStr(Data1.Recordset(a_Browse001(i, 1)))
   Else
      Me.GrdSubHelp.TextMatrix(i + 1, 0) = a_Browse001(i, 0)
      Me.GrdSubHelp.TextMatrix(i + 1, 1) = Data1.Recordset(a_Browse001(i, 1))
   End If
   Me.GrdSubHelp.row = i + 1
   Me.GrdSubHelp.Col = 1
   Me.GrdSubHelp.CellFontBold = True
   Me.GrdSubHelp.row = 0
Next
End Sub

Private Sub cmdFILTRO_Click()
Dim wl_Campo
Dim wl_Index As Integer
Dim wl_Tabela As String
Dim wl_Campos As String
On Error Resume Next
wl_Index = Mid(Me.cmbCampo.Text, 1, 3)
If Me.txtFILTRO.Text = "" Then
   InformaaoUsuario "É necessário informar o conteúdo do filtro"
   txtFILTRO.SetFocus
   Exit Sub
End If
If InStr(UCase(Data1.RecordSource), "FROM") = 0 Then
   wl_Tabela = Data1.RecordSource
Else
   wl_Tabela = ""
   For i = InStr(UCase(Data1.RecordSource), "FROM") + 5 To Len(Data1.RecordSource)
      If Mid(Data1.RecordSource, i, 1) = " " Then
         If i = Len(Data1.RecordSource) Then
            Exit For
         Else
            If Mid(Data1.RecordSource, i + 1, 1) <> "," Then Exit For
         End If
      End If
      wl_Tabela = wl_Tabela + Mid(Data1.RecordSource, i, 1)
   Next
End If
Data1.RecordsetType = "1"
Data1.Refresh
wl_Campos = "*"
Data1.RecordSource = "SELECT " + wl_Campos + " FROM " + wl_Tabela + " WHERE " + a_Browse001(wl_Index, 1) + " LIKE '" + IIf(Option1.Value, "*", "") + Me.txtFILTRO.Text + IIf(Option2.Value, "*", IIf(Option1.Value, "*", "")) + "' ORDER BY " + a_Browse001(wl_Index, 1)
Data1.Refresh
If Err <> 0 Then
   Data1.RecordSource = "SELECT " + wl_Campos + " FROM " + wl_Tabela + " WHERE " + a_Browse001(wl_Index, 1) + " LIKE " + Me.txtFILTRO.Text + "' ORDER BY " + a_Browse001(wl_Index, 1)
   Data1.Refresh
End If
For i = 0 To UBound(a_Browse001)
   Me.TDBGrid1.Columns(i).Width = a_Browse001(i, 2)
Next
Me.TDBGrid1.SetFocus
Me.box_Filtro.Visible = False
Me.SubHelp.Visible = False
wp_SubHelp = False
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
MsgBox "Erro ao abrir tabela " + TabeladaConsulta, vbCritical, "Erro :" + Str(DataErr)
End Sub

Private Sub Form_Activate()
Dim i As Integer
On Error Resume Next
If Err <> 0 Then
   Unload Me
   Exit Sub
End If
lblTitulo.Visible = False
wp_Top = Top
wp_Left = Left
If IndicedaConsulta = "" Then
   Timer1.Enabled = False
   Label1.Visible = False
End If
Me.Visible = True
lblTitulo.Visible = True
SubHelp.ZOrder 0
SubHelp.Height = 0
For i = 0 To UBound(a_Browse001)
   If GrdSubHelp.ColWidth(0) + a_Browse001(i, 2) > SubHelp.Width Then
      SubHelp.Width = GrdSubHelp.ColWidth(0) + a_Browse001(i, 2)
   End If
   SubHelp.Height = SubHelp.Height + GrdSubHelp.RowHeight(i)
Next
SubHelp.Height = SubHelp.Height + GrdSubHelp.RowHeight(i)
GrdSubHelp.Height = SubHelp.Height
GrdSubHelp.Width = SubHelp.Width
SubHelp.Left = Width - SubHelp.Width - 200
SubHelp.Top = Me.TDBGrid1.Top + TDBGrid1.Height - SubHelp.Height
Me.box_Filtro.Left = (Me.Width / 2) - (Me.box_Filtro.Width / 2)
wp_SubHelp = False
'Timer1.Enabled = True
If Width = 90 Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim c As Long
Dim l As Long
If KeyCode = vbKeyF2 Then
   KeyCode = 0
   wp_SubHelp = Not wp_SubHelp
   If wp_SubHelp Then Refresh_SubHelp
   SubHelp.Visible = wp_SubHelp
ElseIf KeyCode = vbKeyF3 Then
   KeyCode = 0
   If Not box_Filtro.Visible Then
      box_Filtro.Top = 0 - box_Filtro.Height
      box_Filtro.Visible = True
      For i = box_Filtro.Top To 45 Step 70
         box_Filtro.Top = i
         DoEvents
      Next
      box_Filtro.Top = 45
      txtFILTRO.SetFocus
      HomeEnd
   Else
      For i = 45 To 0 - box_Filtro.Height Step -70
         box_Filtro.Top = i
         DoEvents
      Next
      Me.box_Filtro.Visible = False
   End If
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   If Me.box_Filtro.Visible Then
      Me.TDBGrid1.SetFocus
      Me.box_Filtro.Visible = False
   End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
lblTitulo.Caption = pb_TitulodaConsulta
Width = 0
If IndicedaConsulta <> "" Then
   Data1.RecordsetType = 0
End If
Data1.DatabaseName = BancodeDadosdaConsulta
Data1.RecordSource = TabeladaConsulta
Data1.Refresh
DoEvents
Call Redesenha
Me.Visible = False
End Sub


Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
End Sub


Private Sub Form_Paint()
Call Efeito
End Sub

Private Sub Image1_Click()
Call TDBGrid1_KeyDown(vbKeyF1, False)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IndicedaConsulta = "" Then
   Exit Sub
End If
If Not Shape1.Visible Then
   Exit Sub
End If
wp_opcao = 0
Call Atu_Opcao
If Not Shape4.Visible Then
   Shape4.Visible = True
End If
If Shape1.Visible Then
   Shape1.Visible = False
End If
End Sub


Private Sub Image2_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
   Data1.Recordset.MoveLast
End If
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Shape3.Visible Then
   Exit Sub
End If
wp_opcao = 0
Call Atu_Opcao
If Not Shape4.Visible Or Not Shape6.Visible Then
   Shape4.Visible = True
   Shape6.Visible = True
End If
If Shape3.Visible Then
   Shape3.Visible = False
End If
End Sub


Private Sub Image3_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF Then
   Data1.Recordset.MoveFirst
End If
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Shape4.Visible Then
   Exit Sub
End If
wp_opcao = 0
Call Atu_Opcao
If Not Shape3.Visible Or Not Shape1.Visible Then
   Shape3.Visible = True
   Shape1.Visible = True
End If
Shape4.Visible = False
End Sub


Private Sub Image4_Click()
Data1.Recordset.MoveLast
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Shape5.Visible Then
   Exit Sub
End If
wp_opcao = 0
Call Atu_Opcao
If Not Shape6.Visible Then
   Shape6.Visible = True
End If
Shape5.Visible = False
End Sub


Private Sub Image5_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Shape6.Visible Then
   Exit Sub
End If
wp_opcao = 0
Call Atu_Opcao
If Not Shape3.Visible Or Not Shape5.Visible Then
   Shape3.Visible = False
   Shape5.Visible = False
End If
Shape6.Visible = False
End Sub


Private Sub GrdSubHelp_GotFocus()
Me.TDBGrid1.SetFocus
SendKeys "{F2}"
End Sub


Private Sub TDBGrid1_DblClick()
SendKeys "{enter}"
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim wl_Retorno
Dim wl_RetornoVideo
On Error Resume Next
Err = 0
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   PESQUISA = InputBox("Informe na caixa abaixo a informação a pesquisar.", "O QUE PESQUISAR ?")
   If PESQUISA = "" Then TDBGrid1.SetFocus: Exit Sub
   Timer1.Enabled = True
   Label1.Visible = True
   If IndicedaConsulta <> "" Then
      Data1.Recordset.Seek ">=", PESQUISA
   Else
      If pb_CampodePesquisa <> "" Then
         Data1.Recordset.FindFirst pb_CampodePesquisa + " >= '" + (PESQUISA) + "'"
      End If
   End If
   TDBGrid1.Refresh
   TDBGrid1.SetFocus
   If Err <> 0 Then
      Call MsgBox("Erro ao pesquisar !", vbExclamation, "Mensagem do Sistema")
      Err = 0
   End If
   TDBGrid1.SetFocus
   SendKeys "{RIGHT}"
   Label1 = "<F1> Consulta / <F2> Detalhes"
   DoEvents
ElseIf KeyCode = 13 Then
   KeyCode = 0
   If Not IsArray(pbCampodeRetorno) Then
      If pbCampodeRetorno = "" Then
         Unload Me
         Exit Sub
      End If
      pbRetornoVideo = Data1.Recordset(pbCampodeRetorno)
   Else
      For i = 0 To UBound(pbCampodeRetorno)
         wl_Retorno = Data1.Recordset(pbCampodeRetorno(i))
         aadd wl_RetornoVideo, wl_Retorno
      Next
      pbRetornoVideo = wl_RetornoVideo
   End If
   Unload Me
   Exit Sub
ElseIf KeyCode = 27 Then
   KeyCode = 0
   If Not IsArray(pbCampodeRetorno) Then
      pbRetornoVideo = ""
   Else
      For i = 0 To UBound(pbCampodeRetorno)
         aadd wl_RetornoVideo, ""
      Next
      pbRetornoVideo = wl_RetornoVideo
   End If
   Unload Me
ElseIf KeyCode = vbKeyEnd Then
   KeyCode = 0
   Data1.Recordset.MoveLast
ElseIf KeyCode = vbKeyHome Then
   KeyCode = 0
   Data1.Recordset.MoveFirst
   Data1.Recordset.MoveFirst
End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Refresh_SubHelp
End Sub

Private Sub Timer2_Timer()

End Sub


Private Sub txtFILTRO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   cmdFILTRO_Click
End If
End Sub


