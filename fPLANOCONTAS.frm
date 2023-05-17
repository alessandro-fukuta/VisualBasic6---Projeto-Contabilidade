VERSION 5.00
Object = "{5756E734-2046-400A-BC65-0E105EC5876E}#1.0#0"; "CAIXATEX.OCX"
Object = "{BA676A3D-9505-4A77-87DC-76025E082864}#1.0#0"; "ETIQUETA.OCX"
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fPLANOCONTAS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plano de Contas"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "fPLANOCONTAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbNATUREZA 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "fPLANOCONTAS.frx":0442
      Left            =   2535
      List            =   "fPLANOCONTAS.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   255
      Width           =   2415
   End
   Begin Etiq.Etiqueta lblGRAU 
      Height          =   300
      Left            =   2055
      TabIndex        =   16
      Top             =   270
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      BackColor       =   -2147483624
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483625
   End
   Begin Mascara.Máscara txtNIVEL1 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   270
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "#"
      Text            =   ""
      ÉValor          =   -1  'True
   End
   Begin Mascara.Máscara txtNIVEL2 
      Height          =   300
      Left            =   270
      TabIndex        =   1
      Top             =   270
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "##"
      Text            =   ""
      ÉValor          =   -1  'True
   End
   Begin Mascara.Máscara txtNIVEL3 
      Height          =   300
      Left            =   585
      TabIndex        =   2
      Top             =   270
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "##"
      Text            =   ""
      ÉValor          =   -1  'True
   End
   Begin Mascara.Máscara txtNIVEL4 
      Height          =   300
      Left            =   900
      TabIndex        =   3
      Top             =   270
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "##"
      Text            =   ""
      ÉValor          =   -1  'True
   End
   Begin Mascara.Máscara txtCONTA 
      Height          =   300
      Left            =   1245
      TabIndex        =   4
      Top             =   270
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      BackColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "#####"
      Text            =   ""
      ÉValor          =   -1  'True
   End
   Begin VB.Frame boxDADOS 
      Enabled         =   0   'False
      Height          =   4275
      Left            =   45
      TabIndex        =   18
      Top             =   480
      Width           =   4890
      Begin Mascara.Máscara txtPROPORCAO 
         Height          =   300
         Left            =   3615
         TabIndex        =   11
         Top             =   3450
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00"
         Text            =   ""
         ÉValor          =   -1  'True
      End
      Begin Etiq.Etiqueta lblRATEIO 
         Height          =   300
         Left            =   645
         TabIndex        =   25
         Top             =   3450
         Visible         =   0   'False
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   529
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483625
      End
      Begin Mascara.Máscara txtCONTARATEIO 
         Height          =   300
         Left            =   75
         TabIndex        =   10
         Top             =   3450
         Visible         =   0   'False
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "#####"
         Text            =   ""
         ÉValor          =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid grdRATEIO 
         Height          =   1440
         Left            =   75
         TabIndex        =   9
         Top             =   1995
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   2540
         _Version        =   393216
         Rows            =   50
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "Conta | Descrição                                           | Proporção %"
      End
      Begin VB.CommandButton cmdGRAVA 
         BackColor       =   &H00C0C000&
         Caption         =   "&Grava"
         Height          =   405
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3795
         Width           =   1005
      End
      Begin VB.CommandButton cmdDELETA 
         BackColor       =   &H00C0C000&
         Caption         =   "&Deleta"
         Height          =   405
         Left            =   1245
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3795
         Width           =   1005
      End
      Begin CaixaTexto.Caixa_Texto txtDESCRICAO 
         Height          =   300
         Left            =   75
         TabIndex        =   6
         Top             =   360
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   529
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
      Begin Mascara.Máscara txtSALDOABERTURA 
         Height          =   300
         Left            =   75
         TabIndex        =   7
         Top             =   885
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,###,##0.00"
         Text            =   "_____"
         ÉValor          =   -1  'True
      End
      Begin Etiq.Etiqueta lblCREDITO 
         Height          =   300
         Left            =   1830
         TabIndex        =   19
         Top             =   885
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483625
      End
      Begin Etiq.Etiqueta lblDEBITO 
         Height          =   300
         Left            =   3120
         TabIndex        =   20
         Top             =   885
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483625
      End
      Begin Etiq.Etiqueta lblATIVO 
         Height          =   300
         Left            =   660
         TabIndex        =   27
         Top             =   1410
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   529
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483625
      End
      Begin Mascara.Máscara txtATIVO 
         Height          =   300
         Left            =   90
         TabIndex        =   8
         Top             =   1410
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "#####"
         Text            =   ""
         ÉValor          =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Conta Ativo/ Passivo"
         Height          =   195
         Left            =   75
         TabIndex        =   28
         Top             =   1230
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Rateio de Contas"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   75
         TabIndex        =   26
         Top             =   1755
         Width           =   4650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Saldo de Abertura"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Crédito"
         Height          =   195
         Left            =   3105
         TabIndex        =   23
         Top             =   705
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Débito"
         Height          =   195
         Left            =   1815
         TabIndex        =   22
         Top             =   690
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   150
         Width           =   720
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Lançamentos Permitidos"
      Height          =   195
      Left            =   2535
      TabIndex        =   17
      Top             =   60
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Grau"
      Height          =   195
      Left            =   2040
      TabIndex        =   15
      Top             =   75
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta"
      Height          =   195
      Left            =   45
      TabIndex        =   14
      Top             =   75
      Width           =   420
   End
End
Attribute VB_Name = "fPLANOCONTAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tbPlano As Recordset
Private tbRateio As Recordset
Private wp_Cria As Boolean

Private Function Grava_Plano() As Boolean
Dim wl_Conta As String
Dim i As Integer
Dim wl_Index As String
wl_Index = tbPlano.Index
If wp_Cria Then
   If Not add_reg(tbPlano) Then
      Exit Function
   End If
   wl_Conta = Format(txtNIVEL1.VALOR, "0")
   wl_Conta = wl_Conta + IIf(txtNIVEL2.VALOR > 0, "." + Trim(Format(txtNIVEL2.VALOR, "##")), "")
   wl_Conta = wl_Conta + IIf(txtNIVEL3.VALOR > 0, "." + Trim(Format(txtNIVEL3.VALOR, "##")), "")
   wl_Conta = wl_Conta + IIf(txtNIVEL4.VALOR > 0, "." + Format(txtNIVEL4.VALOR, "00"), "")
   wl_Conta = wl_Conta + IIf(txtconta.VALOR > 0, "." + Format(txtconta.VALOR, "00000"), "")
   tbPlano("CONTA") = wl_Conta
Else
   If txtconta.VALOR <> 0 Then
      tbPlano.Index = "iTRADUTOR"
      Loca_Contas tbPlano, Me.txtconta.VALOR
   End If
   If Not edit_reg(tbPlano) Then
      Exit Function
   End If
End If
tbPlano("TRADUTOR") = txtconta.VALOR
tbPlano("DESCRICAO") = txtDESCRICAO.Text
tbPlano("SALDOABERTURA") = txtSALDOABERTURA.VALOR
tbPlano("TIPO") = Mid(cmbNATUREZA.Text, 1, 1)
tbPlano("ATIVOPASSIVO") = Me.txtATIVO.VALOR
If Not update_reg(tbPlano) Then
   Exit Function
End If
tbPlano.Index = wl_Index
GoSub Grava_Rateio
Grava_Plano = True
Exit Function

Grava_Rateio:
tbRateio.Seek "=", txtconta.VALOR
If Not tbRateio.NoMatch Then
   Do While Not tbRateio.EOF
      If tbRateio("CONTAPRINCIPAL") <> txtconta.VALOR Then Exit Do
      If edit_reg(tbRateio) Then tbRateio.Delete
      tbRateio.MoveNext
   Loop
End If
For i = 1 To Me.grdRATEIO.rows
   If grdRATEIO.TextMatrix(i, 0) = "" Then Exit For
   If add_reg(tbRateio) Then
      tbRateio("CONTAPRINCIPAL") = txtconta.VALOR
      tbRateio("CONTARATEIO") = Me.grdRATEIO.TextMatrix(i, 0)
      tbRateio("PROPORCAO") = Me.grdRATEIO.TextMatrix(i, 2)
      update_reg tbRateio
   End If
Next
Return
End Function


Private Sub LimpaGrade()
Dim i As Integer
For i = 1 To Me.grdRATEIO.rows
   If Me.grdRATEIO.TextMatrix(i, 0) = "" Then Exit For
   grdRATEIO.TextMatrix(i, 0) = ""
   grdRATEIO.TextMatrix(i, 1) = ""
   grdRATEIO.TextMatrix(i, 2) = ""
Next
End Sub

Private Sub Mon_Plano()
Dim i As Integer
Dim wl_Conta As String
Dim wl_Index As String
If txtNIVEL1.VALOR = 0 And txtNIVEL2.VALOR = 0 And txtNIVEL3.VALOR = 0 And txtNIVEL4.VALOR = 0 Then
   txtNIVEL1.Text = Mid(tbPlano("CONTA"), 1, 1)
   If pb_NivelPlano = 3 Then
      wl_Conta = ""
      i = 3
      Do While True
         If Mid(tbPlano("CONTA"), i, 1) = "." Then Exit Do
         wl_Conta = wl_Conta + Mid(tbPlano("CONTA"), i, 1)
         i = i + 1
      Loop
      i = i + 1
      txtNIVEL2.Text = wl_Conta
   ElseIf pb_NivelPlano = 4 Then
      wl_Conta = ""
      i = 3
      Do While True
         If Mid(tbPlano("CONTA"), i, 1) = "." Then Exit Do
         wl_Conta = wl_Conta + Mid(tbPlano("CONTA"), i, 1)
         i = i + 1
      Loop
      i = i + 1
      txtNIVEL2.Text = wl_Conta
      wl_Conta = ""
      Do While True
         If Mid(tbPlano("CONTA"), i, 1) = "." Then Exit Do
         wl_Conta = wl_Conta + Mid(tbPlano("CONTA"), i, 1)
         i = i + 1
      Loop
      i = i + 1
      txtNIVEL3.Text = wl_Conta
   ElseIf pb_NivelPlano = 5 Then
      wl_Conta = ""
      i = 3
      Do While True
         If Mid(tbPlano("CONTA"), i, 1) = "." Then Exit Do
         wl_Conta = wl_Conta + Mid(tbPlano("CONTA"), i, 1)
         i = i + 1
      Loop
      i = i + 1
      txtNIVEL2.Text = wl_Conta
      wl_Conta = ""
      Do While True
         If Mid(tbPlano("CONTA"), i, 1) = "." Then Exit Do
         wl_Conta = wl_Conta + Mid(tbPlano("CONTA"), i, 1)
         i = i + 1
      Loop
      i = i + 1
      txtNIVEL3.Text = wl_Conta
      wl_Conta = ""
      Do While True
         If Mid(tbPlano("CONTA"), i, 1) = "." Then Exit Do
         wl_Conta = wl_Conta + Mid(tbPlano("CONTA"), i, 1)
         i = i + 1
      Loop
      i = i + 1
      txtNIVEL4.Text = wl_Conta
   End If
End If
If tbPlano("TIPO") = "D" Then
   cmbNATUREZA.Text = "Despesa"
ElseIf tbPlano("TIPO") = "R" Then
   cmbNATUREZA.Text = "Receita"
ElseIf tbPlano("TIPO") = "A" Then
   cmbNATUREZA.Text = "Ambos"
End If
txtDESCRICAO.Text = tbPlano("DESCRICAO")
txtSALDOABERTURA.Text = tbPlano("SALDOABERTURA")
If VtoP(txtconta.Text) > 0 Then
   txtATIVO.Text = tbPlano("ATIVOPASSIVO")
   If txtATIVO.VALOR > 0 Then
      wl_Index = tbPlano.Index
      tbPlano.Index = "iTRADUTOR"
      If Loca_Contas(tbPlano, txtATIVO.VALOR) Then
         lblATIVO.Caption = tbPlano("DESCRICAO")
      End If
      Loca_Plano tbPlano, txtconta.VALOR
      tbPlano.Index = wl_Index
   End If
End If
i = 1
tbRateio.Seek "=", txtconta.VALOR
If Not tbRateio.NoMatch Then
   Do While Not tbRateio.EOF
      If tbRateio("CONTAPRINCIPAL") <> txtconta.VALOR Then Exit Do
      grdRATEIO.TextMatrix(i, 0) = tbRateio("CONTARATEIO")
      If Loca_Contas(tbPlano, tbRateio("CONTARATEIO")) Then
         grdRATEIO.TextMatrix(i, 1) = tbPlano("DESCRICAO")
      End If
      grdRATEIO.TextMatrix(i, 2) = Format(tbRateio("PROPORCAO"), "##0.00")
      tbRateio.MoveNext
      i = i + 1
   Loop
   wl_Index = tbPlano.Index
   tbPlano.Index = "iTRADUTOR"
   Loca_Contas tbPlano, txtconta.VALOR
   tbPlano.Index = wl_Index
End If
End Sub


Private Sub cmbNATUREZA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   boxDados.Enabled = True
   txtDESCRICAO.SetFocus
End If
End Sub

Private Sub cmbNATUREZA_LostFocus()
cmbNATUREZA.Enabled = False
End Sub


Private Sub cmdDELETA_Click()
Dim wl_Index As String
If MsgBox("Confirma a deleção?", vbYesNo, "Mensagem do Sistema") = vbYes Then
   If Me.txtconta.VALOR > 0 Then
      tbRateio.Seek "=", txtconta.VALOR
      If Not tbRateio.NoMatch Then
         Do While Not tbRateio.EOF
            If tbRateio("CONTAPRINCIPAL") <> txtconta.VALOR Then Exit Do
            If edit_reg(tbRateio) Then tbRateio.Delete
            tbRateio.MoveNext
         Loop
      End If
      wl_Index = tbPlano.Index
      tbPlano.Index = "iTRADUTOR"
      Loca_Contas tbPlano, txtconta.VALOR
   Else
      wl_Index = tbPlano.Index
      tbPlano.Index = "iCONTA"
      If Not Loca_Plano(tbPlano, txtNIVEL1.VALOR, txtNIVEL2.VALOR, txtNIVEL3.VALOR, txtNIVEL4.VALOR) Then
         InformaaoUsuario "A deleção não foi feita ..."
         tbPlano.Index = wl_Index
         Exit Sub
      End If
   End If
   If Not edit_reg(tbPlano) Then
      InformaaoUsuario "A deleção não foi feita ..."
      Exit Sub
   End If
   tbPlano.Delete
   tbPlano.Index = wl_Index
   SendKeys "{ESC}"
End If
End Sub

Private Sub cmdGrava_Click()
If Mid(cmbNATUREZA.Text, 1, 1) = "(" And txtNIVEL2.VALOR = 0 Then
   InformaaoUsuario "Qual o Tipo da Operação"
   Me.cmbNATUREZA.Enabled = True
   cmbNATUREZA.SetFocus
   Exit Sub
End If
If Me.txtNIVEL1.VALOR = 3 And Mid(cmbNATUREZA.Text, 1, 1) <> "D" Then
   If Not Confirme("O tipo de lançamento pertimido está em desacordo com o sistema. Continua?") Then
      Me.cmbNATUREZA.Enabled = True
      cmbNATUREZA.SetFocus
      Exit Sub
   End If
ElseIf Me.txtNIVEL1.VALOR = 4 And Mid(cmbNATUREZA.Text, 1, 1) <> "R" Then
   If Not Confirme("O tipo de lançamento pertimido está em desacordo com o sistema. Continua?") Then
      Me.cmbNATUREZA.Enabled = True
      cmbNATUREZA.SetFocus
      Exit Sub
   End If
ElseIf Me.txtNIVEL1.VALOR <> 3 And Me.txtNIVEL1.VALOR <> 4 And Mid(cmbNATUREZA.Text, 1, 1) <> "A" Then
   If Not Confirme("O tipo de lançamento pertimido está em desacordo com o sistema. Continua?") Then
      Me.cmbNATUREZA.Enabled = True
      cmbNATUREZA.SetFocus
      Exit Sub
   End If
End If
If Grava_Plano Then
   txtNIVEL1.SetFocus
End If
End Sub

Private Sub Form_Activate()
If Not Abre_PlanoContas(tbPlano) Or _
   Not Abre_RateioContas(tbRateio) Then
   Unload Me
   Exit Sub
End If
tbPlano.Index = "iCONTA"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   KeyAscii = 0
   If UCase(Me.ActiveControl.Name) = "TXTNIVEL1" Then
      Unload Me
   ElseIf UCase(Me.ActiveControl.Name) = "GRDRATEIO" Then
      SendKeys "{TAB}"
   ElseIf Me.txtCONTARATEIO.Visible Then
      If UCase(Me.ActiveControl.Name) = "TXTCONTARATEIO" Then
         Me.grdRATEIO.SetFocus
      Else
         txtCONTARATEIO.SetFocus
      End If
   Else
      If wp_Cria And Me.boxDados.Enabled Then
         If Not Confirme("Você está criando uma nova conta! Deseja Abandonar ?") Then Exit Sub
      End If
      txtDESCRICAO.Text = ""
      txtSALDOABERTURA.VALOR = 0
      Me.lblDebito.Caption = ""
      Me.lblCredito.Caption = ""
      If txtconta.VALOR > 0 Then
         txtconta.Text = ""
         txtconta.SetFocus
         Exit Sub
      End If
      If txtNIVEL2.VALOR = 0 Then txtNIVEL1.Text = "": txtNIVEL1.SetFocus: Exit Sub
      If txtNIVEL3.VALOR = 0 Then txtNIVEL2.Text = "": txtNIVEL2.SetFocus: Exit Sub
      If txtNIVEL4.VALOR = 0 Then txtNIVEL3.Text = "": txtNIVEL3.SetFocus: Exit Sub
      If txtconta.VALOR = 0 Then txtNIVEL4.Text = "": txtNIVEL4.SetFocus: Exit Sub
   End If
End If
End Sub

Private Sub Form_Load()
centraobj Me
If pb_NivelPlano = 4 Then
   txtNIVEL4.Visible = False
ElseIf pb_NivelPlano = 3 Then
   txtNIVEL3.Visible = False
   txtNIVEL4.Visible = False
ElseIf pb_NivelPlano = 2 Then
   txtNIVEL2.Visible = False
   txtNIVEL3.Visible = False
   txtNIVEL4.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbPlano.Close
tbRateio.Close
Me.Visible = False
End Sub

Private Sub grdRATEIO_GotFocus()
If txtconta.VALOR = 0 Or Me.cmbNATUREZA.Text = "Ambos" Then SendKeys "{TAB}"
Me.txtCONTARATEIO.Visible = False
Me.lblRATEIO.Visible = False
Me.txtPROPORCAO.Visible = False
End Sub

Private Sub grdRATEIO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
   KeyCode = 0
   If grdRATEIO.TextMatrix(grdRATEIO.row, 0) = "" Then Exit Sub
   If Confirme("Confirma a deleção da linha?") Then
      Call EliminaLinhadaGrade(grdRATEIO)
   End If
End If
End Sub


Private Sub grdRATEIO_KeyPress(KeyAscii As Integer)
Dim wl_Conta As Long
Dim wl_Proporcao As Currency
If KeyAscii = 13 Then
   KeyAscii = 0
   Me.txtCONTARATEIO.Visible = True
   Me.lblRATEIO.Visible = True
   Me.txtPROPORCAO.Visible = True
   If grdRATEIO.TextMatrix(grdRATEIO.row, 0) <> "" Then
      wl_Conta = grdRATEIO.TextMatrix(grdRATEIO.row, 0)
      If Loca_Contas(tbPlano, wl_Conta) Then
         Me.lblRATEIO.Caption = tbPlano("DESCRICAO")
      End If
      wl_Proporcao = grdRATEIO.TextMatrix(grdRATEIO.row, 2)
      Me.txtCONTARATEIO.Text = wl_Proporcao
      Me.txtPROPORCAO.Text = wl_Proporcao
   Else
      Me.txtCONTARATEIO.Text = ""
      Me.lblRATEIO.Caption = ""
      Me.txtPROPORCAO.Text = ""
   End If
   txtCONTARATEIO.Visible = True
   Me.lblRATEIO.Visible = True
   Me.txtPROPORCAO.Visible = True
   Me.txtCONTARATEIO.SetFocus
End If
End Sub


Private Sub txtATIVO_GotFocus()
If txtconta.VALOR = 0 Then SendKeys "{TAB}"
End Sub

Private Sub txtATIVO_KeyPress(KeyAscii As Integer)
Dim wl_Index As String
If KeyAscii = 13 Then
   KeyAscii = 0
   If VtoP(txtATIVO.Text) = 0 Then
      Me.lblATIVO.Caption = ""
   Else
      wl_Index = tbPlano.Index
      tbPlano.Index = "iTRADUTOR"
      If Not Loca_Contas(tbPlano, txtATIVO.VALOR) Then
         InformaaoUsuario "Conta não encontrada"
         HomeEnd
         Exit Sub
      End If
      Me.lblATIVO.Caption = tbPlano("DESCRICAO")
   End If
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtATIVO_LostFocus()
grdRATEIO.row = 1
End Sub

Private Sub txtCONTA_GotFocus()
If txtNIVEL4.Text = "" And txtNIVEL4.Visible And txtNIVEL1.VALOR <> 0 Then
   InformaaoUsuario "Informe o nível 4"
   txtNIVEL4.SetFocus
   Exit Sub
End If
txtconta.Text = ""
txtDESCRICAO.Text = ""
txtSALDOABERTURA.Text = ""
lblDebito.Caption = ""
lblCredito.Caption = ""
cmdGrava.Enabled = False
cmdDELETA.Enabled = False
Me.txtATIVO.Text = ""
Me.lblATIVO.Caption = ""
LimpaGrade
boxDados.Enabled = False
lblGRAU.Caption = pb_NivelPlano - 1
cmbNATUREZA.Enabled = False
If txtNIVEL1.VALOR <> 0 Then
   If Loca_Plano(tbPlano, txtNIVEL1.VALOR, txtNIVEL2.VALOR, txtNIVEL3.VALOR, txtNIVEL4.VALOR) Then
      tbPlano.Index = "iTRADUTOR"
      If tbPlano.RecordCount = 0 Then
         txtconta.Text = 1
      Else
         tbPlano.MoveLast
         txtconta.Text = tbPlano("TRADUTOR") + 1
      End If
      HomeEnd
      tbPlano.Index = "iCONTA"
   Else
      wp_Cria = True
      boxDados.Enabled = True
      txtDESCRICAO.SetFocus
   End If
Else
   lblGRAU.Caption = ""
End If
End Sub

Private Sub txtconta_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtconta.SetFocus
   If wl_Retorno <> "" Then
      txtconta.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If
End Sub


Private Sub txtconta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtconta.Text = "" Or txtconta.Text = "0" Then
      If txtNIVEL1.VALOR = 0 Then
         InformaaoUsuario "Para criar uma nova conta, informe todos os níveis"
         txtNIVEL1.SetFocus
         Exit Sub
      End If
      If Loca_Plano(tbPlano, txtNIVEL1.VALOR, txtNIVEL2.VALOR, txtNIVEL3.VALOR, txtNIVEL4.VALOR) Then
         Mon_Plano
         wp_Cria = False
      Else
         wp_Cria = True
      End If
   Else
      lblGRAU.Caption = IIf(txtNIVEL1.VALOR > 0, Format(pb_NivelPlano, "00"), "")
      tbPlano.Index = "iTRADUTOR"
      If Loca_Contas(tbPlano, txtconta.Text) Then
         If txtNIVEL1.VALOR > 0 Or txtNIVEL2.VALOR > 0 Or txtNIVEL3.VALOR > 0 Or txtNIVEL4.VALOR > 0 Then
            txtNIVEL1.Text = ""
            txtNIVEL2.Text = ""
            txtNIVEL3.Text = ""
            txtNIVEL4.Text = ""
         End If
         Mon_Plano
         wp_Cria = False
      Else
         If txtNIVEL1.VALOR = 0 Then
            InformaaoUsuario "Conta não encontrada"
            txtNIVEL1.SetFocus
            Exit Sub
         End If
         wp_Cria = True
      End If
      tbPlano.Index = "iCONTA"
      If Not wp_Cria Then
         Loca_Plano tbPlano, txtNIVEL1.VALOR, txtNIVEL2.VALOR, txtNIVEL3.VALOR, txtNIVEL4.VALOR, txtconta.Text
      End If
   End If
   boxDados.Enabled = True
   SendKeys "{TAB}"
End If
End Sub

Private Sub txtCONTARATEIO_GotFocus()
Me.lblRATEIO.Caption = ""
Me.txtPROPORCAO.Text = ""
End Sub

Private Sub txtCONTARATEIO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   If wl_Retorno <> "" Then
      txtCONTARATEIO.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If
End Sub


Private Sub txtCONTARATEIO_KeyPress(KeyAscii As Integer)
Dim wl_Index As String
If KeyAscii = 13 Then
   KeyAscii = 0
   If VtoP(txtCONTARATEIO.Text) = 0 Then
      Me.grdRATEIO.SetFocus
      Exit Sub
   End If
   wl_Index = tbPlano.Index
   tbPlano.Index = "iTRADUTOR"
   If Loca_Contas(tbPlano, txtCONTARATEIO.Text) Then
      Me.lblRATEIO.Caption = tbPlano("DESCRICAO")
   Else
      tbPlano.Index = wl_Index
      InformaaoUsuario "Conta não encontrada ..."
      HomeEnd
      Exit Sub
   End If
   tbPlano.Index = wl_Index
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtDESCRICAO_GotFocus()
Me.lblGRAU.Caption = IIf(txtconta.VALOR > 0, pb_NivelPlano, lblGRAU.Caption)
cmdGrava.Enabled = True
cmdDELETA.Enabled = Not wp_Cria

End Sub



Private Sub txtNIVEL1_GotFocus()
cmbNATUREZA.Text = "(Nenhum)"
boxDados.Enabled = False
lblGRAU.Caption = "00"
cmbNATUREZA.Enabled = False
cmdGrava.Enabled = False
cmdDELETA.Enabled = False
txtNIVEL2.Text = ""
txtNIVEL3.Text = ""
txtNIVEL4.Text = ""
txtconta.Text = ""
txtDESCRICAO.Text = ""
Me.txtATIVO.Text = ""
Me.lblATIVO.Caption = ""
LimpaGrade
End Sub

Private Sub txtNIVEL1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   Most_PlanodeContas
End If
End Sub


Private Sub txtNIVEL1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtNIVEL1.Text = "" Or txtNIVEL1.Text = "0" Then
      txtconta.SetFocus
      HomeEnd
      Exit Sub
   End If
   If Loca_Plano(tbPlano, Val(txtNIVEL1.Text), 0, 0, 0, 0) Then
      If tbPlano("TIPO") = "D" Then
         cmbNATUREZA.Text = "Despesa"
      ElseIf tbPlano("TIPO") = "R" Then
         cmbNATUREZA.Text = "Receita"
      ElseIf tbPlano("TIPO") = "A" Then
         cmbNATUREZA.Text = "Ambos"
      End If
   Else
      If txtNIVEL1.Text = "1" Then cmbNATUREZA.Text = "Ambos"
      If txtNIVEL1.Text = "2" Then cmbNATUREZA.Text = "Ambos"
      If txtNIVEL1.Text = "3" Then cmbNATUREZA.Text = "Despesa"
      If txtNIVEL1.Text = "4" Then cmbNATUREZA.Text = "Receita"
   End If
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtNIVEL2_GotFocus()
If txtNIVEL1.VALOR = 0 Then
   InformaaoUsuario "Informe o nível 1"
   txtNIVEL1.SetFocus
   Exit Sub
End If
txtNIVEL3.Text = ""
txtNIVEL4.Text = ""
txtconta.Text = ""
txtDESCRICAO.Text = ""
txtSALDOABERTURA.Text = ""
lblDebito.Caption = ""
lblCredito.Caption = ""
boxDados.Enabled = False
Me.txtATIVO.Text = ""
Me.lblATIVO.Caption = ""
lblGRAU.Caption = "01"
cmbNATUREZA.Enabled = False
If Not Loca_Plano(tbPlano, txtNIVEL1.VALOR) Then
   wp_Cria = True
   boxDados.Enabled = True
   txtDESCRICAO.SetFocus
End If
LimpaGrade
End Sub

Private Sub txtNIVEL2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Most_PlanodeContas
End Sub


Private Sub txtNIVEL2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
'   If txtNIVEL2.Text = "" Or txtNIVEL2.Text = "0" Then
    If txtNIVEL2.Text = "" Then
      If Loca_Plano(tbPlano, txtNIVEL1.VALOR) Then
         Mon_Plano
         wp_Cria = False
      Else
         wp_Cria = True
      End If
      cmbNATUREZA.Enabled = True
      cmbNATUREZA.SetFocus
   Else
      SendKeys "{TAB}"
   End If
End If
End Sub


Private Sub txtNIVEL3_GotFocus()
'If txtNIVEL2.VALOR = "" Then
'   InformaaoUsuario "Informe o nível 2"
'   txtNIVEL2.SetFocus
'End If
txtNIVEL4.Text = ""
txtconta.Text = ""
txtDESCRICAO.Text = ""
txtSALDOABERTURA.Text = ""
lblDebito.Caption = ""
lblCredito.Caption = ""
Me.txtATIVO.Text = ""
Me.lblATIVO.Caption = ""
cmdGrava.Enabled = False
cmdDELETA.Enabled = False
boxDados.Enabled = False
lblGRAU.Caption = "02"
cmbNATUREZA.Enabled = False
If Not Loca_Plano(tbPlano, txtNIVEL1.VALOR, txtNIVEL2.VALOR) Then
   wp_Cria = True
   boxDados.Enabled = True
   txtDESCRICAO.SetFocus
End If
LimpaGrade
End Sub

Private Sub txtNIVEL3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Most_PlanodeContas
End Sub


Private Sub txtNIVEL3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
  ' If txtNIVEL3.Text = "" Or txtNIVEL3.Text = "0" Then
    If txtNIVEL3.Text = "" Then
      If Loca_Plano(tbPlano, txtNIVEL1.VALOR, txtNIVEL2.VALOR) Then
         Mon_Plano
         wp_Cria = False
      Else
         wp_Cria = True
      End If
      boxDados.Enabled = True
      txtDESCRICAO.SetFocus
   Else
      SendKeys "{TAB}"
   End If
End If
End Sub


Private Sub txtNIVEL4_GotFocus()
'If txtNIVEL3.VALOR = 0 Then
'   InformaaoUsuario "Informe o nível 3"
'   txtNIVEL3.SetFocus
'End If
txtconta.Text = ""
txtDESCRICAO.Text = ""
txtSALDOABERTURA.Text = ""
lblDebito.Caption = ""
lblCredito.Caption = ""
Me.txtATIVO.Text = ""
Me.lblATIVO.Caption = ""
cmdGrava.Enabled = False
cmdDELETA.Enabled = False
boxDados.Enabled = False
lblGRAU.Caption = "03"
cmbNATUREZA.Enabled = False
If Not Loca_Plano(tbPlano, txtNIVEL1.VALOR, txtNIVEL2.VALOR, txtNIVEL3.VALOR) Then
   wp_Cria = True
   boxDados.Enabled = True
   txtDESCRICAO.SetFocus
End If
LimpaGrade
End Sub

Private Sub txtNIVEL4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Most_PlanodeContas
End Sub


Private Sub txtNIVEL4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
  'If txtNIVEL4.Text = "" Or txtNIVEL4.Text = "0" Then
   If txtNIVEL4.Text = "" Then
      If Loca_Plano(tbPlano, txtNIVEL1.VALOR, txtNIVEL2.VALOR, txtNIVEL3.VALOR) Then
         Mon_Plano
         wp_Cria = False
      Else
         wp_Cria = True
      End If
      boxDados.Enabled = True
      txtDESCRICAO.SetFocus
   Else
      SendKeys "{TAB}"
   End If
End If
End Sub


Private Sub txtPROPORCAO_KeyPress(KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtPROPORCAO.VALOR = 0 Then
      txtCONTARATEIO.Text = ""
      Me.txtCONTARATEIO.SetFocus
      Exit Sub
   End If
   GoSub atualiza_grade
   txtCONTARATEIO.Text = ""
   Me.txtCONTARATEIO.SetFocus
End If
Exit Sub

atualiza_grade:
For i = 1 To Me.grdRATEIO.rows
   If Me.grdRATEIO.TextMatrix(i, 0) = "" Then Exit For
Next
grdRATEIO.TextMatrix(i, 0) = Str(Me.txtCONTARATEIO.VALOR)
grdRATEIO.TextMatrix(i, 1) = Me.lblRATEIO.Caption
grdRATEIO.TextMatrix(i, 2) = Format(Me.txtPROPORCAO.VALOR, "##0.00")
grdRATEIO.row = i + 1
Return
End Sub


Private Sub txtSALDOABERTURA_GotFocus()
If txtconta.VALOR = 0 Then SendKeys "{TAB}"
End Sub

