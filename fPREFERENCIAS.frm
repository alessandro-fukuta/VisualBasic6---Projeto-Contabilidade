VERSION 5.00
Object = "{BA676A3D-9505-4A77-87DC-76025E082864}#1.0#0"; "ETIQUETA.OCX"
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fPREFERENCIAS 
   Caption         =   "Preferências"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   Icon            =   "fPREFERENCIAS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "&Confirma"
      Height          =   375
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3945
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "C&ancela"
      Height          =   375
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3945
      Width           =   1230
   End
   Begin VB.Frame box_Contabilidade 
      Height          =   2445
      Left            =   420
      TabIndex        =   9
      Top             =   630
      Width           =   6495
      Begin Mascara.Máscara txtNIVEL1 
         Height          =   300
         Left            =   105
         TabIndex        =   4
         Top             =   1605
         Width           =   225
         _ExtentX        =   397
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
      Begin VB.CheckBox chkREGIME 
         Caption         =   "Usa Regime de Competência como padrão"
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   1140
         Value           =   1  'Checked
         Width           =   3345
      End
      Begin VB.CheckBox chkOnline 
         Caption         =   "Lançamentos Contábeis On-Line"
         Enabled         =   0   'False
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   915
         Width           =   3345
      End
      Begin VB.CheckBox chkPLANO 
         Caption         =   "Abre todos os níveis do plano de contas"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   675
         Width           =   3345
      End
      Begin VB.ComboBox cmbNIVEL 
         Height          =   315
         ItemData        =   "fPREFERENCIAS.frx":000C
         Left            =   2010
         List            =   "fPREFERENCIAS.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   900
      End
      Begin Mascara.Máscara txtNIVEL2 
         Height          =   300
         Left            =   330
         TabIndex        =   5
         Top             =   1605
         Width           =   330
         _ExtentX        =   582
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
         Left            =   675
         TabIndex        =   6
         Top             =   1605
         Width           =   330
         _ExtentX        =   582
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
         Left            =   1005
         TabIndex        =   7
         Top             =   1605
         Width           =   330
         _ExtentX        =   582
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   """Disponível"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1230
         TabIndex        =   43
         Top             =   1425
         Width           =   1110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nível do Grupo"
         Height          =   195
         Left            =   90
         TabIndex        =   42
         Top             =   1425
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nível do Plano de Contas"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   300
         Width           =   1830
      End
   End
   Begin VB.Frame box_CLIENTES 
      Height          =   2955
      Left            =   420
      TabIndex        =   26
      Top             =   630
      Visible         =   0   'False
      Width           =   6495
      Begin Etiq.Etiqueta lblCLIENTE 
         Height          =   300
         Left            =   795
         TabIndex        =   27
         Top             =   375
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   529
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Mascara.Máscara txtCliente 
         Height          =   300
         Left            =   90
         TabIndex        =   28
         Top             =   375
         Width           =   690
         _ExtentX        =   1217
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
      Begin Etiq.Etiqueta lblDESCONTO_C 
         Height          =   300
         Left            =   795
         TabIndex        =   33
         Top             =   900
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   529
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Mascara.Máscara txtDESCONTO_C 
         Height          =   300
         Left            =   90
         TabIndex        =   29
         Top             =   900
         Width           =   690
         _ExtentX        =   1217
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
      Begin Etiq.Etiqueta lblJUROS_C 
         Height          =   300
         Left            =   795
         TabIndex        =   34
         Top             =   1425
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   529
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Mascara.Máscara txtJUROS_C 
         Height          =   300
         Left            =   90
         TabIndex        =   30
         Top             =   1425
         Width           =   690
         _ExtentX        =   1217
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
      Begin Etiq.Etiqueta lblCORRECAO_C 
         Height          =   300
         Left            =   795
         TabIndex        =   35
         Top             =   1965
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   529
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Mascara.Máscara txtCORRECAO_C 
         Height          =   300
         Left            =   90
         TabIndex        =   31
         Top             =   1965
         Width           =   690
         _ExtentX        =   1217
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
      Begin Etiq.Etiqueta lblTAXA 
         Height          =   300
         Left            =   795
         TabIndex        =   40
         Top             =   2490
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   529
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Mascara.Máscara txtTAXA 
         Height          =   300
         Left            =   90
         TabIndex        =   32
         Top             =   2490
         Width           =   705
         _ExtentX        =   1244
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Taxa Desconto de Título"
         Height          =   195
         Left            =   75
         TabIndex        =   41
         Top             =   2280
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Correção Recebida"
         Height          =   195
         Left            =   75
         TabIndex        =   39
         Top             =   1755
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Juros Recebidos"
         Height          =   195
         Left            =   75
         TabIndex        =   38
         Top             =   1215
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descontos Concedidos"
         Height          =   195
         Left            =   75
         TabIndex        =   37
         Top             =   690
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clientes Diversos"
         Height          =   195
         Left            =   75
         TabIndex        =   36
         Top             =   165
         Width           =   1215
      End
   End
   Begin VB.Frame box_Fornecedores 
      Height          =   2445
      Left            =   420
      TabIndex        =   11
      Top             =   630
      Visible         =   0   'False
      Width           =   6495
      Begin Etiq.Etiqueta lblFORNECEDOR 
         Height          =   300
         Left            =   795
         TabIndex        =   12
         Top             =   375
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   529
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Mascara.Máscara txtFORNECEDOR 
         Height          =   300
         Left            =   90
         TabIndex        =   13
         Top             =   375
         Width           =   690
         _ExtentX        =   1217
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
      Begin Etiq.Etiqueta lblDESCONTO_F 
         Height          =   300
         Left            =   795
         TabIndex        =   17
         Top             =   900
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   529
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Mascara.Máscara txtDESCONTO_F 
         Height          =   300
         Left            =   90
         TabIndex        =   14
         Top             =   900
         Width           =   690
         _ExtentX        =   1217
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
      Begin Etiq.Etiqueta lblJURO_F 
         Height          =   300
         Left            =   795
         TabIndex        =   18
         Top             =   1425
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   529
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Mascara.Máscara txtJURO_F 
         Height          =   300
         Left            =   90
         TabIndex        =   15
         Top             =   1425
         Width           =   690
         _ExtentX        =   1217
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
      Begin Etiq.Etiqueta lblCORRECAO_F 
         Height          =   300
         Left            =   795
         TabIndex        =   19
         Top             =   1965
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   529
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Mascara.Máscara txtCORRECAO_F 
         Height          =   300
         Left            =   90
         TabIndex        =   16
         Top             =   1965
         Width           =   690
         _ExtentX        =   1217
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedores Diversos"
         Height          =   195
         Left            =   75
         TabIndex        =   23
         Top             =   165
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descontos Obtidos"
         Height          =   195
         Left            =   75
         TabIndex        =   22
         Top             =   690
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Juros Pagos"
         Height          =   195
         Left            =   75
         TabIndex        =   21
         Top             =   1215
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Correção Paga"
         Height          =   195
         Left            =   75
         TabIndex        =   20
         Top             =   1755
         Width           =   1065
      End
   End
   Begin MSComctlLib.TabStrip Abas 
      Height          =   4275
      Left            =   240
      TabIndex        =   8
      Top             =   195
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   7541
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilidade Geral"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fornecedores"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clientes"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fPREFERENCIAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tbPlano As Recordset
Private tbPreferencias As Recordset
Private wp_Entrada As Boolean




Sub Habilita_Niveis()
If pb_NivelPlano = 2 Then
   txtNIVEL2.Visible = False
   txtNIVEL3.Visible = False
   txtNIVEL4.Visible = False
ElseIf pb_NivelPlano = 3 Then
   txtNIVEL3.Visible = False
   txtNIVEL4.Visible = False
ElseIf pb_NivelPlano = 4 Then
   txtNIVEL4.Visible = False
End If
End Sub

Private Sub Inicializa()
Dim wl_Retorno As String
Dim i As Integer
Dim wl_Nivel As String
Dim wl_ContaNivel As Integer
chkOnline.Value = 0
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "RegimeCompetencia")
Me.chkREGIME.Value = IIf(wl_Retorno = "", 1, wl_Retorno)
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "NivelPlanodeContas")
If wl_Retorno = "" Then
   cmbNIVEL.Text = "5"
Else
   cmbNIVEL.Text = wl_Retorno
End If
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "PlanodeContasAberto")
chkPLANO.Value = IIf(wl_Retorno = "", 1, wl_Retorno)
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "Online")
chkOnline.Value = IIf(wl_Retorno = "", 1, wl_Retorno)
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "Disponivel")
wl_Nivel = ""
For i = 1 To Len(wl_Retorno)
   If Mid(wl_Retorno, i, 1) <> "." Then
      wl_Nivel = wl_Nivel + Mid(wl_Retorno, i, 1)
   Else
      GoSub Monta_Nivel
   End If
Next
GoSub Monta_Nivel
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "FornecedoresDiversos")
txtFORNECEDOR.Text = wl_Retorno
If txtFORNECEDOR.Text <> "" And txtFORNECEDOR.Text <> "0" Then
   If Loca_Contas(tbPlano, txtFORNECEDOR.Text) Then
      lblFORNECEDOR.Caption = tbPlano("DESCRICAO")
   End If
End If
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "JURO_F")
txtJURO_F.Text = wl_Retorno
If txtJURO_F.Text <> "" And txtJURO_F.Text <> "0" Then
   If Loca_Contas(tbPlano, txtJURO_F.Text) Then
      lblJURO_F.Caption = tbPlano("DESCRICAO")
   End If
End If
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "DESCONTO_F")
txtDESCONTO_F.Text = wl_Retorno
If txtDESCONTO_F.Text <> "" And txtDESCONTO_F.Text <> "0" Then
   If Loca_Contas(tbPlano, txtDESCONTO_F.Text) Then
      lblDESCONTO_F.Caption = tbPlano("DESCRICAO")
   End If
End If
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "CORRECAO_F")
txtCORRECAO_F.Text = wl_Retorno
If txtCORRECAO_F.Text <> "" And txtCORRECAO_F.Text <> "0" Then
   If Loca_Contas(tbPlano, txtCORRECAO_F.Text) Then
      lblCORRECAO_F.Caption = tbPlano("DESCRICAO")
   End If
End If

wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "ClientesDiversos")
txtCLIENTE.Text = wl_Retorno
If txtCLIENTE.Text <> "" And txtCLIENTE.Text <> "0" Then
   If Loca_Contas(tbPlano, txtCLIENTE.Text) Then
      lblCLIENTE.Caption = tbPlano("DESCRICAO")
   End If
End If
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "JURO_C")
txtJUROS_C.Text = wl_Retorno
If txtJUROS_C.Text <> "" And txtJUROS_C.Text <> "0" Then
   If Loca_Contas(tbPlano, txtJUROS_C.Text) Then
      lblJUROS_C.Caption = tbPlano("DESCRICAO")
   End If
End If
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "DESCONTO_C")
txtDESCONTO_C.Text = wl_Retorno
If txtDESCONTO_C.Text <> "" And txtDESCONTO_C.Text <> "0" Then
   If Loca_Contas(tbPlano, txtDESCONTO_C.Text) Then
      lblDESCONTO_C.Caption = tbPlano("DESCRICAO")
   End If
End If
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "CORRECAO_C")
txtCORRECAO_C.Text = wl_Retorno
If txtCORRECAO_C.Text <> "" And txtCORRECAO_C.Text <> "0" Then
   If Loca_Contas(tbPlano, txtCORRECAO_C.Text) Then
      lblCORRECAO_C.Caption = tbPlano("DESCRICAO")
   End If
End If
wl_Retorno = RetornaConfiguracao("PREFERENCIAS_" + Format(pb_Empresa, "000"), "TAXA")
txtTAXA.Text = wl_Retorno
If txtTAXA.Text <> "" And txtTAXA.Text <> "0" Then
   If Loca_Contas(tbPlano, txtTAXA.Text) Then
      lblTAXA.Caption = tbPlano("DESCRICAO")
   End If
End If
Exit Sub

Monta_Nivel:
wl_ContaNivel = wl_ContaNivel + 1
If wl_ContaNivel = 1 Then txtNIVEL1.Text = wl_Nivel
If wl_ContaNivel = 2 Then txtNIVEL2.Text = wl_Nivel
If wl_ContaNivel = 3 Then txtNIVEL3.Text = wl_Nivel
If wl_ContaNivel = 4 Then txtNIVEL4.Text = wl_Nivel
wl_Nivel = ""
Return
End Sub



Private Sub Abas_Click()
Dim wl_Retorno
If UCase(Abas.SelectedItem.Caption) = "CONTABILIDADE GERAL" Then
   box_Fornecedores.Visible = False
   box_CLIENTES.Visible = False
   box_Contabilidade.Visible = True
   cmbNIVEL.SetFocus
ElseIf UCase(Abas.SelectedItem.Caption) = "CLIENTES" Then
   box_Fornecedores.Visible = False
   box_Contabilidade.Visible = False
   box_CLIENTES.Visible = True
   If txtCLIENTE.Text <> "" And txtCLIENTE.Text <> "0" Then
      If Loca_Contas(tbPlano, txtCLIENTE.Text) Then
         lblCLIENTE.Caption = tbPlano("DESCRICAO")
      End If
   End If
   If txtJUROS_C.Text <> "" And txtJUROS_C.Text <> "0" Then
      If Loca_Contas(tbPlano, txtJUROS_C.Text) Then
         lblJUROS_C.Caption = tbPlano("DESCRICAO")
      End If
   End If
   If txtDESCONTO_C.Text <> "" And txtDESCONTO_C.Text <> "0" Then
      If Loca_Contas(tbPlano, txtDESCONTO_C.Text) Then
         lblDESCONTO_C.Caption = tbPlano("DESCRICAO")
      End If
   End If
   If txtCORRECAO_C.Text <> "" And txtCORRECAO_C.Text <> "0" Then
      If Loca_Contas(tbPlano, txtCORRECAO_C.Text) Then
         lblCORRECAO_C.Caption = tbPlano("DESCRICAO")
      End If
   End If
   txtCLIENTE.SetFocus
ElseIf UCase(Abas.SelectedItem.Caption) = "FORNECEDORES" Then
   box_Contabilidade.Visible = False
   box_CLIENTES.Visible = False
   box_Fornecedores.Visible = True
   If txtFORNECEDOR.Text <> "" And txtFORNECEDOR.Text <> "0" Then
      If Loca_Contas(tbPlano, txtFORNECEDOR.Text) Then
         lblFORNECEDOR.Caption = tbPlano("DESCRICAO")
      End If
   End If
   If txtJURO_F.Text <> "" And txtJURO_F.Text <> "0" Then
      If Loca_Contas(tbPlano, txtJURO_F.Text) Then
         lblJURO_F.Caption = tbPlano("DESCRICAO")
      End If
   End If
   If txtDESCONTO_F.Text <> "" And txtDESCONTO_F.Text <> "0" Then
      If Loca_Contas(tbPlano, txtDESCONTO_F.Text) Then
         lblDESCONTO_F.Caption = tbPlano("DESCRICAO")
      End If
   End If
   If txtCORRECAO_F.Text <> "" And txtCORRECAO_F.Text <> "0" Then
      If Loca_Contas(tbPlano, txtCORRECAO_F.Text) Then
         lblCORRECAO_F.Caption = tbPlano("DESCRICAO")
      End If
   End If
   txtFORNECEDOR.SetFocus
End If
End Sub

Private Sub box_Sistema_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub chkOnline_Click()
pb_Online = IIf(chkOnline.Value = 1, True, False)
End Sub

Private Sub Command1_Click()
Dim wl_Disponivel As String
wl_Disponivel = Str(txtNIVEL1.VALOR) + IIf(txtNIVEL2.VALOR > 0, "." + Trim(Str(txtNIVEL2.VALOR)), ".") + IIf(txtNIVEL3.VALOR > 0, "." + Trim(Str(txtNIVEL3.VALOR)), ".") + IIf(txtNIVEL4.VALOR > 0, "." + Format(txtNIVEL4.VALOR, "00"), ".")
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "NivelPlanodeContas", cmbNIVEL.Text
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "PlanodeContasAberto", Me.chkPLANO.Value
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "Online", Me.chkOnline.Value
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "RegimeCompetencia", Me.chkREGIME.Value
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "Disponivel", wl_Disponivel

Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "FornecedoresDiversos", Me.txtFORNECEDOR.VALOR
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "Juro_F", Me.txtJURO_F.VALOR
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "Desconto_F", Me.txtDESCONTO_F.VALOR
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "Correcao_F", Me.txtCORRECAO_F.VALOR

Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "ClientesDiversos", Me.txtCLIENTE.VALOR
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "Juro_C", Me.txtJUROS_C.VALOR
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "Desconto_C", Me.txtDESCONTO_C.VALOR
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "Correcao_C", Me.txtCORRECAO_C.VALOR
Grava_Configuracoes "PREFERENCIAS_" + Format(pb_Empresa, "000"), "Taxa", Me.txtTAXA.VALOR
pb_NivelPlano = cmbNIVEL.Text
pb_RegimeCompetencia = IIf(Me.chkREGIME.Value = 1, True, False)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If Not wp_Entrada Then
   If Not Abre_PlanoContas(tbPlano) Then
      Unload Me
      Exit Sub
   End If
   Inicializa
   Habilita_Niveis
   wp_Entrada = True
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub


Private Sub lstNIVEL_Scroll()

End Sub


Private Sub Form_Load()
wp_Entrada = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbPlano.Close
End Sub

Private Sub txtCLIENTE_GotFocus()
aviso "<F1> Plano de Contas"
End Sub

Private Sub txtCLIENTE_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtCLIENTE.SetFocus
   If wl_Retorno <> "" Then
      txtCLIENTE.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If
End Sub

Private Sub txtCLIENTE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtCLIENTE.Text <> "" And txtCLIENTE.Text <> "0" Then
      If Not Loca_Contas(tbPlano, txtCLIENTE.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtCLIENTE.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblCLIENTE.Caption = tbPlano("DESCRICAO")
   Else
      lblCLIENTE.Caption = ""
   End If
   SendKeys "{TAB}"
End If
End Sub


Private Sub txtCLIENTE_LostFocus()
aviso
End Sub

Private Sub txtCORRECAO_C_GotFocus()
aviso "<F1> Plano de Contas"
End Sub

Private Sub txtCORRECAO_C_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtCORRECAO_C.SetFocus
   If wl_Retorno <> "" Then
      txtCORRECAO_C.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If

End Sub

Private Sub txtCORRECAO_C_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtCORRECAO_C.Text <> "" And txtCORRECAO_C.Text <> "0" Then
      If Not Loca_Contas(tbPlano, txtCORRECAO_C.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtCORRECAO_C.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblCORRECAO_C.Caption = tbPlano("DESCRICAO")
   Else
      lblCORRECAO_C.Caption = ""
   End If
   SendKeys "{TAB}"
End If

End Sub


Private Sub txtCORRECAO_C_LostFocus()
aviso
End Sub

Private Sub txtCORRECAO_F_GotFocus()
aviso "<F1> Plano de Contas"
End Sub

Private Sub txtCORRECAO_F_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtCORRECAO_F.SetFocus
   If wl_Retorno <> "" Then
      txtCORRECAO_F.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If

End Sub

Private Sub txtCORRECAO_F_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtCORRECAO_F.Text <> "" And txtCORRECAO_F.Text <> "0" Then
      If Not Loca_Contas(tbPlano, txtCORRECAO_F.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtCORRECAO_F.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblCORRECAO_F.Caption = tbPlano("DESCRICAO")
   Else
      lblCORRECAO_F.Caption = ""
   End If
   SendKeys "{TAB}"
End If
End Sub

Private Sub txtCORRECAO_F_LostFocus()
aviso
End Sub

Private Sub txtDESCONTO_C_GotFocus()
aviso "<F1> Plano de Contas"
End Sub

Private Sub txtDESCONTO_C_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtDESCONTO_C.SetFocus
   If wl_Retorno <> "" Then
      txtDESCONTO_C.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If
End Sub

Private Sub txtDESCONTO_C_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtDESCONTO_C.Text <> "" And txtDESCONTO_C.Text <> "0" Then
      If Not Loca_Contas(tbPlano, txtDESCONTO_C.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtDESCONTO_C.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblDESCONTO_C.Caption = tbPlano("DESCRICAO")
   Else
      lblDESCONTO_C.Caption = ""
   End If
   SendKeys "{TAB}"
End If

End Sub


Private Sub txtDESCONTO_C_LostFocus()
aviso
End Sub

Private Sub txtDESCONTO_F_GotFocus()
aviso "<F1> Plano de Contas"
End Sub

Private Sub txtDESCONTO_F_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtDESCONTO_F.SetFocus
   If wl_Retorno <> "" Then
      txtDESCONTO_F.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If

End Sub


Private Sub txtDESCONTO_F_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtDESCONTO_F.Text <> "" And txtDESCONTO_F.Text <> "0" Then
      If Not Loca_Contas(tbPlano, txtDESCONTO_F.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtDESCONTO_F.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblDESCONTO_F.Caption = tbPlano("DESCRICAO")
   Else
      lblDESCONTO_F.Caption = ""
   End If
   SendKeys "{TAB}"
End If

End Sub

Private Sub txtDESCONTO_F_LostFocus()
aviso
End Sub

Private Sub txtFORNECEDOR_GotFocus()
aviso "<F1> Plano de Contas"
End Sub

Private Sub txtFORNECEDOR_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtFORNECEDOR.SetFocus
   If wl_Retorno <> "" Then
      txtFORNECEDOR.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If
End Sub


Private Sub txtFORNECEDOR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtFORNECEDOR.Text <> "" And txtFORNECEDOR.Text <> "0" Then
      If Not Loca_Contas(tbPlano, txtFORNECEDOR.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtFORNECEDOR.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblFORNECEDOR.Caption = tbPlano("DESCRICAO")
   Else
      lblFORNECEDOR.Caption = ""
   End If
   SendKeys "{TAB}"
End If
End Sub

Private Sub txtFORNECEDOR_LostFocus()
aviso
End Sub


Private Sub txtJURO_F_GotFocus()
aviso "<F1> Plano de Contas"

End Sub

Private Sub txtJURO_F_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtJURO_F.SetFocus
   If wl_Retorno <> "" Then
      txtJURO_F.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If

End Sub


Private Sub txtJURO_F_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtJURO_F.Text <> "" And txtJURO_F.Text <> "0" Then
      If Not Loca_Contas(tbPlano, txtJURO_F.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtJURO_F.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblJURO_F.Caption = tbPlano("DESCRICAO")
   Else
      lblJURO_F.Caption = ""
   End If
   SendKeys "{TAB}"
End If

End Sub


Private Sub txtJURO_F_LostFocus()
aviso
End Sub


Private Sub txtJUROS_C_GotFocus()
aviso "<F1> Plano de Contas"
End Sub

Private Sub txtJUROS_C_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtJUROS_C.SetFocus
   If wl_Retorno <> "" Then
      txtJUROS_C.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If

End Sub


Private Sub txtJUROS_C_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtJUROS_C.Text <> "" And txtJUROS_C.Text <> "0" Then
      If Not Loca_Contas(tbPlano, txtJUROS_C.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtJUROS_C.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblJUROS_C.Caption = tbPlano("DESCRICAO")
   Else
      lblJUROS_C.Caption = ""
   End If
   SendKeys "{TAB}"
End If

End Sub


Private Sub txtJUROS_C_LostFocus()
aviso
End Sub


Private Sub txtNIVEL1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Most_PlanodeContas
End Sub

Private Sub txtNIVEL2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Most_PlanodeContas

End Sub


Private Sub txtNIVEL3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Most_PlanodeContas

End Sub


Private Sub txtNIVEL4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Most_PlanodeContas

End Sub


Private Sub txtTAXA_GotFocus()
aviso "<F1> Plano de Contas"
End Sub

Private Sub txtTAXA_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wl_Retorno
If KeyCode = vbKeyF1 Then
   KeyCode = 0
   wl_Retorno = Most_PlanodeContas
   txtTAXA.SetFocus
   If wl_Retorno <> "" Then
      txtTAXA.Text = wl_Retorno
      SendKeys "{ENTER}"
   End If
End If

End Sub


Private Sub txtTAXA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If txtTAXA.Text <> "" And txtTAXA.Text <> "0" Then
      If Not Loca_Contas(tbPlano, txtTAXA.Text) Then
         InformaaoUsuario "Conta não encontrada"
         txtTAXA.SetFocus
         HomeEnd
         Exit Sub
      End If
      lblTAXA.Caption = tbPlano("DESCRICAO")
   Else
      lblTAXA.Caption = ""
   End If
   SendKeys "{TAB}"
End If

End Sub


Private Sub txtTAXA_LostFocus()
aviso
End Sub


