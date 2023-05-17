VERSION 5.00
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Begin VB.Form fCONFEREMOVIMENTO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conferência do Movimento de Contas"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   ClipControls    =   0   'False
   Icon            =   "fCONFEREMOVIMENTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "&Imprime"
      Height          =   375
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1245
   End
   Begin Mascara.Máscara txtDATAINICIAL 
      Height          =   300
      Left            =   75
      TabIndex        =   1
      Top             =   315
      Width           =   1020
      _ExtentX        =   1799
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
      Format          =   "dd/mmm/yyyy"
      Mask            =   "##/##/####"
      Text            =   ""
      ÉData           =   -1  'True
   End
   Begin Mascara.Máscara txtDATAFINAL 
      Height          =   300
      Left            =   1155
      TabIndex        =   3
      Top             =   315
      Width           =   1020
      _ExtentX        =   1799
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
      Format          =   "dd/mmm/yyyy"
      Mask            =   "##/##/####"
      Text            =   ""
      ÉData           =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
      Height          =   195
      Left            =   1140
      TabIndex        =   2
      Top             =   105
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   795
   End
End
Attribute VB_Name = "fCONFEREMOVIMENTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tbMoviCaixa As Recordset
Private tbContas As Recordset

Private Sub Imprime_Conferencia()
Dim wl_Linha As Currency
Dim wl_Imprime As Boolean
Dim pp As Currency
Dim aCampo
Dim aReferencia

aadd aCampo, Array("Data", "Movto.", "Conta", "Valor")
aadd aReferencia, Array("99.99.9999", "99999", "99999.MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM", "9,999,999.99")

tbMoviCaixa.Seek ">=", CDate(txtDATAINICIAL.Pacote), 0
If tbMoviCaixa.NoMatch Then
   InformaaoUsuario "Período não encontrado"
   Exit Sub
End If
If tbMoviCaixa("DATA") > CDate(txtDATAFINAL.Pacote) Then
   InformaaoUsuario "Período não encontrado"
   Exit Sub
End If
If Not PadraodeImpressao Then
   Exit Sub
End If
Do While Not tbMoviCaixa.EOF
   If tbMoviCaixa("DATA") > CDate(txtDATAFINAL.Pacote) Then
      Exit Do
   End If
   If wl_Linha = 0 Then GoSub CABECALHO
   Imprime wl_Linha, 0, "Data: " + CStr(tbMoviCaixa("DATA")) + "   -   Movimento: " + Str(tbMoviCaixa("MOVIMENTO"))
   wl_Linha = wl_Linha + 0.5
   If tbMoviCaixa("DEBITO") > 0 Then
      If Loca_Contas(tbContas, tbMoviCaixa("DEBITO")) Then
         Imprime wl_Linha, 0, "Débito :" + Format(tbMoviCaixa("DEBITO"), "00000") + " - " + tbContas("DESCRICAO")
      Else
         Imprime wl_Linha, 0, "Débito :" + Format(tbMoviCaixa("DEBITO"), "00000")
      End If
      wl_Linha = wl_Linha + 0.5
   End If
   If tbMoviCaixa("CREDITO") > 0 Then
      If Loca_Contas(tbContas, tbMoviCaixa("CREDITO")) Then
         Imprime wl_Linha, 0, "Crédito:" + Format(tbMoviCaixa("CREDITO"), "00000") + " - " + tbContas("DESCRICAO")
      Else
         Imprime wl_Linha, 0, "Crédito:" + Format(tbMoviCaixa("CREDITO"), "00000")
      End If
      wl_Linha = wl_Linha + 0.5
   End If
   Imprime wl_Linha, 0, "Valor: " + Format(tbMoviCaixa("VALOR"), "##,###,##0.00")
   wl_Linha = wl_Linha + 0.5
   Imprime wl_Linha, 0, "Historico: " + tbMoviCaixa("HISTORICO")
   wl_Linha = wl_Linha + 1
   If wl_Linha > IIf(pb_ImpressaoMatricial, 29, 26) Then
      Salta_Pagina
      wl_Linha = 0
   End If
   tbMoviCaixa.MoveNext
Loop
Finaliza_Impressao
Exit Sub

CABECALHO:
pp = pp + 1
Imprime 0, 1, pb_RAZAOSOCIAL ', imp_Condensado_NEGRITO
Imprime 0, 16, "Folha .." + Str(pp), , "D"
Imprime 0.5, 1, "Relatório Simples Conferência Movimento de Contas" ', imp_Condensado
Imprime 1, 1, "Emissao em " + CStr(Date) + "  -> Período de " + CStr(txtDATAINICIAL.Pacote) + "  a  " + CStr(txtDATAFINAL.Pacote) ', imp_Condensado

wl_Linha = wl_Linha + 2
'Monta_Cabecalho aCampo, aReferencia, 3, wl_Linha ', imp_Condensado
'wl_Linha = wl_Linha + 1
Return
End Sub


Private Sub Command1_Click()
If Not IsDate(txtDATAINICIAL.Pacote) Then
   InformaaoUsuario "Informe a data inicial corretamente"
   txtDATAINICIAL.SetFocus
   Exit Sub
End If
If Not IsDate(txtDATAFINAL.Pacote) Then
   InformaaoUsuario "Informe a data final corretamente"
   txtDATAFINAL.SetFocus
   Exit Sub
End If
If MsgBox("Confirma o período?", vbQuestion + vbYesNo, "Mensagem do Sistema") = vbYes Then
   Imprime_Conferencia
End If
txtDATAINICIAL.SetFocus
End Sub


Private Sub Form_Activate()
If Not Abre_MoviCaixa(tbMoviCaixa) Or _
   Not Abre_PlanoContas(tbContas) Then
   Unload Me
   Exit Sub
End If



End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
   KeyAscii = 0
   If UCase(Me.ActiveControl.Name) = "TXTDATAINICIAL" Then
      Unload Me
   Else
      txtDATAINICIAL.SetFocus
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
tbMoviCaixa.Close
tbContas.Close
End Sub

Private Sub txtDATAINICIAL_GotFocus()
txtDATAINICIAL.Text = Date
txtDATAFINAL.Text = Date

End Sub

