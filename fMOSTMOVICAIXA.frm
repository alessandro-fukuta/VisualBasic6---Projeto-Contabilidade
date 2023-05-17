VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fMostMoviCaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fluxo de Caixa"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grdFLUXO 
      Height          =   3645
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   6429
      _Version        =   393216
      Rows            =   15
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483638
      GridLines       =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "fMostMoviCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tbMoviCaixa As Recordset
Private Function Desenha_Grade()
Dim wl_Linhas As Integer
Dim wl_Saldo As Currency
tbMoviCaixa.Seek ">=", CDate(fCAIXA.txtDATA.Pacote), 1
If tbMoviCaixa.NoMatch Then
   Unload Me
   Exit Function
End If
If tbMoviCaixa("DATA") <> CDate(fCAIXA.txtDATA.Pacote) Then
   Unload Me
   Exit Function
End If
grdFluxo.ColWidth(0) = 1000
grdFluxo.TextMatrix(0, 0) = "Lançto."
grdFluxo.ColWidth(1) = 1000

grdFluxo.TextMatrix(0, 1) = IIf(pb_InverteOperacao, "Débito", "Crédito")
grdFluxo.ColWidth(2) = 1500
grdFluxo.TextMatrix(0, 2) = IIf(pb_InverteOperacao, "Crédito", "Débito")
grdFluxo.ColWidth(3) = 7000
grdFluxo.TextMatrix(0, 3) = "Histórico"
grdFluxo.ColAlignment(3) = cnEsquerda
wl_Linhas = 1
Do While Not tbMoviCaixa.EOF
   If tbMoviCaixa("DATA") <> CDate(fCAIXA.txtDATA.Pacote) Then
      Exit Do
   End If
   If wl_Linhas = grdFluxo.rows Then
      grdFluxo.rows = grdFluxo.rows + 1
   End If
   grdFluxo.TextMatrix(wl_Linhas, 0) = tbMoviCaixa("MOVIMENTO")
   If tbMoviCaixa("CREDITO") <> 0 Then
      grdFluxo.TextMatrix(wl_Linhas, 1) = Format(tbMoviCaixa("VALOR"), "##,###,##0.00;(#,###,##0.00)")
      If pb_InverteOperacao Then
         grdFluxo.row = wl_Linhas
         grdFluxo.Col = 2
         grdFluxo.CellForeColor = VERMELHO
         wl_Saldo = wl_Saldo - tbMoviCaixa("VALOR")
      Else
         wl_Saldo = wl_Saldo + tbMoviCaixa("VALOR")
      End If
   End If
   If tbMoviCaixa("DEBITO") <> 0 Then
      grdFluxo.TextMatrix(wl_Linhas, 2) = Format(tbMoviCaixa("VALOR"), "##,###,##0.00;(#,###,##0.00)")
      If Not pb_InverteOperacao Then
         grdFluxo.row = wl_Linhas
         grdFluxo.Col = 2
         grdFluxo.CellForeColor = VERMELHO
         wl_Saldo = wl_Saldo - tbMoviCaixa("VALOR")
      Else
         wl_Saldo = wl_Saldo + tbMoviCaixa("VALOR")
      End If
   End If
   grdFluxo.TextMatrix(wl_Linhas, 3) = tbMoviCaixa("HISTORICO")
   tbMoviCaixa.MoveNext
   wl_Linhas = wl_Linhas + 1
Loop
grdFluxo.row = wl_Linhas - 1
grdFluxo.Col = 0
SendKeys "{DOWN}"
grdFluxo.Visible = True
End Function


Private Sub Form_Activate()
If Not Abre_MoviCaixa(tbMoviCaixa) Then
   Unload Me
   Exit Sub
End If
Caption = "Movimento de Caixa. Movimento do dia " + CStr(fCAIXA.txtDATA.Pacote)
aviso "Aguarde. Montando grade de pesquisa ..."

Me.MousePointer = 11
DoEvents
Call Desenha_Grade
aviso
Me.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tbMoviCaixa.Close
End Sub

Private Sub grdFLUXO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   pbRetornoVideo = grdFluxo.TextMatrix(grdFluxo.row, 0)
   Unload Me
ElseIf KeyAscii = 27 Then
   KeyAscii = 0
   pbRetornoVideo = ""
   Unload Me
End If
End Sub


