VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fMOSTUSUARIOS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Usuários"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grdUSUARIOS 
      Height          =   3645
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   6429
      _Version        =   393216
      Rows            =   15
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorSel    =   -2147483633
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483638
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "fMOSTUSUARIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function Inicia_Grade()
Dim tbUsuarios As Recordset
Dim row As Integer
If Not Abre_Usuarios(tbUsuarios) Then
   Unload Me
End If
If tbUsuarios.RecordCount = 0 Then
   Unload Me
End If
grdUSUARIOS.TextMatrix(0, 0) = "Nome do Usuário"
row = 1
Do While Not tbUsuarios.EOF
   grdUSUARIOS.TextMatrix(row, 0) = xtoc(tbUsuarios("NOME"))
   row = row + 1
   tbUsuarios.MoveNext
Loop
End Function


Private Sub Form_Activate()
Call Inicia_Grade
End Sub

Private Sub Form_Load()
grdUSUARIOS.ColWidth(0) = 6000
End Sub


Private Sub grdUSUARIOS_DblClick()
SendKeys "{enter}"
End Sub

Private Sub grdUSUARIOS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   pbRetornoVideo = grdUSUARIOS.TextMatrix(grdUSUARIOS.row, 0)
   ShowRetorno = True
   Unload Me
ElseIf KeyAscii = 27 Then
   pbRetornoVideo = ""
   Unload Me
End If
End Sub


