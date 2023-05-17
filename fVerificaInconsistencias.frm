VERSION 5.00
Object = "{9FDDA49F-0DDF-4F9B-AEFC-DAFB8A5CDE9E}#1.0#0"; "MASCARA.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fVerificaInconsistencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verifica Inconsistências"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   200
      Cols            =   3
      FixedCols       =   0
      ForeColor       =   255
      TextStyleFixed  =   3
      GridLines       =   3
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   "DATA       | Lançamento | Erro encontrado                                                                              "
   End
   Begin Mascara.Máscara txtinicio 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
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
   Begin Mascara.Máscara txtfim 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial e Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "fVerificaInconsistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tblanca As Recordset
Dim xcredito As Long
Dim xdebito As Long
Dim tbconta As Recordset
Dim werro As Boolean
Dim erro As String

Private Sub Form_Load()
centraobj Me
If Not Abre_MoviCaixa(tblanca) Or Not _
       Abre_PlanoContas(tbconta) Then

    MsgBox "Impossível verifica inconsistência nesse momento !", vbInformation
    Unload Me
    Exit Sub

End If


End Sub


Private Sub txtfim_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   
   If txtfim.Pacote < txtinicio.Pacote Then
      MsgBox "Datas Inválidas !", vbInformation
      Me.txtinicio.SetFocus
      Exit Sub
   End If
   
   tblanca.Index = "iDATA"
   
   tblanca.Seek ">=", txtinicio.Pacote, 0
   
   Do While Not tblanca.EOF
   
        If tblanca.EOF Or tblanca.NoMatch Then
           Exit Do
        End If
        
        If tblanca("data") >= txtinicio.Pacote And tblanca("data") <= txtfim.Pacote Then
        
           werro = False
           erro = ""
         
           xcredito = tblanca("credito")
           xdebito = tblanca("debito")
           
           tbconta.Index = "iTRADUTOR"
           tbconta.Seek "=", xcredito
           If tbconta.NoMatch Then
              werro = True
              erro = "Conta crédito informada não existe."
           End If
           
           tbconta.Index = "iTRADUTOR"
           tbconta.Seek "=", xdebito
           If tbconta.NoMatch Then
              werro = True
              erro = "Conta débito informada não existe."
           End If
                    
           If werro Then
              Grid1.TextMatrix(Grid1.row, 0) = tblanca("data")
              Grid1.TextMatrix(Grid1.row, 1) = tblanca("movimento")
              Grid1.TextMatrix(Grid1.row, 2) = erro
              If Grid1.row + 1 > 199 Then
                 Grid1.rows = Grid1.row + 2
              End If
              Grid1.row = Grid1.row + 1
           End If
           
        
        End If
   
   
        tblanca.MoveNext
         
        werro = False
   
   Loop
   
   
End If
End Sub

Private Sub txtinicio_GotFocus()
LimpaCaixasTexto Me
End Sub

