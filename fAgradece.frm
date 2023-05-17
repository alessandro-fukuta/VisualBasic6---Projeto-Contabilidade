VERSION 5.00
Begin VB.Form fAgradece 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7260
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fAgradece.frx":0000
   ScaleHeight     =   7260
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Fecha"
      Height          =   420
      Left            =   8040
      TabIndex        =   1
      Top             =   6480
      Width           =   1410
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parabéns,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"fAgradece.frx":D5D0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   2160
      TabIndex        =   0
      Top             =   1680
      Width           =   5595
   End
End
Attribute VB_Name = "fAgradece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

