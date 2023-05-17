VERSION 5.00
Begin VB.Form fInformacao 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   ClientHeight    =   645
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fInformacao.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   45
      ScaleHeight     =   525
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   60
      Width           =   8055
      Begin VB.Label lblMensagem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<< Mensagem >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   1
         Top             =   135
         Width           =   8055
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "fInformacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

