VERSION 5.00
Begin VB.Form frmEditarLivro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEditarLivro 
      Caption         =   "Editar Livros"
      Height          =   3855
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      Begin VB.CommandButton cmdSalvarEdicoes 
         Caption         =   "Salvar Edições"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txtEditarDataPubli 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtEditarAutor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtEditarEdicao 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtEditarNome 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblEditarNome 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblEditaroEdicao 
         BackStyle       =   0  'Transparent
         Caption         =   "Edição"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblEditarAutor 
         BackStyle       =   0  'Transparent
         Caption         =   "Autor"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   135
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   15
      End
      Begin VB.Label lblEditarDataPubli 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Publicação"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmEditarLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
