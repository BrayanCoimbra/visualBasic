VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Text            =   "Digite a altura do losango"
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton btnLosango 
      Caption         =   "Formar Losango"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3- Escreva um programa que leia um número inteiro maior do que 4 que seja capaz de produzir um losango como saída:
'Exemplo:
'        *
'       ***
'      *****
'       ***
'        *

Private Sub btnLosango_Click()
   Dim colLinhas() As String
   Dim altura As String
   Dim Linha As String
   Dim Inicio As Integer
   Dim Fim As Integer
   Dim Tamanho As Integer
   Dim i As Integer
   Dim j As Integer
   i = 1
   j = 1
      
   For i = 1 To CInt(Me.Text2.Text)
      For j = 1 To i
         Me.Text1.Text = Me.Text1.Text & "*"
      Next j
      Me.Text1.Text = Me.Text1.Text & vbNewLine
   Next i

   For i = CInt(Me.Text2.Text) - 1 To 1 Step -1
      For j = i To 1 Step -1
         Me.Text1.Text = Me.Text1.Text & "*"
      Next j
      Me.Text1.Text = Me.Text1.Text & vbNewLine
   Next i
    
End Sub

Private Sub btnSair_Click()
   End
End Sub

Private Sub Text2_Click()
   Text2.Text = " "
End Sub
