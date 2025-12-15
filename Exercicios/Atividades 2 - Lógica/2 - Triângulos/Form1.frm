VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   6510
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnFormarFiguras 
      Caption         =   "Formar Figuras"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2- Escreva um programa que leia um inteiro de entrada e utilize laço de repetição
'capaz de produzir as seguintes formas geométricas de saída simultaneamente.
'*       ***      ***       *
'**      **        **      **
'***     *          *     ***

Private Sub btnFormarFiguras_Click()
Dim colLinhas() As String
Dim Linha As String
Dim Inicio As Integer
Dim Fim As Integer
Dim Tamanho As Integer
Dim Numero As Integer
Dim i As Integer
Dim j As Integer
i = 1
j = 1


   For i = 1 To 3
      For k = 1 To i
         Label1.Caption = Label1.Caption & " "
      Next k
      
      For j = 1 To i
         Label1.Caption = Label1.Caption & "*"
      Next j
      
      Label1.Caption = Label1.Caption & vbNewLine
   Next i
   
   For i = 1 To 3
      For j = 3 To i Step -1
         Label2.Caption = Label2.Caption & "*"
      Next j
      
      For k = i To 1 Step -1
         Label2.Caption = Label2.Caption & " "
      Next k
      Label2.Caption = Label2.Caption & vbNewLine
   Next i
   
   For i = 1 To 3
      For k = i To 1 Step -1
         Label3.Caption = Label3.Caption & " "
      Next k
      
      For j = 3 To i Step -1
         Label3.Caption = Label3.Caption & "*"
      Next j
      
      Label3.Caption = Label3.Caption & vbNewLine
   Next i
   
   For i = 1 To 3
      For j = 1 To i
         Label4.Caption = Label4.Caption & "*"
      Next j
      
      For k = 1 To i
         Label4.Caption = Label4.Caption & " "
      Next k
      
      Label4.Caption = Label4.Caption & vbNewLine
   Next i
   
End Sub

Private Sub btnSair_Click()
   End
End Sub

Private Sub txtBox_Click()
   txtBox.Text = " "
   Label1.Caption = " "
   Label2.Caption = " "
   Label3.Caption = " "
   Label4.Caption = " "
End Sub

