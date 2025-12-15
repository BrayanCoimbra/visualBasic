VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnLosango 
      Caption         =   "Losango"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton btnSeta 
      Caption         =   "Seta"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton btnElipse 
      Caption         =   "Elipse"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton btnRetangulo 
      Caption         =   "Retângulo"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "A figura aparecerá aqui"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'1- Escreva um aplicativo que apresente como saída as seguintes figuras geométricas:
'a) um retângulo
'b) uma elipse
'c) uma seta para cima
'd) um losango

'obs.: Todas as figuras devem ser realizadas com o caractere '*'.

'************           ***               *               *
'*          *         *     *            ***             * *
'*          *        *       *          *****           *   *
'*          *        *       *            *            *     *
'*          *        *       *            *           *       *
'*          *        *       *            *            *     *
'*          *        *       *            *             *   *
'*          *         *     *             *              * *
'************           ***                               *

Private Sub btnElipse_Click()
   Dim colLinhas(1 To 8) As String
   Dim Linha As String
   Dim Inicio As Integer
   Dim Fim As Integer
   Dim Tamanho As Integer
   Dim i As Integer
   Dim j As Integer
   i = 1
   j = 1
   'Limpa o Label
   
   Label1.Caption = ""
   
   Label1.Caption = "*"
   
   For i = 1 To 5
      Label1.Caption = Label1.Caption & vbNewLine
      Label1.Caption = Label1.Caption & "*"
      For j = 1 To i
         Label1.Caption = Label1.Caption & " "
      Next j
      Label1.Caption = Label1.Caption & "*"
   Next i
   
   For i = 5 To 1 Step -1
      Label1.Caption = Label1.Caption & vbNewLine
      Label1.Caption = Label1.Caption & "*"
      For j = i To 1 Step -1
         Label1.Caption = Label1.Caption & " "
      Next j
      Label1.Caption = Label1.Caption & "*"
   Next i
   
   Label1.Caption = Label1.Caption & vbNewLine & "*"
   
End Sub

Private Sub btnLosango_Click()
   Dim colLinhas(1 To 5) As String
   Dim Linha As String
   Dim Inicio As Integer
   Dim Fim As Integer
   Dim Tamanho As Integer
   Dim i As Integer
   Dim j As Integer
   i = 1
   j = 1
   
   'Limpa o Label
   Label1.Caption = ""
   
   For i = 1 To 6
      For j = 1 To i
         Label1.Caption = Label1.Caption & "*"
      Next j
      Label1.Caption = Label1.Caption & vbNewLine
   Next i

   For i = 6 To 1 Step -1
      For j = i To 1 Step -1
         Label1.Caption = Label1.Caption & "*"
      Next j
      Label1.Caption = Label1.Caption & vbNewLine
   Next i
        
End Sub

Private Sub btnRetangulo_Click()
   Dim colLinhas(1 To 10) As String
   Dim Linha As String
   Dim Inicio As Integer
   Dim Fim As Integer
   Dim Tamanho As Integer
   Dim i As Integer
   Dim j As Integer
   i = 1
   j = 1
   
   Label1.Caption = ""
   Label1.Caption = "**********" & vbNewLine
   
   For i = 1 To 9
      Label1.Caption = Label1.Caption & "*        *" & vbNewLine
   Next i
   
   Label1.Caption = vbNewLine & Label1.Caption & "**********"

End Sub


Private Sub btnSeta_Click()
   Dim Linha, linha1, linha2, linha3, linha4, linha5 As String
   Dim colLinhas(1 To 3) As String
   Dim Tamanho As Integer
   Dim Inicio As Integer
   Dim Fim As Integer
   Dim i As Integer
   Dim j As Integer
   i = 1
   j = 1
    
   'Limpa o Label
   Label1.Caption = ""
   
   For i = 1 To 3
      For j = 1 To i
         Label1.Caption = Label1.Caption & "*"
      Next j
      Label1.Caption = Label1.Caption & vbNewLine
   Next i
      
   For i = 1 To 7
     Label1.Caption = Label1.Caption & "     *     " & vbNewLine
   Next i

End Sub
