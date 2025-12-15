VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "frm1"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   4335
   End
   Begin VB.TextBox txtBox 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4- Escreva um programa que calcule os quadrados e os cubos dos números de 0 a 10 e imprima
'os valores resultantes em forma de tabela.
'Obs_1: o programa não possui entrada de dados.
'Obs_2: obrigatório o uso de laço de repetição
'Obs_3: obrigatório o uso de array ou de collection
'saída:
'  número quadrado cubos
'  0 0 0
'  1 1 1
'  2 4 8
'  3 9 27
'  4 16 64
'  5 25 125
'  6 36 216
'  7 49 343
'  8 64 512
'  9 81 729
'  10 100 1000

Private Sub Command1_Click()
   Dim colNumeros(0 To 9)
   Dim i, j As Integer
   Dim varAux As String
   
   For i = 0 To 9
      For j = 0 To i + 1
         varAux = CStr(j & "   " & (j ^ 2) & "   " & (j ^ 3) & vbCrLf)
         colNumeros(i) = colNumeros(i) & varAux
      Next j
   txtBox.Text = colNumeros(i)
   Next i
End Sub



