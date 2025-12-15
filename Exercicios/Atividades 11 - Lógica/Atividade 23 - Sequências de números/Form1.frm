VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'10- Escreva um algoritmo que imprima as seguintes sequências de números:
'(1, 1 2 3 4 5 6 7 8 9 10)
'(2, 1 2 3 4 5 6 7 8 9 10)
'(3, 1 2 3 4 5 6 7 8 9 10)
'(4, 1 2 3 4 5 6 7 8 9 10)
'e assim sucessivamente, até que o primeiro número (antes da vírgula), também chegue a 10.

Private Sub Form_Load()

   Me.Text1.Text = ""
   
   Dim linha As String
   
   For i = 1 To 10
      linha = linha & i & ", "
      For j = 1 To 10
         linha = linha & " " & j
      Next j
      linha = linha & vbNewLine
   Next i
   
   Me.Text1.Text = linha
End Sub
