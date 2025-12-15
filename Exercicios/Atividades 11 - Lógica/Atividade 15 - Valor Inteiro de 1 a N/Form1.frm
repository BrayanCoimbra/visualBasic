VERSION 5.00
Begin VB.Form frmCount 
   Caption         =   "Count"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private colResultado As Collection

Private Sub Command1_Click()
   Dim N As Integer
   N = Val(InputBox("Digite o valor de N:", "Digite um número"))
   
   Me.Text1.Text = ""
   
   If N > 0 Then
      Set colResultado = New Collection

      ' Adiciona todos os valores inteiros entre 1 e N à coleção
      For i = 1 To N
         Me.Text1.Text = Me.Text1.Text & i & vbNewLine
      Next i
      
      Set colResultado = Nothing
   Else
      MsgBox "O valor de N deve ser maior que zero.", vbInformation, "Atenção"
   End If
End Sub

