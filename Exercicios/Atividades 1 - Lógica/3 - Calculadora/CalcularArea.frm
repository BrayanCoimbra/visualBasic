VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CalcularArea"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   8700
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Text            =   "A area do circulo aparecerá aqui"
      Top             =   2280
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular área do ciruclo"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Text            =   "Digite o valor do Raio"
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtBoxValorPI 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
   If IsNumeric(txtBoxValor1.Text) And (txtBoxValor2.Text) Then
      dblArea = txtBoxValor1 * (Raio ^ 2)
      txtBoxResultado = "A área do circulo é " & dblArea
   Else
      varAux = MsgBox("Averiguar código", 48, "Erro")
   End If
   
End Sub

Private Sub txtBoxValorPI_Change()
   txtBoxValorPI.Text = ValorPI
End Sub
