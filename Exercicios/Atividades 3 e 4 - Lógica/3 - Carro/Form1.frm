VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ferrari"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   2535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCarro 
      Caption         =   "Carro"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton cmdCustoMovimento 
         Caption         =   "Calcular Custo"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Text            =   "(L)"
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdAbastecer 
         Caption         =   "Abastecer"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Text            =   "(KM)"
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton cmdAndar 
         Caption         =   "Andar"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Text            =   "(KM)"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblCustoMovimento 
         Caption         =   "Custo para movimentar"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblAbastece 
         Caption         =   "Abastecer"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMovimento 
         Caption         =   "Movimentar o carro"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1680
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'03- Implemente um carro.
'O tanque de combustível do carro armazena no máximo 50 litros de gasolina. --> If
'O carro consome 15 km/litro. Deve ser possível:
'a) Abastecer o carro com uma certa quantidade de gasolina; OK
'b) Mover o carro em uma determinada distância (medida em km); OK
'c) Retornar a quantidade de combustível e a distância total percorrida. OK
'd) Utilizar um controle para inserir o preço da gasolina em litros por quilômetro e calcular o custo do deslocamento. OK
Dim Carro As New clsCarro

Private Sub cmdAbastecer_Click()
   
   If IsNumeric(Text1.Text) Then
      Carro.Abastece (Text1.Text)
      msg = MsgBox("Carro abastecido com " & Text1.Text & "L", 48, "Carro")
   End If
   
End Sub

Private Sub cmdAndar_Click()
   If IsNumeric(Text2.Text) And (Text2.Text > 0) Then
      Carro.MovimentarCarro (CDbl(Text2.Text))
   Else
      msg = MsgBox("O valor para percorrer deve ser maior que 0", 48, "ATENÇÃO")
   End If
End Sub

Private Sub cmdCustoMovimento_Click()
   If IsNumeric(Text5.Text) And (Text5.Text > 0) Then
      Carro.CustoParaMovimentar (CDbl(Text5.Text))
      msg = MsgBox("Para percorrer " & Text5.Text & "KM, você deve abastecer " & Carro.dblCusto & " Litros", 46, "Informação")
   Else
      msg = MsgBox("O valor para percorrer deve ser maior que 0", 48, "ATENÇÃO")
   End If
End Sub

Private Sub Text2_Click()
    Text2.Text = " "
End Sub

Private Sub Text1_Click()
    Text1.Text = " "
End Sub


Private Sub Text5_Click()
   Text5.Text = ""
End Sub
