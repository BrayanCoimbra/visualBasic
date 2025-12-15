VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000011&
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   3600
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Estimar Massa Muscular"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A reta de regressão para a relação entre as variáveis Y e X."
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular coeficiente de correlação linear R"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim dblcolMassa(1 To 18) As Double ' armazena os valores iniciais de massa
Dim dblcolIdade(1 To 18) As Double ' armazena os valores iniciais de idade
Dim dblcolIdadeQuadrado(1 To 18) As Double ' armazena os valores de idade ao quadrado
Dim dblcolIdadeXMassa(1 To 18) As Double ' armazena os valores de idade X massa
Dim dblcolMassaQuadrado(1 To 18) As Double ' armazena os valores de massa ao quadrado
Dim dblcolResultadoMassaMenosMassaMedia(1 To 18) As Double ' armazena os valores da massa menos a média da massa
Dim dblcolResultadoIdadeMenosIdadeMedia(1 To 18) As Double  ' armazena os valores da idade menos a média da idade
Dim dblcolSomatorioIdadeXMassa(1 To 18) As Double
Dim dblcolSomatorioTotalMassaAoQuadrado(1 To 18) As Double
Dim dblcolSomatorioTotalIdadeAoQuadrado(1 To 18) As Double
Dim intMediaMassaMuscular As Integer
Dim intMediaIdade As Integer
Dim r As Double

intMediaMassaMuscular = 85
intMediaIdade = 61.556

dblcolMassa(1) = 82
dblcolMassa(2) = 91
dblcolMassa(3) = 100
dblcolMassa(4) = 68
dblcolMassa(5) = 87
dblcolMassa(6) = 73
dblcolMassa(7) = 78
dblcolMassa(8) = 80
dblcolMassa(9) = 65
dblcolMassa(10) = 84
dblcolMassa(11) = 116
dblcolMassa(12) = 76
dblcolMassa(13) = 97
dblcolMassa(14) = 100
dblcolMassa(15) = 105
dblcolMassa(16) = 77
dblcolMassa(17) = 73
dblcolMassa(18) = 78

dblcolIdade(1) = 71
dblcolIdade(2) = 64
dblcolIdade(3) = 43
dblcolIdade(4) = 67
dblcolIdade(5) = 56
dblcolIdade(6) = 73
dblcolIdade(7) = 68
dblcolIdade(8) = 56
dblcolIdade(9) = 73
dblcolIdade(10) = 65
dblcolIdade(11) = 45
dblcolIdade(12) = 58
dblcolIdade(13) = 45
dblcolIdade(14) = 53
dblcolIdade(15) = 49
dblcolIdade(16) = 78
dblcolIdade(17) = 73
dblcolIdade(18) = 68

For i = 1 To 18
   'armazena Xi - Xmédia
   dblcolResultadoMassaMenosMassaMedia(i) = dblcolMassa(i) - intMediaMassaMuscular
Next i

For i = 1 To 18
   'armazena Yi - Ymédia
   dblcolResultadoIdadeMenosIdadeMedia(i) = dblcolIdade(i) - intMediaIdade
Next i

For i = 1 To 18
   'armazena (Xi - Xmédia) X (Yi - Ymédia)
   dblcolIdadeXMassa(i) = dblcolResultadoIdadeMenosIdadeMedia(i) * dblcolResultadoMassaMenosMassaMedia(i)
Next i

For i = 1 To 18
   'armazena (Xi - Xmédia)^2
   dblcolIdadeQuadrado(i) = dblcolResultadoIdadeMenosIdadeMedia(i) ^ 2
Next i

For i = 1 To 18
   'armazena (Yi - Ymédia)^2
   dblcolMassaQuadrado(i) = dblcolResultadoMassaMenosMassaMedia(i) ^ 2
Next i

For i = 1 To 18
   Var = Var + dblcolIdadeXMassa(i)
   dblcolSomatorioIdadeXMassa(i) = Var
Next i

For i = 1 To 18
   var2 = var2 + dblcolMassaQuadrado(i)
   dblcolSomatorioTotalMassaAoQuadrado(i) = var2
Next i

For i = 1 To 18
   var3 = var3 + dblcolIdadeQuadrado(i)
   dblcolSomatorioTotalIdadeAoQuadrado(i) = var3
Next i

r = dblcolSomatorioIdadeXMassa(18) / (Sqr(dblcolSomatorioTotalIdadeAoQuadrado(18) * dblcolSomatorioTotalMassaAoQuadrado(18)))

Text1.Text = CStr(r)

End Sub

Private Sub Command2_Click()
Dim dblcolMassa(1 To 18) As Double ' armazena os valores iniciais de massa
Dim dblcolIdade(1 To 18) As Double ' armazena os valores iniciais de idade
Dim dblcolIdadeQuadrado(1 To 18) As Double ' armazena os valores de idade ao quadrado
Dim dblcolIdadeXMassa(1 To 18) As Double ' armazena os valores de idade X massa
Dim dblcolMassaQuadrado(1 To 18) As Double ' armazena os valores de massa ao quadrado
Dim dblcolResultadoMassaMenosMassaMedia(1 To 18) As Double ' armazena os valores da massa menos a média da massa
Dim dblcolResultadoIdadeMenosIdadeMedia(1 To 18) As Double  ' armazena os valores da idade menos a média da idade
Dim dblcolSomatorioIdadeXMassa(1 To 18) As Double
Dim dblcolSomatorioTotalMassaAoQuadrado(1 To 18) As Double
Dim dblcolSomatorioTotalIdadeAoQuadrado(1 To 18) As Double
Dim intMediaMassaMuscular As Integer
Dim intMediaIdade As Integer
Dim r As Double

intMediaMassaMuscular = 85
intMediaIdade = 61.556

dblcolMassa(1) = 82
dblcolMassa(2) = 91
dblcolMassa(3) = 100
dblcolMassa(4) = 68
dblcolMassa(5) = 87
dblcolMassa(6) = 73
dblcolMassa(7) = 78
dblcolMassa(8) = 80
dblcolMassa(9) = 65
dblcolMassa(10) = 84
dblcolMassa(11) = 116
dblcolMassa(12) = 76
dblcolMassa(13) = 97
dblcolMassa(14) = 100
dblcolMassa(15) = 105
dblcolMassa(16) = 77
dblcolMassa(17) = 73
dblcolMassa(18) = 78

dblcolIdade(1) = 71
dblcolIdade(2) = 64
dblcolIdade(3) = 43
dblcolIdade(4) = 67
dblcolIdade(5) = 56
dblcolIdade(6) = 73
dblcolIdade(7) = 68
dblcolIdade(8) = 56
dblcolIdade(9) = 73
dblcolIdade(10) = 65
dblcolIdade(11) = 45
dblcolIdade(12) = 58
dblcolIdade(13) = 45
dblcolIdade(14) = 53
dblcolIdade(15) = 49
dblcolIdade(16) = 78
dblcolIdade(17) = 73
dblcolIdade(18) = 68

For i = 1 To 18
   'armazena Xi - Xmédia
   dblcolResultadoMassaMenosMassaMedia(i) = dblcolMassa(i) - intMediaMassaMuscular
Next i

For i = 1 To 18
   'armazena Yi - Ymédia
   dblcolResultadoIdadeMenosIdadeMedia(i) = dblcolIdade(i) - intMediaIdade
Next i

For i = 1 To 18
   'armazena (Xi - Xmédia) X (Yi - Ymédia)
   dblcolIdadeXMassa(i) = dblcolResultadoIdadeMenosIdadeMedia(i) * dblcolResultadoMassaMenosMassaMedia(i)
Next i

For i = 1 To 18
   'armazena (Xi - Xmédia)^2
   dblcolIdadeQuadrado(i) = dblcolResultadoIdadeMenosIdadeMedia(i) ^ 2
Next i

For i = 1 To 18
   'armazena (Yi - Ymédia)^2
   dblcolMassaQuadrado(i) = dblcolResultadoMassaMenosMassaMedia(i) ^ 2
Next i

For i = 1 To 18
   Var = Var + dblcolIdadeXMassa(i)
   dblcolSomatorioIdadeXMassa(i) = Var
Next i

For i = 1 To 18
   var2 = var2 + dblcolMassaQuadrado(i)
   dblcolSomatorioTotalMassaAoQuadrado(i) = var2
Next i

For i = 1 To 18
   var3 = var3 + dblcolIdadeQuadrado(i)
   dblcolSomatorioTotalIdadeAoQuadrado(i) = var3
Next i

b1 = dblcolSomatorioIdadeXMassa(18) / dblcolSomatorioTotalIdadeAoQuadrado(18)
b0 = intMediaMassaMuscular - (b1 * intMediaIdade)
resultado = b0 & " " & b1 & "X"

Text2.Text = CStr(resultado)

End Sub

Private Sub Command3_Click()
Dim dblcolMassa(1 To 18) As Double ' armazena os valores iniciais de massa
Dim dblcolIdade(1 To 18) As Double ' armazena os valores iniciais de idade
Dim dblcolIdadeQuadrado(1 To 18) As Double ' armazena os valores de idade ao quadrado
Dim dblcolIdadeXMassa(1 To 18) As Double ' armazena os valores de idade X massa
Dim dblcolMassaQuadrado(1 To 18) As Double ' armazena os valores de massa ao quadrado
Dim dblcolResultadoMassaMenosMassaMedia(1 To 18) As Double ' armazena os valores da massa menos a média da massa
Dim dblcolResultadoIdadeMenosIdadeMedia(1 To 18) As Double  ' armazena os valores da idade menos a média da idade
Dim dblcolSomatorioIdadeXMassa(1 To 18) As Double
Dim dblcolSomatorioTotalMassaAoQuadrado(1 To 18) As Double
Dim dblcolSomatorioTotalIdadeAoQuadrado(1 To 18) As Double
Dim intMediaMassaMuscular As Integer
Dim intMediaIdade As Integer
Dim r As Double

intMediaMassaMuscular = 85
intMediaIdade = 61.556

dblcolMassa(1) = 82
dblcolMassa(2) = 91
dblcolMassa(3) = 100
dblcolMassa(4) = 68
dblcolMassa(5) = 87
dblcolMassa(6) = 73
dblcolMassa(7) = 78
dblcolMassa(8) = 80
dblcolMassa(9) = 65
dblcolMassa(10) = 84
dblcolMassa(11) = 116
dblcolMassa(12) = 76
dblcolMassa(13) = 97
dblcolMassa(14) = 100
dblcolMassa(15) = 105
dblcolMassa(16) = 77
dblcolMassa(17) = 73
dblcolMassa(18) = 78

dblcolIdade(1) = 71
dblcolIdade(2) = 64
dblcolIdade(3) = 43
dblcolIdade(4) = 67
dblcolIdade(5) = 56
dblcolIdade(6) = 73
dblcolIdade(7) = 68
dblcolIdade(8) = 56
dblcolIdade(9) = 73
dblcolIdade(10) = 65
dblcolIdade(11) = 45
dblcolIdade(12) = 58
dblcolIdade(13) = 45
dblcolIdade(14) = 53
dblcolIdade(15) = 49
dblcolIdade(16) = 78
dblcolIdade(17) = 73
dblcolIdade(18) = 68

For i = 1 To 18
   'armazena Xi - Xmédia
   dblcolResultadoMassaMenosMassaMedia(i) = dblcolMassa(i) - intMediaMassaMuscular
Next i

For i = 1 To 18
   'armazena Yi - Ymédia
   dblcolResultadoIdadeMenosIdadeMedia(i) = dblcolIdade(i) - intMediaIdade
Next i

For i = 1 To 18
   'armazena (Xi - Xmédia) X (Yi - Ymédia)
   dblcolIdadeXMassa(i) = dblcolResultadoIdadeMenosIdadeMedia(i) * dblcolResultadoMassaMenosMassaMedia(i)
Next i

For i = 1 To 18
   'armazena (Xi - Xmédia)^2
   dblcolIdadeQuadrado(i) = dblcolResultadoIdadeMenosIdadeMedia(i) ^ 2
Next i

For i = 1 To 18
   'armazena (Yi - Ymédia)^2
   dblcolMassaQuadrado(i) = dblcolResultadoMassaMenosMassaMedia(i) ^ 2
Next i

For i = 1 To 18
   Var = Var + dblcolIdadeXMassa(i)
   dblcolSomatorioIdadeXMassa(i) = Var
Next i

For i = 1 To 18
   var2 = var2 + dblcolMassaQuadrado(i)
   dblcolSomatorioTotalMassaAoQuadrado(i) = var2
Next i

For i = 1 To 18
   var3 = var3 + dblcolIdadeQuadrado(i)
   dblcolSomatorioTotalIdadeAoQuadrado(i) = var3
Next i

b1 = dblcolSomatorioIdadeXMassa(18) / dblcolSomatorioTotalIdadeAoQuadrado(18)
b0 = intMediaMassaMuscular - b1 * intMediaIdade
resultado = b0 + b1 * 50

Text3.Text = CStr(resultado)
End Sub

