VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Text            =   "0010"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Chama a função com o número de dígitos na placa (4)
    ContabilizarPlacas 4, UCase(Me.Text1.Text)
End Sub

Private Sub ContabilizarPlacas(ByVal intQtdDigitosNaPlaca As Integer, strPlaca As String)
   Dim i As Long
   Dim maxNum As Long
   Dim strPlacaAtual As String
   Dim strPlacaDesejada As String
   Dim PlacaEncontrada As Boolean
   
   strPlacaDesejada = InverterString(strPlaca)
   
   ' Calcula o maior número para o número de bits especificado
   maxNum = 2 ^ intQtdDigitosNaPlaca - 1
   
   ' Loop para gerar todos os números binários de 0 até o maior número
   For i = 1 To maxNum
      strPlacaAtual = DecToBin(i, intQtdDigitosNaPlaca)
      
      If strPlacaDesejada = strPlacaAtual Then
         MsgBox i & " dias"
         PlacaEncontrada = True
         Exit Sub
      End If
   Next i
   
   If Not PlacaEncontrada Then
      MsgBox "Não é possível"
   End If
   
End Sub

Private Function DecToBin(ByVal num As Long, ByVal numBits As Integer) As String
   ' Converte o parâmetro num em binário usando divisão e módulo
   Dim binaryStr As String
   Dim remainder As Integer
   Dim i As Integer

   binaryStr = ""
   
   ' Loop para obter cada bit, começando do menos significativo para o mais significativo
   For i = 0 To numBits - 1
       remainder = num Mod 2
       num = num \ 2
       binaryStr = CStr(remainder) & binaryStr
   Next i

   ' Adiciona zeros à esquerda se a representação binária for menor que numBits
   If Len(binaryStr) < numBits Then
       binaryStr = String(numBits - Len(binaryStr), "0") & binaryStr
   End If
   
   DecToBin = InverterString(binaryStr)
End Function

Private Function InverterString(strTermo As String) As String
'Inverte a string passada como parametro
   For i = Len(strTermo) To 1 Step -1
      InverterString = InverterString & Mid(strTermo, i, 1)
   Next i
End Function
