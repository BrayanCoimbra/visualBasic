VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "10 11"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "1 = Ímpar"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "0 = Par"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'04- Interferências, ruídos e outros fenômenos que prejudicam a integridade dos dados são problemas fundamentais quando computadores se comunicam em rede.
'Para detectar alterações em bits, os dados são sempre enviados com redundâncias computadas a partir dos bits originais.'Este tipo de técnica de detecção de erros costuma receber o nome de checksum, e segue o mesmo princípio dos dígitos verificadores presentes em diversos
'documentos e identificadores numéricos (por exemplo, números de contas e agências bancárias).

'Uma das técnicas de detecção de erros mais simples e mais usadas é o teste de paridade.

'Cada byte (1 byte é igual a 8 bits) é enviado junto com um bit adicional, que indica se o número de bits com valor 1 no byte é par (bit redundante = 0)
'ou ímpar (bit redundante = 1).

'Por exemplo um byte com o valor 8 tem os bits 00001000, ou seja, apenas 1 bit "setado", portanto a sua paridade é 1.

'Já um byte com o valor 0x55 é representado pelos bits 01010101 - 4 bits "setados", portanto a sua paridade é 0.

'Portanto, escreva um programa que possui a função CalculaParidade(decimal as Integer)
'que dado uma sequência inicial de números decimais calcule e informe a paridade de elemento da entrada.

Private Sub Command1_Click()
   Dim numeroDecimal As Long
   Dim numeroBinario As String
   Dim Resposta As String
   Dim colValores As Collection
   
   Set colValores = Split(Me.Text1.Text, " ")
   
   For i = 1 To colValores.Count
      Me.Text2.Text = Me.Text2.Text & CStr(colValores.Item(i)) & " = " & CStr(CalculaParidade(colValores.Item(i))) & vbNewLine
   Next i
      
   Set colValores = Nothing
End Sub

Public Function CalculaParidade(pDecimal As Integer) As Integer
   Dim binario As String
   Dim i As Integer
   Dim resto As Integer
   
   If pDecimal = 0 Then
       DecimalParaBinario = "0"
       Exit Function
   End If
   
   binario = ""
   
   Do While pDecimal > 0
      resto = pDecimal Mod 2
      binario = CStr(resto) & binario
      pDecimal = pDecimal \ 2
   Loop
    
   For i = 1 To Len(binario)
      If Mid(binario, i, 1) = 1 Then
         contUm = contUm + 1
      End If
   Next i
   
   If contUm Mod 2 = 0 Then
      CalculaParidade = 0
   Else
      CalculaParidade = 1
   End If
End Function

Private Function Split(stringParaSplit As String, caractereSeparador As String) As Collection
   Dim strCaractere As String
   Dim strTermo As String
   
   Set Split = New Collection
   
   For i = 1 To Len(stringParaSplit) + 1
      strCaractere = Mid(stringParaSplit, i, 1)
      
      If strCaractere = caractereSeparador Or strCaractere = "" And strTermo <> "" Then
         Split.Add strTermo
         strTermo = ""
      Else
         strTermo = strTermo & strCaractere
      End If
   Next i
End Function

