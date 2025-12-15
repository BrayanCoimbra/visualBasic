VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   2385
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Verificar Palavra"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'01- Problema da palavra mágica
'Uma palavra P de tamanho K (2<=K<=100) é mágica se ela tem um número par de letras
'e se ao ordenarmos em ordem alfabética as K/2 primeiras letras obtém-se as K/2 letras
'finais.
'Exemplo de palavras mágicas:
'asas.
'gogo.
'gluglu.
'chocho.
'adfsadfs.
'aaaa.
'Exemplo de palavras que não são mágicas:
'xixi (ordenando as letras xi em ordem alfabética ficaria ix que não é igual a xi).
'xoxo (ordenando as letras xo em ordem alfabética ficaria ox que não é igual a xo).
'muximuxi (ordenando as letras muxi em ordem alfabética ficaria imux que não é
'igual a muxi).
'asdffdsa (ordenando as letras asdf em ordem alfabética ficaria adfs que não é igual a fdsa).
'aaaaa (não é mágica, pois tem um número ímpar de caracteres).
'Faça um programa que recebendo uma palavra seja capaz de identificar se ela é ou não uma palavra mágica
'e apresente conforme o caso:
'É mágica
'Não é mágica

Private arr() As String
Private i As Integer
Private primeiraMetade As String
Private segundaMetade As String
Private OrdenouEmOrdemAlfabetica As Boolean
Private palavra As String
Private metadeTamanho As Integer

Private Function OrdenarLetras(ByVal str As String) As String
   
   ' Converte a string em um array de caracteres.
   ReDim arr(1 To Len(str))
   For i = 1 To Len(str)
      arr(i) = Mid(str, i, 1)
   Next i
   
   ' Ordena o array em ordem alfabética.
   For i = 1 To UBound(arr)
      For j = i + 1 To UBound(arr)
         If arr(i) > arr(j) Then
            Dim temp As String
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            OrdenouEmOrdemAlfabetica = True
         End If
      Next j
   Next i
   
End Function

Private Sub Command1_Click()
   palavra = ""
   metadeTamanho = 0
   OrdenouEmOrdemAlfabetica = False
   primeiraMetade = ""
   segundaMetade = ""

   palavra = InputBox("Digite uma palavra:")
   
   ' Verifica se o tamanho da palavra é par.
   If Len(palavra) Mod 2 = 0 Then
   
      ' Divide a palavra ao meio.
      metadeTamanho = Len(palavra) / 2
      primeiraMetade = Left(palavra, metadeTamanho)
      segundaMetade = Right(palavra, metadeTamanho)
      
      ' Ordena as duas metades da palavra em ordem alfabética.
      primeiraMetade = OrdenarLetras(primeiraMetade)
      segundaMetade = OrdenarLetras(segundaMetade)
      
      ' Verifica se as duas metades são iguais.
      If primeiraMetade = segundaMetade And OrdenouEmOrdemAlfabetica = False Then
         MsgBox "A palavra é mágica!"
      Else
         MsgBox "A palavra não é mágica."
      End If
      
   Else
      MsgBox "A palavra não é mágica, pois possui um número ímpar de caracteres."
   End If
End Sub
