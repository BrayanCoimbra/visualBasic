VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'02- Problema da moda
'A moda é o conjunto formado pelos elementos com a maior frequência em uma amostra.
'Por exemplo:
'na amostra (1,2,3,3,3,4,4,4,4,5) a moda é {4},
'na amostra (1,1,1,2,2,2,3,4,5,5,6) a moda é {1,2}.
'Faça um programa capaz de calcular a moda de uma amostra de dados K.
'A entrada será um número indefinido de elementos de K um por linha.
'A saída deverá ser os elementos K pertencentes à moda em ordem crescente um em cada linha.
'K terá como elementos números inteiros positivos entre 0 e 256,
'sendo que o número zero identifica o fim da entrada de dados.

Private arrNmr() As Integer
Private colAux As Collection
Private colModa As Collection
Private ModaResposta As String
Private resp As String

Private Sub Command1_Click()
   Set colAux = New Collection
   Set colModa = New Collection
   
   If (Split(Text1.Text)) And colAux.Count <= 256 Then
      Sort
      ContabilizarModa
   Else
      GoTo DestruirObjetos
   End If
   
   Me.Text1.Text = ""
   
   GoTo DestruirObjetos
   
DestruirObjetos:
   Set colAux = Nothing
   Set colModa = Nothing
End Sub

Private Sub ContabilizarModa()
   Dim contador As Integer
   Dim moda As String
   Dim modaContador As Integer
   Dim modaEncontrada As Boolean
   ModaResposta = ""
   
   For i = 1 To colAux.Count
      contador = 0
         For j = 1 To colAux.Count
            If arrNmr(j) = arrNmr(i) Then
               contador = contador + 1
            End If
         Next j
      If contador > modaContador Then
         ModaResposta = CStr(arrNmr(i))
         modaContador = contador
         modaEncontrada = True
         
      ElseIf contador = modaContador Then
         ModaResposta = ModaResposta & ", " & CStr(arrNmr(i))
      End If
   Next i
   
   MontarResposta
   
   If modaEncontrada Then
      MsgBox "A moda é: {" & resp & "}."
   Else
      MsgBox "Não há moda."
   End If
   
End Sub

Private Sub MontarResposta()
   Dim strCaractereAtual As String
   Dim strProximoCaractere As String
   resp = ""
   ' Percorre a string ModaResposta para capturar apenas os caracteres únicos
   For i = 1 To Len(ModaResposta)
      strCaractereAtual = Mid(ModaResposta, i, 1)
      
      ' Verifica se o caractere atual não é uma vírgula ou espaço e se não foi capturado anteriormente
      If strCaractereAtual <> "," And strCaractereAtual <> " " And InStr(resp, strCaractereAtual) = 0 Then
         resp = resp & strCaractereAtual & ";"
      End If
   Next i
End Sub

Private Function Split(strString As String) As Boolean
   Split = False
   
   Dim strCaractere As String
   Dim strNmr As String
   Dim EndPos As Integer
   
   EndPos = Len(Text1.Text) + 1
   
   For i = 1 To Len(Text1.Text) + 1
      strCaractere = Mid(Text1.Text, i, 1)
      
      If strCaractere = " " Or i = EndPos Then
         If Not strNmr < 0 Then
            colAux.Add strNmr
            strNmr = ""
         Else
            Split = False
            MsgBox "Digite apenas números inteiros positivos.", vbExclamation
            Exit Function
         End If
      Else
         strNmr = strNmr & strCaractere
      End If
   Next i
   
   ReDim arrNmr(1 To colAux.Count)
    
   For i = 1 To colAux.Count
      arrNmr(i) = colAux(i)
   Next i
   
   Split = True
   
End Function

Private Sub Sort()
   Dim temp As Integer
   
   For i = 1 To colAux.Count
      For j = i + 1 To colAux.Count
         If arrNmr(i) > arrNmr(j) Then
            temp = arrNmr(i)
            arrNmr(i) = arrNmr(j)
            arrNmr(j) = temp
         End If
      Next j
   Next i
End Sub
