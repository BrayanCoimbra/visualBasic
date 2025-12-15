VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCraps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Soma As Integer
Private Ponto As Integer
Private Dado1 As Integer
Private Dado2 As Integer
Private Mensagem As String
Private CaminhoDado1 As String
Private CaminhoDado2 As String
Private NumeroDeExecucoes As Integer

Public Property Get gNumeroDeExecucoes() As Integer
   gNumeroDeExecucoes = NumeroDeExecucoes
End Property

Public Property Let lNumeroDeExecucoes(Valor As Integer)
   NumeroDeExecucoes = Valor
End Property

Public Property Get gSoma() As Integer
   gSoma = Soma
End Property

Public Property Let lSoma(Valor As Integer)
   Soma = Valor
End Property

Public Property Get gPonto() As Integer
   gPonto = Ponto
End Property

Public Property Let lPonto(Valor As Integer)
   Ponto = Valor
End Property

Public Property Get gDado1() As Integer
   gDado1 = Dado1
End Property

Public Property Let lDado1(Valor As Integer)
   Dado1 = Valor
End Property

Public Property Get gDado2() As Integer
   gDado2 = Dado2
End Property

Public Property Let lDado2(Valor As Integer)
   Dado2 = Valor
End Property

Public Property Get gCaminhoDado1() As String
   gCaminhoDado1 = CaminhoDado1
End Property

Public Property Let lCaminhoDado1(Valor As String)
   CaminhoDado1 = Valor
End Property

Public Property Get gCaminhoDado2() As String
   gCaminhoDado2 = CaminhoDado2
End Property

Public Property Let lCaminhoDado2(Valor As String)
   CaminhoDado2 = Valor
End Property

Public Property Get gMensagem() As String
   gMensagem = Mensagem
End Property

Public Property Let lMensagem(Valor As String)
   Mensagem = Valor
End Property

Public Function LancarDados() As Boolean
    
   On Error GoTo LancarDados_E
    
   LancarDados = False
    
   Randomize
         
   Dim caminhoImagens As String
   caminhoImagens = "C:\Imagens\"

   ' Declaração e inicialização do array de caminhos das imagens
   Dim colCaminho(1 To 6) As String
   For i = 1 To 6
       colCaminho(i) = caminhoImagens & "Face_" & i & ".jpg"
   Next i
    
   ' Lançar os dados
   Dado1 = (Int((6 * Rnd) + 1))
   Dado2 = (Int((6 * Rnd) + 1))
   CaminhoDado1 = colCaminho(Dado1)
   CaminhoDado2 = colCaminho(Dado2)
       
   Soma = Dado1 + Dado2
   
   ' Checa se há possibilidade de vitória na primeira jogada
   If NumeroDeExecucoes = 1 Then
      If Soma = 11 Or Soma = 7 Then
           Mensagem = CStr("Você venceu, a soma dos dados deu " & Soma & " de primeira! " & "Parabéns!")
      
      'Se não, checa possibilidade de derrota
      ElseIf Soma = 2 Or Soma = 3 Or Soma = 12 Then
           Mensagem = CStr("Craps! Você perdeu. A soma dos dados deu " & Soma & "! " & "Ops!")
           
      'Se não não der vitória na primeira jogada e não der derrota, guarda o número como o Ponto
      Else
           GuardaPonto Dado1, Dado2
      End If
   End If

   'Cecha possibilidade de derrota se a soma for 7 após a segunda jogada
   If NumeroDeExecucoes > 1 Then
      If Soma = 7 Then
         Mensagem = CStr("Craps! Você perdeu. A soma dos dados deu " & Soma & "!")
      End If
   End If

    'Checa possibilidade de vitória se a soma der o ponto, após a primeira jogada, que define o ponto
   If NumeroDeExecucoes > 1 Then
     If Soma = Ponto Then
       Mensagem = CStr("No ponto! Você venceu. A soma dos dados deu " & Ponto & " !")
     End If
   End If

   NumeroDeExecucoes = NumeroDeExecucoes + 1
    
   LancarDados = True
    
   GoTo DestruirObjetos
    
LancarDados_E:
DestruirObjetos:
End Function

Public Function GuardaPonto(Dado1 As Integer, Dado2) As Boolean
    
    On Error GoTo GuardaPonto_E
    
    GuardaPonto = False
   
   ' Calcular a soma dos valores dos dados
    Soma = Dado1 + Dado2
    
    ' Se a soma for diferente de 7, 11, 2, 3 ou 12, é considerado o "ponto" do jogador
    If Soma = 4 Or Soma = 5 Or Soma = 6 Or Soma = 8 Or Soma = 9 Or Soma = 10 Then
        Ponto = Soma
    End If
    
    GuardaPonto = True
    
    GoTo DestruirObjetos
    
GuardaPonto_E:
DestruirObjetos:
End Function

