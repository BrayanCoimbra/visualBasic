VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCombinacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Combinacao As Integer
Private Parenteses As String
Public str As String

Public Property Get gParenteses() As Integer
   gParenteses = Parenteses
End Property

Public Property Get gCombinacao() As Integer
   gCombinacao = Combinacao
End Property

Public Function Combinar(num As Integer) As Boolean

   On Error GoTo Combinar_E

   Combinar = False

   Combinacao = (Fatorial(2 * num)) / ((Fatorial(num + 1)) * (Fatorial(num)))
      
   Combinar = True

   GoTo DestruirObjetos

Combinar_E:
   MsgBox "Erro na função clsCombinacoes.Combinar!", 16, "ERRO"
DestruirObjetos:
   'Limpar objetos
End Function

Private Function Fatorial(numero As Integer) As Double
   If numero = 0 Or numero = 1 Then
      Fatorial = numero
   Else
      Fatorial = numero * Fatorial(numero - 1)
   End If
End Function

Public Function MontarParenteses(intParesParenteses As Integer, intPA As Integer, intPF As Integer, strResultado As String, colParenteses As Collection) As Boolean

'PA = Parenteses abertos
'PF = Parenteses fechados
   
   On Error GoTo MontarParenteses_E
         
   MontarParenteses = False
         
   'Finaliza as chamadas recursivas
   If intPA = intParesParenteses And intPF = intParesParenteses Then
      colParenteses.Add (strResultado)
      GoTo DestruirObjetos
   End If

   If intPA < intParesParenteses Then
      MontarParenteses intParesParenteses, intPA + 1, intPF, strResultado & "(", colParenteses
   End If
      
   If intPF < intPA Then
      MontarParenteses intParesParenteses, intPA, intPF + 1, strResultado & ")", colParenteses
   End If
         
   MontarParenteses = True
      
   GoTo DestruirObjetos

MontarParenteses_E:
   MsgBox "Erro na função clsCombinacoes.MontarParenteses!", 16, "ERRO"
   
DestruirObjetos:

End Function

'Exemplo_1: Entrada igual a 3
'() () ()
'() (())
'(()) ()
'((()))
'(()())
