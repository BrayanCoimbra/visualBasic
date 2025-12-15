VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTuring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Estado As String
Private Resultado As String
Private colArmazenaCaractere(1 To 100) As String

Public Property Get gEstado() As String
   gEstado = Estado
End Property

Public Property Let lEstado(Valor As String)
   Estado = Valor
End Property

Public Property Get gResultado() As String
   gResultado = Resultado
End Property

Public Property Let lResultado(Valor As String)
   Resultado = Valor
End Property

Public Property Get gcolArmazenaCaractere() As String
   gcolArmazenaCaractere = colArmazenaCaractere
End Property

Public Property Let lcolArmazenaCaractere(Valor As Collection)
   colArmazenaCaractere = Valor
End Property

Public Function Iniciar(Texto As String) As Boolean
      
   On Error GoTo Iniciar_E
   
   Iniciar = False
      
   Resultado = ""
     
   For k = 1 To Len(Texto)
      colArmazenaCaractere(k) = ""
   Next k
     
   For j = 1 To Len(Texto)
      colArmazenaCaractere(j) = Mid(Texto, j, 1)
   Next j
       
   For i = 1 To Len(Texto)
      If Estado = "O" And colArmazenaCaractere(i) = "@" Then
         colArmazenaCaractere(i) = "."
         Estado = "halt"
         ElseIf Estado = "O" And colArmazenaCaractere(i) = "*" Then
           Estado = "O"
         ElseIf Estado = "O" And colArmazenaCaractere(i) = "/" Then
            colArmazenaCaractere(i) = "x"
            Estado = "y"
         
         ElseIf Estado = "y" And colArmazenaCaractere(i) = "@" Then
            colArmazenaCaractere(i) = "."
            Estado = "halt"
         ElseIf Estado = "y" And colArmazenaCaractere(i) = "*" Then
            Estado = "y"
         ElseIf Estado = "y" And colArmazenaCaractere(i) = "/" Then
            colArmazenaCaractere(i) = "y"
            Estado = "O"
      End If
   Resultado = Resultado + colArmazenaCaractere(i)
   Next i
   
   Iniciar = True
   
   GoTo DestruirObjetos
   
Iniciar_E:
DestruirObjetos:
End Function
