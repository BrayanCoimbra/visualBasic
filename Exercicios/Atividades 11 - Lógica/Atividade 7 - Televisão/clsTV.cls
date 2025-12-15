VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Volume As Integer
Private CanalAtual As Integer

Public Property Get gVolume() As Integer
   gVolume = Volume
End Property

Public Property Get gCanal() As Integer
   gCanal = Canal
End Property

Public Property Get gCanalAtual() As Integer
   gCanalAtual = CanalAtual
End Property

Public Function AumentarVolume() As Boolean

   On Error GoTo AumentarVolume_E
   
   Volume = Volume + 1
   
   AumentarVolume = True
   
   Exit Function
   
AumentarVolume_E:
   AumentarVolume = False
End Function

Public Function DiminuirVolume() As Boolean
   
   On Error GoTo DiminuirVolume_E

   Volume = Volume - 1
   
   DiminuirVolume = True
   
   Exit Function

DiminuirVolume_E:
   DiminuirVolume = False
End Function

Public Function AumentarCanal() As Boolean

   On Error GoTo AumentarCanal_E
   
   CanalAtual = CanalAtual + 1
   
   AumentarCanal = True
   
   Exit Function
   
AumentarCanal_E:
   AumentarCanal = False
End Function

Public Function DiminuirCanal() As Boolean
   
   On Error GoTo DiminuirCanal_E
     
   CanalAtual = CanalAtual - 1
   
   DiminuirCanal = True
   
   Exit Function

DiminuirCanal_E:
   DiminuirCanal = False
End Function

Public Function TrocarCanal(numero As Integer) As Boolean

   On Error GoTo TrocarCanal_E
   
   Select Case numero
       Case 1 To 10
           CanalAtual = numero
           TrocarCanal = True ' Retorna True para indicar que o canal foi trocado com sucesso
       Case Else
           ' Não faz nada, mantém o canal atual
           TrocarCanal = False ' Retorna False para indicar que o número do canal não é válido
   End Select
   
   TrocarCanal = True
   
   Exit Function
   
TrocarCanal_E:
   TrocarCanal = False
End Function
