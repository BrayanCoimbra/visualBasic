VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pDblValorResultante As Double
Private pDblValorResultanteBask As Variant
Private strErro As String

Public Property Get getErro() As String
    getErro = strErro
End Property

Public Property Get ValorResultante() As Double
    ValorResultante = pDblValorResultante
End Property

Public Property Get ValorResultanteBask() As Variant
    ValorResultanteBask = pDblValorResultanteBask
End Property
  
Public Function Somar(dblValor1 As Double, dblValor2 As Double) As Boolean
      
   On Error GoTo Somar_E
   
   Somar = False
      
   pDblValorResultante = dblValor1 + dblValor2
       
   Somar = True
   
   GoTo DestruirObjetos
       
Somar_E:
    strErro = Err.Description
    
DestruirObjetos:
End Function

Public Function Subtracao(dblValor1 As Double, dblValor2 As Double) As Boolean
   
   On Error GoTo Subtracao_E

   Subtracao = False
                
   pDblValorResultante = dblValor1 - dblValor2
    
   Subtracao = True
    
   GoTo DestruirObjetos

Subtracao_E:
    strErro = Err.Description
    
DestruirObjetos:
End Function

Public Function Multiplicacao(dblValor1 As Double, dblValor2 As Double) As Boolean
   
   On Error GoTo Multiplicacao_E

   Multiplicacao = False
                
   pDblValorResultante = dblValor1 * dblValor2
    
   Multiplicacao = True
    
   GoTo DestruirObjetos

Multiplicacao_E:
    strErro = Err.Description
    
DestruirObjetos:
End Function

Public Function Divisao(dblValor1 As Double, dblValor2 As Double) As Boolean
   
   On Error GoTo Divisao_E

   Divisao = False
                
   pDblValorResultante = dblValor1 / dblValor2
    
   Divisao = True
    
   GoTo DestruirObjetos

Divisao_E:
    strErro = Err.Description
    
DestruirObjetos:
End Function

Public Function DivisaoInteira(dblValor1 As Double, dblValor2 As Double) As Boolean
   
   On Error GoTo DivisaoInteira_E

   DivisaoInteira = False
                
   pDblValorResultante = dblValor1 \ dblValor2
    
   DivisaoInteira = True
    
   GoTo DestruirObjetos

DivisaoInteira_E:
    strErro = Err.Description
    
DestruirObjetos:
End Function

Public Function DivisaoMod(dblValor1 As Double, dblValor2 As Double) As Boolean
   
   On Error GoTo DivisaoMod_E

   DivisaoMod = False
                
   pDblValorResultante = dblValor1 Mod dblValor2
    
   DivisaoMod = True
    
   GoTo DestruirObjetos

DivisaoMod_E:
    strErro = Err.Description
    
DestruirObjetos:
End Function

Public Function Exponenciacao(dblValor1 As Double, dblValor2 As Double) As Boolean
   
   On Error GoTo Exponenciacao_E

   Exponenciacao = False
                
   pDblValorResultante = dblValor1 ^ dblValor2
    
   Exponenciacao = True
    
   GoTo DestruirObjetos

Exponenciacao_E:
    strErro = Err.Description
    
DestruirObjetos:
End Function

Public Function CalcularAreaCirculo(dblVlrRaio As Double) As Boolean
   
   On Error GoTo CalcularAreaCirculo_E

   CalcularAreaCirculo = False
                
   pDblValorResultante = ValorPI * (dblVlrRaio ^ 2)
    
   CalcularAreaCirculo = True
    
   GoTo DestruirObjetos

CalcularAreaCirculo_E:
    strErro = Err.Description
    
DestruirObjetos:
End Function

Public Function CalcularBhaskara(dblValorA As Double, dblValorB As Double, dblValorC As Double) As Boolean
   
    On Error GoTo CalcularBhaskara_E

    CalcularBhaskara = False
   
    Dim delta As Double
    Dim raiz1 As Double
    Dim raiz2 As Double
    
    If dblValorA = 0 Then
        blnSucesso = False
        CalcularBhaskaraCorrigidoFuncao = "Erro: O valor de 'a' não pode ser zero."
        Exit Function
    End If
   
    delta = (dblValorB ^ 2) - (4 * dblValorA * dblValorC)
    
    If delta >= 0 Then
        raiz1 = (-dblValorB + Sqr(delta)) / (2 * dblValorA)
        raiz2 = (-dblValorB - Sqr(delta)) / (2 * dblValorA)
        pDblValorResultanteBask = "Raiz 1: " & CStr(raiz1) & "  Raiz 2: " & CStr(raiz2)
    Else
        pDblValorResultanteBask = "O delta (" & CStr(delta) & ") é negativo, logo, não há raízes reais."
    End If
       
    CalcularBhaskara = True
       
    GoTo DestruirObjetos
   
CalcularBhaskara_E:
    strErro = Err.Description
    
DestruirObjetos:
End Function

