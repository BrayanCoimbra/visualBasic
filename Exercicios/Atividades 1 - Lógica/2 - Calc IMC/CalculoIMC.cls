VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalculoIMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dblPeso As Double
Private dblAltura As Double
Private dblIMC As Double
Private strErro As String

Public Property Get Erro() As Double
    Erro = strErro
End Property
    
Public Property Get IMC() As Double
    IMC = dblIMC
End Property
    
Public Property Let IMC(ByVal Valor As Double)
    dblIMC = Valor
End Property
Public Property Get altura() As Double
    altura = dblAltura
End Property
    
Public Property Let altura(ByVal Valor As Double)
    dblAltura = Valor
End Property
    
Public Property Get peso() As Double
    peso = dblPeso
End Property
    
Public Property Let peso(ByVal Valor As Double)
    dblPeso = Valor
End Property
    
Private Sub Class_Initialize()
    dblPeso = 0
    dblAltura = 0
    dblIMC = 0
End Sub
    
Function CalcularIMC() As Boolean
    On Error GoTo TratarErro
    
    CalcularIMC = False
    
    dblIMC = (dblPeso / (dblAltura ^ 2))
    
    CalcularIMC = True
    
    GoTo LimparObjetos
    
TratarErro:
    strErro = Err.Description
    
LimparObjetos:

End Function

Private Sub Class_Terminate()
    'limpar objetos se necessario
End Sub
