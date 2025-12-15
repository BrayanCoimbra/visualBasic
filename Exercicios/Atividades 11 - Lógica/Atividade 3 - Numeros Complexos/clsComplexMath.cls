VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComplexMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'03- 1- Escreva um programa (com formulário) que realize as seguintes operações de números complexos:
'   a) soma
'   b) subtração
'   c) multiplicação
'   d) módulo (valor absoluto)
'
'Obs.: Utilize uma classe chamada Complex como modelo de número complexo.
'      Utilize uma classe chamada ComplexMath que realiza os cálculos de matématica complexa.
'
' Na classe Complex:
' Forneça formas que permitam que os objetos dessa classe sejam inicializados (passados valores) das seguintes maneiras:
' I)  sem parâmetros que inicializa o objeto como zero
' II) um parâmetro representando a parte real; a parte imaginária será zero
' III)com dois parâmetros representando as partes real e imaginária

Private Z1 As clsComplex
Private Z2 As clsComplex
Public ZResult As clsComplex
Public Denominador As Double

Private ResultadoReal As Double
Private ResultadoImaginario As Double

Public Function Somar(dblParteRealN1 As Double, dblParteImagN1 As Double, dblParteRealN2 As Double, dblParteImagN2 As Double) As Boolean
   
   On Error GoTo Somar_E
   
   Set Z1 = New clsComplex
   Set Z2 = New clsComplex
   
   With Z1
      .dblParteReal = dblParteRealN1
      .dblParteImaginaria = dblParteImagN1
   End With
   
   With Z2
      .dblParteReal = dblParteRealN2
      .dblParteImaginaria = dblParteImagN2
   End With
   
   With ZResult
      .dblParteReal = (Z1.gParteReal + Z2.gParteReal)
      .dblParteImaginaria = (Z1.gParteImaginaria + Z2.gParteImaginaria)
   End With
   
   Somar = True
   
   GoTo DestruirObjetos
   
Somar_E:
   Somar = False

DestruirObjetos:
   Set Z1 = Nothing
   Set Z2 = Nothing
End Function

Public Function Subtrair(dblParteRealN1 As Double, dblParteImagN1 As Double, dblParteRealN2 As Double, dblParteImagN2 As Double) As Boolean
   
   On Error GoTo Subtrair_E
      
   Dim ResultadoReal As Double
   Dim ResultadoImaginario As Double
   Dim Resultado1 As Double
   Dim Resultado2 As Double
   Dim Z1 As clsComplex
   Dim Z2 As clsComplex
   
   Set Z1 = New clsComplex
   Set Z2 = New clsComplex
   
   With Z1
      .dblParteReal = dblParteRealN1
      .dblParteImaginaria = dblParteImagN1
   End With
   
   With Z2
      .dblParteReal = dblParteRealN2
      .dblParteImaginaria = dblParteImagN2
   End With
   
   'Real
   Resultado1 = (Z1.dblParteReal - Z2.dblParteReal)
   
   'Imaginario
   Resultado2 = (Z1.dblParteImaginaria - Z2.dblParteImaginaria)
   
   ResultadoReal = Resultado1 + Resultado4
   ResultadoImaginario = Resultado2 + Resultado3
   
   With ZResult
      .dblParteReal = ResultadoReal
      .dblParteImaginaria = ResultadoImaginario
   End With
   
   Subtrair = True
   
   GoTo DestruirObjetos
   
Subtrair_E:
   Subtrair = False

DestruirObjetos:
   Set Z1 = Nothing
   Set Z2 = Nothing
End Function

Public Function Multiplicar(dblParteRealN1 As Double, dblParteImagN1 As Double, dblParteRealN2 As Double, dblParteImagN2 As Double) As Boolean
'(2 + 5i) * (1 – 2i)
'2 – 4i + 5i – 10i² (lembrando que i² = – 1)
'2 – 4i + 5i – 10 * (–1)
'2 – 4i + 5i + 10
'12 + i
'real real = real, real imag = imag,
'imag real = imag, imag imag = imag

   On Error GoTo Multiplicar_E
   
   Dim ResultadoReal As Double
   Dim ResultadoImaginario As Double
   Dim Resultado1 As Double
   Dim Resultado2 As Double
   Dim Resultado3 As Double
   Dim Resultado4 As Double
   Dim Z1 As clsComplex
   Dim Z2 As clsComplex

   Set Z1 = New clsComplex
   Set Z2 = New clsComplex
   Set ZResult = New clsComplex
   
   With Z1
      .dblParteReal = dblParteRealN1
      .dblParteImaginaria = dblParteImagN1
   End With
   
   With Z2
      .dblParteReal = dblParteRealN2
      .dblParteImaginaria = dblParteImagN2
   End With
   
   'Real
   Resultado1 = (Z1.dblParteReal * Z2.dblParteReal)
   
   'Imaginario
   Resultado2 = (Z1.dblParteReal * Z2.dblParteImaginaria)
   
   'Imaginario
   Resultado3 = (Z1.dblParteImaginaria * Z2.dblParteReal)
   
   'Real
   Resultado4 = ((Z1.dblParteImaginaria * Z2.dblParteImaginaria) * (-1))
   
   ResultadoReal = Resultado1 + Resultado4
   ResultadoImaginario = Resultado2 + Resultado3
   
   With ZResult
      .dblParteReal = ResultadoReal
      .dblParteImaginaria = ResultadoImaginario
   End With
   
   Multiplicar = True
   
   GoTo DestruirObjetos
   
Multiplicar_E:
   Multiplicar = False

DestruirObjetos:
   Set Z1 = Nothing
   Set Z2 = Nothing
End Function

Public Function Modulo(dblParteRealN1 As Double, dblParteImagN1 As Double, dblParteRealN2 As Double, dblParteImagN2 As Double) As Boolean
   
   On Error GoTo Modulo_E
   
   Dim ResultadoReal As Double
   Dim ResultadoImaginario As Double
   Dim Resultado1 As Double
   Dim Resultado2 As Double
   Dim Resultado3 As Double
   Dim Resultado4 As Double
   Dim Resultado5 As Double
   Dim Resultado6 As Double
   Dim Resultado7 As Double
   Dim Resultado8 As Double

   
   Dim Z1 As clsComplex
   Dim Z2 As clsComplex
   
   Set Z1 = New clsComplex
   Set Z2 = New clsComplex
   Set ZResult = New clsComplex
   
   With Z1
      .dblParteReal = dblParteRealN1
      .dblParteImaginaria = dblParteImagN1
   End With
   
   With Z2
      .dblParteReal = dblParteRealN2
      .dblParteImaginaria = dblParteImagN2
   End With
   
   'Numerador
   Resultado1 = Z1.dblParteReal * Z2.dblParteReal
   Resultado2 = Z1.dblParteReal * (Z2.dblParteImaginaria * -1)
   Resultado3 = Z1.dblParteImaginaria * Z2.dblParteReal
   Resultado4 = Z1.dblParteImaginaria * ((Z2.dblParteImaginaria * -1) * -1)

   ResultadoRealNumerador = Resultado1 + Resultado4
   ResultadoImaginarioNumerador = Resultado2 + Resultado3
   
   'Denominador
   Resultado5 = Z2.dblParteReal * Z2.dblParteReal
   Resultado6 = Z2.dblParteReal * (Z2.dblParteImaginaria * -1)
   Resultado7 = Z2.dblParteImaginaria * Z2.dblParteReal
   Resultado8 = Z2.dblParteImaginaria * ((Z2.dblParteImaginaria * -1) * -1)

   Denominador = Resultado5 + Resultado6 + Resultado7 + Resultado8
   
   With ZResult
      .dblParteReal = ResultadoRealNumerador Mod Denominador
      .dblParteImaginaria = ResultadoImaginarioNumerador Mod Denominador
   End With
   
   Modulo = True
   
   GoTo DestruirObjetos
   
Modulo_E:
   Modulo = False

DestruirObjetos:
   Set Z1 = Nothing
   Set Z2 = Nothing
End Function

Private Sub Class_Initialize()
   Set ZResult = New clsComplex
End Sub

Private Sub Class_Terminate()
   Set ZResult = Nothing
End Sub
