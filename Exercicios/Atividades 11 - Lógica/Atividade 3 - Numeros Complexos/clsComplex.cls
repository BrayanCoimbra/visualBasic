VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComplex"
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

Public dblParteReal As Double
Public dblParteImaginaria As Double

Public Property Get gParteReal() As Double
   gParteReal = dblParteReal
End Property

Public Property Get gParteImaginaria() As Double
   gParteImaginaria = dblParteImaginaria
End Property

Public Property Let lParteReal(Valor As Double)
   dblParteReal = Valor
End Property

Public Property Get lParteImaginaria(Valor As Double)
   dblParteImaginaria = Valor
End Property

