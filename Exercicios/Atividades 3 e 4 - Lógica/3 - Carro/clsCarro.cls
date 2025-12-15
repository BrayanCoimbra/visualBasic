VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCarro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'03- Implemente um carro.
'O tanque de combustível do carro armazena no máximo 50 litros de gasolina.

'O carro consome 15 km/litro. Deve ser possível:
'a) Abastecer o carro com uma certa quantidade de gasolina;
'b) Mover o carro em uma determinada distância (medida em km);
'c) Retornar a quantidade de combustível e a distância total percorrida.
'd) Utilizar um controle para inserir o preço da gasolina em litros por quilômetro e calcular o custo do deslocamento.

Public dblTanque As Double
Public dblMovimento As Double
Public dblCusto As Double

Public Property Get ValorCusto() As Double
   ValorCusto = dblCusto
End Property
Public Property Let ValorCusto(ByVal Valor As Double)
   dblCusto = Valor
End Property

Public Property Get ValorTanque() As Double
   ValorTanque = dblTanque
End Property

Public Property Let ValorTanque(ByVal Valor As Double)
   dblTanque = Valor
End Property

Public Property Get ValorMovimento() As Double
   ValorMovimento = dblMovimento
End Property

Public Property Let ValorMovimento(ByVal Valor As Double)
   dblMovimento = Valor
End Property

Public Function MovimentarCarro(dblValorMovimentado As Double) As Boolean

   On Error GoTo MovimentarCarro_E

   MovimentarCarro = False
   
   If (dblValorMovimentado <= (dblTanque * 15)) Then
      dblTanque = (dblTanque - (dblValorMovimentado / 15))
      msg = MsgBox("Você andou " & dblValorMovimentado & "KM e ainda possui " & dblTanque & "L de combustível", 48, "CARRO MOVIMENTADO")
      dblTanque = dblTanque - (dblValorMovimentado / 15)
   Else
      msg = MsgBox("Você precisa de mais gasolina", 48, "FALTOU COMBUSTÍVEL")
   End If
   
   MovimentarCarro = True
   
   GoTo DestruirObjetos
   
MovimentarCarro_E:
DestruirObjetos:
End Function

Public Function Abastece(dblValorTanqueAbastece As Double) As Boolean

   On Error GoTo Abastece_E

   Abastece = False
   
   dblTanque = dblValorTanqueAbastece
   
   Abastece = True
   
   GoTo DestruirObjetos

Abastece_E:
DestruirObjetos:
End Function

Public Function CustoParaMovimentar(dblDistancia As Double) As Boolean

   On Error GoTo CustoParaMovimentar_E

   CustoParaMovimentar = False
   
   dblCusto = CDbl((dblDistancia) / (15))
   
   CustoParaMovimentar = True
   
   GoTo DestruirObjetos

CustoParaMovimentar_E:
DestruirObjetos:
End Function
