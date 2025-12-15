VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private colBuffer() As Variant
Private ResultadoAddValorNoBuffer As Variant
Private ResultadoRmvValorNoBuffer As Variant
Private ResultadoSubstituiValorNoBuffer As Variant
Private ResultadoAtualizarExibicaoBuffer As Variant
Dim i, j, k As Integer

Public Property Get gResultadoAddValorNoBuffer()
   gResultadoAddValorNoBuffer = ResultadoAddValorNoBuffer
End Property

Public Property Let lResultadoAddValorNoBuffer(Valor As Variant)
   gResultadoAddValorNoBuffer = Valor
End Property

Public Property Get gResultadoRmvValorNoBuffer()
   gResultadoRmvValorNoBuffer = ResultadoRmvValorNoBuffer
End Property

Public Property Let lResultadoRmvValorNoBuffer(Valor As Variant)
   gResultadoRmvValorNoBuffer = Valor
End Property

Public Property Get gResultadoSubstituiValorNoBuffer()
   gResultadoSubstituiValorNoBuffer = ResultadoSubstituiValorNoBuffer
End Property

Public Property Let lResultadoSubstituiValorNoBuffer(Valor As Variant)
   gResultadoSubstituiValorNoBuffer = Valor
End Property

Public Property Get gResultadoAtualizarExibicaoBuffer()
   gResultadoAtualizarExibicaoBuffer = ResultadoAtualizarExibicaoBuffer
End Property

Public Property Let lResultadoAtualizarExibicaoBuffer(Valor As Variant)
   gResultadoAtualizarExibicaoBuffer = Valor
End Property

Public Function DefinirTamanhoBuffer(intNmr As Integer) As Boolean
   
   On Error GoTo DefinirTamanhoBuffer_E
   
   DefinirTamanhoBuffer = False
   
   i = 0
   j = 0
   k = 0
   
   ReDim colBuffer(0 To 0)
   
   ReDim colBuffer(1 To intNmr)
   
   DefinirTamanhoBuffer = True
   
   GoTo DestruirObjetos
   
DefinirTamanhoBuffer_E:
DestruirObjetos:
End Function

Public Function AddValorNoBuffer(intValorAdicionado As Integer) As Boolean
   
   On Error GoTo AddValorNoBuffer_E
   
   AddValorNoBuffer = False
   
   k = 0
   j = j + 1
      
   If j <= UBound(colBuffer) Then
      colBuffer(j) = (intValorAdicionado)
   Else
      j = 1
      colBuffer(j) = (intValorAdicionado)
   End If
   
   AddValorNoBuffer = True
   
   GoTo DestruirObjetos

AddValorNoBuffer_E:
DestruirObjetos:
End Function

Public Function RmvValorNoBuffer() As Boolean
   
   On Error GoTo RmvValorNoBuffer_E
   
   RmvValorNoBuffer = False
   
   k = k + 1
   If k <= UBound(colBuffer) Then
      colBuffer(k) = ""
   End If
      
   RmvValorNoBuffer = True
   
   GoTo DestruirObjetos

RmvValorNoBuffer_E:
DestruirObjetos:
End Function

Public Function SubstituiValorNoBuffer(varNovoCaractere As Variant) As Boolean
   
   On Error GoTo SubstituiValor_E
   
   SubstituiValor = False
   
   k = 1
   If k <= UBound(colBuffer) Then
      For k = k To UBound(colBuffer)
         If colBuffer(k) <> "" Then
            colBuffer(k) = varNovoCaractere
            Exit For
         Else
         End If
      Next k
   End If
   
   k = k Mod (UBound(colBuffer))
 
   SubstituiValor = True
   
   GoTo DestruirObjetos
  
SubstituiValor_E:
DestruirObjetos:
End Function

Public Function AtualizarExibicaoBuffer(txtAtualizarBuffer As TextBox) As Boolean
 
   On Error GoTo AtualizarExibicaoBuffer_E
   
   AtualizarExibicaoBuffer = False
    
   ' Limpar o texto anterior antes de exibir o novo resultado
   txtAtualizarBuffer.Text = ""
    
   ' Adicionar os valores atualizados do buffer ao texto
   For i = 1 To UBound(colBuffer)
      txtAtualizarBuffer.Text = txtAtualizarBuffer.Text & " " & "[" & colBuffer(i) & "]"
   Next i
    
   AtualizarExibicaoBuffer = True
   
   GoTo DestruirObjetos
   
AtualizarExibicaoBuffer_E:
DestruirObjetos:
End Function

