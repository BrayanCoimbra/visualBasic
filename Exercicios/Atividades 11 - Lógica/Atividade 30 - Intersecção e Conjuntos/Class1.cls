VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OperacoesConjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private arrAux()
Private Const Uniao = 2
Private Const Inserccao = 1
Private Retorno As String
Private NmrCompare As Integer
Private colConjuntoUm As Collection
Private colConjuntoDois As Collection
Private colConjuntoResultante As Collection

Public Property Get gRetorno() As String
   gRetorno = Retorno
End Property

Public Function ProcessarConjuntos(operacao As Integer, ConjuntoA As String, ConjuntoB As String) As Boolean
   
   On Error GoTo ProcessarConjuntos_E
   
   ProcessarConjuntos = False

   Set colConjuntoUm = PopularCollection(ConjuntoA)
   Set colConjuntoDois = PopularCollection(ConjuntoB)
   Set colConjuntoResultante = New Collection

   Select Case operacao
   Case 1
      Retorno = "Intersecção: "
      IntersectarConjuntos
      
   Case 2
      Retorno = "União: "
      UnirConjuntos colConjuntoUm, colConjuntoDois
   End Select
         
   If colConjuntoResultante.Count > 0 Then
      Retorno = Retorno & "{ " & ConjuntoResultante & " }"
   Else
      Retorno = "Não foi possível realizar a operação. Cheque seus conjuntos."
   End If
   
   ProcessarConjuntos = True
   
   GoTo DestuirObjetos
   
ProcessarConjuntos_E:
   ProcessarConjuntos = False
   MsgBox Err.Description
   
DestuirObjetos:
   Set colConjuntoUm = Nothing
   Set colConjuntoDois = Nothing
   Set colConjuntoResultante = Nothing
End Function

Private Sub UnirConjuntos(ParamArray pCollections())
'Para cada conjunto, verifica se há repetições no conjunto resultante
'Se sim, não adiciona o item, pois é uma repetição
'Se não, adiciona, realizando a união

   Dim Col As Variant
   
   For Each Col In pCollections
      For i = 1 To Col.Count
         If Not Repetiu(Col.Item(i), colConjuntoResultante) Then
            colConjuntoResultante.Add Col.Item(i)
         End If
      Next i
   Next
End Sub

Private Sub IntersectarConjuntos()
'Percorre cada item dos conjuntos e compara se são iguais, se sim, verifica se há repetições, se não houver repetição, adiciona, se houver, não adiciona
   For i = 1 To colConjuntoUm.Count
      For j = 1 To colConjuntoDois.Count
         If colConjuntoUm.Item(i) = colConjuntoDois.Item(j) And (Not Repetiu(colConjuntoUm.Item(i), colConjuntoResultante)) Then
            colConjuntoResultante.Add colConjuntoUm.Item(i)
         End If
      Next j
   Next i
End Sub

Private Function Repetiu(NmrCompare As Integer, pCol As Collection) As Boolean
'Percorre a pCol comparando cada item com o NmrCompare a fim de definir se o valor já existe no conjunto resultante
   For i = 1 To pCol.Count
      If NmrCompare = pCol.Item(i) Then
         Repetiu = True
         Exit For
      Else
         Repetiu = False
      End If
   Next i
End Function

Private Function ConjuntoResultante() As String
'Monta o resultado e devolve numa string
   SortColResultante
   
   For i = 1 To UBound(arrAux)
      ConjuntoResultante = ConjuntoResultante & arrAux(i) & " "
   Next i
End Function

Private Function PopularCollection(Conjunto As String) As Collection
   Set PopularCollection = New Collection
   
   For i = 1 To Len(Conjunto) + 1
      strCaractere = Mid(Conjunto, i, 1)
      If strCaractere = " " Or strCaractere = "" Then
         PopularCollection.Add strTermo
         strTermo = ""
      Else
         strTermo = strTermo & strCaractere
      End If
   Next i
End Function

Private Sub SortColResultante()
'Popula o arrAux com os valores da colConjuntoResultante para organizar em ordem crescente

   ReDim arrAux(1 To colConjuntoResultante.Count)
   Dim temp As String

   For i = 1 To colConjuntoResultante.Count
      arrAux(i) = colConjuntoResultante.Item(i)
   Next i

   For i = 1 To UBound(arrAux)
      For j = i + 1 To UBound(arrAux)
         If CInt(arrAux(i)) > CInt(arrAux(j)) Then
            temp = arrAux(i)
            arrAux(i) = arrAux(j)
            arrAux(j) = temp
         End If
      Next j
   Next i
End Sub
