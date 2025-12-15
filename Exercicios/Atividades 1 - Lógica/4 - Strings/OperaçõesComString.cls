VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OperaçõesComString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public strResultado As String

Public Function InverterString(strTermo As String) As Boolean

   On Error GoTo InverterString_E

   InverterString = False
   
   Dim strArmazenaPalavra As String
   Dim colPalavra As New Collection
   Set colPalavra = New Collection

   strResultado = ""
   
   For i = 1 To Len(strTermo)
      strLetra = Mid(strTermo, i, 1)

      If strLetra <> " " Then
         strArmazenaPalavra = strArmazenaPalavra + Mid(strTermo, i, 1)
      Else
         colPalavra.Add strArmazenaPalavra
         strArmazenaPalavra = ""
      End If
   Next i

   For i = colPalavra.Count To 1 Step -1
      strResultado = strResultado & colPalavra(i) & " "
   Next i
   
   Set colPalavra = Nothing

   InverterString = True
   
   Exit Function
   
InverterString_E:
   strResultado = Err.Description

End Function

Public Function RemoverLetraE(strTermo As String) As Boolean

   On Error GoTo RemoverLetraE_E
   
   RemoverLetraE = False
   
   Dim strLetra As String
   
   strResultado = ""
   
   For i = 1 To Len(strTermo)
       strLetra = Mid(strTermo, i, 1)
  
      If Not strLetra = "e" Or strLetra = "E" Then
         strResultado = strResultado & strLetra
      End If
   Next i
   
   RemoverLetraE = True
   
   Exit Function
   
RemoverLetraE_E:
   strResultado = Err.Description
End Function

Public Function ContabilizarVogais(strTermo As String) As Boolean
   
   On Error GoTo ContabilizarVogais_E
   
   ContabilizarVogais = False
   
   Dim contadorDeVogal(4) As Integer
   Dim strLetra As String
  
   strResultado = ""
      
   For i = 1 To Len(strTermo)
      strLetra = LCase(Mid(strTermo, i, 1))
        
      Select Case strLetra
         Case "a"
            contadorDeVogal(0) = contadorDeVogal(0) + 1
         Case "e"
            contadorDeVogal(1) = contadorDeVogal(1) + 1
         Case "i"
            contadorDeVogal(2) = contadorDeVogal(2) + 1
         Case "o"
            contadorDeVogal(3) = contadorDeVogal(3) + 1
         Case "u"
            contadorDeVogal(4) = contadorDeVogal(4) + 1
      End Select
    Next i
    
    strResultado = "Quantidade de vogais: " & vbNewLine & _
                "A: " & contadorDeVogal(0) & ", " & vbNewLine & _
                "E: " & contadorDeVogal(1) & ", " & vbNewLine & _
                "I: " & contadorDeVogal(2) & ", " & vbNewLine & _
                "O: " & contadorDeVogal(3) & ", " & vbNewLine & _
                "U: " & contadorDeVogal(4)
    
    ContabilizarVogais = True
    
    Exit Function
    
ContabilizarVogais_E:
   strResultado = Err.Description
End Function
