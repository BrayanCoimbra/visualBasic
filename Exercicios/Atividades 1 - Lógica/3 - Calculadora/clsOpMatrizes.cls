VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpMatrizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private M3() As Double

Private Enum Operacoes
    enSoma = 1
    enSubtracao = 2
    enMultiplicacao = 3
    enMultiplicacaoPorQualquerNumero = 4
    enTransporMatriz = 5
End Enum

Public Property Get MatrizResultante(linha As Integer, coluna As Integer) As Double
    MatrizResultante = M3(linha, coluna)
End Property

Private Function ConstruirMatrizResultante(Operacao As Integer, linhas As Integer, Optional colunas As Integer = 0) As Boolean
    On Error GoTo ConstruirMatrizResultante_E
    
    ConstruirMatrizResultante = False
    
    Select Case Operacao
    
        Case enSoma, enSubtracao, enMultiplicacaoPorQualquerNumero
            ReDim M3(0 To linhas, 0 To colunas)
        
        Case enMultiplicacao
            ReDim M3(0 To linhas, 0 To colunas)
        
        Case enTransporMatriz
            ReDim M3(0 To colunas, 0 To linhas)
    End Select
    
    ConstruirMatrizResultante = True
    
    Exit Function
    
ConstruirMatrizResultante_E:
   MsgBox Err.Description

End Function

Public Function Somar(ParamArray Matrizes()) As Boolean
   On Error GoTo Somar_E
    
   Somar = False
   
   Dim linha As Integer
   Dim coluna As Integer
   Dim QtdLinhas As Integer
   Dim QtdColunas As Integer
    
   QtdLinhas = UBound(Matrizes(0), 1)

   QtdColunas = UBound(Matrizes(0), 2)
    
   If Not ConstruirMatrizResultante(enSoma, QtdLinhas, QtdColunas) Then
      GoTo Somar_E
   End If
   
   For linha = 0 To QtdLinhas
      For coluna = 0 To QtdColunas
         M3(linha, coluna) = Matrizes(0)(linha, coluna) + Matrizes(1)(linha, coluna)
      Next coluna
   Next linha
       
   Somar = True
   
   Exit Function
    
Somar_E:
    MsgBox "Erro ao somar matrizes."
End Function

Public Function Subtrair(ParamArray Matrizes()) As Boolean
   On Error GoTo Subtrair_E
   
   Subtrair = False

   Dim linha As Integer
   Dim coluna As Integer
   Dim QtdLinhas As Integer
   Dim QtdColunas As Integer
   
   QtdLinhas = UBound(Matrizes(0), 1)
   
   QtdColunas = UBound(Matrizes(0), 2)

   If Not ConstruirMatrizResultante(enSubtracao, QtdLinhas, QtdColunas) Then
      GoTo Subtrair_E
   End If

   For linha = 0 To QtdLinhas
      For coluna = 0 To QtdColunas
         M3(linha, coluna) = Matrizes(0)(linha, coluna) - Matrizes(1)(linha, coluna)
      Next coluna
   Next linha
   
   Subtrair = True
   
   Exit Function

Subtrair_E:
    MsgBox "Erro ao subtrair matrizes."
End Function

Public Function Multiplicacao(ParamArray Matrizes()) As Boolean
   On Error GoTo Multiplicacao_E
   
   Multiplicacao = False
   
   Dim M1 As Variant
   Dim M2 As Variant
   Dim linha As Integer
   Dim coluna As Integer
   Dim aux As Integer
   Dim QtdLinhasMatriz1 As Integer
   Dim QtdColunasMatriz1 As Integer
   Dim QtdLinhasMatriz2 As Integer
   Dim QtdColunasMatriz2 As Integer

   M1 = Matrizes(0)
   
   M2 = Matrizes(1)

   QtdLinhasMatriz1 = UBound(M1, 1)
   
   QtdColunasMatriz1 = UBound(M1, 2)

   QtdLinhasMatriz2 = UBound(M2, 1)

   QtdColunasMatriz2 = UBound(M2, 2)

   If Not ConstruirMatrizResultante(enMultiplicacao, QtdLinhasMatriz1, QtdColunasMatriz2) Then
      GoTo Multiplicacao_E
   End If

   For linha = 0 To QtdLinhasMatriz1
       For coluna = 0 To QtdColunasMatriz2
           M3(linha, coluna) = 0
           For aux = 0 To QtdColunasMatriz1
               M3(linha, coluna) = M3(linha, coluna) + (M1(linha, aux) * M2(aux, coluna))
           Next aux
       Next coluna
   Next linha
   
   Multiplicacao = True
    
   Exit Function

Multiplicacao_E:
    MsgBox "Erro na multiplicação: " & Err.Description
End Function

Public Function MultiplicacaoPorNumeroQualquer(pMatriz As Variant, n As Double) As Boolean
   On Error GoTo MultiplicacaoPorNumeroQualquer_E

   MultiplicacaoPorNumeroQualquer = False

   Dim linha As Integer
   Dim coluna As Integer
   Dim QtdLinhas As Integer
   Dim QtdColunas As Integer
   
   QtdLinhas = UBound(pMatriz, 1)

   QtdColunas = UBound(pMatriz, 2)

   If Not ConstruirMatrizResultante(enMultiplicacaoPorQualquerNumero, QtdLinhas, QtdColunas) Then
      GoTo MultiplicacaoPorNumeroQualquer_E
   End If

   For linha = 0 To QtdLinhas
      For coluna = 0 To QtdColunas
         M3(i, coluna) = pMatriz(i, coluna) * n
      Next coluna
   Next linha

   MultiplicacaoPorNumeroQualquer = True

   Exit Function

MultiplicacaoPorNumeroQualquer_E:
    MsgBox "Erro ao multiplicar matriz por número."
End Function

Public Function Transpor(pMatriz As Variant) As Boolean
   On Error GoTo Transpor_E
   
   Transpor = False

   Dim QtdLinhas As Integer
   QtdLinhas = UBound(pMatriz, 1)

   Dim QtdColunas As Integer
   QtdColunas = UBound(pMatriz, 2)

   If Not ConstruirMatrizResultante(enTransporMatriz, QtdLinhas, QtdColunas) Then
      GoTo Transpor_E
   End If

   Dim linhas As Integer
   Dim colunas As Integer

   For linhas = 0 To QtdLinhas
      For colunas = 0 To QtdColunas
         M3(colunas, linhas) = pMatriz(linhas, colunas)
      Next colunas
   Next linhas

   Transpor = True
   
   Exit Function

Transpor_E:
   MsgBox "Erro ao transpor matriz."
End Function
