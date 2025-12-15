VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KNII"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Interceptar Mensagem"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'03- Problema da transmissão de rádio
' Deseja-se criar um programa capaz de identificar uma mensagem inimiga que está sendo transmitida
' em ondas de rádio acima de 100Mhz. O programa de computador espião Kni já captou a transmissão
' e é necessário que seja construído outro software capaz de interpretar e extrair a mensagem.
' O Kni dá como saída uma cadeia como a seguinte:

' 90c87esd67uj,./';*&^120lin87uj101gu87km102a77jh150gem..&

' Onde, da esquerda para direita:
' 90 é a frequência em Mhz.
' c é o código lido na frequência de 90Mhz.

' 87 é a frequência do próximo código.
' esd é o código lido na frequência de 87Mhz

' 67 é a frequência do próximo código.
' uj é o código lido na frequência de 67Mhz

' ,./';&^ foi uma interferência que ocorreu quando lia-se o código da frequência de 67Mhz.
' ...
'Assim, no fragmento acima, a mensagem transmitida acima de 100Mhz foi: linguagem.
'Pois, lin foi transmitido em 120Mhz, gu em 101Mhz, a em 102Mhz, gem em 150Mhz.
'Construa um programa capaz de recebendo uma cadeia de no máximo 250 caracteres retornar
' a mensagem transmitida acima de 100Mhz.

'Considere que:
' a freqüência estará sempre entre 1 e 200Mhz.
' toda a interferência deverá ser ignorada. Deve-se considerar interferência todo
' caractere diferente de uma letra ou um número.
' não existirá espaços na cadeia de entrada (produzida pelo Kni).
' o tamanho máximo da mensagem será de 100 caracteres.

'Entrada:
' 90c87esd67uj,./';*&^120lin87uj101gu87km102a77jh150gem..&
' Saída:
' linguagem

'Entrada:
' *(12*23qualquer130i120n87j102t87ejh104er*&^_)(105n7k122e33kw140t**
' Saída:
' internet

Private blnCriandoMsg As Boolean
Private colMensagemMontada As Collection
Private colStrSaida As Collection
Private colFrequencias As Collection
Private strEntrada, strSaida, strCaractere, strFrequencia, strMensagem, strMensagemMontada As String
Private i As Integer

Private Sub Command1_Click()
   Set colMensagemMontada = New Collection
   Set colFrequencias = New Collection
   Set colStrSaida = New Collection
   strSaida = ""
   strCaractere = ""
   strFrequencia = ""
   strMensagem = ""
   strMensagemMontada = ""
   
'   strEntrada = "90c87esd67uj,./';*&^120lin87uj101gu87km102a77jh150gem..&"
   strEntrada = "*(12*23qualquer130i120n87j102t87ejh104er*&^_)(105n7k122e33kw140t**"
   
   i = 1
   
   IdentificarFrequencia
   
   For j = 1 To colMensagemMontada.Count + 1
      If Not j = colMensagemMontada.Count + 1 Then
         colStrSaida.Add (VerificaCaracteresEspeciais(colMensagemMontada.Item(j)))
         strSaida = strSaida & colStrSaida.Item(j)
      End If
   Next j
   
   MsgBox strSaida, vbInformation, "Mensagem Inimiga Interceptada"
   
   Set colMensagemMontada = Nothing
   Set colFrequencias = Nothing
   Set colStrSaida = Nothing
End Sub

Private Sub IdentificarFrequencia()
   For i = i To Len(strEntrada)
      strCaractere = Mid(strEntrada, i, 1)
      
      If IsNumeric(strCaractere) Then
         strFrequencia = strFrequencia & strCaractere
      
      ElseIf Not IsNumeric(strCaractere) Then
         If strFrequencia <> "" Then
            colFrequencias.Add strFrequencia
            CapturarMensagem
            strFrequencia = ""
            i = i - 1
         End If
      End If
   Next i
End Sub

Private Function VerificaCaracteresEspeciais(strTermoVerificado As String) As String
   Dim strCaractere As String
   Dim strCaracteresNormais As String
   Dim strCaracteresEspeciais As String
   
   For i = 1 To Len(strTermoVerificado)
      strCaractere = Mid(strTermoVerificado, i, 1)
      
      If InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", strCaractere) > 0 Then
         strCaracteresNormais = strCaracteresNormais & strCaractere
         VerificaCaracteresEspeciais = strCaracteresNormais
      End If
   Next i
   
   VerificaCaracteresEspeciais = VerificaCaracteresEspeciais ' Retorna a lista de caracteres especiais encontrados
End Function


Private Sub CapturarMensagem()
   endpos = Len(strEntrada) + 1
   Dim strCaractereLocal As String
   
   For i = i To endpos
      strCaractere = Mid(strEntrada, i, 1)
      
      If Not IsNumeric(strCaractere) And Not strCaractere = "" Then
         strMensagem = strMensagem & strCaractere

      ElseIf IsNumeric(strCaractere) Or i = endpos Then
         If CInt(strFrequencia) >= 100 And CInt(strFrquencia) <= 200 Then
            colMensagemMontada.Add strMensagem
         End If

         strMensagem = ""
         Exit For
      End If
   Next

End Sub
