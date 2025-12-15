VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Descodificar 
      Caption         =   "Descodificar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton Codificar 
      Caption         =   "Codificar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Codificar_Click()
   Dim strResultado As String
   Dim i As Integer

   strResultado = ""

   For i = 1 To Len(Text1.Text)
      Dim strCaractere As String
      strCaractere = UCase(Mid(Text1.Text, i, 1))

      Select Case strCaractere
      Case "A"
      strResultado = strResultado & "  .-"
      Case "B"
      strResultado = strResultado & "  -..."
      Case "C"
      strResultado = strResultado & "  -.-."
      Case "D"
      strResultado = strResultado & "  -.."
      Case "E"
      strResultado = strResultado & "  ."
      Case "F"
      strResultado = strResultado & "  ..-."
      Case "G"
      strResultado = strResultado & "  --."
      Case "H"
      strResultado = strResultado & "  ...."
      Case "I"
      strResultado = strResultado & "  .."
      Case "J"
      strResultado = strResultado & "  .---"
      Case "K"
      strResultado = strResultado & "  -.-"
      Case "L"
      strResultado = strResultado & "  .-.."
      Case "M"
      strResultado = strResultado & "  --"
      Case "N"
      strResultado = strResultado & "  -."
      Case "O"
      strResultado = strResultado & "  ---"
      Case "P"
      strResultado = strResultado & "  .--."
      Case "Q"
      strResultado = strResultado & "  --.-"
      Case "R"
      strResultado = strResultado & "  .-."
      Case "S"
      strResultado = strResultado & "  ..."
      Case "T"
      strResultado = strResultado & "  -"
      Case "U"
      strResultado = strResultado & "  ..-"
      Case "V"
      strResultado = strResultado & "  ...-"
      Case "W"
      strResultado = strResultado & "  .--"
      Case "X"
      strResultado = strResultado & "  -..-"
      Case "Y"
      strResultado = strResultado & "  -.--"
      Case "Z"
      strResultado = strResultado & "  --.."
      Case "0"
      strResultado = strResultado & "  -----"
      Case "1"
      strResultado = strResultado & "  .----"
      Case "2"
      strResultado = strResultado & "  ..---"
      Case "3"
      strResultado = strResultado & "  ...--"
      Case "4"
      strResultado = strResultado & "  ....-"
      Case "5"
      strResultado = strResultado & "  ....."
      Case "6"
      strResultado = strResultado & "  -...."
      Case "7"
      strResultado = strResultado & "  --..."
      Case "8"
      strResultado = strResultado & "  ---.."
      Case "9"
      strResultado = strResultado & "  ----."
      Case Else
      ' Adicione manipulação para outros caracteres, se necessário
      strResultado = strResultado
      End Select
   Next i

   ' Exibir o resultado
   MsgBox "Código Morse: " & strResultado
End Sub

Private Sub Descodificar_Click()
    
   Dim strResultado As String
   Dim strPalavra As String
   Dim i As Integer
   
   strResultado = ""
   strPalavra = ""
   
   For i = 1 To Len(Text1.Text)
      Dim strCaractere As String
      strCaractere = UCase(Mid(Text1.Text, i, 1))
         
      If strCaractere <> " " Then
         strPalavra = strPalavra & strCaractere
      Else
         If strPalavra <> "" Then
            strResultado = strResultado & MorseToLetra(strPalavra)
            strPalavra = ""
         End If
      strResultado = strResultado & " " ' Adiciona um espaço em branco entre palavras
      End If
   Next i
      ' Adiciona a última palavra codificada, se houver
      If strPalavra <> "" Then
         strResultado = strResultado & MorseToLetra(strPalavra)
      End If
   ' Exiba o resultado decodificado
   MsgBox "Mensagem Decodificada: " & strResultado
End Sub

Function MorseToLetra(morse As String) As String
    Select Case morse
        Case ".-"
            MorseToLetra = "A"
        Case "-..."
            MorseToLetra = "B"
        Case "-.-."
            MorseToLetra = "C"
        Case "-.."
            MorseToLetra = "D"
        Case "."
            MorseToLetra = "E"
        Case "..-."
            MorseToLetra = "F"
        Case "--."
            MorseToLetra = "G"
        Case "...."
            MorseToLetra = "H"
        Case ".."
            MorseToLetra = "I"
        Case ".---"
            MorseToLetra = "J"
         Case "-.-"
            MorseToLetra = "K"
        Case ".-.."
            MorseToLetra = "L"
        Case "--"
            MorseToLetra = "M"
        Case "-."
            MorseToLetra = "N"
        Case "---"
            MorseToLetra = "O"
        Case ".--."
            MorseToLetra = "P"
        Case "--.-"
            MorseToLetra = "Q"
        Case ".-."
            MorseToLetra = "R"
        Case "..."
            MorseToLetra = "S"
        Case "-"
            MorseToLetra = "T"
        Case "..-"
            MorseToLetra = "U"
        Case "...-"
            MorseToLetra = "V"
        Case ".--"
            MorseToLetra = "W"
        Case "-..-"
            MorseToLetra = "X"
        Case "-.--"
            MorseToLetra = "Y"
        Case "--.."
            MorseToLetra = "Z"
        Case Else
            MorseToLetra = "?" ' Retorna ? para caracteres desconhecidos
    End Select
End Function
