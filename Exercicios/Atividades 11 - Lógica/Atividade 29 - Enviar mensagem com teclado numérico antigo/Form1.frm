VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   585
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "ABC"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strPalavra As String
Private strRetorno As String

Private Sub Command1_Click()
   strRetorno = ""
   ContabilizarPalavra Me.Text1.Text
   MsgBox strRetorno
End Sub

Private Sub ContabilizarPalavra(pStrPalavra As String)
   Dim strCaractere As String
   Dim QtdCliques As Integer
   Dim NumeroBotao As Integer

   For i = 1 To Len(pStrPalavra)
      strCaractere = UCase(Mid(pStrPalavra, i, 1))
      
      NumeroBotao = ObterNumeroBotao(strCaractere)
      QtdCliques = Contador(NumeroBotao, Asc(strCaractere))
      strRetorno = strRetorno & "#" & NumeroBotao & "=" & QtdCliques & vbNewLine
   Next i
End Sub

Private Function ObterNumeroBotao(strCaractere As String) As Integer
   Select Case Asc(strCaractere)
   Case Is <= 67
      ObterNumeroBotao = 2
   Case Is <= 70
      ObterNumeroBotao = 3
   Case Is <= 73
      ObterNumeroBotao = 4
   Case Is <= 76
      ObterNumeroBotao = 5
   Case Is <= 79
      ObterNumeroBotao = 6
   Case Is <= 83
      ObterNumeroBotao = 7
   Case Is <= 86
      ObterNumeroBotao = 8
   Case Is <= 90
      ObterNumeroBotao = 9
   End Select
End Function

Private Function Contador(NumeroBotao As Integer, Letra As Integer) As Integer
   Dim i As Integer, j As Integer
   Dim intIni As Integer, intFim As Integer
   
   Select Case NumeroBotao
   Case 2
      intIni = Asc("A")
      intFim = Asc("C")
   Case 3
      intIni = Asc("D")
      intFim = Asc("F")
   Case 4
      intIni = Asc("G")
      intFim = Asc("I")
   Case 5
      intIni = Asc("J")
      intFim = Asc("L")
   Case 6
      intIni = Asc("M")
      intFim = Asc("O")
   Case 7
      intIni = Asc("P")
      intFim = Asc("S")
   Case 8
      intIni = Asc("T")
      intFim = Asc("V")
   Case 9
      intIni = Asc("W")
      intFim = Asc("Z")
   End Select

   For i = intIni To intFim
      j = j + 1
      If i = Letra Then
         Contador = j
         Exit For
      End If
   Next i
End Function

