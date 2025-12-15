VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "..."
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Modo 1"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modo 2"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
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
Private strArray() As String
Private intContador As Integer

Private Sub Command1_Click()
'Circular

   Dim strPalavra As String
   Dim strNovaPalavra As String
   Dim colLetra As Collection

   Dim i As Integer
   Dim temp As String
   Dim letraAtual As String
   
   Me.Text1.Text = ""
   strPalavra = InputBox("Digite uma palavra ou número.", "Ordenation.")
   
   Me.Caption = strPalavra
   
   OrganizarPalavra strPalavra

   For i = 1 To UBound(strArray)
      For j = 1 To UBound(strArray) - 1
         letraAtual = UCase(strArray(j))
         strArray(j) = strArray(j + 1)
         strArray(j + 1) = letraAtual
         PrintPalavra
      Next j
      Debug.Print VbevwLine
   Next i
   
End Sub

Private Sub Command2_Click()
   Dim strPalavra As String
   Dim strNovaPalavra As String
   Dim colLetra As Collection

   Dim i As Integer
   Dim temp As String
   Dim letraAtual As String
   
   Me.Text1.Text = ""
   strPalavra = InputBox("Digite uma palavra ou número.", "Ordenation.")
   
   Me.Caption = strPalavra
   
   OrganizarPalavraComOrdenacao strPalavra
   
   For i = 1 To UBound(strArray)
      PrintPalavra
      For j = 2 To UBound(strArray) - 1
         letraAtual = UCase(strArray(j))
         strArray(j) = strArray(j + 1)
         strArray(j + 1) = letraAtual
         PrintPalavra
      Next j
      
      Debug.Print VbewLine
      OrganizarPalavraComOrdenacao strPalavra
   Next i
   
   intContador = 0
End Sub

Private Sub OrganizarPalavraComOrdenacao(strTermo As String)

   intContador = intContador + 1

   ReDim strArray(1 To Len(strTermo))

   'Reiniciar o termo
   For i = 1 To Len(strTermo)
      strArray(i) = Mid(strTermo, i, 1)
   Next i
   
   If intContador <= UBound(strArray) Then
      For i = 2 To intContador
          temp = strArray(1)
          strArray(1) = strArray(i)
          strArray(i) = temp
      Next i
   End If
   
End Sub

Private Sub OrganizarPalavra(strTermo As String)

   ReDim strArray(1 To Len(strTermo))
   
   For i = 1 To Len(strTermo)
      strArray(i) = Mid(strTermo, i, 1)
   Next i
   
End Sub

Private Sub PrintPalavra()

   For k = 1 To UBound(strArray)
      strNovaPalavra = strNovaPalavra & strArray(k)
   Next k
   
   Text1.Text = Text1.Text & strNovaPalavra & vbNewLine & vbNewLine
   
   Debug.Print strNovaPalavra
   strNovaPalavra = ""
   
End Sub
