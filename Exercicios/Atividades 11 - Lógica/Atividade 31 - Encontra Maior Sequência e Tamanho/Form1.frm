VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "5 10 3 2 4 7 9 8 5"
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'03- Escreva um programa no qual o usuário entra com uma quantidade arbitrária de números inteiros positivos.
'O programa deve indicar a maior sequencia indicada e seu o tamanho.
'Por exemplo, se os números digitados forem 5, 10, 3, 2, 4, 7, 9, 8, 5
'A maior sequência crescente é 2, 4, 7, 9 com tamanho 4.

Private arrNmr()
Private colAux As Collection

Private Sub Command1_Click()
   Dim ContSeqAtual As Integer
   Dim ContSeqAnterior As Integer
   Dim SeqAtual As String
   Dim SeqAnterior As String
   
   Set colAux = Split(Me.Text1.Text, " ")
   
   ReDim arrNmr(1 To colAux.Count)
   
   For i = 1 To colAux.Count
       arrNmr(i) = colAux.Item(i)
   Next i

   ContSeqAtual = 1
   SeqAtual = arrNmr(1)
   
   For i = 2 To colAux.Count
       If CInt(arrNmr(i)) > CInt(arrNmr(i - 1)) Then
           SeqAtual = SeqAtual & " " & arrNmr(i)
           ContSeqAtual = ContSeqAtual + 1
       Else
           If ContSeqAtual > ContSeqAnterior Then
               ContSeqAnterior = ContSeqAtual
               SeqAnterior = SeqAtual
           End If
           SeqAtual = arrNmr(i)
           ContSeqAtual = 1
       End If
   Next i
   
   ' Verifica a última sequência
   If ContSeqAtual > ContSeqAnterior Then
       ContSeqAnterior = ContSeqAtual
       SeqAnterior = SeqAtual
   End If
   
   Label1.Caption = ""
   Label1.Caption = "A maior sequência crescente é " & SeqAnterior & " com tamanho " & CStr(ContSeqAnterior)
   
   Set colAux = Nothing
End Sub

Private Function Split(stringParaSplit As String, caractereSeparador As String) As Collection
   Dim strCaractere As String
   Dim strTermo As String
   
   Set Split = New Collection
   
   For i = 1 To Len(stringParaSplit) + 1
       strCaractere = Mid(stringParaSplit, i, 1)
       
       If strCaractere = caractereSeparador Or strCaractere = "" Then
           If strTermo <> "" Then
               Split.Add strTermo
           End If
           strTermo = ""
       Else
           strTermo = strTermo & strCaractere
       End If
   Next i
End Function

