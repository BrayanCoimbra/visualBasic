VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
                   
   Dim qtdnmr As Integer
   qtdnmr = InputBox("Digite quantos números você quer digitar.")

   If qtdnmr > 0 Then
   
   Dim colNmr() As Integer
   ReDim colNmr(1 To qtdnmr)
   
      For i = 1 To qtdnmr
         numero = InputBox("Digite o número")
         
         If IsNumeric(numero) Then
            colNmr(i) = CInt(numero)
         Else
            MsgBox "Digite apenas números"
            Exit Sub ' Saia do sub se um número inválido for digitado
         End If
      Next i
    
      ' Exibir os números digitados
      For i = 1 To qtdnmr
         Text1.Text = Text1.Text & CStr(colNmr(i)) & " "
      Next i
   
      ' Ordenar os números
      For i = 1 To qtdnmr
         For j = i + 1 To qtdnmr
            If colNmr(i) > colNmr(j) Then
               ' Trocar os números se estiverem fora de ordem
               Dim temp As Integer
               temp = colNmr(i)
               colNmr(i) = colNmr(j)
               colNmr(j) = temp
            End If
         Next j
      Next i

   ' Exibir os números ordenados
   For i = 1 To qtdnmr
       Text2.Text = Text2.Text & CStr(colNmr(i)) & " "
   Next i
   
   Else
      MsgBox "Quantidade inválida de números."
   End If

End Sub

Private Sub Text1_Click()
   Text1.Text = ""
   Text2.Text = ""
End Sub
