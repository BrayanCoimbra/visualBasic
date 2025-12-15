VERSION 5.00
Begin VB.Form Strings 
   Caption         =   "Form1"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Teste"
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Text2.Text = ""
   Str (Text1.Text)
   
'   If Text2.Text = Text1.Text Then
'      MsgBox "Palindromo"
'   Else
'     MsgBox "Não é palindromo"
'   End If
End Sub

Public Function Str(termo As String) As String
   Dim Tamanho As Integer
   Tamanho = Len(termo)
   
   Dim Caractere As String
   Caractere = UCase(Mid(termo, Tamanho, 1))
      
   Text2.Text = Text2.Text & Caractere
   
   If Len(Caractere) = Tamanho Then
      Exit Function
   Else
      Str (Mid(termo, 1, Tamanho - 1))
   End If
   
End Function

