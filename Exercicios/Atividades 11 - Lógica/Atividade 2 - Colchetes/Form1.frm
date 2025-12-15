VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Bracket Matcher"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Comparar 
      BackColor       =   &H00404040&
      Caption         =   "Comparar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
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
'02- Comparador de colchetes
'Faça com que a função BracketMatcher(str) pegue o parâmetro str que está sendo passado e retorne 1 se os colchetes corresponderem corretamente e cada um for contabilizado, caso contrário, retorne 0.
'Por exemplo:
'1- se str for "(hello (world))", então a saída deve ser 1
'2- se str for "((hello (world))" a saída deve ser 0 porque os colchetes não correspondem corretamente.
'3- se str não contiver colchetes, retorne 1

Private Function BracketMatcher(strTermo As String) As Boolean

   On Error GoTo BracketMatcher_E
   
   BracketMatcher = False
   
   Dim count As Integer
   Dim i As Integer
   Dim strCaractere As String

   count = 0
    
   ' Iterar sobre cada caractere na string
   For i = 1 To Len(strTermo)
       strCaractere = Mid(strTermo, i, 1)
       
       ' Incrementar ou decrementar o contador com base no caractere atual
       If strCaractere = "(" Then
           count = count + 1
       ElseIf strCaractere = ")" Then
           count = count - 1
       End If
        
       ' Se o contador ficar negativo, significa que um colchete fechado foi encontrado antes de um aberto
       If count < 0 Then
           BracketMatcher = False
           Exit Function
       End If
   Next i
       
   ' Se o contador for zero no final, significa que os colchetes correspondem corretamente
   If count = 0 Then
      BracketMatcher = True
      GoTo DestruirObjetos
   Else
      BracketMatcher = False
      GoTo DestruirObjetos
   End If
   
BracketMatcher_E:
      BracketMatcher = False
      
DestruirObjetos:

End Function


Private Sub Comparar_Click()

   If Text1.Text = "" Then
      MsgBox "Digite algo para verificar!"
   Else
      If BracketMatcher(Text1.Text) Then
         MsgBox "1 - Certo", , "Resultado"
      Else
         MsgBox "0 - Errado", , "Resultado"
      End If
   End If
   
End Sub

