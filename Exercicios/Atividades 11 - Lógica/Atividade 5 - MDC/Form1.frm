VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   2520
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "MDC"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then

       Dim numero1 As Integer
       Dim numero2 As Integer
       numero1 = CInt(Text1.Text)
       numero2 = CInt(Text2.Text)
       
       'Chama a função Mdc para calcular o MDC dos números
       Dim resultado As Integer
       resultado = Mdc(numero1, numero2)
       
      'Exibe o resultado
      MsgBox "O MDC de " & numero1 & " e " & numero2 & " é: " & resultado
   Else
      MsgBox "Insira números válidos nos campos de texto.", vbExclamation, "Erro de entrada"
   End If
End Sub

Public Function Mdc(x As Integer, y As Integer) As Integer
   If y = 0 Then
       Mdc = x ' Retorna x quando y é zero
   Else
       Mdc = Mdc(y, x Mod y)
   End If
End Function


