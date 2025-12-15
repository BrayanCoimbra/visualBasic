VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
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
    Dim soma As Double
    soma = 0
    
    ' Loop para somar os números positivos até 1000
    For i = 1 To 1000
        soma = soma + i
    Next i
    
    MsgBox "A soma dos números positivos até 1000 é: " & soma, vbInformation, "Soma"
End Sub

