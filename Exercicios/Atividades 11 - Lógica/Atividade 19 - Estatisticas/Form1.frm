VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Congresso de Médicos"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Label Label3 
         Caption         =   "C) Quantos homens participaram desse congresso? "
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "A) Quantas mulheres pediatras participaram desse congresso? "
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "B) Quantas mulheres participaram desse congresso? "
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   5055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'06- Dentre os 1.250 médicos que participaram de um congresso, 48% eram mulheres.
'Dentre as mulheres 9% eram pediatras. Faça um programa que possa responder
'as perguntas abaixo;
'a) Quantas mulheres pediatras participaram desse congresso?
'b) Quantas mulheres participaram desse congresso?
'c) Quantos homens participaram desse congresso?

Private Const QtdMedicos = 1250
Private QtdMedicas As Integer
Private QtdMedicoss As Integer
Private QtdMedicasPediatras As Integer

Private Function Percent(nmr As Double) As Double
   Percent = (nmr * QtdMedicos) / 100
End Function

Private Sub Label1_Click()
'a) Quantas mulheres pediatras participaram desse congresso?
   
   QtdMedicasPediatras = Percent(9)
   MsgBox "Total de Médicas Pediátras: " & QtdMedicasPediatras
End Sub

Private Sub Label2_Click()
'b) Quantas mulheres participaram desse congresso?
   
   QtdMedicas = Percent(48)
   MsgBox "Total de Médicas: " & QtdMedicas
End Sub

Private Sub Label3_Click()
'c) Quantos homens participaram desse congresso?
   
   QtdMedicoss = Percent(52)
   MsgBox "Total de Médicos: " & QtdMedicoss
End Sub
