VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Conjuntos"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInter 
      Caption         =   "Intersecção"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "3 4"
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "1 2"
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdUni 
      Caption         =   "União"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OperarConjuntos As OperacoesConjuntos
Private Const Interseccao = 1
Private Const Uniao = 2


Private Sub CmdUni_Click()
   On Error GoTo CmdUni_Click_E
   
   With OperarConjuntos
      If (.ProcessarConjuntos(Uniao, Me.Text1.Text, Me.Text2)) Then
         MsgBox .gRetorno
      Else
         GoTo CmdUni_Click_E
      End If
   End With
   
   Exit Sub
   
CmdUni_Click_E:
   MsgBox "Houve um erro ao processar os conjuntos!", vbCritical, "Erro"
End Sub

Private Sub CmdInter_Click()
   On Error GoTo CmdInter_Click_E
   
   With OperarConjuntos
      If (.ProcessarConjuntos(Interseccao, Me.Text1.Text, Me.Text2)) Then
         MsgBox .gRetorno
      Else
         GoTo CmdInter_Click_E
      End If
   End With
   
   Exit Sub
   
CmdInter_Click_E:
   MsgBox "Houve um erro ao processar os conjuntos!", vbCritical, "Erro"
End Sub

Private Sub Form_Load()
   Set OperarConjuntos = New OperacoesConjuntos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set OperarConjuntos = Nothing
End Sub
