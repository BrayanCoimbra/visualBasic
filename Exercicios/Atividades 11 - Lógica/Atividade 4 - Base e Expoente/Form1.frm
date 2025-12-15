VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   1800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Power(base As Double, expoente As Double) As Double
   If expoente = 1 Then
      Power = base
      Exit Function
   Else
      Power = base * Power(base, expoente - 1)
   End If
End Function

Private Sub Command1_Click()
   Text3.Text = CStr(Power(CDbl(Text1.Text), CDbl(Text2.Text)))
End Sub
