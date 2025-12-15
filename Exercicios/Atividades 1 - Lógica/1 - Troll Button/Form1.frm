VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFechar 
      BackColor       =   &H8000000D&
      Caption         =   "Fechar"
      Height          =   600
      Left            =   1440
      MaskColor       =   &H00808080&
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFechar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Random As Integer
    Random = Int((2000 * Rnd) + 500)
    
    cmdFechar.Left = Random
    cmdFechar.Top = Random
    
End Sub
