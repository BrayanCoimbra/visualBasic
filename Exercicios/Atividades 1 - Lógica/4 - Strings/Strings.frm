VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Strings"
   ClientHeight    =   2040
   ClientLeft      =   4125
   ClientTop       =   1530
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   9225
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   7335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Strings.frx":0000
      Top             =   120
      Width           =   7335
   End
   Begin VB.CommandButton cmdRmvE 
      Caption         =   "Remove E"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdReverte 
      Caption         =   "Reverter texto"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdVogal 
      Caption         =   "Quantidade Vogais"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsOperadorDeString As OperaçõesComString

Private Sub cmdReverte_Click()

   Text2.Text = ""
   
   With clsOperadorDeString
      If (.InverterString(Me.Text1)) Then
         Me.Text2.Text = .strResultado
      Else
         MsgBox .strResultado, vbCritical, "Erro"
      End If
   End With
   
End Sub

Private Sub cmdRmvE_Click()
   
   Text2.Text = ""
   
   With clsOperadorDeString
      If (.RemoverLetraE(Me.Text1)) Then
         Me.Text2.Text = .strResultado
      Else
         MsgBox .strResultado, vbCritical, "Erro"
      End If
   End With

End Sub

Private Sub cmdVogal_Click()

   Text2.Text = ""
   
   With clsOperadorDeString
      If (.ContabilizarVogais(Me.Text1)) Then
         Me.Text2.Text = .strResultado
      Else
         MsgBox .strResultado, vbCritical, "Erro"
      End If
   End With

End Sub

Private Sub Form_Load()
   Set clsOperadorDeString = New OperaçõesComString
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsOperadorDeString = Nothing
End Sub
