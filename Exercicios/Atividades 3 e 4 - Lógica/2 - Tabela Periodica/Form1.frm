VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Elementos"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tabela Peródica"
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      Begin VB.Label lblCarbono 
         BackColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   6600
         TabIndex        =   7
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label lblEstroncio 
         BackColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   1200
         TabIndex        =   6
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label lblRubidio 
         Height          =   15
         Left            =   720
         TabIndex        =   5
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label lblPotassio 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblActinio 
         BackColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   2040
         TabIndex        =   3
         Top             =   4800
         Width           =   135
      End
      Begin VB.Label lblTorio 
         BackColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   2520
         TabIndex        =   2
         Top             =   4800
         Width           =   135
      End
      Begin VB.Label lblUranio 
         BackColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   3480
         TabIndex        =   1
         Top             =   4800
         Width           =   135
      End
      Begin VB.Image Image1 
         Height          =   4980
         Left            =   360
         Picture         =   "Form1.frx":0000
         Top             =   360
         Width           =   8520
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lblActinio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Var = MsgBox("O actínio (Ac) é um elemento químico metálico pertencente à classe dos metais de transição, radioativo, macio, branco-prateado que brilha no escuro e localiza-se no grupo 3 e período 7 da Tabela Periódica.", 46, "INFORMAÇÃO")
End Sub

Private Sub lblCarbono_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Var = MsgBox("O carbono-14 (C-14) é um isótopo radioativo do carbono que ocorre naturalmente na atmosfera terrestre e é usado em datação por radiocarbono para determinar a idade de materiais orgânicos.", 46, "INFORMAÇÃO")
End Sub

Private Sub lblEstroncio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Var = MsgBox("Ambos o rubídio e o estrôncio possuem isótopos radioativos, como o rubídio-87 (Rb-87) e o estrôncio-90 (Sr-90), respectivamente, que são produzidos como produtos de decaimento de elementos mais pesados.", 46, "INFORMAÇÃO")
End Sub

Private Sub lblPotassio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Var = MsgBox("O potássio é um elemento essencial para a vida e é encontrado em muitos alimentos. Uma pequena fração do potássio é composta pelo isótopo radioativo potássio-40 (K-40), que contribui para a radioatividade natural da Terra.", 46, "INFORMAÇÃO")
End Sub

Private Sub lblRubidio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Var = MsgBox("Ambos o rubídio e o estrôncio possuem isótopos radioativos, como o rubídio-87 (Rb-87) e o estrôncio-90 (Sr-90), respectivamente, que são produzidos como produtos de decaimento de elementos mais pesados.", 46, "INFORMAÇÃO")
End Sub

Private Sub lblTorio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Var = MsgBox("O tório é outro elemento radioativo natural que ocorre naturalmente na crosta terrestre. Seu isótopo mais comum é o tório-232 (Th-232). O tório também tem aplicações em energia nuclear e em reatores nucleares.", 46, "INFORMAÇÃO")
End Sub

Private Sub lblUranio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Var = MsgBox("O urânio é talvez o elemento radioativo natural mais conhecido e amplamente estudado. Seus isótopos mais comuns são o urânio-238 (U-238) e o urânio-235 (U-235), ambos utilizados em aplicações nucleares e na produção de energia.", 46, "INFORMAÇÃO")
End Sub
