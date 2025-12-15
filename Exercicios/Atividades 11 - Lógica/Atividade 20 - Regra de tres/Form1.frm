VERSION 5.00
Object = "{263D3036-6BF5-11D5-A656-0080C8BAEF42}#1.5#0"; "LydiansEdit.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin LydiansOcx.vlr vlrZin 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      ForeColor       =   0
      TitNome         =   "Zinco:"
      TitLargura      =   700
      Value           =   9.75
   End
   Begin LydiansOcx.vlr vlrEst 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      ForeColor       =   0
      TitNome         =   "Estanho:"
      TitLargura      =   700
      Value           =   0.25
   End
   Begin LydiansOcx.vlr vlrCob 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      ForeColor       =   0
      TitNome         =   "Cobre:"
      TitLargura      =   700
      Value           =   15
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'07- Para a confecção de uma peça metálica, foram fundidos 15 kg de cobre, 9,75 kg de zinco e 0,25 kg estanho.
'Qual é a porcentagem de cobre dessa peça?
'Faça um programa ue receba os valores de cobre, zinco e estanho e mostre a porcentagem de cobre da peça.

Private PesoTotal As Double
Private PercentEstanho As Double
Private PercentZinco As Double
Private PercentCobre As Double
Private PercentTotal As Double
Private PesoEstanho As Double
Private PesoZinco As Double
Private PesoCobre As Double

Function RegraDeTres(PesoElemento As Double) As Double
   RegraDeTres = (100 * PesoElemento) / PesoTotal
End Function

Private Sub CalcularPesoTotal()
   PesoEstanho = Me.vlrEst
   PesoZinco = Me.vlrZin
   PesoCobre = Me.vlrCob
   PesoTotal = PesoCobre + PesoEstanho + PesoZinco
End Sub

Private Sub Command1_Click()
   CalcularPesoTotal
   
   PercentEstanho = RegraDeTres(PesoEstanho)
   PercentCobre = RegraDeTres(PesoCobre)
   PercentZinco = RegraDeTres(PesoZinco)
   
   MsgBox "Peso total: " & PesoTotal & "Kg - 100%" & vbNewLine & vbNewLine & _
         "Estanho: " & PesoEstanho & "Kg - " & PercentEstanho & " % " & vbNewLine & _
         "Zinco: " & PesoZinco & "Kg - " & PercentZinco & " % " & vbNewLine & _
         "Cobre: " & PesoCobre & "Kg - " & PercentCobre & " %" & vbNewLine _
         , vbInformation, "Dados sobre o material informado."
End Sub
