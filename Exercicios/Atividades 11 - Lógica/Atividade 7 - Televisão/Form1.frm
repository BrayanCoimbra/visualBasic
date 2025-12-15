VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   3450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Trocar Canal"
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton CONSULTACANAL 
      Caption         =   "CAN"
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton CONSULTAVOL 
      Caption         =   "VOL"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton VOLMENOS 
      Caption         =   "V -"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton VOLMAIS 
      Caption         =   "V +"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton CNMENOS 
      Caption         =   "CN -"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton CNMAIS 
      Caption         =   "CN +"
      Height          =   615
      Left            =   480
      Picture         =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   -2040
      Picture         =   "Form1.frx":1D4A
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TV As clsTV

Private Sub CNMAIS_Click()
   If TV.AumentarCanal Then
      MsgBox "Canal trocado, você está no canal " & TV.gCanalAtual, vbInformation, "Televisão"
   Else
      MsgBox "Erro ao trocar canal.", vbInformation, "Televisão"
   End If
End Sub

Private Sub CNMENOS_Click()
   If TV.DiminuirCanal Then
      MsgBox "Canal trocado, você está no canal " & TV.gCanalAtual, vbInformation, "Televisão"
   Else
      MsgBox "Erro ao trocar canal.", vbInformation, "Televisão"
   End If
End Sub

Private Sub Command1_Click()
   End
End Sub

Private Sub Command2_Click()
   Dim ValorCanal As Integer
   ValorCanal = InputBox("Digite o número do canal", "Canal")
   
   If TV.TrocarCanal(ValorCanal) Then
      MsgBox "Canal trocado, você está no canal " & TV.gCanalAtual, vbInformation, "Televisão"
   Else
      MsgBox "Este canal não existe, você pode acessar os canais de 1 - 10", vbInformation, "Televisão"
   End If
End Sub

Private Sub CONSULTACANAL_Click()
   MsgBox "Canal atual: " & CStr(TV.gCanalAtual), vbInformation, "Televisão"
End Sub

Private Sub CONSULTAVOL_Click()
   MsgBox "Volume atual: " & CStr(TV.gVolume), vbInformation, "Televisão"
End Sub

Private Sub Form_Load()
   Set TV = New clsTV
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set TV = Nothing
End Sub

Private Sub VOLMAIS_Click()
   If (TV.AumentarVolume) Then
      MsgBox "Volume " & TV.gVolume, vbInformation, "Televisão"
   Else
      MsgBox "Erro ao aumentar volume.", vbInformation, "Televisão"
   End If
End Sub

Private Sub VOLMENOS_Click()
   If TV.DiminuirVolume Then
      MsgBox "Volume " & TV.gVolume, vbInformation, "Televisão"
   Else
      MsgBox "Erro ao diminuir volume.", vbInformation, "Televisão"
   End If
End Sub

