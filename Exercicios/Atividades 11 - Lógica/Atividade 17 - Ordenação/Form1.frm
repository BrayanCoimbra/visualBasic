VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sum Pares"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   2280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox vlr3 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   2
      Top             =   840
      Width           =   780
   End
   Begin VB.PictureBox vlr2 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   480
      Width           =   780
   End
   Begin VB.PictureBox vlr1 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Resultado:"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private colPares As Collection
Private colNumeros As Collection
Private arrNmr(1 To 3) As Integer

Private Sub Sort()
   Dim temp As Integer
   
   For i = 1 To UBound(arrNmr)
      For j = i + 1 To UBound(arrNmr)
         If arrNmr(i) > arrNmr(j) Then
            temp = arrNmr(i)
            arrNmr(i) = arrNmr(j)
            arrNmr(j) = temp
         End If
      Next j
   Next i
End Sub

Private Sub Command1_Click()
   Me.Label1.Caption = "Resultado: "
   
   Dim Soma As Integer
   arrNmr(1) = Me.vlr1
   arrNmr(2) = Me.vlr2
   arrNmr(3) = Me.vlr3
   
   Sort
   
   Soma = arrNmr(2) + arrNmr(3)
   Me.Label1.Caption = Label1 & " " & Soma
End Sub
