VERSION 5.00
Begin VB.Form frmIMC 
   Caption         =   "Cálculo | IMC"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcularClass 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtBoxIMC 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Seu IMC aparecerá aqui"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtBoxPeso 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "100"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtBoxAltura 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "1,82"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblIMC 
      Caption         =   "Seu IMC é:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblPeso 
      Caption         =   "Peso"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblAltura 
      Caption         =   "Altura"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmIMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalcularClass_Click()
    On Error GoTo TratarErro
    
    Dim CalculoIMC As CalculoIMC
    Set CalculoIMC = New CalculoIMC
        
    CalculoIMC.altura = CDbl(txtBoxAltura.Text)
    CalculoIMC.peso = CDbl(txtBoxPeso.Text)
    
    If Not CalculoIMC.CalcularIMC Then
        MsgBox CalculoIMC.Erro
    End If
    
    txtBoxIMC.Text = CalculoIMC.IMC
    
    If Val(txtBoxIMC.Text) < 18.5 Then
       txtBoxIMC.BackColor = &H8080FF
       
    ElseIf Val(txtBoxIMC.Text) >= 18.6 And Val(txtBoxIMC.Text) <= 24.9 Then
       txtBoxIMC.BackColor = &HC0FFC0
       
    ElseIf Val(txtBoxIMC.Text) >= 25 And Val(txtBoxIMC.Text) <= 29.9 Then
       txtBoxIMC.BackColor = &H80C0FF
       
    ElseIf Val(txtBoxIMC.Text) >= 30 And Val(txtBoxIMC.Text) <= 34.9 Then
       txtBoxIMC.BackColor = &H8080FF
       
    ElseIf Val(txtBoxIMC.Text) >= 35 And Val(txtBoxIMC.Text) <= 39.9 Then
       txtBoxIMC.BackColor = &HFF&
       
    Else
       txtBoxIMC.BackColor = &HC0&
    End If
    
    GoTo LimparObjetos
    
TratarErro:
    MsgBox Err.Description
    
LimparObjetos:
    Set CalculoIMC = Nothing
End Sub
