VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "Digite um valor decimal"
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtBoxResultado 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Text            =   "Resultado"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Converter para"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3375
      Begin VB.OptionButton optHexad 
         Caption         =   "Hexadecimal"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optOctal 
         Caption         =   "Octal"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optBinario 
         Caption         =   "Binário"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'6- Escreva um programa que converta um número decimal informado para as bases:
'a) binária
'b) octal
'c) hexadecimal

Private Sub optHexad_Click()
   DecimalParaHexad CInt(Text1.Text)
   txtBoxResultado.Text = DecimalParaHexad(CInt(Text1.Text))
End Sub

Private Sub optOctal_Click()
   DecimalParaOctal CInt(Text1.Text)
   txtBoxResultado.Text = DecimalParaOctal(CInt(Text1.Text))
End Sub

Private Sub Text1_Click()
   Text1.Text = ""
End Sub

Private Sub optBinario_Click()
   DecimalParaBinario CInt(Text1.Text)
   txtBoxResultado.Text = DecimalParaBinario(CInt(Text1.Text))
End Sub

Private Function DecimalParaBinario(NumeroDecimal As Integer) As String
Dim ResultadoBinario As String
ResultadoBinario = " "
 
   Do While NumeroDecimal > 0
      ResultadoBinario = CStr(NumeroDecimal Mod 2) & ResultadoBinario
      NumeroDecimal = NumeroDecimal \ 2
   Loop
   
   DecimalParaBinario = ResultadoBinario
End Function

Private Function DecimalParaOctal(NumeroDecimal As Integer) As String
Dim ResultadoOctal As String
ResultadoOctal = " "
 
   Do While NumeroDecimal > 0
      ResultadoOctal = CStr(NumeroDecimal Mod 8) & ResultadoOctal
      NumeroDecimal = NumeroDecimal \ 8
   Loop
   
   DecimalParaOctal = ResultadoOctal
End Function

Private Function DecimalParaHexad(NumeroDecimal As Integer) As String
Dim ResultadoHexad As String
ResultadoHexad = " "
 
   Do While NumeroDecimal > 0
      ResultadoHexad = CStr(NumeroDecimal Mod 16) & ResultadoHexad
      NumeroDecimal = NumeroDecimal \ 16
   Loop
   
   DecimalParaHexad = ResultadoHexad
End Function




