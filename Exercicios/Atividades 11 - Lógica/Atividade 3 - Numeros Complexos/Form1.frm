VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Número Complexos"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404040&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   18.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      MaskColor       =   &H00404040&
      TabIndex        =   18
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   "Operações"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   3240
      TabIndex        =   12
      Top             =   120
      Width           =   2055
      Begin VB.CheckBox chkMod 
         BackColor       =   &H00404040&
         Caption         =   "Divisão Modular"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox chkMult 
         BackColor       =   &H00404040&
         Caption         =   "Multiplicação"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox chkSub 
         BackColor       =   &H00404040&
         Caption         =   "Subtração"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkSom 
         BackColor       =   &H00404040&
         Caption         =   "Somar"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "Z Result"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   3015
      Begin VB.TextBox txtResultado2 
         Alignment       =   2  'Center
         BackColor       =   &H00F3F3F3&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtResultado1 
         Alignment       =   2  'Center
         BackColor       =   &H00F3F3F3&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   18.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Z2"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
      Begin VB.TextBox txt_Nmr2_Z2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_Nmr1_Z2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Z1"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txt_Nmr2_Z1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_Nmr1_Z1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   480
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkMod_Click()
   Dim DivMod As clsComplexMath
   Set DivMod = New clsComplexMath
   
   If IsNumeric(txt_Nmr1_Z1.Text) And IsNumeric(txt_Nmr2_Z1.Text) And IsNumeric(txt_Nmr1_Z2.Text) And IsNumeric(txt_Nmr2_Z2.Text) Then
      If (DivMod.Modulo(CDbl(txt_Nmr1_Z1), CDbl(txt_Nmr2_Z1), CDbl(txt_Nmr1_Z2), CDbl(txt_Nmr2_Z2))) Then
         Label4 = "+"
         txtResultado1.Text = CStr(DivMod.ZResult.dblParteReal) & "/" & CStr(DivMod.Denominador)
         txtResultado2.Text = CStr(DivMod.ZResult.dblParteImaginaria) & "/" & CStr(DivMod.Denominador)
      Else
         MsgBox "Erro"
      End If
   Else
      MsgBox "Você não preencher todos os campos ou preencheu incorretamente. Preencha os campo utilizando apenas números!", , "Atenção"
   End If
   
End Sub

Private Sub chkMult_Click()
   Dim Multiplicacao As clsComplexMath
   Set Multiplicacao = New clsComplexMath
   
   If IsNumeric(txt_Nmr1_Z1.Text) And IsNumeric(txt_Nmr2_Z1.Text) And IsNumeric(txt_Nmr1_Z2.Text) And IsNumeric(txt_Nmr2_Z2.Text) Then
      If (Multiplicacao.Multiplicar(CDbl(txt_Nmr1_Z1), CDbl(txt_Nmr2_Z1), CDbl(txt_Nmr1_Z2), CDbl(txt_Nmr2_Z2))) Then
         Label4 = "+"
         txtResultado1.Text = CStr(Multiplicacao.ZResult.dblParteReal)
         txtResultado2.Text = CStr(Multiplicacao.ZResult.dblParteImaginaria)
      Else
         MsgBox "Erro"
      End If
   Else
      MsgBox "Você não preencher todos os campos ou preencheu incorretamente. Preencha os campo utilizando apenas números!", , "Atenção"
   End If
End Sub

Private Sub chkSom_Click()
   Dim Soma As clsComplexMath
   Set Soma = New clsComplexMath
   
   If IsNumeric(txt_Nmr1_Z1.Text) And IsNumeric(txt_Nmr2_Z1.Text) And IsNumeric(txt_Nmr1_Z2.Text) And IsNumeric(txt_Nmr2_Z2.Text) Then
      If (Soma.Somar(CDbl(txt_Nmr1_Z1), CDbl(txt_Nmr2_Z1), CDbl(txt_Nmr1_Z2), CDbl(txt_Nmr2_Z2))) Then
         Label4 = "+"
         txtResultado1.Text = CStr(Soma.ZResult.dblParteReal)
         txtResultado2.Text = CStr(Soma.ZResult.dblParteImaginaria)
      Else
         MsgBox "Erro"
      End If
   Else
      MsgBox "Você não preencher todos os campos ou preencheu incorretamente. Preencha os campo utilizando apenas números!", , "Atenção"
   End If
End Sub

Private Sub chkSub_Click()
   Dim Subtracao As clsComplexMath
   Set Subtracao = New clsComplexMath
   
   If IsNumeric(txt_Nmr1_Z1.Text) And IsNumeric(txt_Nmr2_Z1.Text) And IsNumeric(txt_Nmr1_Z2.Text) And IsNumeric(txt_Nmr2_Z2.Text) Then
      If (Subtracao.Subtrair(CDbl(txt_Nmr1_Z1), CDbl(txt_Nmr2_Z1), CDbl(txt_Nmr1_Z2), CDbl(txt_Nmr2_Z2))) Then
         txtResultado1.Text = CStr(Subtracao.ZResult.dblParteReal)
         txtResultado2.Text = CStr(Subtracao.ZResult.dblParteImaginaria)
      Else
         MsgBox "Erro"
      End If
   Else
      MsgBox "Você não preencher todos os campos ou preencheu incorretamente. Preencha os campo utilizando apenas números!", , "Atenção"
   End If
End Sub

Private Sub Command1_Click()
   End
End Sub
