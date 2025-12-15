VERSION 5.00
Object = "{263D3036-6BF5-11D5-A656-0080C8BAEF42}#1.4#0"; "LydiansEdit.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parenteses"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraResultado 
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   10815
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
         Height          =   5055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   360
         Width           =   10575
      End
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin LydiansOcx.txt txt1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      TitNome         =   "Nº  Parenteses"
      Text            =   "3"
      TitLargura      =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsParenteses As clsCombinacoes

Private Sub cmdCalcular_Click()
   Text1.Text = ""
   
   Set clsParenteses = New clsCombinacoes
   
   Dim colResultado As New Collection
   Dim strResultado As String
   strResultado = ""
   Dim i As Integer

   On Error GoTo cmdCalcular_Click_E
   
   If CInt(Me.txt1.Text) < 0 Or CInt(Me.txt1.Text) > 6 Then
      MsgBox "Informe um número válido. Número inteiro de 0 a 6!", 48, "ATENÇÃO"
      GoTo DestruirObjetos
   End If
   
   With clsParenteses
   
      If .Combinar(txt1.Text) Then
         fraResultado.Caption = "Número de combinações possíveis: " & .gCombinacao & " "
      Else
         MsgBox "Erro em Form - Validação clsColchetes.Combinar!", 48, "ERRO"
         GoTo DestruirObjetos
      End If
      
      If .MontarParenteses(CInt(txt1.Text), 0, 0, strResultado, colResultado) Then
         For i = CInt(colResultado.Count) To 1 Step -1
            Me.Text1.Text = Me.Text1.Text & colResultado(i) & vbNewLine & vbNewLine
         Next i
      Else
         MsgBox "Erro em Form - Validação clsColchetes.MontarParenteses!", 48, "ERRO"
         GoTo DestruirObjetos
      End If

   End With
   
   GoTo DestruirObjetos

cmdCalcular_Click_E:
   MsgBox "Erro em Form - Sub cmdCalcular_Click!", 16, "ERRO"
DestruirObjetos:
   Set clsParenteses = Nothing
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_E
   
   Set clsParenteses = New clsCombinacoes
   
   GoTo DestruirObjetos

Form_Load_E:
   MsgBox "Erro em Form - Sub Form_Load!", 16, "ERRO"
   
DestruirObjetos:

End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_E

   Set clsParenteses = Nothing
   
   GoTo DestruirObjetos

Form_Unload_E:
   MsgBox "Erro em Form - Sub Form_Unload!", 16, "ERRO"
   
DestruirObjetos:

End Sub
