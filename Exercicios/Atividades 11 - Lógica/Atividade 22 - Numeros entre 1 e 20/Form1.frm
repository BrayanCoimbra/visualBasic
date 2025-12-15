VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   570
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'09- Ler 10 valores e escrever quantos desses valores lidos estão no intervalo [10,20]
'(incluindo os valores 10 e 20 no intervalo) e quantos deles estão fora deste intervalo.

Private colNumeros As Collection

Private Sub Command1_Click()
   Set colNumeros = New Collection
   
   Split Text1.Text
   
   If colNumeros.Count = 10 Then
      VerificarIntervalo
   Else
      MsgBox "Digite apenas 10 números.", vbExclamation, "Atenção"
   End If
   
   Set colNumeros = Nothing
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
' Adiciona o ponto e vírgula quando Enter é pressionado
   
   If KeyCode = vbKeyReturn Then
      Text1.Text = Text1.Text & "; "
      Text1.SelStart = Len(Text1.Text)
      KeyCode = 0
   End If
   
End Sub

Private Sub VerificarIntervalo()
   Dim strNumerosDentroDoIntervalo As String
   Dim strNumerosForaDoIntervalo As String
    
   For i = 1 To colNumeros.Count
      If CInt(colNumeros.Item(i)) >= 10 And CInt(colNumeros.Item(i)) <= 20 Then
         strNumerosDentroDoIntervalo = strNumerosDentroDoIntervalo & CStr(colNumeros.Item(i)) & "; "
      Else
         strNumerosForaDoIntervalo = strNumerosForaDoIntervalo & CStr(colNumeros.Item(i)) & "; "
      End If
   Next i
   
   MsgBox "Números no Intervalo 1 - 20: " & strNumerosDentroDoIntervalo & vbNewLine & _
   "Números fora Intervalo 1 - 20: " & strNumerosForaDoIntervalo, vbInformation
End Sub

Private Sub Split(strString As String)
   Dim strCaractere As String
   
   For i = 1 To Len(Text1.Text)
      strCaractere = Mid(Text1.Text, i, 1)

      If strCaractere = ";" Then
         colNumeros.Add strTermo
         strTermo = ""
      Else
         strTermo = strTermo & strCaractere
      End If
   Next i
End Sub

