VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Palíndromos"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "3010"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "3000"
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim colResultado(1 To 1000) As Integer
   Dim colStr(1 To 1000) As String
   
   'Validação de entradas
   If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) And (CInt(Text1.Text) < CInt(Text2.Text)) And CInt(Text2.Text) <= 4000 Then
      Palindromo CInt(Text1.Text), CInt(Text2.Text)
   Else
      MsgBox "Você deve preencher os campos com um intervalo entre dois números, do menor para o maior. Exemplo: 3000 e 3010." & vbNewLine & vbNewLine & "Não utilize valores maiores que 4K.", vbInformation, "Atenção"
   End If
End Sub

Private Function Palindromo(nmr1 As Integer, nmr2 As Integer) As String
   Dim colResultado(1 To 1000) As Integer
   Dim colStr(1 To 1000) As String
      
   'Encontra o intervalo entre os números para realizar a soma
   Dim intervalo As Integer
   intervalo = nmr2 - nmr1
   
   'Realiza a soma com base no intervalo encontrado
   For i = 1 To intervalo
      If i = 1 Then
         colResultado(i) = colResultado(i) + nmr1 + 1
      Else
         colResultado(i) = colResultado(i - 1) + 1
      End If
   Next i
          
   'Salva todos os inteiros como string em uma nova collection
   For i = 1 To intervalo
      colStr(i) = CStr(colResultado(i))
   Next i
         
   'Utiliza as strings para realizar a manipulação
   For i = 1 To intervalo
   
   'Monta o primeiro número
      For j = 1 To Len(colStr(i))
         Caractere = Caractere & Mid(colStr(i), j, 1)
         Tamanho = Len(Caractere)
      Next j
            
   'Monta o mesmo número invertido para realizar a comparação
      For k = 1 To Len(Caractere)
         If k = 1 Then
            CaractereInvertido = CaractereInvertido + Mid(Caractere, Len(Caractere), 1)
         Else
            CaractereInvertido = CaractereInvertido + Mid(Caractere, Tamanho - 1, 1)
            Tamanho = Tamanho - 1
         End If
      Next k
            
   'Compara as string montadas e salva se forem iguais
      If Caractere = CaractereInvertido Then
         Resultado = Resultado & vbNewLine & Caractere & " na mosca!"
      End If
            
   'Limpa as variáveis auxiliares para reutilizar
      Caractere = ""
      CaractereInvertido = ""
   Next i
         
   'Retorna os valores palíndromos encontrados
   MsgBox "Números palíndromos encontrados: " & Resultado, vbInformation, "Resultado"
End Function

