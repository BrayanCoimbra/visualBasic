VERSION 5.00
Begin VB.Form frmCalculoMatrizes 
   Caption         =   "Cálculos com Matrizes"
   ClientHeight    =   2310
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   7260
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   1920
      TabIndex        =   33
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   1800
      Width           =   1620
   End
   Begin VB.Frame fraM3 
      Caption         =   "M3 - Resultado"
      Height          =   1695
      Left            =   5460
      TabIndex        =   37
      Top             =   0
      Width           =   1695
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   8
         Left            =   960
         TabIndex        =   31
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   960
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame fraM2 
      Caption         =   "M2"
      Height          =   1695
      Left            =   3720
      TabIndex        =   36
      Top             =   0
      Width           =   1650
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   8
         Left            =   960
         TabIndex        =   22
         Text            =   "9"
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   21
         Text            =   "8"
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Text            =   "7"
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   960
         TabIndex        =   19
         Text            =   "6"
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   18
         Text            =   "5"
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Text            =   "4"
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   16
         Text            =   "3"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   15
         Text            =   "2"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Text            =   "1"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame fraM1 
      Caption         =   "M1"
      Height          =   1695
      Left            =   1920
      TabIndex        =   35
      Top             =   0
      Width           =   1695
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Text            =   "1"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Text            =   "2"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   2
         Left            =   960
         TabIndex        =   7
         Text            =   "3"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Text            =   "4"
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   4
         Left            =   600
         TabIndex        =   9
         Text            =   "5"
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   5
         Left            =   960
         TabIndex        =   10
         Text            =   "6"
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Text            =   "7"
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   7
         Left            =   600
         TabIndex        =   12
         Text            =   "8"
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   8
         Left            =   960
         TabIndex        =   13
         Text            =   "9"
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame fraOperações 
      Caption         =   "Operações"
      Height          =   1695
      Left            =   120
      TabIndex        =   34
      Top             =   0
      Width           =   1650
      Begin VB.OptionButton optMultiplicarNumeroQualquer 
         Caption         =   "Mult. Decimal"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton optTransposta 
         Caption         =   "Transposta"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optMultiplicar 
         Caption         =   "Multiplicar"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optSubtrair 
         Caption         =   "Subtrair"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optSomar 
         Caption         =   "Somar"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3840
      TabIndex        =   38
      Top             =   1890
      Width           =   3255
   End
End
Attribute VB_Name = "frmCalculoMatrizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CalculadoraDeMatriz As clsOpMatrizes
Private dblNmrDecimal As Double
Private QtdLinhasM1 As Integer
Private QtdColunasM1 As Integer
Private QtdLinhasM2 As Integer
Private QtdColunasM2 As Integer
Private QtdLinhasM3 As Integer
Private QtdColunasM3 As Integer
Private M1() As Double
Private M2() As Double
Private M3() As Double
Private Enum Operacoes
    enSoma = 1
    enSubtracao = 2
    enMultiplicacao = 3
    enMultiplicacaoPorQualquerNumero = 4
    enTransporMatriz = 5
End Enum

Private Sub Command1_Click()
   
   On Error GoTo Command1_Click_E
   
   If Not ConstruirMatrizes Then
      MsgBox "Erro ao constuir Matrizes.", vbCritical, "Form frmCalculoMatrizes"
      GoTo DestruirObjetos
   End If
   
   With CalculadoraDeMatriz
   
      If Me.optSomar Then
         If Not (.Somar(M1, M2)) Then
            MsgBox "Erro ao validar função .Somar", vbCritical, "Form frmCalculoMatrizes"
         End If

      ElseIf Me.optSubtrair Then
         If Not (.Subtrair(M1, M2)) Then
            MsgBox "Erro ao validar função .Subtrair", vbCritical, "Form frmCalculoMatrizes"
         End If

      ElseIf Me.optMultiplicar Then
         If Not (.Multiplicacao(M1, M2)) Then
            MsgBox "Erro ao validar função .Multiplicacao", vbCritical, "Form frmCalculoMatrizes"
         End If

      ElseIf Me.optTransposta Then
         If Not (.Transpor(M1)) Then
            MsgBox "Erro ao validar função .Transpor", vbCritical, "Form frmCalculoMatrizes"
         End If

      ElseIf Me.optMultiplicarNumeroQualquer Then
         If Not (.MultiplicacaoPorNumeroQualquer(M1, dblNmrDecimal)) Then
            MsgBox "Erro ao validar função .MultiplicacaoPorNumeroQualquer", vbCritical, "Form frmCalculoMatrizes"
         End If
      
      Else
          MsgBox "Selecione um opção.", vbExclamation, "Atenção"

      End If
      
      If Not MontarMatrizResultante Then
         GoTo Command1_Click_E
      End If
      
      GoTo DestruirObjetos
      
   End With
   
Command1_Click_E:
   MsgBox "Não foi possível realizar a operação. Erro - " & Err.Description, vbCritical, "Form frmCalculoMatrizes - Sub Command1_Click"

DestruirObjetos:
   Erase M1, M2, M3
End Sub

Private Function MontarMatrizResultante() As Boolean
'Responsável por popular apenas a matriz resultante
   On Error GoTo MontarMatrizResultante_E
   
   MontarMatrizResultante = False
   
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   
   With CalculadoraDeMatriz
      For i = 0 To QtdLinhasM3 - 1
         For j = 0 To QtdColunasM3 - 1
            For k = 0 To 8
               If Me.txtM3(k).Enabled = True Then
                  Me.txtM3(k) = .MatrizResultante(i, j)
                  Me.txtM3(k).Enabled = False
                  Exit For
               End If
            Next k
         Next j
      Next i
   End With
   
   MontarMatrizResultante = True
   
   Exit Function
   
MontarMatrizResultante_E:
   MsgBox "Não foi possível montar a matriz resultante. Erro - " & Err.Description, vbCritical, "Form frmCalculoMatrizes - Function MontarMatrizResultante"
   
End Function

Private Function ExibirMatrizes(pQtdLinhasM1 As Integer, pQtdColunasM1 As Integer, pQtdLinhasM2 As Integer, pQtdColunasM2 As Integer, pQtdLinhasM3 As Integer, pQtdColunasM3 As Integer) As Boolean
'Função responsável por exibir e construis os textBoxs que representarão as matrizes

   On Error GoTo ExibirMatrizes_E

   ExibirMatrizes = False

   Dim linha As Integer
   Dim coluna As Integer

   For linha = 0 To pQtdLinhasM1 - 1
      For coluna = 0 To pQtdColunasM1 - 1
         Me.txtM1(linha * 3 + coluna).Enabled = True
         Me.txtM1(linha * 3 + coluna).Visible = True
      Next coluna
   Next linha
   
   For linha = 0 To pQtdLinhasM2 - 1
      For coluna = 0 To pQtdColunasM2 - 1
         Me.txtM2(linha * 3 + coluna).Enabled = True
         Me.txtM2(linha * 3 + coluna).Visible = True
      Next coluna
   Next linha
   
   For linha = 0 To pQtdLinhasM3 - 1
      For coluna = 0 To pQtdColunasM3 - 1
         Me.txtM3(linha * 3 + coluna).Enabled = True
         Me.txtM3(linha * 3 + coluna).Visible = True
      Next coluna
   Next linha

   ExibirMatrizes = True

   Exit Function

ExibirMatrizes_E:
   MsgBox "Não foi possível exibir as matrizes para realizar as operações. Erro - " & Err.Description, vbCritical, "Form frmCalculoMatrizes - Function ConstruirMatrizes"

End Function

Private Function ConstruirMatrizes() As Boolean
'Função responsável por popular as variaveis da primeira e segunda matriz a fim de utilizar estes valores nas operações posteriormente

   On Error GoTo ConstruirMatrizes_E

   ConstruirMatrizes = False

   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   
   'Se o respectivo textBox estiver ativado, significa que será utilizado para representar uma matriz, logo, deve populado

   For i = 0 To QtdLinhasM1 - 1
      For j = 0 To QtdColunasM1 - 1
         For k = 0 To 8
            If Me.txtM1(k).Enabled = True Then
               M1(i, j) = CDbl(Me.txtM1(k).Text)
               Me.txtM1(k).Enabled = False 'Desativa o textBox para não haver interferencia durante a operação
               Exit For
            End If
         Next k
      Next j
   Next i

   k = 0

   For i = 0 To QtdLinhasM2 - 1
      For j = 0 To QtdColunasM2 - 1
         For k = 0 To 8
            If Me.txtM2(k).Enabled = True Then
               M2(i, j) = CDbl(Me.txtM2(k).Text)
               Me.txtM2(k).Enabled = False
               Exit For
            End If
         Next k
      Next j
   Next i

   ConstruirMatrizes = True

   Exit Function

ConstruirMatrizes_E:
   MsgBox "Não foi possível construir as matrizes para realizar as operações. Erro - " & Err.Description, vbCritical, "Form frmCalculoMatrizes - Function ConstruirMatrizes"

End Function

Private Function ReceberValidarEntradas(Operacao As Integer) As Boolean
   On Error GoTo ReceberValidarEntradas_E
   
   ReceberValidarEntradas = False
   
   QtdLinhasM1 = InputBox("Digite a quantidade de LINHAS da Matriz 1.", "Operação com Matrizes")
   QtdColunasM1 = InputBox("Digite a quantidade de COLUNAS Matriz 1.", "Operação de Matrizes")
   
   If Not (Operacao = enMultiplicacaoPorQualquerNumero) And Not (Operacao = enTransporMatriz) Then
      QtdLinhasM2 = InputBox("Digite a quantidade de LINHAS da Matriz 2.", "Operação de Matrizes")
      QtdColunasM2 = InputBox("Digite a quantidade de COLUNAS Matriz 2.", "Operação de Matrizes")
   ElseIf (Operacao = enMultiplicacaoPorQualquerNumero) Then
      dblNmrDecimal = InputBox("Digite um número decimal", "Multiplicar por número decimal qualquer", "")
   End If
   
   If Not IsNumeric(QtdLinhasM1) And IsNumeric(QtdColunasM1) And IsNumeric(QtdLinhasM2) And IsNumeric(QtdColunasM2) Then
      MsgBox "Digite apenas números", 48, "ATENÇÃO"
   End If
   
   If QtdLinhasM1 > 3 Or QtdColunasM1 > 3 Or QtdLinhasM2 > 3 Or QtdColunasM2 > 3 Or QtdLinhasM1 <= 0 Or QtdColunasM1 <= 0 Or QtdLinhasM2 <= 0 Or QtdColunasM2 <= 0 Then
      MsgBox "Você pode operar apenas matrizes de 1x1 até 3x3", 48, "ATENÇÃO"
      Exit Function
   End If
      
   Select Case Operacao
      Case enSoma
         If Not QtdLinhasM1 = QtdLinhasM2 And QtdColunasM1 = QtdColunasM2 Then
            MsgBox "Para Somar duas matrizes, você deve utilizar o mesmo número de linhas por colunas. (3x3, 2x2...)", vbInformation
            QtdLinhasM1 = 0
            QtdColunasM1 = 0
            QtdLinhasM2 = 0
            QtdColunasM2 = 0
            Exit Function
         End If
        
         QtdLinhasM3 = QtdLinhasM1
         QtdColunasM3 = QtdColunasM1
         
         ReDim M1(0 To QtdLinhasM1 - 1, 0 To QtdColunasM1 - 1)
         ReDim M2(0 To QtdLinhasM2 - 1, 0 To QtdColunasM2 - 1)
      
      Case enSubtracao
         If QtdLinhasM1 <> QtdColunasM1 Then
            MsgBox "Para subtratir duas matrizes, você deve utilizar o mesmo número de linhas por colunas. (3x3, 2x2...)", vbInformation
            QtdLinhasM1 = 0
            QtdColunasM1 = 0
            Exit Function
         End If

         QtdLinhasM2 = QtdLinhasM1
         QtdColunasM2 = QtdColunasM1
         QtdLinhasM3 = QtdLinhasM1
         QtdColunasM3 = QtdColunasM1
         
         ReDim M1(0 To QtdLinhasM1 - 1, 0 To QtdColunasM1 - 1)
         ReDim M2(0 To QtdLinhasM1 - 1, 0 To QtdColunasM1 - 1)
      
      Case enMultiplicacao
         If CInt(QtdColunasM1) <> CInt(QtdLinhasM2) Then
            MsgBox "A multiplicação de matrizes só é possível quando o número de colunas da primeira matriz é igual ao número de linhas da segunda matriz.", vbInformation
            QtdLinhasM1 = 0
            QtdColunasM1 = 0
            QtdLinhasM2 = 0
            QtdColunasM2 = 0
            Exit Function
         End If
         
         QtdLinhasM3 = QtdLinhasM1
         QtdColunasM3 = QtdColunasM2
         
         ReDim M1(0 To QtdLinhasM1 - 1, 0 To QtdColunasM1 - 1)
         ReDim M2(0 To QtdLinhasM2 - 1, 0 To QtdColunasM2 - 1)
      
      Case enMultiplicacaoPorQualquerNumero
         QtdLinhasM3 = QtdLinhasM1
         QtdColunasM3 = QtdColunasM1
         
         ReDim M1(0 To QtdLinhasM1 - 1, 0 To QtdColunasM1 - 1)
         
      Case enTransporMatriz
         QtdLinhasM3 = QtdColunasM1
         QtdColunasM3 = QtdLinhasM1
         
         ReDim M1(0 To QtdLinhasM1 - 1, 0 To QtdColunasM1 - 1)
         
   End Select
      
   ReceberValidarEntradas = True
   
   Exit Function
   
ReceberValidarEntradas_E:
   MsgBox "Erro ao receber e validas valores para realizar as operações. Erro - " & Err.Description, vbCritical, "Form frmCalculoMatrizes - Function ReceberValidarEntradas"
   
End Function

Private Sub optMultiplicar_Click()
   On Error GoTo optMultiplicar_Click_E
   
   If Not ReceberValidarEntradas(enMultiplicacao) Then
      Exit Sub
   End If
            
   If Not ExibirMatrizes(QtdLinhasM1, QtdColunasM1, QtdLinhasM2, QtdColunasM2, QtdLinhasM3, QtdColunasM3) Then
      GoTo optMultiplicar_Click_E
   End If
   
   Me.Label1.Caption = "M1 = " & QtdLinhasM1 & "x" & QtdColunasM1 & "   M2 = " & QtdLinhasM2 & "x" & QtdColunasM2 & "   M3 = " & QtdLinhasM3 & "x" & QtdColunasM3
                     
   Exit Sub

optMultiplicar_Click_E:
   MsgBox Err.Description, vbCritical, "Form frmCalculoMatrizes - Sub optMultiplicar_Click"
   
End Sub

Private Sub optMultiplicarNumeroQualquer_Click()
   On Error GoTo optMultiplicarNumeroQualquer_Click_E
   
   If Not ReceberValidarEntradas(enMultiplicacaoPorQualquerNumero) Then
      Exit Sub
   End If
         
   If Not ExibirMatrizes(QtdLinhasM1, QtdColunasM1, QtdLinhasM2, QtdColunasM2, QtdLinhasM3, QtdColunasM3) Then
      GoTo optMultiplicarNumeroQualquer_Click_E
   End If
   
   Me.Label1.Caption = "M1 = " & QtdLinhasM1 & "x" & QtdColunasM1 & "   M3 = " & QtdLinhasM3 & "x" & QtdColunasM3
                     
   Exit Sub

optMultiplicarNumeroQualquer_Click_E:
   MsgBox Err.Description, vbCritical, "Form frmCalculoMatrizes - Sub optMultiplicarNumeroQualquer_Click"
   
End Sub

 Private Sub optSomar_Click()
   On Error GoTo optSomar_Click_E
     
   If Not ReceberValidarEntradas(enSoma) Then
      Exit Sub
   End If
      
   If Not ExibirMatrizes(QtdLinhasM1, QtdColunasM1, QtdLinhasM2, QtdColunasM2, QtdLinhasM3, QtdColunasM3) Then
      GoTo optSomar_Click_E
   End If

   Me.Label1.Caption = "M1 = " & QtdLinhasM1 & "x" & QtdColunasM1 & "   M2 = " & QtdLinhasM2 & "x" & QtdColunasM2 & "   M3 = " & QtdLinhasM3 & "x" & QtdColunasM3
   
   Exit Sub
                     
optSomar_Click_E:
   MsgBox Err.Description, vbCritical, "Form frmCalculoMatrizes - Sub optSomar_Click"
   
End Sub

Private Sub optSubtrair_Click()
   On Error GoTo optSubtrair_Click_E

   If Not ReceberValidarEntradas(enSubtracao) Then
      Exit Sub
   End If
   
   If Not ExibirMatrizes(QtdLinhasM1, QtdColunasM1, QtdLinhasM2, QtdColunasM2, QtdLinhasM3, QtdColunasM3) Then
      GoTo optSubtrair_Click_E
   End If
   
   Me.Label1.Caption = "M1 = " & QtdLinhasM1 & "x" & QtdColunasM1 & "   M2 = " & QtdLinhasM2 & "x" & QtdColunasM2 & "   M3 = " & QtdLinhasM3 & "x" & QtdColunasM3
   
   Exit Sub
                     
optSubtrair_Click_E:
   MsgBox Err.Description, vbCritical, "Form frmCalculoMatrizes - Sub optSubtrair_Click"
   
End Sub

Private Sub optTransposta_Click()
   On Error GoTo optTransposta_Click_E

   If Not ReceberValidarEntradas(enTransporMatriz) Then
      Exit Sub
   End If
      
   If Not ExibirMatrizes(QtdLinhasM1, QtdColunasM1, QtdLinhasM2, QtdColunasM2, QtdLinhasM3, QtdColunasM3) Then
      GoTo optTransposta_Click_E
   End If
   
   Me.Label1.Caption = "M1 = " & QtdLinhasM1 & "x" & QtdColunasM1 & "   M3 = " & QtdLinhasM3 & "x" & QtdColunasM3
                     
   Exit Sub
                     
optTransposta_Click_E:
   MsgBox Err.Description, vbCritical, "Form frmCalculoMatrizes - Sub optTransposta_Click"
   
End Sub

Private Sub cmdLimpar_Click()
'Realiza a limpeza do form, desativa e desabila a visualização dos textBoxs que representam as matrizes

   On Error GoTo cmdLimpar_Click_E
   
   Dim i As Integer
   
   For i = 0 To 8
      Me.txtM1(i).Text = 0
      Me.txtM2(i).Text = 0
      Me.txtM3(i).Text = 0

      Me.txtM1(i).Enabled = False
      Me.txtM2(i).Enabled = False
      Me.txtM3(i).Enabled = False
      
      
      Me.txtM1(i).Visible = False
      Me.txtM2(i).Visible = False
      Me.txtM3(i).Visible = False
   Next i
      
   Me.Label1.Caption = ""
   
   QtdLinhasM1 = 0
   QtdColunasM1 = 0
   QtdLinhasM2 = 0
   QtdColunasM2 = 0
   QtdLinhasM3 = 0
   QtdColunasM3 = 0
   
   Me.optSomar.Value = False
   Me.optSubtrair.Value = False
   Me.optTransposta.Value = False
   Me.optMultiplicar.Value = False
   Me.optMultiplicarNumeroQualquer.Value = False
   
   Exit Sub

cmdLimpar_Click_E:
   MsgBox "Erro ao limpar a tela - " & Err.Description, vbCritical, "Form frmCalculoMatrizes - Sub cmdLimpar_Click"
   
End Sub

Private Sub Form_Load()
   Set CalculadoraDeMatriz = New clsOpMatrizes
   cmdLimpar_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set CalculadoraDeMatriz = Nothing
End Sub

