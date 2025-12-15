VERSION 5.00
Begin VB.Form frmCalculadora 
   Caption         =   "Calculadora"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton limpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton CalcCirc 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton calcBhask 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4440
      TabIndex        =   20
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Matrizes 
      Caption         =   "Ops. Matrizes"
      Height          =   375
      Left            =   3000
      TabIndex        =   19
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Fibonacci 
      Caption         =   "Fibonacci"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton areaCirc 
      Caption         =   "Área Circulo"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Bhaskara 
      Caption         =   "Bhaskara"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Concatenacao 
      Caption         =   "Concatenar"
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Exponenciar 
      Caption         =   "Exponenciar"
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton RestoMod 
      Caption         =   "Resto - Mod"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton DivisaoInt 
      Caption         =   "Divisão inteira"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Multiplicar 
      Caption         =   "Multiplicar"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Dividir 
      Caption         =   "Dividir"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Subtrair 
      Caption         =   "Subtrair"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Somar 
      Caption         =   "Somar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtBoxValorC 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   300
      Width           =   1935
   End
   Begin VB.TextBox txtBoxResultado 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   2700
      Width           =   4575
   End
   Begin VB.TextBox txtBoxValor2 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   300
      Width           =   1815
   End
   Begin VB.TextBox txtBoxValor1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   1815
   End
   Begin VB.Label lblBhask 
      Caption         =   "BHASKARA"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   90
      Width           =   1815
   End
   Begin VB.Label lblResultado 
      Caption         =   "RESULTADO"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2745
      Width           =   1095
   End
   Begin VB.Label lblValor2 
      Caption         =   "VALOR 2"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   90
      Width           =   1455
   End
   Begin VB.Label lblValor1 
      Caption         =   "VALOR 1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   1455
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCalculadora As clsOperacoes

Private Sub areaCirc_Click()
   On Error GoTo optCalcAreaCirc_Click_E
   
   txtBoxValorC.Text = ""
   txtBoxValorC.Enabled = False
   txtBoxValor1.Text = ValorPI
   txtBoxValor2.Text = "Digite o valor do Raio"
   Me.CalcCirc.Enabled = True
   
   Exit Sub
   
optCalcAreaCirc_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub Bhaskara_Click()
   On Error GoTo optBhaskara_Click_E
      
   lblValor1.Caption = "Digite o valor de A"
   lblValor2.Caption = "Digite o valor de B"
   lblBhask.Caption = "Digite o valor de C"
   Me.txtBoxValorC.Enabled = True
   Me.calcBhask.Enabled = True
   
   Exit Sub
   
optBhaskara_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub calcBhask_Click()
   On Error GoTo chkCalcularBhaskara_Click_E
   
   If Me.txtBoxValorC.Text = "" Then
      MsgBox "Digite o valor correto para C"
      Exit Sub
   End If
   
   With clsCalculadora
      If IsNumeric(txtBoxValor1.Text) And IsNumeric(txtBoxValor2.Text) Then
         If Not (.CalcularBhaskara(CDbl(Me.txtBoxValor1), CDbl(Me.txtBoxValor2), CDbl(Me.txtBoxValorC))) Then
            GoTo chkCalcularBhaskara_Click_E
         End If
      Else
         MsgBox "Digite apenas números", 48, "Erro"
      End If
      
      txtBoxResultado = .ValorResultanteBask
   End With
   
   txtBoxValor1.Text = ""
   txtBoxValor2.Text = ""
   txtBoxValorC.Text = ""
   lblValor1.Caption = "VALOR 1"
   lblValor2.Caption = "VALOR 1"
   lblBhask.Caption = "BHASKARA"
   txtBoxValor2.Text = ""
   txtBoxValorC.Text = ""
   txtBoxValorC.Enabled = False
   limpar_Click
   
   Exit Sub
   
chkCalcularBhaskara_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub CalcCirc_Click()
   On Error GoTo chkCalcAreaCirc_Click_E
   
   If Me.txtBoxValor1.Text = "" And Me.txtBoxValor2.Text = "" Then
      MsgBox "Digite um valor númerico para operar."
      Exit Sub
   End If

   With clsCalculadora
      If IsNumeric(txtBoxValor1.Text) And IsNumeric(txtBoxValor2.Text) Then
         If Not (.CalcularAreaCirculo(CDbl(txtBoxValor2.Text))) Then
            GoTo chkCalcAreaCirc_Click_E
         End If
      Else
         MsgBox "Digite apenas números", 48, "Erro"
      End If
      
      txtBoxResultado = "A área do circulo é " & .ValorResultante
   End With
   
   txtBoxValor1.Text = ""
   txtBoxValor2.Text = ""
   
   limpar_Click
   
   Exit Sub
   
chkCalcAreaCirc_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub Concatenacao_Click()
   On Error GoTo optConcatenacao_Click_E
   
   If Me.txtBoxValor1.Text = "" And Me.txtBoxValor2.Text = "" Then
      MsgBox "Digite um valor númerico para operar."
      Exit Sub
   End If
   
   txtBoxValorC.Text = ""
   txtBoxValorC.Enabled = False
   txtBoxResultado = txtBoxValor1 & txtBoxValor2
   
   Exit Sub

optConcatenacao_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub Dividir_Click()
   On Error GoTo optDivisão_Click_E
   
   With clsCalculadora
      If IsNumeric(txtBoxValor1.Text) And IsNumeric(txtBoxValor2.Text) Then
         If Not (.Divisao(CDbl(Me.txtBoxValor1), CDbl(Me.txtBoxValor2))) Then
            GoTo optDivisão_Click_E
         End If
      Else
         MsgBox "Digite apenas números", 48, "ATENÇÃO"
      End If
      
      txtBoxResultado = .ValorResultante
   End With
         
   Exit Sub

optDivisão_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub DivisaoInt_Click()
   On Error GoTo optDivisaoInteira_Click_E
   
   If Me.txtBoxValor1.Text = "" And Me.txtBoxValor2.Text = "" Then
      MsgBox "Digite um valor númerico para operar."
      Exit Sub
   End If
   
   With clsCalculadora
      If IsNumeric(txtBoxValor1.Text) And IsNumeric(txtBoxValor1.Text) Then
         If Not (.DivisaoInteira(CDbl(txtBoxValor1.Text), CDbl(txtBoxValor2.Text))) Then
            GoTo optDivisaoInteira_Click_E
         End If
      Else
         MsgBox "Digite apenas números", 48, "ATENÇÃO"
      End If
      
      txtBoxResultado.Text = .ValorResultante
   End With
   
   Exit Sub
   
optDivisaoInteira_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub Exponenciar_Click()
   On Error GoTo optExponenciacao_Click_E
   
   If Me.txtBoxValor1.Text = "" And Me.txtBoxValor2.Text = "" Then
      MsgBox "Digite um valor númerico para operar."
      Exit Sub
   End If

   With clsCalculadora
      If IsNumeric(txtBoxValor1.Text) And IsNumeric(txtBoxValor1.Text) Then
         If Not (.Exponenciacao(CDbl(txtBoxValor1.Text), CDbl(txtBoxValor2.Text))) Then
            GoTo optExponenciacao_Click_E
         End If
      Else
         MsgBox "Digite apenas números", 48, "ATENÇÃO"
      End If
      
      txtBoxResultado.Text = .ValorResultante
   End With
   
   Exit Sub
   
optExponenciacao_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub Fibonacci_Click()
   On Error GoTo optFibonacci_Click_E
   
   Dim anterior, atual, próximo As Integer
   anterior = 0
   atual = 1
   proximo = 1
   
   Me.txtBoxResultado.Text = ""
   
   For i = 1 To 17
      proximo = atual + anterior
      anterior = atual
      atual = proximo
      txtBoxResultado.Text = CStr(txtBoxResultado.Text & " " & proximo)
   Next i
   
   txtBoxResultado.Text = txtBoxResultado.Text + " ..."
      
   Exit Sub

optFibonacci_Click_E:
   MsgBox "Erro ao gerar sequência Fibonacci"
End Sub

Private Sub Form_Load()
   Set clsCalculadora = New clsOperacoes
   Me.Somar.Enabled = True
   Me.Subtrair.Enabled = True
   Me.Dividir.Enabled = True
   Me.Multiplicar.Enabled = True
   Me.DivisaoInt.Enabled = True
   Me.RestoMod.Enabled = True
   Me.Exponenciar.Enabled = True
   Me.Concatenacao.Enabled = True
   Me.areaCirc.Enabled = True
   Me.Bhaskara.Enabled = True
   Me.Fibonacci.Enabled = True
   Me.Matrizes.Enabled = True
   Me.calcBhask.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsCalculadora = Nothing
End Sub

Private Sub limpar_Click()
   txtBoxValor1.Text = ""
   txtBoxValor2.Text = ""
   txtBoxValorC.Text = ""
   lblValor1.Caption = "VALOR 1"
   lblValor2.Caption = "VALOR 1"
   lblBhask.Caption = "BHASKARA"
   Me.calcBhask.Enabled = False
   Me.CalcCirc.Enabled = False
End Sub

Private Sub Matrizes_Click()
   On Error GoTo optMatrizes_Click_E
   
   frmCalculoMatrizes.Show
      
   Exit Sub

optMatrizes_Click_E:
   MsgBox "Erro ao gerar operações com matrizes"
End Sub

Private Sub Multiplicar_Click()
   On Error GoTo optMultiplicacao_Click_E
   
   If Me.txtBoxValor1.Text = "" And Me.txtBoxValor2.Text = "" Then
      MsgBox "Digite um valor númerico para operar."
      Exit Sub
   End If

   With clsCalculadora
      If IsNumeric(txtBoxValor1.Text) And IsNumeric(txtBoxValor1.Text) Then
         If Not (.Multiplicacao(CDbl(txtBoxValor1.Text), CDbl(txtBoxValor2.Text))) Then
            GoTo optMultiplicacao_Click_E
         End If
      Else
         MsgBox "Digite apenas números", 48, "ATENÇÃO"
      End If
      
      txtBoxResultado.Text = .ValorResultante
   End With
      
   Exit Sub

optMultiplicacao_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub RestoMod_Click()
   On Error GoTo optRestoMod_Click_E
   
   If Me.txtBoxValor1.Text = "" And Me.txtBoxValor2.Text = "" Then
      MsgBox "Digite um valor númerico para operar."
      Exit Sub
   End If
   
   With clsCalculadora
      If IsNumeric(txtBoxValor1.Text) And IsNumeric(txtBoxValor1.Text) Then
         If Not (.DivisaoMod(CDbl(txtBoxValor1.Text), CDbl(txtBoxValor2.Text))) Then
            GoTo optRestoMod_Click_E
         End If
      Else
         MsgBox "Digite apenas números", 48, "ATENÇÃO"
      End If
      
      txtBoxResultado.Text = .ValorResultante
   End With
   
   Exit Sub

optRestoMod_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub Somar_Click()
   On Error GoTo optSoma_Click_E
      
   If Me.txtBoxValor1.Text = "" And Me.txtBoxValor2.Text = "" Then
      MsgBox "Digite um valor númerico para operar."
      Exit Sub
   End If
   
   With clsCalculadora
      If IsNumeric(txtBoxValor1.Text) And IsNumeric(txtBoxValor1.Text) Then
         If Not (.Somar(CDbl(txtBoxValor1.Text), CDbl(txtBoxValor2.Text))) Then
            GoTo optSoma_Click_E
         End If
      Else
         MsgBox "Digite apenas números", 48, "ATENÇÃO"
      End If
      
      txtBoxResultado.Text = .ValorResultante
   End With
        
   Exit Sub

optSoma_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub

Private Sub Subtrair_Click()
   On Error GoTo optSubtracao_Click_E
   
   If Me.txtBoxValor1.Text = "" And Me.txtBoxValor2.Text = "" Then
      MsgBox "Digite um valor númerico para operar."
      Exit Sub
   End If

   With clsCalculadora
      If IsNumeric(txtBoxValor1.Text) And IsNumeric(txtBoxValor1.Text) Then
         If Not (.Subtracao(CDbl(txtBoxValor1.Text), CDbl(txtBoxValor2.Text))) Then
            GoTo optSubtracao_Click_E
         End If
      Else
         MsgBox "Digite apenas números", 48, "ATENÇÃO"
      End If
      
      txtBoxResultado.Text = .ValorResultante
   End With
      
   Exit Sub

optSubtracao_Click_E:
   MsgBox "Erro ao gerar operação" & clsCalculadora.getErro, 48, "ERRO - ATENÇÃO"
End Sub
