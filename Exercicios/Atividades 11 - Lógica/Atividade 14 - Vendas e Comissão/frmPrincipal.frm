VERSION 5.00
Begin VB.Form frmVendas 
   Caption         =   "Vendas"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   1920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   1920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraVendas 
      Caption         =   "Opções"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton cmdCadastrarVenda 
         Caption         =   "Registrar"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdConsultarSalario 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SalarioVendedor As Double = 1000
Private Const PercentualDeVenda As Double = 0.005
Private Const PercentualDaVendaTotal As Double = 0.05
Private colTotalVendaComissao As Collection
Private colValorVenda As Collection
Private dblValorTotalVendido As Double

Private Sub cmdCadastrarVenda_Click()
   On Error GoTo cmdCadastrarVenda_Click_E
   
   Dim dblValorVenda As Double
   Dim dblValorCarro As Double
   Dim intQtdvendida As Integer
   
   dblValorCarro = InputBox("Digite o valor do carro vendido.", "Valor Carro")
   intQtdvendida = InputBox("Digite a quantidade vendida.", "Quantidade vendida")
   
   If IsNumeric(dblValorCarro) And IsNumeric(intQtdvendida) Then
      dblValorVenda = dblValorCarro * intQtdvendida
      
      colValorVenda.Add dblValorVenda
      
      dblValorTotalVendido = dblValorTotalVendido + dblValorVenda
      
      MsgBox "Venda cadastradas com sucesso. Parabéns!", vbInformation, "Cadastro de Vendas"
      
   Else
      MsgBox "O valor da venda e da quantidade vendida devem ser preenchidos corretamente.", vbExclamation, "Registro Cancelado"
      
   End If
   
   Exit Sub

cmdCadastrarVenda_Click_E:
   MsgBox "Houve um erro ao cadastrar sua venda. Verifique os valores de entrada.", vbCritical, "Erro"
   
End Sub

Private Sub cmdConsultarSalario_Click()
   If Not dblValorTotalVendido = 0 Then
      ' Calcula o percentual de comissão parcial para cada venda
      For i = 1 To colValorVenda.Count
         colTotalVendaComissao.Add CalculoPercentComissaoParcial(colValorVenda.Item(i))
      Next i
      
      ' Calcula o percentual de comissão total
      ValorTotalVendido = CalculoPercentComissaoTotal(dblValorTotalVendido)
      
      MsgBox "Salário Fixo: R$ " & SalarioVendedor & vbNewLine & _
            "Valor Total Vendido: R$ " & dblValorTotalVendido & vbNewLine & _
            "Valor Total de Comissão: R$" & ValorTotalVendido & vbNewLine & _
            "Salário Total: R$ " & SalarioVendedor + ValorTotalVendido, vbInformation
   Else
      MsgBox "Não há vendas registradas.", vbExclamation, "Vendas"
   End If
End Sub

Private Sub Form_Load()
   Set colValorVenda = New Collection
   Set colTotalVendaComissao = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set colValorVenda = Nothing
   Set colTotalVendaComissao = Nothing
End Sub

Private Function CalculoPercentComissaoTotal(dblNumero As Double) As Double
   CalculoPercentComissaoTotal = dblNumero * PercentualDaVendaTotal
End Function

Private Function CalculoPercentComissaoParcial(dblNumero As Double) As Double
   CalculoPercentComissaoParcial = dblNumero * PercentualDeVenda
End Function


