VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Salary"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFunc 
      Caption         =   "Funcionário"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.Label lblSalExt 
         Caption         =   "."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblSalario 
         Caption         =   "."
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblHrsExt 
         Caption         =   "."
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblHrs 
         Caption         =   "."
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton CalcularSalario 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   90
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'08- A jornada de trabalho semanal de um funcionário é de 40 horas. O funcionário que trabalhar mais de 40 horas
'receberá hora extra, cujo cálculo é o valor da hora regular com um acréscimo de 50%. Escreva um algoritmo
'que leia o número de horas trabalhadas em um mês, o salário por hora e escreva o salário total do funcionário,
'que deverá ser crescido das horas extras, caso tenham sido trabalhadas
'(considere que o mês possua 4 semanas exatas).

Private Sub CalcularSalario_Click()
   
   Dim horasTrabalhadas As Integer
   Dim salarioHora, salarioTotal, salarioHoraExtra As Double
   
   Me.lblHrs.Caption = ""
   Me.lblHrsExt.Caption = ""
   Me.lblSalario.Caption = ""
   Me.lblSalExt.Caption = ""
   
   ' Solicita ao usuário que insira o número de horas trabalhadas no mês e o salário por hora.
   horasTrabalhadas = InputBox("Digite o número de horas trabalhadas no mês:")
   salarioHora = InputBox("Digite o salário por hora:")
   
   ' Verifica se o número de horas trabalhadas é maior que 160 (40 horas por semana * 4 semanas).
   If horasTrabalhadas > 160 Then
      ' Calcula o salário total considerando as horas extras com um acréscimo de 50%.
      salarioHoraExtra = salarioHora * 1.5
      salarioTotal = (40 * salarioHora) + ((horasTrabalhadas - 160) * salarioHoraExtra)
      
      Me.lblHrs.Caption = "Horas Normais: " & horasTrabalhadas & "h. "
      Me.lblHrsExt.Caption = "Horas Extras: " & (horasTrabalhadas - 160) & "h. "
      Me.lblSalario.Caption = "Salário Total: R$ " & salarioTotal
      Me.lblSalExt.Caption = "Salário Extra: R$ " & salarioHoraExtra
   Else
      ' Se não houver horas extras, calcula o salário total apenas com as horas normais.
      salarioTotal = horasTrabalhadas * salarioHora
      Me.lblHrs.Caption = "Horas Normais: " & horasTrabalhadas & "h. "
      Me.lblHrsExt.Caption = "Horas Extras: 0h"
      Me.lblSalario.Caption = "Salário Total: R$ " & salarioTotal
      Me.lblSalExt.Caption = "Salário Extra: R$0,00"
   End If

End Sub

