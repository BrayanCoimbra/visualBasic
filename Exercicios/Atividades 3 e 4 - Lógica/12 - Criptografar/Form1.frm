VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Crips"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   2820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCriptografia 
      Caption         =   "Criptografia"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.TextBox txtMostraDescriptografado 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtMostraCriptografado 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdDescriptografar 
         Caption         =   "Descriptografar"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdCriptografar 
         Caption         =   "Criptografar"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtEntrada 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colGuardaNumeroMaisSete() As Integer
Dim colGuardaNumeroMaisSeteModuloDez() As Integer
Dim colGuardaNumeroMaisSeteModuloDezTrocandoPosicao() As Integer
Dim colGuardaNumeroReverteModulo() As Integer
Dim colGuardaNumeroMenosSete() As Integer

Private Sub cmdCriptografar_Click()
Dim i, j As Integer
        
    nmrParaCriptografar = CInt(txtEntrada.Text)
    ReDim colGuardaNumeroMaisSete(1 To Len(txtEntrada.Text))
    ReDim colGuardaNumeroMaisSeteModuloDez(1 To Len(txtEntrada.Text))
    ReDim colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(1 To Len(txtEntrada.Text))
    
    For i = 1 To Len(txtEntrada.Text)
        guardaNumero = CInt(Mid(txtEntrada.Text, i, 1))
        colGuardaNumeroMaisSete(i) = CInt(guardaNumero + 7)
        colGuardaNumeroMaisSeteModuloDez(i) = CInt(colGuardaNumeroMaisSete(i) Mod 10)
    Next i
    
    'Limitado a 4 posições - averiguar para tornar dinâmico
    colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(1) = colGuardaNumeroMaisSeteModuloDez(3)
    colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(2) = colGuardaNumeroMaisSeteModuloDez(4)
    colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(3) = colGuardaNumeroMaisSeteModuloDez(1)
    colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(4) = colGuardaNumeroMaisSeteModuloDez(2)
    
    For j = 1 To Len(txtEntrada.Text)
        txtMostraCriptografado.Text = txtMostraCriptografado.Text & " " & colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(j)
    Next j
    
End Sub

Private Sub cmdDescriptografar_Click()
Dim i, j, k As Integer

    txtMostraDescriptografado = ""

    nmrParaCriptografar = CInt(txtEntrada.Text)
    ReDim Preserve colGuardaNumeroMaisSete(1 To Len(txtEntrada.Text))
    ReDim Preserve colGuardaNumeroMaisSeteModuloDez(1 To Len(txtEntrada.Text))
    ReDim Preserve colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(1 To Len(txtEntrada.Text))
    ReDim colGuardaNumeroReverteModulo(1 To Len(txtEntrada.Text))
    ReDim colGuardaNumeroMenosSete(1 To Len(txtEntrada.Text))
    
    'Limitado a 4 posições - averiguar para tornar dinâmico
    colGuardaNumeroReverteModulo(1) = colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(3)
    colGuardaNumeroReverteModulo(2) = colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(4)
    colGuardaNumeroReverteModulo(3) = colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(1)
    colGuardaNumeroReverteModulo(4) = colGuardaNumeroMaisSeteModuloDezTrocandoPosicao(2)
    
    For i = 1 To Len(txtEntrada.Text)
    ' Guarda o número original
    guardaNmrOriginal = colGuardaNumeroReverteModulo(i)

    ' Verifica se o número é menor que 7 para poder aplicar o módulo diretamente
        If guardaNmrOriginal < 7 Then
            colGuardaNumeroReverteModulo(i) = guardaNmrOriginal + 3 ' Reverte o módulo
        Else
            ' Se o número original for 7, 8 ou 9, precisamos ajustar para a base correta
            colGuardaNumeroReverteModulo(i) = guardaNmrOriginal - 7 ' Remove o incremento de 7
        End If
    Next i
        
    For j = 1 To Len(txtEntrada.Text)
        txtMostraDescriptografado.Text = txtMostraDescriptografado.Text & " " & colGuardaNumeroReverteModulo(j)
    Next j
    
End Sub
