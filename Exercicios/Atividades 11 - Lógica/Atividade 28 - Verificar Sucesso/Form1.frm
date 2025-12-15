VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Festou"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Verificar Sucesso"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "..."
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   200
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   On Error GoTo Command1_Click_E
   
   Dim NmrConvidados As Integer
   Dim idades() As Integer
   Dim resultado As String
   Dim i As Integer
   
   NmrConvidados = InputBox("Número de convidados:")
   ReDim idades(1 To NmrConvidados)
   
   For i = 1 To NmrConvidados
       idades(i) = InputBox("Idade do convidado " & (i) & ":")
   Next i
   
   resultado = VerificarSucesso(NmrConvidados, idades)
   
   If resultado Then
      MsgBox "A festa será um sucesso!", vbInformation
      Me.Label1.Caption = "Irrá!"
   Else
      MsgBox "A festa será um fracasso!", vbExclamation
      Me.Label1.Caption = "Ops!"
   End If

   Exit Sub
   
Command1_Click_E:
   MsgBox "Houve um erro ao verificar o sucesso da festa.", vbCritical, "Erro"
End Sub

Private Function VerificarSucesso(NmrConvidados As Integer, idades() As Integer) As Boolean
   Dim i As Integer, j As Integer
   
   Dim colHomens As Collection
   Dim colMulheres As Collection
   
   Set colHomens = New Collection
   Set colMulheres = New Collection
   
   Dim ParFormado As Boolean
      
   For i = 1 To NmrConvidados
      If idades(i) Mod 2 = 0 Then
         colMulheres.Add idades(i)
      Else
         colHomens.Add idades(i)
      End If
   Next i
   
   If colHomens.Count > colMulheres.Count Then
      VerificarSucesso = False
   Else
      For i = 1 To colMulheres.Count
         ParFormado = False
         For j = i To colHomens.Count
            If colHomens.Item(j) > colMulheres.Item(i) Then
               ParFormado = True
               Exit For
            Else
               ParFormado = False
               Exit For
            End If
         Next j
      
         If Not ParFormado Then
            VerificarSucesso = False
            Exit Function
         End If
      Next i
      VerificarSucesso = True
   End If
End Function

