VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5- Calcule as primeiras 10 sequências de números que satisfaçam a seguinte equação: a^2 = b^2 + c^2
'Obs.: a, b e c devem ser inteiros positivos

Private Sub Form_Load()
Dim a, b, c As Integer
  For a = 1 To 10 ' Calcula as primeiras 10 sequências
        For b = 1 To 10
            For c = 1 To 10
                If (a ^ 2 = b ^ 2 + c ^ 2) Then
                    MsgBox "Sequência encontrada: a = " & a & ", b = " & b & ", c = " & c
                    Exit For ' Sai do loop interno quando uma sequência é encontrada
                End If
            Next c
        Next b
    Next a
End Sub
