VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "aabcc"
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim letra As String
   
   Dim contadorA As Integer
   Dim contadorB As Integer
   Dim contadorC As Integer
   Dim contadorD As Integer
   Dim contadorE As Integer
   Dim contadorF As Integer
   Dim contadorG As Integer
   Dim contadorH As Integer
   Dim contadorI As Integer
   Dim contadorJ As Integer
   Dim contadorK As Integer
   Dim contadorL As Integer
   Dim contadorM As Integer
   Dim contadorN As Integer
   Dim contadorO As Integer
   Dim contadorP As Integer
   Dim contadorQ As Integer
   Dim contadorR As Integer
   Dim contadorS As Integer
   Dim contadorT As Integer
   Dim contadorU As Integer
   Dim contadorV As Integer
   Dim contadorW As Integer
   Dim contadorX As Integer
   Dim contadorY As Integer
   Dim contadorZ As Integer
   
   Dim PercentA As Double
   Dim PercentB As Double
   Dim PercentC As Double
   Dim PercentD As Double
   Dim PercentE As Double
   Dim PercentF As Double
   Dim PercentG As Double
   Dim PercentH As Double
   Dim PercentI As Double
   Dim PercentJ As Double
   Dim PercentK As Double
   Dim PercentL As Double
   Dim PercentM As Double
   Dim PercentN As Double
   Dim PercentO As Double
   Dim PercentP As Double
   Dim PercentQ As Double
   Dim PercentR As Double
   Dim PercentS As Double
   Dim PercentT As Double
   Dim PercentU As Double
   Dim PercentV As Double
   Dim PercentW As Double
   Dim PercentX As Double
   Dim PercentY As Double
   Dim PercentZ As Double
   
Private Sub cmdCalcular_Click()
       
   Dim colLetras() As String
   ReDim colLetras(1 To Len(Text1.Text)) As String
         
   For i = 1 To Len(Text1.Text)
      letra = UCase(Mid(Text1.Text, i, 1))
      colLetras(i) = letra
      
      Select Case colLetras(i)
         Case "A"
            contadorA = contadorA + 1

         Case "B"
            contadorB = contadorB + 1

         Case "C"
            contadorC = contadorC + 1

         Case "D"
            contadorD = contadorD + 1

         Case "E"
            contadorE = contadorE + 1

         Case "F"
            contadorF = contadorF + 1

         Case "G"
            contadorG = contadorG + 1

         Case "H"
            contadorH = contadorH + 1

         Case "I"
            contadorI = contadorI + 1

         Case "J"
            contadorJ = contadorJ + 1

         Case "K"
            contadorK = contadorK + 1

         Case "L"
            contadorL = contadorL + 1

         Case "M"
            contadorM = contadorM + 1

         Case "N"
            contadorN = contadorN + 1

         Case "O"
            contadorO = contadorO + 1

         Case "P"
            contadorP = contadorP + 1

         Case "Q"
            contadorQ = contadorQ + 1

         Case "R"
            contadorR = contadorR + 1

         Case "S"
            contadorS = contadorS + 1

         Case "T"
            contadorT = contadorT + 1

         Case "U"
            contadorU = contadorU + 1

         Case "V"
            contadorV = contadorV + 1

         Case "W"
            contadorW = contadorW + 1

         Case "X"
            contadorX = contadorX + 1

         Case "Y"
            contadorY = contadorY + 1

         Case "Z"
            contadorZ = contadorZ + 1
      End Select
   Next i
   
   PercentA = Calculo(contadorA, Len(Text1.Text))
   PercentB = Calculo(contadorB, Len(Text1.Text))
   PercentC = Calculo(contadorC, Len(Text1.Text))
   PercentD = Calculo(contadorD, Len(Text1.Text))
   PercentE = Calculo(contadorE, Len(Text1.Text))
   PercentF = Calculo(contadorF, Len(Text1.Text))
   PercentG = Calculo(contadorG, Len(Text1.Text))
   PercentH = Calculo(contadorH, Len(Text1.Text))
   PercentI = Calculo(contadorI, Len(Text1.Text))
   PercentJ = Calculo(contadorJ, Len(Text1.Text))
   PercentK = Calculo(contadorK, Len(Text1.Text))
   PercentL = Calculo(contadorL, Len(Text1.Text))
   PercentM = Calculo(contadorM, Len(Text1.Text))
   PercentN = Calculo(contadorN, Len(Text1.Text))
   PercentO = Calculo(contadorO, Len(Text1.Text))
   PercentP = Calculo(contadorP, Len(Text1.Text))
   PercentQ = Calculo(contadorQ, Len(Text1.Text))
   PercentR = Calculo(contadorR, Len(Text1.Text))
   PercentS = Calculo(contadorS, Len(Text1.Text))
   PercentT = Calculo(contadorT, Len(Text1.Text))
   PercentU = Calculo(contadorU, Len(Text1.Text))
   PercentV = Calculo(contadorV, Len(Text1.Text))
   PercentW = Calculo(contadorW, Len(Text1.Text))
   PercentX = Calculo(contadorX, Len(Text1.Text))
   PercentY = Calculo(contadorY, Len(Text1.Text))
   PercentZ = Calculo(contadorZ, Len(Text1.Text))

   
   If PercentA <> 0 Then
      Resultado = Resultado + "A " + CStr(PercentA) + "%" & vbNewLine
   End If
   
   If PercentB <> 0 Then
      Resultado = Resultado + "B " + CStr(PercentB) + "%" & vbNewLine
   End If
   
   If PercentC <> 0 Then
      Resultado = Resultado + "C " + CStr(PercentC) + "%" & vbNewLine
   End If
   
   If PercentD <> 0 Then
      Resultado = Resultado + "D " + CStr(PercentD) + "%" & vbNewLine
   End If
   
   If PercentE <> 0 Then
      Resultado = Resultado + "E " + CStr(PercentE) + "%" & vbNewLine
   End If
   
   If PercentF <> 0 Then
      Resultado = Resultado + "F " + CStr(PercentF) + "%" & vbNewLine
   End If
   
   If PercentG <> 0 Then
      Resultado = Resultado + "G " + CStr(PercentG) + "%" & vbNewLine
   End If
   
   If PercentH <> 0 Then
      Resultado = Resultado + "H " + CStr(PercentH) + "%" & vbNewLine
   End If
   
   If PercentI <> 0 Then
      Resultado = Resultado + "I " + CStr(PercentI) + "%" & vbNewLine
   End If
   If PercentJ <> 0 Then
      Resultado = Resultado + "J " + CStr(PercentJ) + "%" & vbNewLine
   End If
   
   If PercentK <> 0 Then
      Resultado = Resultado + "K " + CStr(PercentK) + "%" & vbNewLine
   End If
   
   If PercentL <> 0 Then
      Resultado = Resultado + "L " + CStr(PercentL) + "%" & vbNewLine
   End If
   
   If PercentM <> 0 Then
      Resultado = Resultado + "M " + CStr(PercentM) + "%" & vbNewLine
   End If
   
   If PercentN <> 0 Then
      Resultado = Resultado + "N " + CStr(PercentN) + "%" & vbNewLine
   End If
   
   If PercentO <> 0 Then
      Resultado = Resultado + "O " + CStr(PercentO) + "%" & vbNewLine
   End If
   
   If PercentP <> 0 Then
      Resultado = Resultado + "P " + CStr(PercentP) + "%" & vbNewLine
   End If
   
   If PercentQ <> 0 Then
      Resultado = Resultado + "Q " + CStr(PercentR) + "%" & vbNewLine
   End If
   
   If PercentR <> 0 Then
      Resultado = Resultado + "R " + CStr(PercentR) + "%" & vbNewLine
   End If
   
   If PercentS <> 0 Then
      Resultado = Resultado + "S " + CStr(PercentS) + "%" & vbNewLine
   End If
   
   If PercentT <> 0 Then
      Resultado = Resultado + "T " + CStr(PercentT) + "%" & vbNewLine
   End If
   
   If PercentU <> 0 Then
      Resultado = Resultado + "U " + CStr(PercentU) + "%" & vbNewLine
   End If
   
   If PercentV <> 0 Then
      Resultado = Resultado + "V " + CStr(PercentV) + "%" & vbNewLine
   End If
   
   If PercentW <> 0 Then
      Resultado = Resultado + "W " + CStr(PercentW) + "%" & vbNewLine
   End If
   
   If PercentX <> 0 Then
      Resultado = Resultado + "X " + CStr(PercentX) + "%" & vbNewLine
   End If
   
   If PercentY <> 0 Then
      Resultado = Resultado + "Y " + CStr(PercentY) + "%" & vbNewLine
   End If
   
   If PercentZ <> 0 Then
      Resultado = Resultado + "Z " + CStr(PercentZ) + "%" & vbNewLine
   End If
   
   MsgBox Resultado, vbInformation, "Resultado"

' Limpa as variáveis para reutilização
   contadorA = 0
   contadorB = 0
   contadorC = 0
   contadorD = 0
   contadorE = 0
   contadorF = 0
   contadorG = 0
   contadorH = 0
   contadorI = 0
   contadorJ = 0
   contadorK = 0
   contadorL = 0
   contadorM = 0
   contadorN = 0
   contadorO = 0
   contadorP = 0
   contadorQ = 0
   contadorR = 0
   contadorS = 0
   contadorT = 0
   contadorU = 0
   contadorV = 0
   contadorW = 0
   contadorX = 0
   contadorY = 0
   contadorZ = 0
   
   PercentA = 0
   PercentB = 0
   PercentC = 0
   PercentD = 0
   PercentE = 0
   PercentF = 0
   PercentG = 0
   PercentH = 0
   PercentI = 0
   PercentJ = 0
   PercentK = 0
   PercentL = 0
   PercentM = 0
   PercentN = 0
   PercentO = 0
   PercentP = 0
   PercentQ = 0
   PercentR = 0
   PercentS = 0
   PercentT = 0
   PercentU = 0
   PercentV = 0
   PercentW = 0
   PercentX = 0
   PercentY = 0
   PercentZ = 0
End Sub

Private Function Calculo(vlr As Integer, vlrTotal As Integer) As Double
   Calculo = (vlr * 100) / (vlrTotal)
End Function
