VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalculadoraDeDias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private i As Integer
Private dataAtual As Date
Private MesNasci As Boolean
Public DiasVividos As Integer
Private mesIni, mesFim As Integer
Private diaNasc, mesNasc, anoNasc As Integer
Private diaAtual, mesAtual, anoAtual As Integer

Public Function Inicializar(pDiaNasc As Integer, pMesNasc As Integer, pAnoNasc As Integer) As Boolean
   On Error GoTo Inicializar_E
   
   Inicializar = False
   
   MesNasci = False
   DiasVividos = 0
   i = 0
   
   diaNasc = pDiaNasc
   mesNasc = pMesNasc
   anoNasc = pAnoNasc
   
   dataAtual = Date
   diaAtual = Day(dataAtual)
   mesAtual = Month(dataAtual)
   anoAtual = Year(dataAtual)
   
   If ContabilizarAnos Then
      Inicializar = True
      Exit Function
      
   Else
      Inicializar = False
      Exit Function
      
   End If
   
Inicializar_E:
   Inicializar = False
End Function

Private Function ContabilizarAnos() As Boolean
   On Error GoTo ContabilizarAnos_E
   
   ContabilizarAnos = False

   For i = anoNasc To anoAtual
      VerificarCondicoes
      
      For j = mesIni To mesFim
         If j = mesNasc And MesNasci = False Then
            ConTabilizarMesNasc
         
         ElseIf j = mesFim And i = anoAtual Then
            DiasVividos = DiasVividos + diaAtual
         
         Else
            Select Case j
            Case Janeiro, Marco, Maio, Julho, Agosto, Outubro, Dezembro
               DiasVividos = DiasVividos + 31
               
            Case Abril, Junho, Setembro, Novembro
               DiasVividos = DiasVividos + 30
               
            Case Fevereiro
               If AnoBissexto(i) Then
                  DiasVividos = DiasVividos + 29
               Else
                  DiasVividos = DiasVividos + 28
               End If
            End Select
            
         End If
      Next j
   
   Next i
      
   ContabilizarAnos = True
   
   Exit Function
   
ContabilizarAnos_E:
   ContabilizarAnos = False
End Function

Private Sub ConTabilizarMesNasc()
   Select Case mesNasc
   Case Janeiro, Marco, Maio, Julho, Agosto, Outubro, Dezembro
      DiasVividos = DiasVividos + (31 - diaNasc)

   Case Abril, Junho, Setembro, Novembro
      DiasVividos = DiasVividos + (30 - diaNasc)

   Case Fevereiro
      If AnoBissexto(anoNasc) Then
         DiasVividos = DiasVividos + (29 - diaNasc)
      Else
         DiasVividos = DiasVividos + (28 - diaNasc)
      End If
   End Select
   
   MesNasci = True
End Sub

Private Function AnoBissexto(ano As Integer) As Boolean
   If ano Mod 400 = 0 Or ano Mod 4 = 0 Then
      AnoBissexto = True
   Else
      AnoBissexto = False
   End If
End Function

Private Sub VerificarCondicoes()
   If i = anoNasc Then
      mesIni = mesNasc
      mesFim = 12
   ElseIf i = anoAtual Then
      mesIni = 1
      mesFim = mesAtual
   Else
      mesIni = 1
   End If
End Sub
