VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApagarArq 
      Caption         =   "23 - Apagar Arq."
      Height          =   375
      Left            =   2280
      TabIndex        =   22
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdCopiarArq 
      Caption         =   "22 - Copiar Arq. - Temp"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdVerificaLab 
      Caption         =   "21 - Verifica Lab.Txt"
      Height          =   375
      Left            =   4440
      TabIndex        =   20
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdCompactar 
      Caption         =   "20 - Compactar Arq"
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdExtrairHTML 
      Caption         =   "19 - Extrair HTML Pub."
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdQtdArqHTML 
      Caption         =   "18 - Qtd. Arq. HTML Pub."
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalcHashSHA1 
      Caption         =   "17 - Calc. Hash SHA1"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalcHashMD5 
      Caption         =   "16 - Calc. Hash MD5"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdExcluirChave 
      Caption         =   "15 - Excluir Chave "
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdExcluirSemana 
      Caption         =   "14 - Excluir Valor da chave"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdValorSemana 
      Caption         =   "13 - Obter Valor da chave"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdSemana11 
      Caption         =   "12 - Cria valor na chave"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdCriarChave 
      Caption         =   "11 - Criar Chave"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdSis64Bits 
      Caption         =   "10 - Verifica 64 Bits"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdCamPastaTemp 
      Caption         =   "09 - Cam. Pasta Temp."
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCamImgs 
      Caption         =   "08 - Cam. 'Imagens'"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCamMeusDocs 
      Caption         =   "07 - Cam. 'Meus Docs'"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdVersaoSO 
      Caption         =   "06 - Versão SO"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdNomeSO 
      Caption         =   "05 - Nome SO"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdNomePC 
      Caption         =   "04 - Nome PC"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdUserAdm 
      Caption         =   "03 - Usuário Adm"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "02 - Usuário"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdDtHr 
      Caption         =   "01 - Data e Hora"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApagarArq_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Arquivo
   Set objArquivos = CreateObject("LydiansWin.LYD_Arquivo")
   
   'Informa para o objeto o arquivo para manipular
   objArquivos.Arquivo = "C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\Filhote.bmp"
   
   If objArquivos.Existe Then
      If objArquivos.Apagar Then
         'Se objArquivos.Apagar retornar true, então apagou
         MsgBox "Arquivo apagado."
      Else
         'Se não, houve erro e não apagou
         MsgBox "Arquivo não foi apagado. Houve algum erro."
      End If
   Else
      MsgBox "O arquivo não existe para utilizar esta função. Adicione o arquivo no camiinho utilizado na função para continuar.", , "Arquivos"
   End If
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdCalcHashMD5_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Arquivo
   Set objArquivos = CreateObject("LydiansWin.LYD_Arquivo")
   
   'Informa para o objeto o arquivo para manipular
   objArquivos.Arquivo = "C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\HTML_Publico.zip"
   
   'Devolve o cód. Hash MD5 para o usuário na MsgBox
   MsgBox "Cód. Hash: " & objArquivos.Hash(MD5), , "Hash - MD5"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdCalcHashSHA1_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Arquivo
   Set objArquivos = CreateObject("LydiansWin.LYD_Arquivo")
   
   'Informa para o objeto o arquivo para manipular
   objArquivos.Arquivo = "C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\HTML_Publico.zip"
   
   'Devolve o cód. Hash SHA1 para o usuário na MsgBox
   MsgBox "Cód. SHA1: " & objArquivos.Hash(SHA1), , "Hash -  SHA1"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdCamImgs_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
   
   'Devolve a informação para o usuário na MsgBox
   MsgBox "O caminho é " & objWindows.Diretorios.UserMyPictures, , "Imagens"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdCamMeusDocs_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
   
   'Devolve a informação para o usuário na MsgBox
   MsgBox "O caminho é " & objWindows.Diretorios.UserMyDocuments, , "Documentos"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdCamPastaTemp_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
   
   'Devolve a informaçao para o usuario na MsgBox
   MsgBox "O caminho é " & objWindows.Diretorios.WinTemp, , "WinTemp"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdCompactar_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Arquivo
   Set objArquivos = CreateObject("LydiansWin.LYD_Arquivo")
   
   'Informa o arquivo a ser manipulado para o objeto
   objArquivos.Arquivo = "C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\Filhote.bmp"
   
   If objArquivos.Existe Then
      If objArquivos.Compactar_Para("C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11", "Filhote.bmp") Then
         'Se true, compactou o arquivo
         MsgBox "Arquivo compactado."
      Else
         'Se false, houve erro e não compactou o arquivo
         MsgBox "Arquivo não foi compactado. Houve algum erro."
      End If
   Else
      MsgBox "O arquivo não existe para utilizar esta função. Adicione o arquivo no camiinho utilizado para continuar.", , "Arquivos"
   End If
   
   'Limpa o objeto
   Set objArquivos = Nothing
End Sub

Private Sub cmdCopiarArq_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Arquivo
   Set objArquivos = CreateObject("LydiansWin.LYD_Arquivo")
   
   'Informa o arquivo a ser manipulado para o objeto
   objArquivos.Arquivo = "C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\Filhote.bmp"
   
   If objArquivos.Existe Then
      If objArquivos.Copiar_Para("C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\ArqCopia", "Filhote.bmp") Then
         'Se true, copiou o arquivo
         MsgBox "Arquivo copiado para o diretório Temp."
      Else
         'Se false, houve erro e não copiou o arquivo
         MsgBox "Arquivo não foi copiado. Houve erro."
      End If
   Else
      MsgBox "O arquivo não existe para utilizar esta função. Adicione o arquivo no camiinho utilizado para continuar.", , "Arquivos"
   End If
   
   'Limpa o objeto
   Set objArquivos = Nothing
End Sub

Private Sub cmdCriarChave_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Registro
   Set objRegistro = CreateObject("LydiansWin.LYD_Registro")
   
   'Variáveis para auxiliar no algoritmo de busca
   Dim blnValorEncontrado As Boolean
   Dim blnValorCriado As Boolean
   
   blnValorEncontrado = False
   
   'Percorre as chaves a fim de verificar se existe LydiansTeste
   For Each elemento In objRegistro.EnumKeys(HKEY_LOCAL_MACHINE, "Software")
      If elemento = "LydiansTeste" Then
         blnValorEncontrado = True
         Exit For
      End If
   Next
   
   'Se encontrou um elemento LydiansTeste acima, não faz nada, pois já existe
   'Se não encontrar, cria e atribui true a blnValorCriado
   If Not (blnValorEncontrado) Then
      objRegistro.WriteRegKey HKEY_LOCAL_MACHINE, "Software\LydiansTeste"
      blnValorCriado = True
   End If
   
   blnValorEncontrado = False
   
   'Percore a chave criada a fim de verificar se já existe Laboratorio, se não, cria, se sim, alert ao usuário
   For Each elemento In objRegistro.EnumKeys(HKEY_LOCAL_MACHINE, "Software\LydiansTeste")
      If elemento = "Laboratorio" Then
         blnValorEncontrado = True
      End If
   Next

   If (blnValorEncontrado) Then
      'Se ambas as chaves já existem, retorna a informação para o usuário
      MsgBox "Chave LydiansTeste\Laboratorio já existe.", , "Chaves - Registro"
   Else
      'Se ambas as chaves foram criadas, retorna a informação para o usuário
      objRegistro.WriteRegKey HKEY_LOCAL_MACHINE, "Software\LydiansTeste\Laboratorio"
      MsgBox "Chave LydiansTeste\Laboratorio criadas com sucesso.", , "Chaves - Registro"
   End If
   
   'Limpa o objeto
   Set objRegistro = Nothing
End Sub

Private Sub cmdDtHr_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
   
   'Variáveis auxiliares no algoritmo
   Dim strDia As String
   Dim strMes As String
   Dim strAno As String
   Dim strHora As String
   Dim strMinuto As String
   Dim strSegundo As String

   strDia = objWindows.Computador.Tempo.Dia
   strMes = objWindows.Computador.Tempo.Mes
   strAno = objWindows.Computador.Tempo.Ano
   strHora = objWindows.Computador.Tempo.Hora
   strMinuto = objWindows.Computador.Tempo.Minuto
   strSegundo = objWindows.Computador.Tempo.Segundo
   
   'Retorna a informação para o usuário
   MsgBox "Dia: " & strDia & "/" & strMes & "/" & strAno & " | Hora: " & strHora & ":" & strMinuto & ":" & strSegundo, , "Data - Hr"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdExcluirChave_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Registro
   Set objRegistro = CreateObject("LydiansWin.LYD_Registro")
   
   'Variáveis para auxiliar no algoritmo de busca
   Dim blnObjExcluido As Boolean
   
   'Percorre a chave para verificar se já existe
   'Se sim, exlcui e blnObjExcluido = True, segue o fluxo
   For Each elemento In objRegistro.EnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\")
      If elemento = "LydiansTeste" Then
         objRegistro.DeleteKey HKEY_LOCAL_MACHINE, "Software\LydiansTeste"
         blnObjExcluido = True
         Exit For
      End If
   Next
   
   'Verifica a variável blnObjExcluido e retorna o resultado para o usuário.
   If (blnObjExcluido) Then
      MsgBox "Chave excluída: " & elemento, , "Chaves - Valores"
   Else
      MsgBox "Chave não existe.", , "Chaves - Valores"
   End If
   
   'Limpa o objeto
   Set objRegistro = Nothing
End Sub

Private Sub cmdExcluirSemana_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Registro
   Set objRegistro = CreateObject("LydiansWin.LYD_Registro")
   
   Dim blnObjExcluido As Boolean
   
   'Percorre os valores da chave passada e se achar, excluir o valor que está procurando
   For Each elemento In objRegistro.EnumValues(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\LydiansTeste\Laboratorio")
      If elemento = "Semana" Then
         objRegistro.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\LydiansTeste\Laboratorio", "Semana"
         blnObjExcluido = True
         Exit For
      End If
   Next
   
   'Verifica a variável blnObjExcluido e retorna o resultado para o usuário.
   If (blnObjExcluido) Then
      MsgBox "Valor excluído: " & elemento, , "Chaves - Valores"
   Else
      MsgBox "Valor não existe.", , "Chaves - Valores"
   End If
   'Limpa o objeto
   Set objRegistro = Nothing
End Sub

Private Sub cmdExtrairHTML_Click()
'Averiguar com o conte
'   Set objArquivos = CreateObject("LydiansWin.LYD_Zip")
'
'   objArquivos.Extrair "C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\HTML_Publico.zip", "C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\ArqExtraido"
'
'   MsgBox "Arquivo extraido em C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\ArqExtraido"
'
'   Set objArquivos = Nothing
End Sub

Private Sub cmdNomePC_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
   
   Dim strNomePC As String
   
   strNomePC = objWindows.Computador.Nome
   
   'Retorna a informação para o usuário na MsgBox
   MsgBox "O nome do computador é " & strNomePC & "!", , "Computador"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdNomeSO_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
   
   Dim strNomeSO As String
   
   strNomeSO = objWindows.Nome
   
   'Retorna a informação para o usuário na MsgBox
   MsgBox "O nome do SO é " & strNomeSO & "!", , "Computador"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdQtdArqHTML_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Zip
   Set objZip = CreateObject("LydiansWin.LYD_Zip")
   
   objZip.Arquivo = "C:\Program Files (x86)\DevStudio\VB\ProjetosBrayan\Atividades 11\HTML_Publico.zip"
   
   If objZip.Existe Then
      'Retorna a informação para o usuário na MsgBox
      MsgBox "A quantidade de arquivos é: " & objZip.NroArquivos
   Else
      MsgBox "O arquivo não existe para utilizar esta função. Adicione o arquivo no camiinho utilizado para continuar.", , "Arquivos"
   End If
   
   'Limpa o objeto
   Set objZip = Nothing
End Sub

Private Sub cmdSemana11_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Registro
   Set objRegistro = CreateObject("LydiansWin.LYD_Registro")
   
   Dim blnValorEncontrado As Boolean
   
'   If (CBool(objRegistro.EnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\LydiansTeste\Laboratorio"))) Then
      'Percorre as chaves para encontrar o valor especificado no If
      For Each elemento In objRegistro.EnumValues(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\LydiansTeste\Laboratorio")
         If elemento = "Semana" Then
            blnValorEncontrado = True
            Exit For
         End If
      Next
      'Se encontrar ou não, informa ao usuário
      If (blnValorEncontrado) Then
         MsgBox "Valor Semana - 11 já existe.", , "Chaves - Escrita"
      Else
         objRegistro.WriteRegValue HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\LydiansTeste\Laboratorio", "Semana", 11
         MsgBox "Valor Semana - 11 escrito com sucesso.", , "Chaves - Escrita"
      End If
'   Else
'      MsgBox "Não existe a chave especificada para realizar o registro. Crie a chave.", , "Chaves - Registro"
'   End If
'
   'Limpa o objeto
   Set objRegistro = Nothing
End Sub

Private Sub cmdSis64Bits_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
      
   'Informar ao usuário se o sistema é ou não 64Bits
   If (objWindows.Is64Bits) Then
      MsgBox "O computador é 64 Bits!", , "Computador"
   Else
      MsgBox "O computador não é 64 Bits!", , "Computador"
   End If
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdUser_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
   
   'Informar ao usuário qual o ID registrado
   MsgBox "O usuário é " & objWindows.Computador.UsuarioID & "!", , "User"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdUserAdm_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
   
   'Informa ao usuário se o user atual é ADM
   If (objWindows.Computador.UsuarioAdmin) Then
      MsgBox "O usuário é administrador!", , "User"
   Else
      MsgBox "O usuário não é administrador!", , "User"
   End If
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub

Private Sub cmdValorSemana_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Registro
   Set objRegistro = CreateObject("LydiansWin.LYD_Registro")
   
   Dim blnObjEncontrado As Boolean
   
   'Percore os valores das chavas a fim de encontra o valor especificado no If
   For Each elemento In objRegistro.EnumValues(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\LydiansTeste\Laboratorio")
      If elemento = "Semana" Then
         blnObjEncontrado = True
         Exit For
      End If
   Next
   
   'Se encontrar, retorna a informação para o usuário
   If (blnObjEncontrado) Then
      MsgBox "Valor encontrado: " & elemento, , "Chaves - Valores"
   Else
      MsgBox "Valor não existe.", , "Chaves - Valores"
   End If
   
   'Limpa o objeto
   Set objRegistro = Nothing
End Sub

Private Sub cmdVerificaLab_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Arquivo
   Set objArquivos = CreateObject("LydiansWin.LYD_Arquivo")
   
   'Informa o arquivo a ser manipulado para o objeto
   objArquivos.Arquivo = "C:\Temp\laboratorio.txt"
   
   'Se o arquivo especificado existir no diretório Temp, retorna a informação para o usuário
   If objArquivos.Existe Then
      MsgBox "Existe um arquivo laboratorio.txt no diretório C:\Temp"
   Else
      MsgBox "Não existe um arquivo laboratorio em C:\Temp"
   End If
    
   'Limpa o objeto
   Set objArquivos = Nothing
End Sub

Private Sub cmdVersaoSO_Click()
   'Cria um objeto do tipo LydiansWin.LYD_Windows
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
  
  'Retorna a versão do Windows para o usuário
   MsgBox "A versão do SO é " & objWindows.Versao & "!", , "Computador"
   
   'Limpa o objeto
   Set objWindows = Nothing
End Sub
