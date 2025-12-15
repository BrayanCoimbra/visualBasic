VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperacoesArq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Escrever um programa que trabalhe com manipulação de arquivos e diretórios.
'1- Permitir ao operador escrever um texto...
'...O texto deve ser salvo no local (diretório) informado pelo usuário com o seu respectivo nome de arquivo...
'...(inclusive a sua extensão .txt, .json, .xml)
'2 - Permitir a leitura do conteúdo do arquivo texto informado pelo usuário.
'3 - Excluir o arquivo de texto informado pelo usuário
'4 - Listar os arquivos de um determinado diretório incluindo suas pastas e subpastas
'5 - Realizar o backup completo de um determinado diretório...
'...Permitindo ao usuário informar a pasta de origem e de destino

Public Function Excluir(caminhoArq As String) As Boolean

   On Error GoTo Excluir_E

   Excluir = False
   
   'cria o objeto fso do tipo FileSystemObject
   Dim fso As FileSystemObject
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   'verifica se a pasta existe , alerta ao usuário que o caminho está errado
   If (fso.FolderExists(caminhoArq)) Then
       MsgBox "Pasta inexistente."
   End If

   Kill caminhoArq

   Excluir = True

   GoTo DestruirObjetos

Excluir_E:
DestruirObjetos:
End Function

Public Function BKP(caminhoOrigem As String, caminhoDestino As String) As Boolean

   On Error GoTo BKP_E

   BKP = False
   
   'cria o objeto fso do tipo FileSystemObject
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   ' Verifica se o caminho da pasta de origem é válido
   If Not (fso.FolderExists(caminhoOrigem)) Then
      MsgBox "Pasta de origem inválida!"
   End If
   
   ' Verifica se o caminho da pasta de origem é válido
   If Not (fso.FolderExists(caminhoDestino)) Then
      MsgBox "Pasta de destino inválida!"
   End If
    
   ' Obtém o primeiro nome de arquivo na pasta de origem
   nomeArquivo = Dir(caminhoOrigem & "\*.*")
    
   ' Loop para realizar o backup todos os arquivos do caminho origem para o de BKP
   Do While nomeArquivo <> ""
        ' Copia o arquivo de origem para o destino
        fso.CopyFile caminhoOrigem & "\" & nomeArquivo, caminhoDestino & "\" & nomeArquivo
        Debug.Print nomeArquivo
        nomeArquivo = Dir
    Loop
   
   BKP = True

   GoTo DestruirObjetos

BKP_E:
DestruirObjetos:
End Function

Public Function Ler(caminhoArq As String) As Boolean

   On Error GoTo Ler_E

   Ler = False
   
   Dim strLinha As String
   Dim NomeArqler As String
   Dim caminhoPasta As String
   Dim diretórioExistente As Boolean
   Dim fso As FileSystemObject

   'cria o objeto fso do tipo FileSystemObject
   Set fso = CreateObject("Scripting.FileSystemObject")
      
   'verifica se o diretório existe, se não, alerte ao usuário que o caminho está errado
   If (fso.FolderExists(caminhoArq)) Then
      diretórioExistente = True
   Else
      diretórioExistente = False
   End If
   
   Dim numArquivo As Integer
   numArquivo = FreeFile()
   
   Open caminhoArq For Input As #numArquivo
   
   Do While Not EOF(numArquivo)
      Line Input #numArquivo, strLinha
      strFile = strFile + strLinha + vbNewLine
   Loop
   
   Close #numArquivo

   Debug.Print strFile
   
   Ler = True

   GoTo DestruirObjetos

Ler_E:
DestruirObjetos:
End Function

Public Function Salvar(caminhoArq As String, nomeArq As String, registroNoArq As String) As Boolean

   On Error GoTo Salvar_E

   Salvar = False
   
   Dim fso As FileSystemObject

   'cria o objeto fso do tipo FileSystemObject
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   'verifica se o diretório existe , senão existir então cria
   If Not (fso.FolderExists(caminhoArq)) Then
       fso.CreateFolder (caminhoArq)
   End If
      
   Dim numArquivo As Integer
   numArquivo = FreeFile()
   
   Open caminhoArq & "\" & nomeArq For Append As #numArquivo
   Print #numArquivo, registroNoArq
   Close #numArquivo
    
   Salvar = True

   GoTo DestruirObjetos

Salvar_E:
DestruirObjetos:
End Function

Private Function ListarSubPastasEArquivos(caminhoRaiz As String) As Boolean

   On Error GoTo ListarSubPastasEArquivos_E
   
   ListarSubPastasEArquivos = False
   
   Dim fso As Object
   Dim pasta As Object
   Dim subPasta As Object
   Dim arquivo As Object
    
   ' Cria o objeto FileSystemObject
   Set fso = CreateObject("Scripting.FileSystemObject")
    
   ' Obtém a pasta raiz
   Set pasta = fso.GetFolder(caminhoRaiz)
       
   ' Itera sobre as subpastas dentro da pasta raiz
   For Each subPasta In pasta.SubFolders
      Debug.Print "Nome da Subpasta: " & subPasta.Path
       
      ' Itera sobre os arquivos dentro da subpasta
      For Each arquivo In subPasta.Files
         Debug.Print "Nome do Arquivo: " & arquivo.Path
      Next arquivo
        
      ' Chama recursivamente a mesma função para listar arquivos dentro desta subpasta
      ListarSubPastasEArquivos subPasta.Path
    Next subPasta
    
    ListarSubPastasEArquivos = True
    
    GoTo DestruirObjetos
    
ListarSubPastasEArquivos_E:
DestruirObjetos:
End Function

Public Function ListarPastasSubPastasArquivos(caminhoRaiz As String) As Boolean

   On Error GoTo ListarPastasSubPastasArquivos_E
   
   ListarPastasSubPastasArquivoss = False
   
   Dim fso As Object
   Dim pasta As Object
   Dim subPasta As Object
   Dim arquivo As Object
    
   ' Cria o objeto FileSystemObject
   Set fso = CreateObject("Scripting.FileSystemObject")
    
   ' Obtém a pasta raiz
   Set pasta = fso.GetFolder(caminhoRaiz)
    
   ' Itera sobre os arquivos na pasta raiz
   For Each arquivo In pasta.Files
      Debug.Print "Arquivo: " & arquivo.Path
   Next arquivo
    
   ListarSubPastasEArquivos caminhoRaiz
    
   ListarPastasSubPastasArquivos = True
    
   GoTo DestruirObjetos
    
ListarPastasSubPastasArquivos_E:
DestruirObjetos:
End Function
