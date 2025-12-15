VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private colLivro As Collection
'Private strNome As String
'Private strEdicao As String
'Private strAutor As String
'Private dtDataPubli As Date
'
'Public Property Get gNome()
'    gNome = strNome
'End Property
'
'Public Property Let lNome(ByVal Valor As String)
'    strNome = Valor
'End Property
'
'Public Property Get gEdicao()
'    gEdicao = strEdicao
'End Property
'
'Public Property Let lEdicao(ByVal Valor As String)
'    strEdicao = Valor
'End Property
'
'Public Property Get gAutor()
'    gAutor = strAutor
'End Property
'
'Public Property Let lAutor(ByVal Valor As String)
'    strAutor = Valor
'End Property
'
'Public Property Get gDataPubli()
'    gDataPubli = dtDataPubli
'End Property
'
'Public Property Let lDataPubli(ByVal Valor As Date)
'    dtDataPubli = Valor
'End Property
'
'Public Function NovoLivro(strNome As String, strEdicao As String, strAutor As String, dtDataPubli As Date) As Boolean
'
'   On Error GoTo NovoLivro_E
'   NovoLivro = False
'
'   ' Cria uma nova coleção se ainda não existe
'   If colLivro Is Nothing Then
'       Set colLivro = New Collection
'   End If
'
'   ' Adiciona um novo livro à coleção
'   Dim livro As New Collection
'   livro.Add strNome
'   livro.Add strAutor
'   livro.Add strEdicao
'   livro.Add dtDataPubli
'   colLivro.Add livro
'
'NovoLivro = True
'GoTo DestruirObjetos
'
'NovoLivro_E:
'DestruirObjetos:
'End Function
'
'Public Function EditarItem(listLivros As ListBox) As Boolean
'
'   On Error GoTo EditarItem_E
'
'   EditarItem = False
'
'   If listLivros.ListIndex <> -1 Then
'
'      Dim indexSelecionado As Integer
'      indexSelecionado = listLivros.ListIndex
'
'      Dim textoItem As String
'      textoItem = listLivros.List(indexSelecionado)
'
'      Dim novoTexto As String
'      novoTexto = InputBox("Digite suas alterações:", "Editar Item", textoItem)
'
'      If novoTexto <> "" Then
'         listLivros.RemoveItem indexSelecionado
'         listLivros.AddItem novoTexto, indexSelecionado
'         listLivros.ListIndex = indexSelecionado
'      End If
'   Else
'        MsgBox "Por favor, selecione um item para editar.", vbExclamation, "Editar Item"
'   End If
'
'   EditarItem = True
'
'   GoTo DestruirObjetos
'
'EditarItem_E:
'DestruirObjetos:
'End Function
'
'Public Function Excluir(listLivros As ListBox) As Boolean
'
'    On Error GoTo Excluir_E
'
'    Excluir = False
'
'    For i = 0 To 100 - 1
'            If (listLivros.Selected(i)) Then
'                listLivros.RemoveItem (i)
'                msg = MsgBox("Livro removido", 48, "Atenção")
'                Exit For
'            End If
'        Next i
'
'    Excluir = True
'
'    GoTo DestruirObjetos
'
'Excluir_E:
'DestruirObjetos:
'End Function
'
'Public Function Pesquisar(listLivros As ListBox, ByVal termoPesquisa As String) As Boolean
'
'    On Error GoTo Pesquisar_E
'
'    Pesquisar = False
'
'    Dim resultado As String
'    Dim Encontrado As Boolean
'
'    For i = 0 To listLivros.ListCount - 1
'        ' Obtém o texto do item atual
'        Dim textoItem As String
'        textoItem = listLivros.List(i)
'
'        ' Verifica se o texto do item contém o termo de pesquisa
'        If InStr(1, textoItem, termoPesquisa, vbTextCompare) > 0 Then
'            ' Se o termo de pesquisa for encontrado, adiciona o item ao resultado
'            resultado = resultado & textoItem & vbCrLf
'            Encontrado = True
'        End If
'    Next i
'
'    ' Exibe o resultado na caixa de texto
'    If Encontrado Then
'        msg = MsgBox("O livro encontrado foi " & resultado, vbInformation, "Informação")
'    Else
'       msg = MsgBox("Nenhum livro encontrado", 48, "Atenção")
'    End If
'
'    Pesquisar = True
'
'    GoTo DestruirObjetos
'
'Pesquisar_E:
'DestruirObjetos:
'End Function
