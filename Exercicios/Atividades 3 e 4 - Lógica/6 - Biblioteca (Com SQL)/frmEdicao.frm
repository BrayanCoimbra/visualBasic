VERSION 5.00
Begin VB.Form frmEdicao 
   Caption         =   "Livro Encontrado"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdicaoExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdEdicaoSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdEdicaoSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdEdicaoEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame fraEdicao 
      Caption         =   "Edição"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      Begin VB.TextBox txtEdicaoDTPubli 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtEdicaoAutor 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtEdicaoEdicao 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtEdicaoNome 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblEdicaoDTPubli 
         Caption         =   "Data de Publicação"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lbEdicaoAutor 
         Caption         =   "Autor"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblEdicaoEdicao 
         Caption         =   "Edição"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblEdicaoNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmEdicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEdicaoEditar_Click()
   txtEdicaoNome.Enabled = True
   txtEdicaoEdicao.Enabled = True
   txtEdicaoAutor.Enabled = True
   txtEdicaoDTPubli.Enabled = True
   
   cmdEdicaoSalvar.Visible = True
   cmdEdicaoSalvar.Enabled = True
End Sub

Private Sub cmdEdicaoExcluir_Click()
   Set Library = New clsBD
   
   If frmBiblioteca.lstLivros.SelectedItem.Selected Then
      Library.PesquisarComListView frmBiblioteca.lstLivros
      If Library.gEncontrado Then
         If Library.Excluir Then
            MsgBox Library.gmsg, vbInformation, "Exclusão"
         End If
      End If
   ElseIf (Library.Pesquisar(frmBiblioteca.lstLivros, frmBiblioteca.txtPesquisarLivros.Text)) Then
      If Library.Excluir Then
         MsgBox Library.gmsg, vbInformation, "Exclusão"
      End If
   Else
      MsgBox "Erro no botão Edicao", vbInformation, "Pesquisar"
   End If
   
   frmBiblioteca.AtualizarListView
   
   Unload frmEdicao
End Sub

Private Sub cmdEdicaoSair_Click()
    Unload frmEdicao
End Sub

Private Sub cmdEdicaoSalvar_Click()
   If Trim(txtEdicaoNome.Text) <> "" And Not IsNumeric(txtEdicaoNome.Text) And IsNumeric(txtEdicaoEdicao.Text) And Trim(txtEdicaoAutor.Text) <> "" And Not IsNumeric(txtEdicaoAutor.Text) And IsDate(txtEdicaoDTPubli.Text) Then
         Set Library = New clsBD
         If frmBiblioteca.lstLivros.SelectedItem.Selected Then
            Library.PesquisarComListView frmBiblioteca.lstLivros
            If Library.gEncontrado Then
               If Library.AtualizarRegistro(Library.gLivroID(Library.gGuardaIndiceLivroEncontrado), txtEdicaoNome.Text, txtEdicaoAutor.Text, txtEdicaoEdicao.Text, txtEdicaoDTPubli.Text) Then
                  MsgBox "Informações sobre o livro '" & txtEdicaoNome.Text & "' foram atualizadas!", vbInformation, "Atualização"
                  frmBiblioteca.AtualizarListView
                  Unload frmEdicao
               End If
            End If
         ElseIf Library.Pesquisar(frmBiblioteca.lstLivros, CStr(frmBiblioteca.txtPesquisarLivros.Text)) Then
            If Library.AtualizarRegistro(Library.gLivroID(Library.gGuardaIndiceLivroEncontrado), txtEdicaoNome.Text, txtEdicaoAutor.Text, txtEdicaoEdicao.Text, txtEdicaoDTPubli.Text) Then
            MsgBox "Informações sobre o livro '" & txtEdicaoNome.Text & "' foram atualizadas!", vbInformation, "Atualização"
            frmBiblioteca.AtualizarListView
            Unload frmEdicao
            End If
         Else
            MsgBox "Erro no botão Edição", vbInformation, "Pesquisar"
         End If
      Else: MsgBox "Não foi possível salvar as alterações, pois os campos foram preenchidos incorretamente!", vbInformation, "Pesquisar"
   End If
End Sub

Private Sub Form_Load()
   txtEdicaoNome.Enabled = False
   txtEdicaoEdicao.Enabled = False
   txtEdicaoAutor.Enabled = False
   txtEdicaoDTPubli.Enabled = False
   
   cmdEdicaoSalvar.Enabled = False
   cmdEdicaoSalvar.Visible = False
   
   AtualizaCamposComLivroEncontrado
End Sub

Private Sub AtualizaCamposComLivroEncontrado()
   Set Library = New clsBD
   
   If frmBiblioteca.lstLivros.SelectedItem.Selected Then
      Library.PesquisarComListView frmBiblioteca.lstLivros
      If Library.gEncontrado Then
         txtEdicaoNome.Text = Library.gLivroNome.Item(Library.gGuardaIndiceLivroEncontrado)
         txtEdicaoAutor.Text = Library.gLivroAutor.Item(Library.gGuardaIndiceLivroEncontrado)
         txtEdicaoEdicao.Text = Library.gLivroEdicao.Item(Library.gGuardaIndiceLivroEncontrado)
         txtEdicaoDTPubli.Text = Library.gLivroDTPubli.Item(Library.gGuardaIndiceLivroEncontrado)
      End If
   ElseIf (Library.Pesquisar(frmBiblioteca.lstLivros, CStr(frmBiblioteca.txtPesquisarLivros.Text))) Then
      If Library.gEncontrado Then
         txtEdicaoNome.Text = Library.gLivroNome.Item(Library.gGuardaIndiceLivroEncontrado)
         txtEdicaoAutor.Text = Library.gLivroAutor.Item(Library.gGuardaIndiceLivroEncontrado)
         txtEdicaoEdicao.Text = Library.gLivroEdicao.Item(Library.gGuardaIndiceLivroEncontrado)
         txtEdicaoDTPubli.Text = Library.gLivroDTPubli.Item(Library.gGuardaIndiceLivroEncontrado)
      Else
         MsgBox Library.gmsg, vbInformation, "Nenhum livro foi encontrado"
         Unload frmEdicao
      End If
   End If
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'   Me.Close
'End Sub
