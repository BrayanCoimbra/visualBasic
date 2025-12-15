VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "COMCTL32.OCX"
Begin VB.Form frmBiblioteca 
   BackColor       =   &H80000004&
   Caption         =   "Biblioteca"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLivros 
      BackColor       =   &H80000004&
      Caption         =   "Livros"
      Height          =   3975
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   7455
      Begin ComctlLib.ListView lstLivros 
         Height          =   2415
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4260
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         _Version        =   327682
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdIr 
         Caption         =   "Ir"
         Height          =   400
         Left            =   6480
         TabIndex        =   14
         Top             =   590
         Width           =   735
      End
      Begin VB.TextBox txtPesquisarLivros 
         Height          =   405
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label Label2 
         Caption         =   "Biblioteca"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   6975
      End
      Begin VB.Label lblPesquisar 
         Caption         =   "Pesquisar"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame fraCadastroLivro 
      BackColor       =   &H80000004&
      Caption         =   "Cadastro de Livros"
      ForeColor       =   &H80000009&
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdCadastarLivro 
         Caption         =   "Cadastar"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtCadastroDataPubli 
         Enabled         =   0   'False
         Height          =   405
         Left            =   240
         TabIndex        =   4
         Text            =   "12-12-12"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtCadastroAutor 
         Enabled         =   0   'False
         Height          =   405
         Left            =   240
         TabIndex        =   3
         Text            =   "TesteAutor"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtCadastroEdicao 
         Enabled         =   0   'False
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Text            =   "1"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtCadastroNome 
         Enabled         =   0   'False
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Text            =   "TesteNome"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblCadastroNome 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblCadastroEdicao 
         BackStyle       =   0  'Transparent
         Caption         =   "Edição"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblCadastroAutor 
         BackStyle       =   0  'Transparent
         Caption         =   "Autor"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   135
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   15
      End
      Begin VB.Label lblCadastroDataPubli 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Publicação"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmBiblioteca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim livro As New clsLivro
Private Library As clsBD
Private i As Integer

Private Sub cmdCadastarLivro_Click()
   txtCadastroNome.BackColor = &H8000000E
   txtCadastroEdicao.BackColor = &H8000000E
   txtCadastroAutor.BackColor = &H8000000E
   txtCadastroDataPubli.BackColor = &H8000000E
   txtCadastroNome.Enabled = True
   txtCadastroEdicao.Enabled = True
   txtCadastroAutor.Enabled = True
   txtCadastroDataPubli.Enabled = True
   txtCadastroNome.Text = "TesteNome"
   txtCadastroEdicao.Text = 3
   txtCadastroAutor.Text = "TesteAutor"
   txtCadastroDataPubli.Text = "12-12-26"
   cmdSalvar.Enabled = True
End Sub

Private Sub cmdSalvar_Click()

   If Trim(txtCadastroNome.Text) <> "" And Not IsNumeric(txtCadastroNome.Text) And IsNumeric(txtCadastroEdicao.Text) And Trim(txtCadastroAutor.Text) <> "" And Not IsNumeric(txtCadastroAutor.Text) And IsDate(txtCadastroDataPubli.Text) Then
      
      Set Library = New clsBD
      With Library
         .Inserir txtCadastroNome.Text, txtCadastroEdicao.Text, txtCadastroAutor.Text, CDate(txtCadastroDataPubli.Text)
         .Selecionar
      End With
   
      lstLivros.ListItems.Clear
   
      For i = 1 To Library.gLivroNome.Count()
         Set itmX = lstLivros.ListItems.Add()
         itmX.Text = Library.gLivroID.Item(i)
         itmX.SubItems(1) = Library.gLivroNome.Item(i)
         itmX.SubItems(2) = Library.gLivroAutor.Item(i)
         itmX.SubItems(3) = Library.gLivroEdicao.Item(i)
         itmX.SubItems(4) = Library.gLivroDTPubli.Item(i)
      Next i
   
      txtCadastroEdicao.Text = ""
      txtCadastroAutor.Text = ""
      txtCadastroDataPubli.Text = ""
      txtCadastroNome.BackColor = &H80000004
      txtCadastroEdicao.BackColor = &H80000004
      txtCadastroAutor.BackColor = &H80000004
      txtCadastroDataPubli.BackColor = &H80000004
      
      txtCadastroNome.Enabled = False
      txtCadastroEdicao.Enabled = False
      txtCadastroAutor.Enabled = False
      txtCadastroDataPubli.Enabled = False
      cmdSalvar.Enabled = False
   
      MsgBox "Livro " & txtCadastroNome.Text & " adicionado!", vbInformation, "Informação"
      
      txtCadastroNome.Text = ""
   Else
      MsgBox "Você preencheu os campos incorretamente!", vbInformation, "Informação"
   End If
End Sub

Private Sub cmdIr_Click()
   If lstLivros.ListItems.Count = 0 Then
      MsgBox "Você não possui livros na biblioteca para pesquisar!", vbInformation, "Atenção"
   ElseIf lstLivros.SelectedItem.Selected Then
      frmEdicao.Show
   ElseIf txtPesquisarLivros.Text = "" Then
      MsgBox "Você não digitou algo para pesquisar!", vbInformation, "Atenção"
   Else
      If (Library.Pesquisar(lstLivros, CStr(txtPesquisarLivros.Text))) Then
         If Library.gEncontrado Then
            frmEdicao.Show
         Else
            MsgBox Library.gmsg, vbInformation, "Nenhum livro foi encontrado"
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   Dim clmX As ColumnHeader
   Dim itmX As ListItem
   
   Set clmX = lstLivros.ColumnHeaders.Add()
   clmX.Text = "ID"
   
   Set clmX = lstLivros.ColumnHeaders.Add()
   clmX.Text = "Nome"
   
   Set clmX = lstLivros.ColumnHeaders.Add()
   clmX.Text = "Autor"
   
   Set clmX = lstLivros.ColumnHeaders.Add()
   clmX.Text = "Edição"
   
   Set clmX = lstLivros.ColumnHeaders.Add()
   clmX.Text = "Data de Publicação"
   
   Set Library = New clsBD
   
   With Library
      .Selecionar
   End With
      
   lstLivros.ListItems.Clear
      
   For i = 1 To Library.gLivroNome.Count()
      Set itmX = lstLivros.ListItems.Add()
      itmX.Text = Library.gLivroID.Item(i)
      itmX.SubItems(1) = Library.gLivroNome.Item(i)
      itmX.SubItems(2) = Library.gLivroAutor.Item(i)
      itmX.SubItems(3) = Library.gLivroEdicao.Item(i)
      itmX.SubItems(4) = Library.gLivroDTPubli.Item(i)
   Next i
End Sub

Public Sub AtualizarListView()
   Set Library = New clsBD
   
   Library.Selecionar
      
   lstLivros.ListItems.Clear
      
   For i = 1 To Library.gLivroNome.Count()
      Set itmX = lstLivros.ListItems.Add()
      itmX.Text = Library.gLivroID.Item(i)
      itmX.SubItems(1) = Library.gLivroNome.Item(i)
      itmX.SubItems(2) = Library.gLivroAutor.Item(i)
      itmX.SubItems(3) = Library.gLivroEdicao.Item(i)
      itmX.SubItems(4) = Library.gLivroDTPubli.Item(i)
   Next i
End Sub

