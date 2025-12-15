VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Existe Dir"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Existe Arq"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copia Arq"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Remove Dir"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Criar Dir"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove Arq"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Apenda"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Escreve"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Le"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dir + Arq"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'http://www.macoratti.net/d170801.htm

'https://talkcode.wordpress.com/2011/06/16/como-listar-os-arquivos-de-um-diretorio-em-vb/
'https://www.tomasvasquez.com.br/blog/microsoft-office/vba-listar-arquivos-de-um-diretorio/
'https://www.excelpraontem.com.br/listar-arquivos-e-pastas-em-um-diretorio-e-salvar-a-listagem-em-uma-planilha/
'https://www.hashtagtreinamentos.com/percorrendo-arquivos-de-uma-pasta

'https://stackoverflow.com/questions/1404758/how-to-read-a-file-and-write-into-a-text-file

    Dim fso As FileSystemObject
    Dim pasta As String

    'cria o objeto fso do tipo FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    'define o caminho da pasta
    pasta = App.Path & "\Dados"

    'verifica se a pasta \Dados existe , senão existir entao cria
    If Not (fso.FolderExists(pasta)) Then
        fso.CreateFolder (pasta)
    End If

    existeDB = fso.FileExists(strArquivo)
    
End Sub

Private Sub Command10_Click()

   Dim strFolderName As String
   Dim strFolderExists As String
 
   strFolderName = App.Path & "\Teste"
   strFolderExists = Dir(strFolderName, vbDirectory)
     
   If strFolderExists = "" Then
      MsgBox "The selected folder doesn't exist"
   Else
      MsgBox "The selected folder exists"
   End If

End Sub

Private Sub Command2_Click()


'https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/open-statement
'https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/inputstatement

   Dim strLinha As String
   Dim intArq As Integer
   Dim strArquivo As String
   Dim strFile As String


   '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'
   'Ler arquivo
   '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'
   strArquivo = App.Path & "\arquivo_teste.txt"
   intArq = FreeFile()
   Open strArquivo For Input As #intArq
   Do While Not EOF(intArq)
      Line Input #intArq, strLinha
      strFile = strFile + strLinha + vbNewLine
   Loop

   Close #intArq


   Debug.Print strFile

End Sub

Private Sub Command3_Click()
   
   
'https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/printstatement
'https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/put-statement

   Dim strLinha1, strLinha2 As String
   Dim intArq As Integer
   Dim strArquivo As String
   Dim strFile As String

   strLinha1 = "teste1 teste1 teste1"
   strLinha2 = "teste2 teste2 teste2"

   strArquivo = App.Path & "\teste_1.txt"
   intArq = FreeFile()
   Open strArquivo For Output As #intArq
   Print #intArq, strLinha1
   Write #intArq, strLinha2
   Close #intArq
   Close intArq

   intArq = 0

End Sub

Private Sub Command4_Click()


   Dim strLinha1, strLinha2 As String
   Dim intArq As Integer
   Dim strArquivo As String
   Dim strFile As String

   strLinha1 = "ultima linha"

   strArquivo = App.Path & "\teste_1.txt"
   intArq = FreeFile()
   Open strArquivo For Append As #intArq
   Print #intArq, strLinha1
   Close #intArq
   Close intArq

   intArq = 0

End Sub

Private Sub Command5_Click()

   Dim strArquivo As String

   strArquivo = App.Path & "\teste_1.txt"
   Kill strArquivo

End Sub

Private Sub Command6_Click()

'https://stackoverflow.com/questions/10803834/create-a-folder-and-sub-folder-in-excel-vba

   MkDir App.Path & "\Teste"
   

End Sub

Private Sub Command7_Click()

   RmDir App.Path & "\Teste"
   
     
End Sub

Private Sub Command8_Click()


Dim SourceFile, DestinationFile
SourceFile = App.Path & "\arquivo_teste.txt" ' Define source file name.
DestinationFile = App.Path & "\teste2.txt" ' Define target file name.
FileCopy SourceFile, DestinationFile ' Copy source to target.



End Sub

Private Sub Command9_Click()

   Dim strFileName As String
   Dim strFileExists As String
 
'   strFileName = App.Path & "\Teste_3.txt"
   strFileName = App.Path & "\arquivo_teste.txt"
   strFileExists = Dir(strFileName)
 
   If strFileExists = "" Then
      MsgBox "The selected file doesn't exist"
   Else
      MsgBox "The selected file exists"
   End If
    
End Sub
