VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

' Adicionar no projeto
' Microsoft ActiveX Data Object 6.1 Library

Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim connString As String

'connString = "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Lab;Trusted_Connection=yes"
connString = "Driver=SQL Server;Server=10.11.12.7\lydsqlhml;Database=Desenv1;User Id=lydians_mdc;Password=lydians_mdc"
conn.Open connString

rs.Open "Select * from livros", conn, adOpenKeyset

' conn.Execute ("INSERT INTO ...")


 Text1.Text = rs("nome").Value
 Text1.Text = rs(2).Value

While Not rs.EOF
   Text1.Text = rs.Fields.Item(1).Value
   rs.MoveNext
Wend

 
rs.MoveFirst

If rs.RecordCount > 0 Then
   Text1.Text = rs.Fields.Item(1).Value
End If


rs.Close
conn.Close

End Sub
