VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salmos 92"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Escolha para mim."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   4455
   End
   Begin VB.Frame fraSalmo 
      BackColor       =   &H80000004&
      Caption         =   "Versículo do Dia"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4455
      Begin VB.TextBox txtSalmo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Escolher um versículo"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variáveis auxiliares no algoritmo
Dim strDia As String

Private Sub cmdNovo_Click()
   fraSalmo.Caption = ""
   
   Randomize
   
   Dim contador As Integer
   contador = Int(Rnd() * 15)

   If contador >= 0 And contador < Combo1.ListCount Then
      txtSalmo.Text = Combo1.List(contador)
   End If
   
   fraSalmo.Caption = "Versículo escolhido para você"
End Sub

Private Sub Combo1_Click()
   Dim contador As Integer
   contador = Combo1.ListIndex

   If contador >= 0 And contador < Combo1.ListCount Then
      txtSalmo.Text = Combo1.List(contador)
   End If
   
   fraSalmo.Caption = "Versículo escolhido por você"
End Sub

Private Sub Form_Load()
   Combo1.AddItem "1. Como é bom render graças ao Senhor e cantar louvores ao teu nome, ó Altíssimo;"
   Combo1.AddItem "2. anunciar de manhã o teu amor leal e de noite a tua fidelidade,"
   Combo1.AddItem "3. ao som da lira de dez cordas e da cítara, e da melodia da harpa."
   Combo1.AddItem "4. Tu me alegras, Senhor, com os teus feitos; as obras das tuas mãos levam-me a cantar de alegria."
   Combo1.AddItem "5. Como são grandes as tuas obras, Senhor, como são profundos os teus propósitos!"
   Combo1.AddItem "6. O insensato não entende, o tolo não vê"
   Combo1.AddItem "7. que, embora os ímpios brotem como a erva e floresçam todos os malfeitores, eles serão destruídos para sempre."
   Combo1.AddItem "8. Pois tu, Senhor, és exaltado para sempre."
   Combo1.AddItem "9. Mas os teus inimigos, Senhor, os teus inimigos perecerão; serão dispersos todos os malfeitores!"
   Combo1.AddItem "10. Tu aumentaste a minha força como a do boi selvagem; derramaste sobre mim óleo novo."
   Combo1.AddItem "11. Os meus olhos contemplaram a derrota dos meus inimigos; os meus ouvidos escutaram a debandada dos meus maldosos agressores."
   Combo1.AddItem "12. Os justos florescerão como a palmeira, crescerão como o cedro do Líbano;"
   Combo1.AddItem "13. plantados na casa do Senhor, florescerão nos átrios do nosso Deus."
   Combo1.AddItem "14. Mesmo na velhice darão fruto, permanecerão viçosos e verdejantes,"
   Combo1.AddItem "15. para proclamar que o Senhor é justo. Ele é a minha Rocha; nele não há injustiça."
   
   'Cria um objeto do tipo LydiansWin.LYD_Windows para pegar o dia atual
   Set objWindows = CreateObject("LydiansWin.LYD_Windows")
   
   strDia = objWindows.Computador.Tempo.Dia
   txtSalmo.Text = CStr(Combo1.List(CInt(strDia - 1)))

   Label1.Caption = "Versículo do dia"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Limpa o objeto
   Set objWindows = Nothing
End Sub
