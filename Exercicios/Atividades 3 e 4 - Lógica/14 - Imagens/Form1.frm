VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Processamento de Imagem"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   3480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   18
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   1800
      Picture         =   "Form1.frx":7572
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   17
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   120
      Picture         =   "Form1.frx":EAE4
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   16
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Rot. 180°"
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Rot. Esquerda "
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Rot. Direita"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Reset"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   7440
      Width           =   4935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Azul"
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Verde"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Vermelho"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   3480
      Picture         =   "Form1.frx":16056
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   1800
      Picture         =   "Form1.frx":1D5C8
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   120
      Picture         =   "Form1.frx":24B3A
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Negativo"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   3480
      Picture         =   "Form1.frx":2C0AC
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   1800
      Picture         =   "Form1.frx":3361E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   120
      Picture         =   "Form1.frx":3AB90
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cinza Invertido"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cinza"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ExtractR(ByVal CurrentColor As Long) As Byte
   ExtractR = CurrentColor And 255
End Function

Public Function ExtractG(ByVal CurrentColor As Long) As Byte
   ExtractG = (CurrentColor \ 256) And 255
End Function

Public Function ExtractB(ByVal CurrentColor As Long) As Byte
   ExtractB = (CurrentColor \ 65536) And 255
End Function

Public Function ExtractGray(ByVal CurrentColor As Long) As Integer
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    
    R = ExtractR(CurrentColor)
    G = ExtractG(CurrentColor)
    B = ExtractB(CurrentColor)
    
    ExtractGray = (R + G + B) \ 3
End Function

Public Function ExtractInvertGray(ByVal CurrentColor As Long) As Integer
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    
    R = 255 - ExtractR(CurrentColor)
    G = 255 - ExtractG(CurrentColor)
    B = 255 - ExtractB(CurrentColor)
    
    ExtractInvertGray = (R + G + B) \ 3
End Function

Public Function ExtractNegative(ByVal CurrentColor As Long) As Integer
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    
    R = 255 - ExtractR(CurrentColor)
    G = 255 - ExtractG(CurrentColor)
    B = 255 - ExtractB(CurrentColor)
        
End Function

Public Function ExtractRed(ByVal CurrentColor As Long) As Integer
    Dim R As Integer
       
    R = ExtractR(CurrentColor)
        
    ExtractRed = R
End Function

Public Function ExtractGreen(ByVal CurrentColor As Long) As Integer
    Dim G As Integer
       
    G = ExtractG(CurrentColor)
        
    ExtractGreen = G
End Function

Public Function ExtractBlue(ByVal CurrentColor As Long) As Integer
    Dim B As Integer
       
    B = ExtractB(CurrentColor)
        
    ExtractBlue = B
End Function

'Cinza
Private Sub Command1_Click()
    Dim CurrentColor As Long
    Dim p As Byte
    
    For i = 0 To 99
        For j = 0 To 99
            CurrentColor = GetPixel(Picture1.hDC, i, j)
            p = ExtractGray(CurrentColor)
            SetPixel Picture1.hDC, i, j, RGB(p, p, p)
        Next j
    Next i
End Sub

'Cinza invertido
Private Sub Command2_Click()
    Dim CurrentColor As Long
    Dim p As Byte
    
    For i = 0 To 99
        For j = 0 To 99
            CurrentColor = GetPixel(Picture2.hDC, i, j)
            p = ExtractInvertGray(CurrentColor)
            SetPixel Picture2.hDC, i, j, RGB(p, p, p)
        Next j
    Next i
End Sub

'Negativo
Private Sub Command3_Click()
    Dim CurrentColor As Long
    Dim p As Integer
    
    For i = 0 To 99
        For j = 0 To 99
         CurrentColor = GetPixel(Picture3.hDC, i, j)
         R = 255 - ExtractR(CurrentColor)
         G = 255 - ExtractG(CurrentColor)
         B = 255 - ExtractB(CurrentColor)
         SetPixel Picture3.hDC, i, j, RGB(R, G, B)
        Next j
    Next i
End Sub

'Vermelho
Private Sub Command4_Click()
    Dim CurrentColor As Long
    Dim p As Integer
    
    For i = 0 To 99
        For j = 0 To 99
            CurrentColor = GetPixel(Picture4.hDC, i, j)
            p = ExtractRed(CurrentColor)
            SetPixel Picture4.hDC, i, j, RGB(p, 0, 0)
        Next j
    Next i
End Sub

'Verde
Private Sub Command5_Click()
    Dim CurrentColor As Long
    Dim p As Integer
    
    For i = 0 To 99
        For j = 0 To 99
            CurrentColor = GetPixel(Picture5.hDC, i, j)
            p = ExtractGreen(CurrentColor)
            SetPixel Picture5.hDC, i, j, RGB(0, p, 0)
        Next j
    Next i
End Sub

'Azul
Private Sub Command6_Click()
    Dim CurrentColor As Long
    Dim p As Integer
    
    For i = 0 To 99
        For j = 0 To 99
            CurrentColor = GetPixel(Picture6.hDC, i, j)
            p = ExtractBlue(CurrentColor)
            SetPixel Picture6.hDC, i, j, RGB(0, 0, p)
        Next j
    Next i
End Sub

'Reset
Private Sub Command7_Click()
    End
End Sub

'Rot. Direita
Private Sub Command8_Click()
Dim colGuardaPixel(99, 99) As Long

   For i = 0 To 99
      For j = 0 To 99
         colGuardaPixel(j, i) = GetPixel(Picture6.hDC, i, j)
      Next j
   Next i
   
   For i = 0 To 99
      For j = 0 To 99
         SetPixel Picture8.hDC, j, i, colGuardaPixel(99 - j, i)
      Next j
   Next i
   
   'Color2 = SetPixel(Picture2.hDC, 99 - j, 99 - i, RGB(R, G, B))
   
   'Picture8.Picture = LoadPicture("C:\Program Files (x86)\DevStudio\VB\ProjBrayan\Atividades3\Imagem\FilhoteDir.bmp")
End Sub

'Rot. Esquerda
Private Sub Command9_Click()
Dim colGuardaPixel(99, 99) As Long
Dim colGuardaPixelRotacionado(99, 99) As Long

   For i = 0 To 99
      For j = 0 To 99
         colGuardaPixel(j, i) = GetPixel(Picture6.hDC, i, j)
      Next j
   Next i
   
   For i = 0 To 99
      For j = 0 To 99
         SetPixel Picture7.hDC, i, j, colGuardaPixel(i, j)
      Next j
   Next i
 
   'Picture7.Picture = LoadPicture("C:\Program Files (x86)\DevStudio\VB\ProjBrayan\Atividades3\Imagem\FilhoteEsq.bmp")
End Sub

'Rot. 180
Private Sub Command10_Click()
Dim colGuardaPixel(99, 99) As Long

   For i = 0 To 99
      For j = 0 To 99
         colGuardaPixel(j, i) = GetPixel(Picture6.hDC, i, j)
      Next j
   Next i
   
   For i = 0 To 99
      For j = 0 To 99
         SetPixel Picture9.hDC, i, j, colGuardaPixel(99 - j, i)
      Next j
   Next i
   'Picture9.Picture = LoadPicture("C:\Program Files (x86)\DevStudio\VB\ProjBrayan\Atividades3\Imagem\Filhote180.bmp")
End Sub
