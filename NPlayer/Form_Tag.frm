VERSION 5.00
Begin VB.Form Form_Tag 
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   Caption         =   "Meta-dados"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Tag.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Barra_Menu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   593
      TabIndex        =   41
      Top             =   480
      Width           =   8895
      Begin VB.Image Pagina_Anterior 
         Height          =   330
         Left            =   120
         Top             =   120
         Width           =   390
      End
      Begin VB.Image Pagina_Seguinte 
         Height          =   330
         Left            =   480
         Top             =   120
         Width           =   375
      End
      Begin VB.Image Fundo_Menu 
         Enabled         =   0   'False
         Height          =   600
         Left            =   0
         Picture         =   "Form_Tag.frx":57E2
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   633
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      Width           =   9495
      Begin VB.PictureBox Pic_Capa_Album 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00101010&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   6000
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
         Width           =   3000
         Begin VB.PictureBox Frame_Opcoes 
            Appearance      =   0  'Flat
            BackColor       =   &H00313131&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   201
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label Botao_Seguinte 
               BackColor       =   &H00EEEEEE&
               BackStyle       =   0  'Transparent
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   2640
               TabIndex        =   39
               ToolTipText     =   "Capa seguinte"
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label Botao_Antes 
               BackColor       =   &H00EEEEEE&
               BackStyle       =   0  'Transparent
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   2280
               TabIndex        =   38
               ToolTipText     =   "Capa anterior"
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label Botao_Eliminar 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   360
               TabIndex        =   37
               ToolTipText     =   "Eliminar capa"
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Botao_Adicionar 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   36
               ToolTipText     =   "Adicionar capa"
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.PictureBox Image_Sem_Capa 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00222222&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3000
            Left            =   120
            ScaleHeight     =   200
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   200
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   360
            Width           =   3000
         End
      End
      Begin VB.PictureBox Image_Moldura 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   5880
         ScaleHeight     =   217
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   217
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   600
         Width           =   3255
      End
      Begin VB.PictureBox Barra_Text_Comentario 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   240
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3120
         Width           =   5475
         Begin VB.TextBox Text_Comentario 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   720
            TabIndex        =   5
            Top             =   0
            Width           =   1380
         End
         Begin VB.Shape Contorno_Comentario 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Genero 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3120
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   171
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2565
         Begin VB.TextBox Text_Genero 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   720
            TabIndex        =   4
            Top             =   0
            Width           =   780
         End
         Begin VB.Shape Contorno_Genero 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Ano 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   240
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   171
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2565
         Begin VB.TextBox Text_Ano 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   720
            TabIndex        =   3
            Top             =   0
            Width           =   660
         End
         Begin VB.Shape Contorno_Ano 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Album 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3120
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   171
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2565
         Begin VB.TextBox Text_Album 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   720
            TabIndex        =   2
            Top             =   0
            Width           =   780
         End
         Begin VB.Shape Contorno_Album 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Artista 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   240
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   171
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2565
         Begin VB.TextBox Text_Artista 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   720
            TabIndex        =   1
            Top             =   0
            Width           =   780
         End
         Begin VB.Shape Contorno_Artista 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Titulo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   240
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Width           =   5475
         Begin VB.TextBox Text_Titulo 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   720
            TabIndex        =   0
            Top             =   0
            Width           =   780
         End
         Begin VB.Shape Contorno_Titulo 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.TextBox Text_Ficheiro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5760
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbPictureType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Form_Tag.frx":7F84
         Left            =   6720
         List            =   "Form_Tag.frx":7FC7
         TabIndex        =   17
         Top             =   3960
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox Text_Ficheiro_Capa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7080
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape Shape_Centro 
         BorderColor     =   &H00212121&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblIndex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capas:"
         ForeColor       =   &H00FF80FF&
         Height          =   195
         Left            =   6000
         TabIndex        =   32
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label_Remover 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         BackStyle       =   0  'Transparent
         Caption         =   "Remover as tags da música"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   3840
         Width           =   2385
      End
      Begin VB.Label Label_Artista 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Artista"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label_Musica 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label_Album 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Album"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   22
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label_Ano 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Width           =   330
      End
      Begin VB.Label Label_Genero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Género"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   20
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label_Comentario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comentário"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2880
         Width           =   1005
      End
   End
   Begin VB.PictureBox Frame_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5760
      Width           =   9735
      Begin VB.PictureBox Botao_Actualizar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3840
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   6
         Top             =   120
         Width           =   1740
         Begin VB.Label Label_Actualizar 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actualizar"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   450
            TabIndex        =   14
            Top             =   45
            Width           =   840
         End
         Begin VB.Shape Contorno_Actualizar 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Botao_Ok 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   5880
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   7
         Top             =   120
         Width           =   1740
         Begin VB.Shape Contorno_Ok 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label_Ok 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ok"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   750
            TabIndex        =   13
            Top             =   45
            Width           =   240
         End
      End
      Begin VB.PictureBox Botao_Cancelar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   7800
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   8
         Top             =   120
         Width           =   1740
         Begin VB.Shape Contorno_Cancelar 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label_Cancelar 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   465
            TabIndex        =   12
            Top             =   45
            Width           =   780
         End
      End
      Begin VB.Image Fundo_Frame_Botoes 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   593
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   8895
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   8520
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Editor de tags"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   75
         TabIndex        =   10
         Top             =   120
         Width           =   1365
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00212121&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Tag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   NPlayer
'   Copyright © 2011-2012 Nikyts software ™ - Informática e tecnologia
'   www.nikyts.com / nikyts@hotmail.com
'   Desenvolvido por: Nelson do Carmo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Chamar a classe para ler as tags dos ficheiros
Dim cFile As New Classe_Tag_Editor

'Variáveis para visualizar as capas dos albums
Private MP3Path    As String
Private TPic       As IPictureDisp
Private CurIndex   As Long
Private MaxIndex   As Long

'Variável para criar um novo common dialog
Dim Explorador As New Class_Dialog

Private Sub Botao_Actualizar_Click()
    'Atalho para
    Label_Actualizar_Click
End Sub

Private Sub Botao_Actualizar_GotFocus()
    'Colocar o focus no botao
    Contorno_Actualizar.Visible = True
End Sub

Private Sub Botao_Actualizar_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Actualizar_Click
    If KeyCode = vbKeyRight Then Botao_Actualizar_LostFocus: Botao_Ok_GotFocus: Botao_Ok.SetFocus
End Sub

Private Sub Botao_Actualizar_LostFocus()
    'Ao perder o focus no botao
    Contorno_Actualizar.Visible = False
End Sub

Private Sub Botao_Adicionar_Click()
    'Adicinar a capa do album
    Text_Ficheiro_Capa.Text = ""
    
    With Explorador
        .Filter = ("*.bmp,*.gif,*.jpg") ',*.png
        
        ' Decide a pasta inicial
        If Text_Ficheiro_Capa <> "" Then
            If Dir(Text_Ficheiro_Capa, vbDirectory) <> "" Then
                ' Se for uma pasta, é ela mesma
                .Path = Text_Ficheiro_Capa
            Else
                ' Se for um arquivo, extraia só o caminho
                .Path = left(Text_Ficheiro_Capa, InStrRev(Text_Ficheiro_Capa, "\"))
            End If
        End If
        
        .FileFlags = PATHMUSTEXIST
        .FileFlags = .FileFlags + EXPLORER
        
        ' Mostra o diálogo
        .DialogFile OpenFile
        If .cancel = True Then Exit Sub
        Text_Ficheiro_Capa.Text = .FullPath
        
        'Caso tenha sido selecionado algum ficheiro então adiciona-o á lista
        If Len(.FileName) <> 0 Then
            Image_Sem_Capa.Visible = False
            If Not ID3Exist(MP3Path) Then
            End If
            Set TPic = LoadPicture(Text_Ficheiro_Capa.Text)
            If WriteAlbumArt(MP3Path, CurIndex, TPic, cmbPictureType.ListIndex) Then
                Pic_Capa_Album.Cls
                ResizePic
                MaxIndex = MaxIndex + 1
                CurIndex = CurIndex + 1
                lblIndex = "Capas: " & CurIndex & " / " & MaxIndex
            End If
        End If
    End With
End Sub

Private Sub Botao_Cancelar_Click()
    'Atalho para
    Label_Cancelar_Click
End Sub

Private Sub Botao_Cancelar_GotFocus()
    'Colocar o focus no botao
    Contorno_Cancelar.Visible = True
End Sub

Private Sub Botao_Cancelar_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Cancelar_Click
    If KeyCode = vbKeyLeft Then Botao_Cancelar_LostFocus: Botao_Ok_GotFocus: Botao_Ok.SetFocus
End Sub

Private Sub Botao_Cancelar_LostFocus()
    'Remover o focus no botao
    Contorno_Cancelar.Visible = False
End Sub

Private Sub Botao_Eliminar_Click()
    'Remover a capa do album
    Mensagem_de_Aviso "Question", "Pretende realmente elimnar a capa do album?"
    
    'Caso o utilizador confirmar a aliminação da capa
    If Resposta = "Sim" Then
        Dim k As Long
        If LenB(Dir(MP3Path)) Then
            If DeleteAlbumArt(MP3Path, CurIndex) Then
                Pic_Capa_Album.Cls
                MaxIndex = MaxIndex - 1
                If CurIndex - 1 > 0 Then
                    CurIndex = CurIndex - 1
                    If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
                        cmbPictureType.ListIndex = k
                        ResizePic
                    End If
                    lblIndex = "Capas: " & CurIndex & " / " & MaxIndex
                ElseIf MaxIndex > 0 Then
                    If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
                        cmbPictureType.ListIndex = k
                        ResizePic
                    End If
                Else
                    CurIndex = 0
                    Image_Sem_Capa.Visible = True
                End If
                lblIndex = "Capas: " & CurIndex & " / " & MaxIndex
            End If
        End If
    End If
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Botao_Ok_Click()
    'Atalho para
    Label_Ok_Click
End Sub

Private Sub Botao_Ok_GotFocus()
    'Colocar o focus no botao
    Contorno_Ok.Visible = True
End Sub

Private Sub Botao_Ok_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Ok_Click
    If KeyCode = vbKeyLeft Then Botao_Ok_LostFocus: Botao_Actualizar_GotFocus: Botao_Actualizar.SetFocus
    If KeyCode = vbKeyRight Then Botao_Ok_LostFocus: Botao_Cancelar_GotFocus: Botao_Cancelar.SetFocus
End Sub

Private Sub Botao_Ok_LostFocus()
    'Ao perder o focus no botao
    Contorno_Ok.Visible = False
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Frame_Centro.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Centro.BorderColor = .Cor_Contorno_Frame_Centro.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Contorno_Titulo.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Artista.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Album.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Ano.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Genero.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Comentario.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Musica.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Artista.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Album.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Ano.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Genero.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Comentario.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Frame_Opcoes.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Botao_Adicionar.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Botao_Eliminar.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Botao_Antes.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Botao_Seguinte.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        'Barra_Text_Titulo.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Titulo.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Titulo.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Titulo.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Titulo.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Titulo.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Titulo.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Titulo.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Artista.Picture = .TextBox_Intermediate.Picture
        Barra_Text_Artista.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Artista.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Artista.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Artista.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Artista.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Artista.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Artista.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Album.Picture = .TextBox_Intermediate.Picture
        Barra_Text_Album.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Album.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Album.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Album.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Album.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Album.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Album.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Ano.Picture = .TextBox_Intermediate.Picture
        Barra_Text_Ano.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Ano.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Ano.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Ano.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Ano.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Ano.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Ano.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Genero.Picture = .TextBox_Intermediate.Picture
        Barra_Text_Genero.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Genero.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Genero.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Genero.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Genero.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Genero.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Genero.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Comentario.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Comentario.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Comentario.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Comentario.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Comentario.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Comentario.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Comentario.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Comentario.ForeColor = .Cor_Letra_Textbox.backcolor
        Fundo_Menu.Picture = .Bar_Menu.Picture
        Pagina_Anterior.Picture = .Button_Arrow_Right_Normal.Picture
        Pagina_Seguinte.Picture = .Button_Arrow_Left_Normal.Picture
        Image_Moldura.Picture = Nothing
        Image_Moldura.backcolor = .Cor_Fundo_Textbox.backcolor
        Image_Moldura.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 10, 0, 0, 10, 10
        Image_Moldura.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Image_Moldura.ScaleWidth, 10, 10, 0, 40, 10
        Image_Moldura.PaintPicture Form_Skin.Pic_TextBox.Picture, (Image_Moldura.ScaleWidth - 10), 0, 10, 10, 51, 0, 10, 10
        Image_Moldura.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 10, 10, (Image_Moldura.ScaleHeight - 20), 0, 10, 10, 10
        Image_Moldura.PaintPicture Form_Skin.Pic_TextBox.Picture, (Image_Moldura.ScaleWidth - 10), 10, 10, (Image_Moldura.ScaleHeight - 20), 51, 10, 10, 10
        Image_Moldura.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, (Image_Moldura.ScaleHeight - 10), 10, 10, 0, 17, 10, 10
        Image_Moldura.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, (Image_Moldura.ScaleHeight - 10), (Image_Moldura.ScaleWidth - 20), 10, 10, 17, 40, 10
        Image_Moldura.PaintPicture Form_Skin.Pic_TextBox.Picture, (Image_Moldura.ScaleWidth - 10), (Image_Moldura.ScaleHeight - 10), 10, 10, 51, 17, 10, 10
        Pic_Capa_Album.backcolor = .Cor_Fundo_Textbox.backcolor
        Image_Sem_Capa.backcolor = .Cor_Fundo_Textbox.backcolor
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Actualizar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Actualizar.Picture = .Pic_Button.Picture
        Botao_Ok.Picture = .Pic_Button.Picture
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Actualizar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Remover.ForeColor = .Cor_Contorno_Caixas.backcolor
    End With
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Tag_Editor", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Tag_Editor", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Musica.Caption = ReadINI("Tag_Editor", "Label_Music", Localizacao_Ficheiro_Lingua)
    Label_Artista.Caption = ReadINI("Tag_Editor", "Label_Artist", Localizacao_Ficheiro_Lingua)
    Label_Album.Caption = ReadINI("Tag_Editor", "Label_Album", Localizacao_Ficheiro_Lingua)
    Label_Ano.Caption = ReadINI("Tag_Editor", "Label_Year", Localizacao_Ficheiro_Lingua)
    Label_Genero.Caption = ReadINI("Tag_Editor", "Label_Gender", Localizacao_Ficheiro_Lingua)
    Label_Comentario.Caption = ReadINI("Tag_Editor", "Label_Comment", Localizacao_Ficheiro_Lingua)
    Label_Remover.Caption = ReadINI("Tag_Editor", "Label_Remove", Localizacao_Ficheiro_Lingua)
    Pagina_Anterior.ToolTipText = ReadINI("Tag_Editor", "Button_Music_Previous", Localizacao_Ficheiro_Lingua)
    Pagina_Seguinte.ToolTipText = ReadINI("Tag_Editor", "Button_Music_Next", Localizacao_Ficheiro_Lingua)
    Botao_Adicionar.ToolTipText = ReadINI("Tag_Editor", "Icon_Add_Cover", Localizacao_Ficheiro_Lingua)
    Botao_Eliminar.ToolTipText = ReadINI("Tag_Editor", "Icon_Remove_Cover", Localizacao_Ficheiro_Lingua)
    Label_Actualizar.Caption = ReadINI("Tag_Editor", "Button_Update", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Tag_Editor", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Tag_Editor", "Button_Cancel", Localizacao_Ficheiro_Lingua)
End Sub

Private Sub Label_Adicionar_Click()
    'Atalho para
    Botao_Adicionar_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver opções da capa do album
    Frame_Opcoes.Visible = False
End Sub

Private Sub Frame_Centro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver opções da capa do album
    Frame_Opcoes.Visible = False
End Sub

Private Sub Image_Sem_Capa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver opções da capa do album
    Frame_Opcoes.Visible = True
End Sub

Private Sub Label_Cancelar_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Label_Actualizar_Click()
    'Conclui operação
    If Text_Ficheiro.Text = "" Then Exit Sub
    Me.MousePointer = 11
    Actualizar_Tags
    Me.MousePointer = 0
End Sub

Private Sub Label_Eliminar_Click()
    'Atalho para
    Botao_Eliminar_Click
End Sub

Private Sub Label_Ok_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Label_Remover_Click()
    'Procedimento para actualizar as tags do ficheiro selecionado
    If Text_Ficheiro.Text = "" Then Exit Sub
    
    'Remover a capa do album
    Mensagem_de_Aviso "Question", "Esta opção irá remover todas as tags do arquivo." & vbNewLine & "Pretende mesmo assim continuar?"
    
    'Caso o utilizador confirmar a aliminação da capa
    If Resposta = "Sim" Then
        With cFile
            .FileName = Text_Ficheiro.Text
            .Title = ""
            .Artist = ""
            .album = ""
            .Year = ""
            .Genre = ""
            .Comments = ""
            .EncodedBy = ""
            .DeleteID3Tags
        End With
        
        Text_Ficheiro.Text = ""
        
        'Limpar as caixas de texto com as tags do arquivo
        Text_Titulo.Text = ""
        Text_Artista.Text = ""
        Text_Album.Text = ""
        Text_Ano.Text = ""
        Text_Genero.Text = ""
        Text_Comentario.Text = ""
    End If
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Tag
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Tag
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Tag
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para construir o formulario, ajustando os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Titulo.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * (Barra_Text_Titulo.ScaleWidth + (3 * 16) + Image_Moldura.ScaleWidth + 16)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Form_Skin.Bar_Menu.Height + 16 + Form_Skin.Moldura_Pic_Tag_Editor.Height _
                + 16 + (2 * Fundo_Frame_Botoes.Height))
    End With
    
    Ajustar_Formulario_com_Menu Form_Tag, False, False, True, True
    
    Ajustar_Botao Form_Tag, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Tag, Botao_Ok, Label_Ok, True, Contorno_Ok
    Ajustar_Botao Form_Tag, Botao_Actualizar, Label_Actualizar, True, Contorno_Actualizar
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With
    With Botao_Actualizar
        .left = Botao_Ok.left - .ScaleWidth - .top
    End With
    
    With Barra_Menu
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Height = Form_Skin.Bar_Menu.Height
        .Width = Barra_ControlBox.ScaleWidth
        .left = 0
    End With
    
    With Fundo_Menu
        .Stretch = True
        .top = 0
        .Width = Barra_Menu.ScaleWidth
        .left = 0
    End With
    
    Ajustar_Caixa_Texto Barra_Text_Titulo, Text_Titulo, Contorno_Titulo, False
    Ajustar_Caixa_Texto_Media Barra_Text_Artista, Text_Artista, Contorno_Artista
    Ajustar_Caixa_Texto_Media Barra_Text_Album, Text_Album, Contorno_Album
    Ajustar_Caixa_Texto_Media Barra_Text_Ano, Text_Ano, Contorno_Ano
    Ajustar_Caixa_Texto_Media Barra_Text_Genero, Text_Genero, Contorno_Genero
    Ajustar_Caixa_Texto Barra_Text_Comentario, Text_Comentario, Contorno_Comentario, False
    
    With Label_Musica
        .left = 16
        .top = .left
    End With
    
    With Barra_Text_Titulo
        .top = Label_Musica.top + Label_Musica.Height + 3
        .left = Label_Musica.left
    End With
    
    With Label_Artista
        .top = Barra_Text_Titulo.top + Barra_Text_Titulo.ScaleHeight + 6
        .left = Label_Musica.left
    End With
        
    With Barra_Text_Artista
        .top = Label_Artista.top + Label_Artista.Height + 3
        .left = Label_Musica.left
    End With
    
    With Label_Album
        .top = Label_Artista.top
        .left = Barra_Text_Titulo.left + Barra_Text_Titulo.ScaleWidth - Barra_Text_Album.Width
    End With
    
    With Barra_Text_Album
        .top = Barra_Text_Artista.top
        .left = Label_Album.left
    End With
    
    With Label_Ano
        .top = Barra_Text_Album.top + Barra_Text_Album.ScaleHeight + 6
        .left = Label_Musica.left
    End With
        
    With Barra_Text_Ano
        .top = Label_Ano.top + Label_Ano.Height + 3
        .left = Label_Musica.left
    End With
    
    With Label_Genero
        .top = Label_Ano.top
        .left = Barra_Text_Album.left
    End With
    
    With Barra_Text_Genero
        .top = Barra_Text_Ano.top
        .left = Barra_Text_Album.left
    End With
    
    With Label_Comentario
        .top = Barra_Text_Ano.top + Barra_Text_Ano.ScaleHeight + 6
        .left = Label_Musica.left
    End With
    
    With Barra_Text_Comentario
        .top = Label_Comentario.top + Label_Comentario.Height + 3
        .left = Label_Musica.left
    End With
    
    With Label_Remover
        .top = Barra_Text_Comentario.top + Barra_Text_Comentario.ScaleHeight + 16
        .left = Label_Musica.left
    End With
     
    With Image_Moldura
        .Height = Form_Skin.Moldura_Pic_Tag_Editor.Height
        .top = Barra_Text_Titulo.top
        .Width = Form_Skin.Moldura_Pic_Tag_Editor.Width
        .left = Barra_Text_Titulo.left + Barra_Text_Titulo.ScaleWidth + 20
    End With
   
    With Pic_Capa_Album
        .Height = Form_Skin.Moldura_Pic_Tag_Editor.Height - 4
        .top = Image_Moldura.top + 2
        .Width = Form_Skin.Moldura_Pic_Tag_Editor.Width - 4
        .left = Image_Moldura.left + 2
    End With
    
    With Pic_Capa_Album
        .Height = Form_Skin.Image_Sem_Capa.Height
        .top = Image_Moldura.top + ((Image_Moldura.ScaleHeight - .ScaleHeight) / 2)
        .Width = Form_Skin.Image_Sem_Capa.Width
        .left = Image_Moldura.left + ((Image_Moldura.ScaleWidth - .ScaleWidth) / 2)
    End With
    
    With Image_Sem_Capa
        .top = 0
        .left = 0
    End With
    
    With Frame_Opcoes
        .top = 0
        .Width = Pic_Capa_Album.ScaleWidth
        .left = 0
    End With
    
    With Botao_Adicionar
        .top = (Frame_Opcoes.ScaleHeight - .Height) / 2
        .left = .top
    End With
    
    With Botao_Eliminar
        .top = Botao_Adicionar.top
        .left = Botao_Adicionar.left + Botao_Adicionar.Width + 6
    End With
    
    With Botao_Seguinte
        .top = Botao_Adicionar.top
        .left = Frame_Opcoes.ScaleWidth - .Width - .top
    End With
    
    With Botao_Antes
        .top = Botao_Adicionar.top
        .left = Botao_Seguinte.left - .Width - .top
    End With
    
    With Pagina_Anterior
        .top = (Barra_Menu.ScaleHeight - .Height) / 2
        .left = 10
    End With
    
    With Pagina_Seguinte
        .top = (Barra_Menu.ScaleHeight - .Height) / 2
        .left = Pagina_Anterior.left + Pagina_Anterior.Width
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Tag
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Tag
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Tag
End Sub

Public Sub Ler_Tags()
    'Procedimento para ler as tags do ficheiro selecionado
    Dim sFile As String, sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
    
    cFile.FileName = Text_Ficheiro.Text
    sTitle = Replace(cFile.Title, "'", " ", , , vbTextCompare)
    sArtist = Replace(cFile.Artist, "'", " ", , , vbTextCompare)
    sAlbum = Replace(cFile.album, "'", " ", , , vbTextCompare)
    sYear = Replace(cFile.Year, "'", " ", , , vbTextCompare)
    sGenre = Replace(cFile.Genre, "'", " ", , , vbTextCompare)
    sComment = Replace(cFile.Comments, "'", " ", , , vbTextCompare)
    
    'Left(Text_Ficheiro.Text, Len(Text_Ficheiro.Text) - 4)
    If sTitle = "" Then sTitle = Form_Principal.Grelha_Visivel.TextMatrix(Form_Principal.Grelha_Visivel.Row, 1)
    If sArtist = "" Then sArtist = ""
    If sAlbum = "" Then sAlbum = ""
    If sYear = "" Then sYear = ""
    If sGenre = "" Then sGenre = ""
    If sComment = "" Then sComment = ""
    
    'Preencher as caixas de texto com as tags do arquivo
    Text_Titulo.Text = sTitle
    Text_Artista.Text = sArtist
    Text_Album.Text = sAlbum
    Text_Ano.Text = sYear
    Text_Genero.Text = sGenre
    Text_Comentario.Text = sComment
    
    Ver_Capa_Album
End Sub

Public Sub Actualizar_Tags()
    On Error GoTo Corrige_Erro
    'Procedimento para actualizar as tags do ficheiro selecionado
    With cFile
        .FileName = Text_Ficheiro.Text
        .Title = Text_Titulo.Text
        .Artist = Text_Artista.Text
        .album = Text_Album.Text
        .Year = Text_Ano.Text
        .Genre = Text_Genero.Text
        .Comments = Text_Comentario.Text
        .EncodedBy = "NPlayer"
        .UpdateID3Tags
    End With
    
    'Actualizar os campos da grelha biblioteca referente ao arquivo em questão
    With Form_Principal
        .Grelha_Visivel.TextMatrix(.Grelha_Visivel.Row, 1) = Text_Titulo.Text
        .Grelha_Visivel.TextMatrix(.Grelha_Visivel.Row, 2) = Text_Artista.Text
        .Grelha_Visivel.TextMatrix(.Grelha_Visivel.Row, 3) = Text_Album.Text
        .Grelha_Visivel.TextMatrix(.Grelha_Visivel.Row, 4) = Text_Ano.Text
        .Grelha_Visivel.TextMatrix(.Grelha_Visivel.Row, 5) = Text_Genero.Text
        .Grelha_Visivel.TextMatrix(.Grelha_Visivel.Row, 6) = Text_Comentario.Text
            
        'Actualizar a base de dados
        If .Grelha_Musica.Visible = True Then
            Dim Chave As String
            Chave = .Grelha_Visivel.TextMatrix(.Grelha_Visivel.Row, 8)
            .Cnn_Biblioteca.Execute "Update Tabela_Musica Set Titulo = '" & Text_Titulo.Text & "', Artista = '" & Text_Artista.Text & "', Album = '" & Text_Album.Text & "', Ano = '" & Text_Ano.Text & "', Genero = '" & Text_Genero.Text & "', Comentario = '" & Text_Comentario.Text & "'   where Id = '" & Chave & "'"
            .Rs_Musica.Requery 1
            
        ElseIf .Grelha_Listas.Visible = True Then
            .Actualizar_Lista
        End If
    End With
    
    
Exit Sub
Corrige_Erro:
Unload Me
End Sub

Private Sub Pagina_Anterior_Click()
    'Passar para a linha seguinte e ler tags do arquivo
    With Form_Principal.Grelha_Visivel
        If .Row = 1 Then Exit Sub
        .Row = .Row - 1
        .ColSel = .Cols - 1
        Text_Ficheiro.Text = .TextMatrix(.Row, 0)
    End With
    Ler_Tags
End Sub

Private Sub Pagina_Anterior_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver a imagem down
    Pagina_Anterior.Picture = Form_Skin.Button_Arrow_Right_Down.Picture
End Sub

Private Sub Pagina_Anterior_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver a imagem normal
    Pagina_Anterior.Picture = Form_Skin.Button_Arrow_Right_Normal.Picture
End Sub

Private Sub Pagina_Seguinte_Click()
    'Passar para a linha seguinte e ler tags do arquivo
    With Form_Principal.Grelha_Visivel
        If .Row = .Rows - 1 Then Exit Sub
        .Row = .Row + 1
        .ColSel = .Cols - 1
        Text_Ficheiro.Text = .TextMatrix(.Row, 0)
    End With
    Ler_Tags
End Sub

Public Sub Ver_Capa_Album()
    'Procedimento para ver a capa do album
    Dim k As Long
    MP3Path = Text_Ficheiro.Text
    Pic_Capa_Album.Cls
    If ID3Exist(MP3Path) Then
        MaxIndex = GetAlbumArtCount(MP3Path)
        If MaxIndex > 0 Then
            CurIndex = 1
            If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
                cmbPictureType.ListIndex = 0
                ResizePic
                Image_Sem_Capa.Visible = False
            End If
        Else
            CurIndex = 0
            Image_Sem_Capa.Visible = True
        End If
    Else
        CurIndex = 0
        MaxIndex = 0
        Image_Sem_Capa.Visible = True
    End If
    lblIndex = "Capas: " & CurIndex & " / " & MaxIndex
On Error GoTo 0
End Sub

Private Sub ResizePic()
    'Ajustar a capa do album
    Dim nWidth  As Long
    Dim nHeight As Long

    On Error Resume Next
    nWidth = ScaleX(TPic.Width, vbHimetric, vbPixels)
    nHeight = ScaleY(TPic.Height, vbHimetric, vbPixels)
    With Pic_Capa_Album
        If .ScaleWidth < (nWidth * (.ScaleHeight / nHeight)) Then
            nHeight = nHeight * (.ScaleWidth / nWidth)
            nWidth = .ScaleWidth
        Else
            nWidth = nWidth * (.ScaleHeight / nHeight)
            nHeight = .ScaleHeight
        End If
        TPic.Render .hdc, (.ScaleWidth - CLng(nWidth)) / 2, (.ScaleHeight - CLng(nHeight)) / 2, CLng(nWidth), CLng(nHeight), 0, TPic.Height, TPic.Width, -TPic.Height, ByVal 0&
    End With
    On Error GoTo 0
End Sub

Private Sub Botao_Antes_Click()
    Dim k As Long
    If CurIndex > 1 Then
        Pic_Capa_Album.Cls
        CurIndex = CurIndex - 1
        lblIndex = "Capas: " & CurIndex & " / " & MaxIndex
        If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
            cmbPictureType.ListIndex = k
            ResizePic
        End If
    End If
End Sub

Private Sub Botao_Seguinte_Click()
    Dim k As Long
    If CurIndex < MaxIndex Then
        Pic_Capa_Album.Cls
        CurIndex = CurIndex + 1
        lblIndex = "Capas: " & CurIndex & " / " & MaxIndex
        If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
            cmbPictureType.ListIndex = k
            ResizePic
        End If
    End If
End Sub

Private Sub Pagina_Seguinte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver a imagem down
    Pagina_Seguinte.Picture = Form_Skin.Button_Arrow_Left_Down.Picture
End Sub

Private Sub Pagina_Seguinte_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver a imagem normal
    Pagina_Seguinte.Picture = Form_Skin.Button_Arrow_Left_Normal.Picture
End Sub

Private Sub Pic_Capa_Album_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver opções da capa do album
    Frame_Opcoes.Visible = True
End Sub

Private Sub Text_Titulo_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Titulo.Visible = True
End Sub

Private Sub Text_Titulo_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Titulo.Visible = False
End Sub

Private Sub Text_Artista_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Artista.Visible = True
End Sub

Private Sub Text_Artista_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Artista.Visible = False
End Sub

Private Sub Text_Album_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Album.Visible = True
End Sub

Private Sub Text_Album_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Album.Visible = False
End Sub

Private Sub Text_Ano_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Ano.Visible = True
End Sub

Private Sub Text_Ano_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Ano.Visible = False
End Sub

Private Sub Text_Genero_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Genero.Visible = True
End Sub

Private Sub Text_Genero_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Genero.Visible = False
End Sub

Private Sub Text_Comentario_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Comentario.Visible = True
End Sub

Private Sub Text_Comentario_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Comentario.Visible = False
End Sub
'
'Private Sub Barra_Menu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    'Capturar a posição de x e y
'    Capturar_Posicao_Formulario Form_Tag
'End Sub
'
'Private Sub Barra_Menu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    'Mover o formulário e obter a posição de x e y
'    Mover_Formulario Form_Tag
'End Sub
'
'Private Sub Barra_Menu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    'Largar o formulário para a posição final
'    Largar_Formulario Form_Tag
'End Sub

