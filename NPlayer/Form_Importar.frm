VERSION 5.00
Begin VB.Form Form_Importar 
   Appearance      =   0  'Flat
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   5730
   ClientLeft      =   90
   ClientTop       =   0
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Importar media"
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
         TabIndex        =   7
         Top             =   120
         Width           =   1545
      End
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   5880
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   195
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   465
      Width           =   6255
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   3960
         TabIndex        =   22
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5160
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   5160
         TabIndex        =   20
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   3960
         TabIndex        =   19
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Pic_Sim 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   210
      End
      Begin VB.PictureBox Pic_Nao 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1800
         Width           =   210
      End
      Begin VB.PictureBox Barra_Text_Pesquisar_Musica 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   480
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   5475
         Begin VB.PictureBox Botao_Pesquisar 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   4800
            ScaleHeight     =   12
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   12
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Selecionar pasta"
            Top             =   120
            Width           =   180
         End
         Begin VB.TextBox Text_Pesquisar_Musica 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00808080&
            Height          =   300
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   30
            Width           =   1140
         End
         Begin VB.Shape Contorno_Pesquisar_Musica 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.OptionButton Opcao_Nao 
         Appearance      =   0  'Flat
         BackColor       =   &H00313131&
         Caption         =   "Não importar mídia agora"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   5415
      End
      Begin VB.OptionButton Opcao_Sim 
         Appearance      =   0  'Flat
         BackColor       =   &H00313131&
         Caption         =   "Procurar ficheiros no meu computador"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   5415
      End
      Begin NPlayer.NProgressBar ProgressBar1 
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   3480
         Visible         =   0   'False
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   450
      End
      Begin VB.Label Label_Contador 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   2760
         Width           =   5115
      End
      Begin VB.Label Label_Adicionando 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Adicionando..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2415
         TabIndex        =   16
         Top             =   3120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label_Informacao 
         BackColor       =   &H00313131&
         BackStyle       =   0  'Transparent
         Caption         =   "(Uma vez inicializado, o programa permite que você adicione novamente ficheiros mídia de várias formas)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Width           =   5520
      End
      Begin VB.Label Label_Ficheiro 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   5475
      End
      Begin VB.Shape Shape_Centro 
         BorderColor     =   &H00212121&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Frame_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   408
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4920
      Width           =   6120
      Begin VB.PictureBox Botao_Iniciar_Pesquisa 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2160
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   2
         Top             =   120
         Width           =   1740
         Begin VB.Shape Contorno_Iniciar_Pesquisa 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label_Iniciar_Pesquisa 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Iniciar pesquisa"
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   195
            TabIndex        =   15
            Top             =   45
            Width           =   1350
         End
      End
      Begin VB.PictureBox Botao_Cancelar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   4080
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   3
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
            Left            =   480
            TabIndex        =   14
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
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00404040&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Importar"
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

'Variável para chamar a classe que lê as tags dos ficheiros
Dim cFile As New Classe_Tag_Editor

'Variáveis da tabela de música
Dim RS_Music As ADODB.Recordset
Dim Cnn_Music As ADODB.Connection
Dim SQL_Music As String

'Variáveis da tabela de filmes
Dim Rs_Film As ADODB.Recordset
Dim Cnn_Film As ADODB.Connection
Dim SQL_Film As String

'Variavel de pesquisa
Dim Pesquisando As Boolean

'Função para identificar qual foi a biblioteca selecionada
Public Biblioteca_Selecionada As String

'Variável para criar um novo common dialog
Dim Explorador As New Class_Dialog

'Variáveis para caregar o idioma
Dim Idioma_Iniciar_Pesquisa As String
Dim Idioma_Parar_Pesquisa As String
Dim Idioma_Ficheiros As String
Dim Idioma_Adicionando_1 As String
Dim Idioma_Adicionando_2 As String
Dim Idioma_Operacao_Cancelada As String
Dim Idioma_Concluir As String

'Variável para saber que extensão é para procurar
Public Pesquisar_pela_Extensao As String

Private Sub Pesquisar_Ficheiro()
    'Procedimento para carregar a biblioteca
    'On Error Resume Next
    If Pesquisando = True Then
        Me.MousePointer = 11
        Label_Iniciar_Pesquisa = Idioma_Parar_Pesquisa
        List2.AddItem Dir1.Path
        Dim a As Integer
        a = 0
        
        Do Until a >= List2.ListCount
            If Pesquisando = False Then Parar_Pesquisa_e_Adicionar: Exit Sub
            DoEvents
            Dir1.Path = List2.List(a)
        
            For b = 0 To File1.ListCount - 1
                If Right(LCase(File1.List(b)), 3) = Pesquisar_pela_Extensao Then
                    List1.AddItem Dir1.Path & "\" & File1.List(b)
                    Label_Contador.Caption = Idioma_Ficheiros & " [" & List1.ListCount - 1 & "]"
                    Label_Ficheiro.Caption = Dir1.Path & "\" & File1.List(b)
                End If
            Next
            
            For i = 0 To Dir1.ListCount - 1
                List2.AddItem Dir1.List(i)
            Next
            a = a + 1
        Loop
        
        Dir1.Path = Text_Pesquisar_Musica.Text & "\"
        Pesquisando = False
        Me.MousePointer = 0
        Label_Iniciar_Pesquisa = Idioma_Iniciar_Pesquisa
        Adicionar_Ficheiros
        
    Else
        'Chamar o procedimento
        Parar_Pesquisa_e_Adicionar
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Adicionar_Ficheiros()
    'Procedimento para adicionar os ficheiros encontrados
    If List1.ListCount = 0 Then
        Text_Pesquisar_Musica.Text = Empty
        Botao_Iniciar_Pesquisa.Enabled = True
        Label_Iniciar_Pesquisa.Enabled = True
        Label_Iniciar_Pesquisa.Caption = Idioma_Iniciar_Pesquisa
        Pesquisando = False
        Label_Contador.Caption = Idioma_Ficheiros & " [0]"
    
    Else
        With Form_Principal
            .Text_Caminho.Text = Text_Pesquisar_Musica.Text & "\"
            Call WriteINI("Library", "Location_Of_Albuns", Text_Pesquisar_Musica.Text & "\", (Localizacao_Ficheiro_Preferencias))
        End With
        
        Dim sFile As String, sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
        Dim i, j As Integer
        
        Botao_Iniciar_Pesquisa.Enabled = False
        Label_Iniciar_Pesquisa.Enabled = False
        
        Me.MousePointer = 11
        Label_Contador.Visible = False
        Label_Ficheiro.Visible = False
        Label_Adicionando.Visible = True
        List1.ListIndex = 0
        
        With ProgressBar1
            '.Min = 0
            .Max = List1.ListCount - 1
            .Value = 0
            ProgressBar1.Visible = True
            cFile.FileName = True
                 
            For j = 0 To List1.ListCount - 1
               Label_Adicionando.Caption = Idioma_Adicionando_1 & " [ " & j & " ] " & Idioma_Adicionando_2 & " "
               DoEvents
               cFile.FileName = List1.List(j)
               
               sTitle = Replace(cFile.Title, "'", " ", , , vbTextCompare)
               sArtist = Replace(cFile.Artist, "'", " ", , , vbTextCompare)
               sAlbum = Replace(cFile.album, "'", " ", , , vbTextCompare)
               sYear = Replace(cFile.Year, "'", " ", , , vbTextCompare)
               sGenre = Replace(cFile.Genre, "'", " ", , , vbTextCompare)
               sComment = Replace(cFile.Comments, "'", " ", , , vbTextCompare)
               
               'Caso o ficheiro não tenha tags então o titulo será o nome do ficheiro, o qual é obtido atrvés do directório
               Dim Arquivo() As String
               Dim DiretorioArq As String
               DiretorioArq = List1.List(j)
               Arquivo = Split(DiretorioArq, "\")
               
               Dim nome_ficheiro As String: nome_ficheiro = Dir(List1.List(j), vbArchive)
               If sTitle = "" Then sTitle = Mid(nome_ficheiro, 1, InStrRev(nome_ficheiro, ".") - 1)
               If sArtist = "" Then sArtist = ""
               If sAlbum = "" Then sAlbum = ""
               If sYear = "" Then sYear = ""
               If sGenre = "" Then sGenre = ""
               If sComment = "" Then sComment = ""
               
               'Adicionar as músicas na base de dados
                If Biblioteca_Selecionada = "Musica" Then
                    RS_Music.AddNew
                    RS_Music!ficheiro = List1.List(j)
                    RS_Music!Titulo = sTitle
                    RS_Music!Artista = sArtist
                    RS_Music!album = sAlbum
                    RS_Music!ano = sYear
                    RS_Music!Genero = sGenre
                    RS_Music!Comentario = sComment
                    RS_Music!Directorio = List1.List(j)
                    RS_Music!Classificacao = "0"
                    RS_Music!ID = j
                    RS_Music.Update
                   
                    If j = List1.ListCount - 1 Then
                        GoTo Concluir_Processo
                    Else
                        iProg = iProg + 1
                        .Value = iProg
                    End If
               End If
               
               'Adicionar os filmes na base de dados
                If Biblioteca_Selecionada = "Filmes" Then
                    Rs_Film.AddNew
                    Rs_Film!ficheiro = List1.List(j)
                    Rs_Film!Titulo = sTitle
                    Rs_Film!ano = sYear
                    Rs_Film!Genero = sGenre
                    Rs_Film!Comentario = sComment
                    Rs_Film!Directorio = List1.List(j)
                    Rs_Film!Classificacao = "0"
                    Rs_Film!ID = j
                    Rs_Film.Update
                   
                    If j = List1.ListCount - 1 Then
                        GoTo Concluir_Processo
                    Else
                        iProg = iProg + 1
                        .Value = iProg
                    End If
               End If
            Next j
                
            ProgressBar1.Visible = False
        End With
        
        'Fechar as conexões à base de dados
        If Biblioteca_Selecionada = "Musica" Then
            Cnn_Music.Close
            Set Cnn_Music = Nothing
        End If

        If Biblioteca_Selecionada = "Filmes" Then
            Cnn_Film.Close
            Set Cnn_Film = Nothing
        End If
        
        'Fechar automaticamente após o carregamento da grelha
        Me.MousePointer = 0
        Form_Principal.Show
        Unload Me
    End If
    
Exit Sub
Concluir_Processo:

'Fechar as conexões à base de dados
If Biblioteca_Selecionada = "Musica" Then
    Cnn_Music.Close
    Set Cnn_Music = Nothing
End If

If Biblioteca_Selecionada = "Filmes" Then
    Cnn_Film.Close
    Set Cnn_Film = Nothing
End If

ProgressBar1.Visible = False
Me.MousePointer = 0
Form_Principal.Show
Unload Me
End Sub

Private Sub Botao_Cancelar_Click()
    'Fechar formulário
    Label_Cancelar_Click
End Sub

Private Sub Botao_Cancelar_GotFocus()
    'Colocar o focus no botao
    Contorno_Cancelar.Visible = True
End Sub

Private Sub Botao_Cancelar_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Cancelar_Click
    If KeyCode = vbKeyLeft Then Botao_Cancelar_LostFocus: Botao_Iniciar_Pesquisa_GotFocus: Botao_Iniciar_Pesquisa.SetFocus
End Sub

Private Sub Botao_Cancelar_LostFocus()
    'Remover o focus no botao
    Contorno_Cancelar.Visible = False
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar a aplicação
    Unload Me
    End
End Sub

Private Sub Botao_Iniciar_Pesquisa_Click()
    'Atalho para
    Label_Iniciar_Pesquisa_Click
End Sub

Private Sub Botao_Iniciar_Pesquisa_GotFocus()
    'Colocar o focus no botao
    Contorno_Iniciar_Pesquisa.Visible = True
End Sub

Private Sub Botao_Iniciar_Pesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Iniciar_Pesquisa_Click
    If KeyCode = vbKeyRight Then Botao_Iniciar_Pesquisa_LostFocus: Botao_Cancelar_GotFocus: Botao_Cancelar.SetFocus
End Sub

Private Sub Botao_Iniciar_Pesquisa_LostFocus()
    'Ao perder o focus no botao
    Contorno_Iniciar_Pesquisa.Visible = False
End Sub

Private Sub Form_Load()
    'Propriedades inicais do formulário
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, True
    
    Pesquisando = False
    List1.Clear
    List2.Clear
    
    'Limpar campos
    Text_Pesquisar_Musica.Text = ""
    Label_Ficheiro.Caption = ""
    Label_Contador.Caption = ""
    Label_Adicionando.Caption = ""
    
    'Alterar cores do progreesbar
    ProgressBar1.backcolor = Form_Skin.Cor_Contorno_Caixas.backcolor
    
    Biblioteca_Selecionada = "Musica"
    Pesquisar_pela_Extensao = "mp3"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Form_Principal.Show: Unload Me
End Sub

Private Sub Form_Resize()
    'Atalho para
    Desenhar_Formulario
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
        Contorno_Pesquisar_Musica.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Informacao.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Contador.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Ficheiro.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Adicionando.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        'Barra_Text_Pesquisar_Musica.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Pesquisar_Musica.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Pesquisar_Musica.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Pesquisar_Musica.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Pesquisar_Musica.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Pesquisar_Musica.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Pesquisar_Musica.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Pesquisar_Musica.ForeColor = .Cor_Letra_Textbox.backcolor
        Botao_Pesquisar.Picture = .Botao_Pesquisar.Picture
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Iniciar_Pesquisa.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Iniciar_Pesquisa.Picture = .Pic_Button.Picture
        Contorno_Iniciar_Pesquisa.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Opcao_Sim.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Opcao_Sim.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Pic_Sim.Picture = .Opcao_Over.Picture
        Pic_Sim.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Opcao_Nao.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Opcao_Nao.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Pic_Nao.Picture = .Opcao_Normal.Picture
        Pic_Nao.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
    End With
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Import", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Import", "Button_Close", Localizacao_Ficheiro_Lingua)
    Opcao_Sim.Caption = ReadINI("Import", "Option_Yes", Localizacao_Ficheiro_Lingua)
    Opcao_Nao.Caption = ReadINI("Import", "Option_No", Localizacao_Ficheiro_Lingua)
    Label_Informacao.Caption = ReadINI("Import", "Label_Info", Localizacao_Ficheiro_Lingua)
    Botao_Pesquisar.ToolTipText = ReadINI("Import", "Button_Find", Localizacao_Ficheiro_Lingua)
    Idioma_Iniciar_Pesquisa = ReadINI("Import", "Button_Search_1", Localizacao_Ficheiro_Lingua)
    Idioma_Parar_Pesquisa = ReadINI("Import", "Button_Search_2", Localizacao_Ficheiro_Lingua)
    Idioma_Concluir = ReadINI("Import", "Button_Search_3", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Import", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    Idioma_Ficheiros = ReadINI("Import", "Info_Files", Localizacao_Ficheiro_Lingua)
    Idioma_Adicionando_1 = ReadINI("Import", "Info_Add_1", Localizacao_Ficheiro_Lingua)
    Idioma_Adicionando_2 = ReadINI("Import", "Info_Add_2", Localizacao_Ficheiro_Lingua)
    Idioma_Operacao_Cancelada = ReadINI("Import", "Operation_Canceled", Localizacao_Ficheiro_Lingua)
    
    Label_Iniciar_Pesquisa.Caption = Idioma_Iniciar_Pesquisa
End Sub

Private Sub Label_Cancelar_Click()
    'Cancelar operação
    Unload Me
    End
End Sub

Private Sub Label_Iniciar_Pesquisa_Click()
    'Iniciar a pesquisa de ficheiros
    If Opcao_Sim.Value = False Then
        importar_media = False
        Form_Principal.Show
        Unload Me
        
    Else
        If Text_Pesquisar_Musica.Text = Empty Then Exit Sub
        importar_media = True
        If Pesquisando = False Then
            'Pesquisa ficheiros media
'            Opcao_Sim.Enabled = False
'            Pic_Sim.Enabled = False
'            Opcao_Nao.Enabled = False
'            Pic_Nao.Enabled = False
            Text_Pesquisar_Musica.Enabled = False
            Botao_Pesquisar.Enabled = False
            Botao_Iniciar_Pesquisa.Enabled = False
            Label_Iniciar_Pesquisa.Enabled = False
            
            Pesquisando = True
            Me.MousePointer = 11
            Label_Iniciar_Pesquisa.Caption = Idioma_Parar_Pesquisa
            
            'Cnn_Biblioteca.Open "provider=microsoft.jet.oledb.4.0;persist security info = false; data source = " & App.Path & "\Library\Library.mdb"
            'Verifica_Rs_Musica
            If Biblioteca_Selecionada = "Musica" Then
                'Conectar á base de dados -> Musica
                Set RS_Music = New ADODB.Recordset
                Set Cnn_Music = New ADODB.Connection
                With Cnn_Music
                    .Provider = "Microsoft.Jet.OLEDB.4.0"
                    .Properties("Data Source") = App.Path & "\Library\Library.mdb"
                    '.CursorLocation = adUseClient
                    .Open
                End With
                'Selecionar a tabela de música
                SQL_Music = "SELECT * FROM Tabela_Musica"
                RS_Music.Open SQL_Music, Cnn_Music, adOpenDynamic, adLockOptimistic
                Cnn_Music.Execute "Delete * from Tabela_Musica"
            End If
        
            If Biblioteca_Selecionada = "Filmes" Then
                'Conectar á base de dados -> Filmes
                Set Rs_Film = New ADODB.Recordset
                Set Cnn_Film = New ADODB.Connection
                With Cnn_Film
                    .Provider = "Microsoft.Jet.OLEDB.4.0"
                    .Properties("Data Source") = App.Path & "\Library\Library.mdb"
                    .CursorLocation = adUseClient
                    .Open
                End With
                'Selecionar a tabela de filmes
                SQL_Film = "SELECT * FROM Tabela_Filmes"
                Rs_Film.Open SQL_Film, Cnn_Film, adOpenDynamic, adLockOptimistic
                Cnn_Film.Execute "Delete * from Tabela_Filmes"
            End If
        
            'Chamar o procedimento para iniciar a pesquisa da media
            Pesquisar_Ficheiro
            
        Else
            'Chamar o procedimento
            Parar_Pesquisa_e_Adicionar
        End If
    End If
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Importar
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Importar
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Importar
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Importar
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Importar
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Importar
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Pesquisar_Musica.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text_Pesquisar_Musica.ScaleWidth) + (2 * Opcao_Sim.left) + 20 + Opcao_Sim.left + Pic_Sim.ScaleWidth + 2)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + 16 + 16 + Opcao_Sim.Height + 3 + Label_Informacao.Height + 3 + (2 * Barra_Text_Pesquisar_Musica.ScaleHeight) + Opcao_Nao.Height + Fundo_Frame_Botoes.Height + Opcao_Nao.top + (2 * Opcao_Sim.left))
    End With
    
    Ajustar_Formulario Form_Importar, False, False, True, True
    
    Ajustar_Botao Form_Importar, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Importar, Botao_Iniciar_Pesquisa, Label_Iniciar_Pesquisa, True, Contorno_Iniciar_Pesquisa
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Iniciar_Pesquisa
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With
    
    Ajustar_Caixa_Texto Barra_Text_Pesquisar_Musica, Text_Pesquisar_Musica, Contorno_Pesquisar_Musica, False
    
    With Botao_Pesquisar
        .Height = Form_Skin.Botao_Pesquisar.Height
        .top = (Barra_Text_Pesquisar_Musica.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Pesquisar.Width
        .left = Barra_Text_Pesquisar_Musica.ScaleWidth - .ScaleWidth - .top
    End With
    
    Ajustar_Option Pic_Nao
    Ajustar_Option Pic_Sim
    
    With Opcao_Sim
        .left = 16
        .top = 2 * .left
    End With

    With Pic_Sim
        .top = Opcao_Sim.top
        .left = Opcao_Sim.left
    End With

    With Label_Informacao
        .top = Opcao_Sim.top + Opcao_Sim.Height + 3
        .Width = Opcao_Sim.Width - Pic_Sim.ScaleWidth - 2
        .left = Opcao_Sim.left + Pic_Sim.ScaleWidth + 2
    End With

    With Barra_Text_Pesquisar_Musica
        .Height = Form_Skin.Caixa_de_Texto.Height
        .top = Label_Informacao.top + Label_Informacao.Height + 3
        .Width = Form_Skin.Caixa_de_Texto.Width
        .left = Label_Informacao.left
    End With

    With Opcao_Nao
        .top = Barra_Text_Pesquisar_Musica.top + Barra_Text_Pesquisar_Musica.ScaleHeight + (Barra_Text_Pesquisar_Musica.Height)
        .Width = Opcao_Sim.Width
        .left = Opcao_Sim.left
    End With

    With Pic_Nao
        .top = Opcao_Nao.top
        .left = Opcao_Nao.left
    End With
    
    With Label_Contador
        .Width = Barra_Text_Pesquisar_Musica.ScaleWidth
        .left = Barra_Text_Pesquisar_Musica.left
    End With
    
    With Label_Ficheiro
        .Width = Barra_Text_Pesquisar_Musica.ScaleWidth
        .left = Barra_Text_Pesquisar_Musica.left
    End With
    
    With Label_Adicionando
        .top = Label_Ficheiro.top
        .Width = Barra_Text_Pesquisar_Musica.ScaleWidth
        .left = Barra_Text_Pesquisar_Musica.left
    End With
    
    With ProgressBar1
        .top = Label_Adicionando.top + Label_Adicionando.Height + 3
        .Width = Barra_Text_Pesquisar_Musica.ScaleWidth
        .left = Barra_Text_Pesquisar_Musica.left
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Opcao_Nao_Click()
    'Activar opcao
    Pic_Sim.Picture = Form_Skin.Opcao_Normal.Picture
    Pic_Nao.Picture = Form_Skin.Opcao_Over.Picture
    Label_Iniciar_Pesquisa.Caption = Idioma_Concluir
    Botao_Iniciar_Pesquisa.Enabled = True
    Label_Iniciar_Pesquisa.Enabled = True
End Sub

Private Sub Opcao_Sim_Click()
    'Activar opcao
    Pic_Sim.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Nao.Picture = Form_Skin.Opcao_Normal.Picture
    Label_Iniciar_Pesquisa.Caption = Idioma_Iniciar_Pesquisa
    
    If Text_Pesquisar_Musica.Text = Empty Then
        Botao_Iniciar_Pesquisa.Enabled = False
        Label_Iniciar_Pesquisa.Enabled = False
    Else
        Botao_Iniciar_Pesquisa.Enabled = True
        Label_Iniciar_Pesquisa.Enabled = True
    End If
End Sub

Private Sub Opcao_Sim_GotFocus()
    'Activar opcao
    Pic_Sim.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Nao.Picture = Form_Skin.Opcao_Normal.Picture
    Label_Iniciar_Pesquisa.Caption = Idioma_Iniciar_Pesquisa
    
    If Text_Pesquisar_Musica.Text = Empty Then
        Botao_Iniciar_Pesquisa.Enabled = False
        Label_Iniciar_Pesquisa.Enabled = False
    Else
        Botao_Iniciar_Pesquisa.Enabled = True
        Label_Iniciar_Pesquisa.Enabled = True
    End If
End Sub

Private Sub Pic_Nao_Click()
    'Activar opcao
    Opcao_Nao.Value = True
    Pic_Sim.Picture = Form_Skin.Opcao_Normal.Picture
    Pic_Nao.Picture = Form_Skin.Opcao_Over.Picture
    Label_Iniciar_Pesquisa.Caption = Idioma_Concluir
    Botao_Iniciar_Pesquisa.Enabled = True
    Label_Iniciar_Pesquisa.Enabled = True
End Sub

Private Sub Pic_Sim_Click()
    'Activar opcao
    Opcao_Sim.Value = True
    Pic_Sim.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Nao.Picture = Form_Skin.Opcao_Normal.Picture
    Label_Iniciar_Pesquisa.Caption = Idioma_Iniciar_Pesquisa
    
    If Text_Pesquisar_Musica.Text = Empty Then
        Botao_Iniciar_Pesquisa.Enabled = False
        Label_Iniciar_Pesquisa.Enabled = False
    Else
        Botao_Iniciar_Pesquisa.Enabled = True
        Label_Iniciar_Pesquisa.Enabled = True
    End If
End Sub

Private Sub Botao_Pesquisar_Click()
    'Abrir explorador de pastas
    On Error Resume Next
    With Explorador
        .BrowseFolder "Selecionar pasta", 0, False
        If .cancel = True Then Exit Sub
        Text_Pesquisar_Musica.Text = .Path
        Dir1.Path = Text_Pesquisar_Musica.Text & "\"
        Label_Contador.Caption = ""
        Botao_Iniciar_Pesquisa.Enabled = True
        Label_Iniciar_Pesquisa.Enabled = True
    End With
End Sub

Private Sub Parar_Pesquisa_e_Adicionar()
    'Procedimento para para a pesquisa e adicionar as músicas encontradas
    Pesquisando = False
    Botao_Iniciar_Pesquisa.Enabled = False
    Label_Iniciar_Pesquisa.Enabled = False
    Adicionar_Ficheiros
    Me.MousePointer = 11
End Sub

Private Sub Text_Pesquisar_Musica_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Pesquisar_Musica.Visible = True
End Sub

Private Sub Text_Pesquisar_Musica_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Pesquisar_Musica.Visible = False
End Sub
