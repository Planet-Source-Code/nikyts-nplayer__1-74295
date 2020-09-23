VERSION 5.00
Begin VB.Form Form_Actualizar_Biblioteca 
   Appearance      =   0  'Flat
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Actualizar_Biblioteca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   6600
      TabIndex        =   10
      Top             =   960
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
      Left            =   7800
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   7800
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   6600
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
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
      ScaleWidth      =   376
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2640
      Width           =   5640
      Begin VB.PictureBox Botao_Iniciar_Pesquisa 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2640
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   1
         Top             =   0
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
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   195
            TabIndex        =   8
            Top             =   45
            Width           =   1350
         End
      End
      Begin VB.Image Fundo_Frame_Botoes 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   6375
      Begin VB.PictureBox Barra_Text_Pesquisar_Musica 
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
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
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   30
            Width           =   1620
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
      Begin NPlayer.NProgressBar ProgressBar1 
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   450
      End
      Begin VB.Shape Shape_Centro 
         BorderColor     =   &H00212121&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label_Adicionando 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Adicionando..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2295
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label_Contador 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   5115
      End
      Begin VB.Label Label_Ficheiro 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   4995
      End
      Begin VB.Label Label_Pesquisar_Em 
         AutoSize        =   -1  'True
         BackColor       =   &H00313131&
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar em:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1230
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
      ScaleWidth      =   417
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   5880
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Nova biblioteca"
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
         TabIndex        =   3
         Top             =   120
         Width           =   1515
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
Attribute VB_Name = "Form_Actualizar_Biblioteca"
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

'Variável para saber que extensão é para procurar
Public Pesquisar_pela_Extensao As String

Private Sub Pesquisar_Ficheiro()
    'Procedimento para carregar a biblioteca
    On Error Resume Next
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
        'Label_Iniciar_Pesquisa = Idioma_Iniciar_Pesquisa
        Adicionar_Ficheiros
        
    Else
        'Chamar o procedimento
        Parar_Pesquisa_e_Adicionar
    End If
End Sub

Private Sub Botao_Iniciar_Pesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho das teclas
    If KeyCode = vbKeyReturn Then Label_Iniciar_Pesquisa_Click
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
        GoTo Concluir_Processo
    
    Else
        'Actualizar o directório onde o programa irá carregar os albuns existentes/ pastas dos artistas
        If Biblioteca_Selecionada = "Musica" Then
            Form_Principal.Text_Caminho.Text = Text_Pesquisar_Musica.Text & "\"
            Call WriteINI("Library", "Location_Of_Albuns", Text_Pesquisar_Musica.Text & "\", (Localizacao_Ficheiro_Preferencias))
            
            Form_Principal.Grelha_Musica.Clear
            Form_Principal.Grelha_Artista.Clear
            Form_Principal.Grelha_Genero.Clear
            Form_Principal.Grelha_Album.Clear
            Form_Principal.Formatar_Grelha_Musica Form_Principal.Grelha_Musica
            Form_Principal.Formatar_Grelha_Artista
            Form_Principal.Formatar_Grelha_Genero
            Form_Principal.Formatar_Grelha_Album
            Dim q As Integer: q = 1
        End If
        
        If Biblioteca_Selecionada = "Filmes" Then
            Form_Principal.Grelha_Filmes.Clear
            Form_Principal.Formatar_Grelha_Filmes
            Dim w As Integer: w = 1
        End If
        
        'Ler as tags do ficheiro, caso existam
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
               
               'Adicionar as músicas na base de dados e respectiva grelha---------------------------------------------------------------------------
                If Biblioteca_Selecionada = "Musica" Then
                    RS_Music.AddNew
                    RS_Music!ficheiro = List1.List(j)
                    RS_Music!Titulo = sTitle
                    RS_Music!Artista = sArtist
                    RS_Music!album = sAlbum
                    RS_Music!ano = sYear
                    RS_Music!Genero = sGenre
                    RS_Music!Comentario = sComment
                    RS_Music!Directorio = Mid(List1.List(j), 1, InStrRev(List1.List(j), "\"))
                    RS_Music!Classificacao = "0"
                    RS_Music!ID = j
                    RS_Music.Update
                    
                    Form_Principal.Grelha_Musica.Rows = Form_Principal.Grelha_Musica.Rows + 1
                    Form_Principal.Grelha_Musica.TextMatrix(q, 0) = List1.List(j)
                    Form_Principal.Grelha_Musica.TextMatrix(q, 1) = sTitle
                    Form_Principal.Grelha_Musica.TextMatrix(q, 2) = sArtist
                    Form_Principal.Grelha_Musica.TextMatrix(q, 3) = sAlbum
                    Form_Principal.Grelha_Musica.TextMatrix(q, 4) = sYear
                    Form_Principal.Grelha_Musica.TextMatrix(q, 5) = sGenre
                    Form_Principal.Grelha_Musica.TextMatrix(q, 6) = sComment
                    Form_Principal.Grelha_Musica.TextMatrix(q, 7) = Mid(List1.List(j), 1, InStrRev(List1.List(j), "\"))
                    Form_Principal.Grelha_Musica.TextMatrix(q, 8) = "0"
                    Form_Principal.Grelha_Musica.TextMatrix(q, 9) = j
                   
                    If j = List1.ListCount - 1 Then
                        Cnn_Music.Close
                        Set Cnn_Music = Nothing
                        Form_Principal.Verifica_Rs_Musica
                        Form_Principal.Rs_Musica.Open "select * from Tabela_Musica order by Titulo", Form_Principal.Cnn_Biblioteca
                        Form_Principal.Carregar_Grelha_Artista
                        Form_Principal.Carregar_Grelha_Genero
                        Form_Principal.Carregar_Grelha_Album
                        GoTo Concluir_Processo
                    Else
                        iProg = iProg + 1
                        .Value = iProg
                        q = q + 1
                    End If
               End If
               
               'Adicionar os filmes na base de dados e respectiva grelha---------------------------------------------------------------------------
                If Biblioteca_Selecionada = "Filmes" Then
                    Rs_Film.AddNew
                    Rs_Film!ficheiro = List1.List(j)
                    Rs_Film!Titulo = sTitle
                    Rs_Film!ano = sYear
                    Rs_Film!Genero = sGenre
                    Rs_Film!Comentario = sComment
                    Rs_Film!Directorio = Mid(List1.List(j), 1, InStrRev(List1.List(j), "\"))
                    Rs_Film!Classificacao = "0"
                    Rs_Film!ID = j
                    Rs_Film.Update
                    
                    Form_Principal.Grelha_Filmes.Rows = Form_Principal.Grelha_Filmes.Rows + 1
                    Form_Principal.Grelha_Filmes.TextMatrix(w, 0) = List1.List(j)
                    Form_Principal.Grelha_Filmes.TextMatrix(w, 1) = sTitle
                    Form_Principal.Grelha_Filmes.TextMatrix(w, 2) = sYear
                    Form_Principal.Grelha_Filmes.TextMatrix(w, 3) = sGenre
                    Form_Principal.Grelha_Filmes.TextMatrix(w, 4) = sComment
                    Form_Principal.Grelha_Filmes.TextMatrix(w, 5) = Mid(List1.List(j), 1, InStrRev(List1.List(j), "\"))
                    Form_Principal.Grelha_Filmes.TextMatrix(w, 6) = "0"
                    Form_Principal.Grelha_Filmes.TextMatrix(w, 7) = j
                   
                    If j = List1.ListCount - 1 Then
                        Cnn_Film.Close
                        Set Cnn_Film = Nothing
                        GoTo Concluir_Processo
                    Else
                        iProg = iProg + 1
                        .Value = iProg
                        w = w + 1
                    End If
               End If
            Next j
        End With
    End If

    
Exit Sub
Concluir_Processo:
ProgressBar1.Visible = False
'Form_Principal.Menu_Ficheiro_Click (1)
Form_Principal.Verificar_Contador
If Biblioteca_Selecionada = "Musica" Then Form_Principal.Criar_Album_Musica
With Form_Principal
    If .Grelha_Reproduzida = .Grelha_Musica Or .Grelha_Reproduzida = .Grelha_Filmes Then
        If .Grelha_Reproduzida.Rows > 1 Then
            .Musica_Linha_Pressionada = 1
            .Musica_Linha_Selecionada = 1
            .Faixa_em_Reproducao = .Grelha_Reproduzida.TextMatrix(.Grelha_Reproduzida.Row, 0)
            Call WriteINI("Library", "Playing_Track", "1", (Localizacao_Ficheiro_Preferencias))
        Else
            .Musica_Linha_Pressionada = 0
            .Musica_Linha_Selecionada = 0
            .Faixa_em_Reproducao = ""
            Call WriteINI("Library", "Playing_Track", "0", (Localizacao_Ficheiro_Preferencias))
        End If
    End If
End With
Me.MousePointer = 0
Unload Me
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar formulário, Fechar as bases de dados
    Unload Me
End Sub

Private Sub Botao_Iniciar_Pesquisa_GotFocus()
    'Ao receber o focus no botao
    Contorno_Iniciar_Pesquisa.Visible = True
End Sub

Private Sub Botao_Iniciar_Pesquisa_LostFocus()
    'Ao receber o focus no botao
    Contorno_Iniciar_Pesquisa.Visible = False
End Sub

Private Sub Combo_Drives_Click()
   sLastPath = "Todas as pastas"
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
    End With
End Sub

Private Sub Botao_Iniciar_Pesquisa_Click()
    'Atalho para
    Label_Iniciar_Pesquisa_Click
End Sub

Private Sub Parar_Pesquisa_e_Adicionar()
    'Procedimento para para a pesquisa e adicionar as músicas encontradas
    Pesquisando = False
    Botao_Iniciar_Pesquisa.Enabled = False
    Label_Iniciar_Pesquisa.Enabled = False
    Adicionar_Ficheiros
    Me.MousePointer = 11
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    'Iniciar o formulário
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    Label_Iniciar_Pesquisa.Caption = Idioma_Iniciar_Pesquisa
    
    'Variáveis para poder mover o formulário
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
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento
    Desenhar_Formulario
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Frame_Centro.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Centro.BorderColor = .Cor_Contorno_Frame_Centro.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Contorno_Pesquisar_Musica.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Pesquisar_Em.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Contador.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Ficheiro.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Adicionando.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Text_Pesquisar_Musica.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Pesquisar_Musica.ForeColor = .Cor_Letra_Textbox.backcolor
        Botao_Pesquisar.Picture = .Botao_Pesquisar.Picture
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Iniciar_Pesquisa.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Iniciar_Pesquisa.Picture = .Pic_Button.Picture
        Contorno_Iniciar_Pesquisa.BorderColor = .Cor_Contorno_Caixas.backcolor
        Barra_Text_Pesquisar_Musica.backcolor = .Cor_Fundo_Textbox.backcolor
        'Barra_Text_Pesquisar_Musica.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Pesquisar_Musica.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Pesquisar_Musica.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Pesquisar_Musica.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Pesquisar_Musica.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Pesquisar_Musica.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
    End With
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("New_Library", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("New_Library", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Pesquisar_Em.Caption = ReadINI("New_Library", "Label_Search", Localizacao_Ficheiro_Lingua)
    Idioma_Iniciar_Pesquisa = ReadINI("New_Library", "Button_Search_1", Localizacao_Ficheiro_Lingua)
    Idioma_Parar_Pesquisa = ReadINI("New_Library", "Button_Search_2", Localizacao_Ficheiro_Lingua)
    Botao_Pesquisar.ToolTipText = ReadINI("New_Library", "Button_Find", Localizacao_Ficheiro_Lingua)
    Idioma_Ficheiros = ReadINI("New_Library", "Info_Files", Localizacao_Ficheiro_Lingua)
    Idioma_Adicionando_1 = ReadINI("New_Library", "Info_Add_1", Localizacao_Ficheiro_Lingua)
    Idioma_Adicionando_2 = ReadINI("New_Library", "Info_Add_2", Localizacao_Ficheiro_Lingua)
    Idioma_Operacao_Cancelada = ReadINI("New_Library", "Operation_Canceled", Localizacao_Ficheiro_Lingua)
End Sub

Private Sub Label_Iniciar_Pesquisa_Click()
    'Iniciar a pesquisa de ficheiros
    If Text_Pesquisar_Musica.Text = Empty Then Exit Sub
    If Pesquisando = False Then
        Pesquisando = True
        Me.MousePointer = 11
        Label_Iniciar_Pesquisa.Caption = Idioma_Parar_Pesquisa
                   
        'Limpar dados
        With Form_Principal
            .Label_Faixa.Caption = ""
            .Faixa_em_Reproducao = ""
            .Tempo_Estimado.Caption = "00:00"
            .Label_Duracao.Caption = "00:00"
            .Label_Contador.Caption = ""
            .Text_Classificacao = "0"
            .Verificar_Classificacao
            .Text_Pesquisar_Musica.Text = ""
        End With
        With Form_Mini_Player
            .Label_Faixa.Caption = ""
            .Tempo_Estimado.Caption = "00:00"
            .Label_Duracao.Caption = "00:00"
        End With
        With Form_PopUp
            .Label_Artista.Caption = ""
            .Label_Contador.Caption = ""
            .Label_Faixa.Caption = ""
        End With
        
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
            'Apagar todos os dados da tabela correspondente
            With Form_Principal
                .Cnn_Biblioteca.Execute "Delete * from Tabela_Musica"
                .Grelha_Musica.Clear
                .Grelha_Musica.Rows = 1
                .Formatar_Grelha_Musica .Grelha_Musica
                
                .Grelha_Artista.Clear
                .Grelha_Artista.Rows = 1
                .Formatar_Grelha_Artista
                .Grelha_Genero.Clear
                .Grelha_Genero.Rows = 1
                .Formatar_Grelha_Genero
                .Grelha_Album.Clear
                .Grelha_Album.Rows = 1
                .Formatar_Grelha_Album
                
                .Apagar_Album_Musica
            End With
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
            'Apagar todos os dados da tabela correspondente
            With Form_Principal
                .Cnn_Biblioteca.Execute "Delete * from Tabela_Filmes"
                .Grelha_Filmes.Clear
                .Grelha_Filmes.Rows = 1
                .Formatar_Grelha_Filmes
            End With
        End If
    
        'Chamar o procedimento para iniciar a pesquisa da media
        Pesquisar_Ficheiro
        
    Else
        'Chamar o procedimento
        Parar_Pesquisa_e_Adicionar
    End If
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Actualizar_Biblioteca
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Actualizar_Biblioteca
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Actualizar_Biblioteca
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Actualizar_Biblioteca
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Actualizar_Biblioteca
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Actualizar_Biblioteca
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    Barra_Text_Pesquisar_Musica.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text_Pesquisar_Musica.ScaleWidth) + (2 * Barra_Text_Pesquisar_Musica.left) + 20)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Label_Pesquisar_Em.left + Label_Pesquisar_Em.Height + 3 _
                + Form_Skin.Caixa_de_Texto.Height + 3 + Label_Contador.Height + 3 + Label_Ficheiro.Height + Fundo_Frame_Botoes.Height _
                + (2 * Label_Pesquisar_Em.left))
    End With
    
    Ajustar_Formulario Form_Actualizar_Biblioteca, False, False, True, True
    
    Ajustar_Botao Form_Actualizar_Biblioteca, Botao_Iniciar_Pesquisa, Label_Iniciar_Pesquisa, True, Contorno_Iniciar_Pesquisa
    With Botao_Iniciar_Pesquisa
        .left = (Frame_Botoes.ScaleWidth - .ScaleWidth) / 2
    End With
    
    Ajustar_Caixa_Texto Barra_Text_Pesquisar_Musica, Text_Pesquisar_Musica, Contorno_Pesquisar_Musica, False
    
    With Label_Contador
        .Width = Barra_Text_Pesquisar_Musica.ScaleWidth
        .left = Barra_Text_Pesquisar_Musica.left
    End With
    
    With Label_Ficheiro
        .Width = Barra_Text_Pesquisar_Musica.ScaleWidth
        .left = Barra_Text_Pesquisar_Musica.left
    End With
    
    With Botao_Pesquisar
        .Height = Form_Skin.Botao_Pesquisar.Height
        .top = (Barra_Text_Pesquisar_Musica.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Pesquisar.Width
        .left = Barra_Text_Pesquisar_Musica.ScaleWidth - .ScaleWidth - .top
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

Private Sub Text_Pesquisar_Musica_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Pesquisar_Musica.Visible = True
End Sub

Private Sub Text_Pesquisar_Musica_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Pesquisar_Musica.Visible = False
End Sub
