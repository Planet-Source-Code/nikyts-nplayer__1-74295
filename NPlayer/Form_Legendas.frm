VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD - Copy.OCX"
Begin VB.Form Form_Legendas 
   Appearance      =   0  'Flat
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   3420
   ClientLeft      =   90
   ClientTop       =   0
   ClientWidth     =   6945
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
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Legendas on-line"
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
         TabIndex        =   2
         Top             =   120
         Width           =   1680
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
   Begin VB.PictureBox Frame_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   416
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2520
      Width           =   6240
      Begin VB.PictureBox Botao_Download 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   4080
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   3
         Top             =   120
         Width           =   1740
         Begin VB.Label Label_Download 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transferir"
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   450
            TabIndex        =   4
            Top             =   45
            Width           =   840
         End
         Begin VB.Shape Contorno_Download 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Image Fundo_Frame_Botoes 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Legendas 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   3223857
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00404040&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Legendas"
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

'Declaração das variáveis
'Option Explicit
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'Api para fazer o download das legendas
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    'Função para fazer o download das legendas
    On erro GoTo Corrige_Erro
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then
        DownloadFile = True 'carrega  a legenda do filme
        With Form_Principal
            'Capturar a posicao actual do slide e a posicao do wmp
            Dim posicao_wmp As Long: posicao = .Wmp.Controls.CurrentPosition
            Dim posicao_slide As Integer: posicao_slide = .Slide.left
            .Parar_o_Player
            
            'Repor a ultima posicao do slide e a posicao do wmp
            .Slide.left = posicao_slide
            .Image_Progresso.Width = .Slide.left
            .Slide_Mini.left = .Slide.left
            Form_Mini_Player.Slide.left = .Slide.left
            Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
            
            Dim offseti As Single
            offseti = (.Slide.left - Form_Principal.Image_Barra_Slide.left - 3) / (Form_Principal.Image_Barra_Slide.Width - 10 - .Slide.Width)
            .Wmp.Controls.CurrentPosition = Int(.Wmp.currentMedia.Duration * offseti)
            Form_Wmp.Wmp.Controls.CurrentPosition = Int(.Wmp.currentMedia.Duration * offseti)
            Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
            .Botao_Play_Click
        End With
        Unload Me
    Else
        'MsgBox "Ocorreu um erro durante a conexão. " & vbNewLine & Grelha_Legendas.TextMatrix(Grelha_Legendas.Row, 0)
        Unload Me
    End If
    
    Me.MousePointer = 0
    Label_Download.Enabled = True
    Botao_Download.Enabled = True
    Unload Me
    
Exit Function
Corrige_Erro:
Unload Me
End Function

Private Sub Botao_Download_Click()
    'Atalho para
    Label_Download_Click
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar a aplicação
    Unload Me
End Sub

Private Sub Form_Load()
    'Propriedades inicais do formulário
    Carregar_Idioma
    Carregar_Skin
    Desenhar_Formulario
    Arredondar_Cantos_do_Form Me, True
    
    Formatar_Grelha_Legendas
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Resize()
    'Atalho para
    Desenhar_Formulario
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Subtitles", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Subtitles", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Download.Caption = ReadINI("Subtitles", "Button_Download", Localizacao_Ficheiro_Lingua)
End Sub

Private Sub Formatar_Grelha_Legendas()
    'Procedimento para formatar a grelha
    With Grelha_Legendas
        .AllowUserResizing = flexResizeColumns
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, 0) = "Localização"
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = ReadINI("Subtitles", "Grid_Col_1", Localizacao_Ficheiro_Lingua)
        .ColWidth(1) = 3600
        .ColAlignment(1) = vbleft
        .TextMatrix(0, 2) = ReadINI("Subtitles", "Grid_Col_2", Localizacao_Ficheiro_Lingua)
        .ColWidth(2) = 1500
        .ColAlignment(2) = vbleft
        .TextMatrix(0, 3) = ReadINI("Subtitles", "Grid_Col_3", Localizacao_Ficheiro_Lingua)
        .ColWidth(3) = 1000
        .ColAlignment(3) = vbleft
        .TextMatrix(0, 4) = ReadINI("Subtitles", "Grid_Col_4", Localizacao_Ficheiro_Lingua)
        .ColWidth(4) = 1500
        .ColAlignment(4) = vbleft
    End With
End Sub

Private Sub Label_Download_Click()
    'Tranferir as legendas e aplica-las ao video
    Dim directorio_do_filme, legenda_do_filme, ficheiro, extensao As String
    
    ficheiro = Form_Principal.Grelha_Filmes.TextMatrix(Form_Principal.Grelha_Filmes.Row, 0)
    directorio_do_filme = Mid(ficheiro, 1, InStrRev(ficheiro, "\"))
    extensao = "." & Grelha_Legendas.TextMatrix(Grelha_Legendas.Row, 3)
    legenda_do_filme = Grelha_Legendas.TextMatrix(Grelha_Legendas.Row, 1) & extensao
    
    If Dir$(directorio_do_filme & legenda_do_filme) <> "" Then
        Mensagem_de_Aviso "Question", "Já existe um ficheiro com este nome." & vbNewLine & "Pretende subtitui-lo?"
        If Resposta = "Sim" Then
            Me.MousePointer = 11
            Label_Download.Enabled = False
            Botao_Download.Enabled = False
            Kill (directorio_do_filme & legenda_do_filme)
            ret = DownloadFile(Grelha_Legendas.TextMatrix(Grelha_Legendas.Row, 0), directorio_do_filme & legenda_do_filme)
        End If
    Else
        Me.MousePointer = 11
        Label_Download.Enabled = False
        Botao_Download.Enabled = False
        ret = DownloadFile(Grelha_Legendas.TextMatrix(Grelha_Legendas.Row, 0), directorio_do_filme & legenda_do_filme)
    End If
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Legendas
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Legendas
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Legendas
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Legendas
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Legendas
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Legendas
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Grelha_Legendas.Width = Form_Skin.Frame_Componentes.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Grelha_Legendas.Width) + (2 * Grelha_Legendas.left) + 60)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + 16 + Form_Skin.Frame_Componentes.Height + Fundo_Frame_Botoes.Height _
                + Grelha_Legendas.left)
    End With
    
    Ajustar_Formulario Form_Legendas, False, False, False, True
    
    Ajustar_Botao Form_Legendas, Botao_Download, Label_Download, True, Contorno_Download
    
    With Botao_Download
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    
    With Grelha_Legendas
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Height = Me.ScaleHeight - Fundo_Barra_ControlBox.Height - Fundo_Frame_Botoes.Height - 2
        .left = 1
        .Width = Me.ScaleWidth - 3
    End With
        
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Download.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Download.Picture = .Pic_Button.Picture
        Contorno_Download.BorderColor = .Cor_Contorno_Caixas.backcolor
        Grelha_Legendas.backcolor = .Cor_Grid_BackColor.backcolor
        Grelha_Legendas.BackColorBkg = .Cor_Grid_BackColorBkg.backcolor
        Grelha_Legendas.BackColorFixed = .Cor_Grid_BackColorFixed.backcolor
        Grelha_Legendas.BackColorSel = .Cor_Grid_BackColorSel.backcolor
        Grelha_Legendas.ForeColor = .Cor_Grid_ForeColor.backcolor
        Grelha_Legendas.ForeColorFixed = .Cor_Grid_ForeColorFixed.backcolor
        Grelha_Legendas.ForeColorSel = .Cor_Grid_ForeColorSel.backcolor
        Grelha_Legendas.GridColor = .Cor_Grid_Color.backcolor
        Grelha_Legendas.GridColorFixed = .Cor_Grid_ColorFixed.backcolor
    End With
End Sub
