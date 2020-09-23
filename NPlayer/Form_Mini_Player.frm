VERSION 5.00
Begin VB.Form Form_Mini_Player 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "NPlayer"
   ClientHeight    =   4560
   ClientLeft      =   9180
   ClientTop       =   2325
   ClientWidth     =   6090
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
   Icon            =   "Form_Mini_Player.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   406
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   5655
      Begin VB.PictureBox pichook 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3360
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "NPlayer - Nikyts software"
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
         Left            =   465
         TabIndex        =   15
         Top             =   120
         Width           =   2505
      End
      Begin VB.Image Icon_do_Programa 
         Enabled         =   0   'False
         Height          =   210
         Left            =   75
         Picture         =   "Form_Mini_Player.frx":57E2
         Top             =   60
         Width           =   210
      End
      Begin VB.Image Botao_Tray 
         Height          =   135
         Left            =   4440
         ToolTipText     =   "Colocar o ícone na bandeja"
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Botao_Restaurar 
         Height          =   135
         Left            =   4920
         ToolTipText     =   "Restaurar"
         Top             =   120
         Width           =   135
      End
      Begin VB.Image Botao_Minimizar 
         Height          =   135
         Left            =   4680
         ToolTipText     =   "Minimizar"
         Top             =   120
         Width           =   135
      End
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   5160
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
   Begin VB.PictureBox Barra_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00B1BEB6&
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
      Height          =   3675
      Left            =   0
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   380
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5700
      Begin VB.PictureBox Picture_Slide_Som 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2700
         Picture         =   "Form_Mini_Player.frx":5A8C
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   110
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2820
         Width           =   1650
         Begin VB.TextBox Text_Slide_Som 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   720
            TabIndex        =   12
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.PictureBox Slide_Som 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   210
         End
      End
      Begin VB.PictureBox Barra_Faixa 
         Appearance      =   0  'Flat
         BackColor       =   &H0082908E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1770
         Left            =   360
         ScaleHeight     =   118
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   349
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   5235
         Begin VB.PictureBox SliderBar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Height          =   150
            Left            =   960
            ScaleHeight     =   10
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   217
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   1440
            Width           =   3255
            Begin VB.PictureBox Image_Progresso 
               Appearance      =   0  'Flat
               BackColor       =   &H00212121&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   135
               Left            =   0
               ScaleHeight     =   9
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   2
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   30
            End
            Begin VB.PictureBox Slide 
               Appearance      =   0  'Flat
               BackColor       =   &H00313131&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   150
               Left            =   0
               ScaleHeight     =   10
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   10
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Image Image_Barra_Slide 
               Enabled         =   0   'False
               Height          =   150
               Left            =   0
               Picture         =   "Form_Mini_Player.frx":6F8E
               Stretch         =   -1  'True
               Top             =   0
               Visible         =   0   'False
               Width           =   3255
            End
         End
         Begin VB.Label Label_Lista 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Lista"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   4695
            TabIndex        =   7
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label_Video 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Video"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   4680
            TabIndex        =   6
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label_Faixa 
            BackColor       =   &H0082908E&
            BackStyle       =   0  'Transparent
            Caption         =   "Eminem - No apologies"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   180
            Width           =   4935
         End
         Begin VB.Label Label_Duracao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H0082908E&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   4650
            TabIndex        =   3
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Tempo_Estimado 
            AutoSize        =   -1  'True
            BackColor       =   &H0082908E&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   1440
            Width           =   480
         End
      End
      Begin VB.Image Botao_Mudo 
         Height          =   150
         Left            =   2400
         ToolTipText     =   "Mudo"
         Top             =   2865
         Width           =   180
      End
      Begin VB.Image Botao_Pausa 
         Height          =   810
         Left            =   720
         ToolTipText     =   "Pausa"
         Top             =   2580
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Image Botao_Seguinte 
         Height          =   585
         Left            =   1440
         ToolTipText     =   "Faixa seguinte"
         Top             =   2640
         Width           =   540
      End
      Begin VB.Image Botao_Antes 
         Height          =   600
         Left            =   240
         ToolTipText     =   "Faixa anterior"
         Top             =   2640
         Width           =   555
      End
      Begin VB.Image Botao_Play 
         Height          =   810
         Left            =   4680
         ToolTipText     =   "Reproduzir"
         Top             =   2580
         Width           =   765
      End
      Begin VB.Image Fundo_Barra_Player 
         Enabled         =   0   'False
         Height          =   525
         Left            =   0
         Top             =   0
         Width           =   465
      End
   End
End
Attribute VB_Name = "Form_Mini_Player"
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

'Declaração de variáveis
'VARIÁVERIS DO SLIDER VIDEO
Dim tx As Integer, Ty As Integer, DN As Boolean
Dim Txa As Integer, DNa As Boolean
Dim Tyb, DNb As Boolean
Dim NewLeft As Integer

'VARIÁVERIS DO SLIDER SOM
Dim TX_Som As Integer, Ty_Som As Integer, DN_Som As Boolean
Dim Txa_Som As Integer, DNa_Som As Boolean
Dim Tyb_Som, Dnb_Som As Boolean
Dim NewLeft_Som As Integer

'tray icon
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA

'Variável para verificar se está em modo hide (stray icon)
Dim Modo_Tray As Boolean

'Declaração das variáveis
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'Variáveis do idioma
Dim Idioma_Ouvir_musica As String

Private Sub Barra_ControlBox_DblClick()
    'Atalho para
    Label_Titulo_DblClick
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Mini_Player
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Mini_Player
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Mini_Player
End Sub

Private Sub Botao_Antes_Click()
    'Atalho para
    Form_Principal.Botao_Antes_Click
    
    'Caso esteja na primeira linha não avança mais
    If Form_Lista.Grelha_Lista_Em_Reproducao.Row = 1 Then
        Exit Sub
    Else
        'Selecionar a linha seguinte
        With Form_Lista.Grelha_Lista_Em_Reproducao
            .Row = .Row - 1
            .Col = 0
            .ColSel = .Cols - 1 'Selecionar a linha por inteiro
        End With
    End If
End Sub

Private Sub Botao_Antes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Antes.Picture = Form_Skin.Botao_Antes_Down.Picture
End Sub

Private Sub Botao_Antes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Antes.Picture = Form_Skin.Botao_Antes_Normal.Picture
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar o programa
    Form_Principal.Botao_Fechar_Click
End Sub

Private Sub Botao_Minimizar_Click()
    'Minimizar o formulário
    If Form_Preferencias.Check_Tray.Value = 1 Then
        Botao_Tray_Click
    Else
        Me.WindowState = 1
        Form_Wmp.Hide
        Form_Lista.Hide
        Label_Video.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
        Label_Lista.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
    End If
End Sub

Private Sub Botao_Mudo_Click()
    'Atalho para
    Form_Principal.Botao_Mudo_Click
End Sub

Private Sub Botao_Pausa_Click()
    'Talho para
    Form_Principal.Botao_Pausa_Click
End Sub

Private Sub Botao_Pausa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Pausa.Picture = Form_Skin.Botao_Pausa_Down.Picture
End Sub

Private Sub Botao_Pausa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Pausa.Picture = Form_Skin.Botao_Pausa_Normal.Picture
End Sub

Private Sub Botao_Play_Click()
    'Talho para
    Form_Principal.Botao_Play_Click
End Sub

Private Sub Botao_Play_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Play.Picture = Form_Skin.Botao_Play_Down.Picture
End Sub

Private Sub Botao_Play_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Play.Picture = Form_Skin.Botao_Play_Normal.Picture
End Sub

Private Sub Botao_Restaurar_Click()
    'Ver biblioteca
    Form_Principal.Show
    Me.Hide
    Form_Wmp.Hide
    Form_Lista.Hide
    Label_Video.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
    Label_Lista.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
End Sub

Private Sub Botao_Seguinte_Click()
    'Atalho para
    Form_Principal.Botao_Seguinte_Click
        
    'Caso esteja na última linha não avança mais
    If Form_Lista.Grelha_Lista_Em_Reproducao.Row = Form_Lista.Grelha_Lista_Em_Reproducao.Rows - 1 Then
        Exit Sub
    Else
        'Selecionar a linha seguinte
        With Form_Lista.Grelha_Lista_Em_Reproducao
            .Row = .Row + 1
            .Col = 0
            .ColSel = .Cols - 1 'Selecionar a linha por inteiro
        End With
    End If
End Sub

Private Sub Botao_Seguinte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Seguinte.Picture = Form_Skin.Botao_Seguinte_Down.Picture
End Sub

Private Sub Botao_Seguinte_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Seguinte.Picture = Form_Skin.Botao_Seguinte_Normal.Picture
End Sub

Private Sub Botao_Tray_Click()
    'Mensagem no icon do projecto/ coloca-lo ao lado do clock
    t.cbSize = Len(t)
    t.hWnd = pichook.hWnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
    t.szTip = "NPlayer" & Chr$(10) 'Texto a ser exibido no icon
    Shell_NotifyIcon NIM_ADD, t
    App.TaskVisible = False
    
    'Colocar o icon do formulário ao lado do clock do windows
    Me.Hide
    Form_Wmp.Hide
    Form_Lista.Hide
    Label_Video.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
    Label_Lista.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
        
    Modo_Tray = True
End Sub

Private Sub Form_Activate()
    Form_Wmp.Wmp.settings.mute = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Botao_Fechar_Click
End Sub

Private Sub Form_Load()
    'Variáveis para poder mover o formulário
    Carregar_Idioma
    Carregar_Skin
    Desenhar_Formulario
    
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, False
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
        
    Botao_Tray.ToolTipText = ReadINI("Mini_Player", "Button_Tray", Localizacao_Ficheiro_Lingua)
    Botao_Minimizar.ToolTipText = ReadINI("Mini_Player", "Button_Minimize", Localizacao_Ficheiro_Lingua)
    Botao_Restaurar.ToolTipText = ReadINI("Mini_Player", "Button_Restore", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Mini_Player", "Button_Close", Localizacao_Ficheiro_Lingua)
    Botao_Mudo.ToolTipText = ReadINI("Mini_Player", "Button_Mute_On", Localizacao_Ficheiro_Lingua)
    Botao_Play.ToolTipText = ReadINI("Mini_Player", "Button_Play", Localizacao_Ficheiro_Lingua)
    Botao_Pausa.ToolTipText = ReadINI("Mini_Player", "Button_Pause", Localizacao_Ficheiro_Lingua)
    Botao_Antes.ToolTipText = ReadINI("Mini_Player", "Button_Previous_Track", Localizacao_Ficheiro_Lingua)
    Botao_Seguinte.ToolTipText = ReadINI("Mini_Player", "Button_Next_Track", Localizacao_Ficheiro_Lingua)
    Label_Video.Caption = ReadINI("Mini_Player", "Label_Video_Display", Localizacao_Ficheiro_Lingua)
    Label_Lista.Caption = ReadINI("Mini_Player", "Label_Playlist", Localizacao_Ficheiro_Lingua)
    
    Idioma_Ouvir_musica = ReadINI("Mini_Player", "Button_Mute_Off", Localizacao_Ficheiro_Lingua)
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Fundo_Barra_Player.Picture = .Fundo_Mini_Player.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Botao_Restaurar.Picture = .Botao_Restaurar_Normal.Picture
        Botao_Minimizar.Picture = .Botao_Minimizar_Normal.Picture
        Botao_Tray.Picture = .Botao_Tray_Normal.Picture
        Label_Faixa.ForeColor = .Cor_Label_Barra_Visor.backcolor
        Tempo_Estimado.ForeColor = .Cor_Label_Barra_Visor.backcolor
        Label_Duracao.ForeColor = .Cor_Label_Barra_Visor.backcolor
        Label_Video.ForeColor = .Cor_Label_Barra_Visor.backcolor
        Label_Lista.ForeColor = .Cor_Label_Barra_Visor.backcolor
        Slide.Picture = .Slide_Musica_Normal.Picture
        SliderBar.Picture = .Image_Barra_Slide.Picture
        Botao_Play.Picture = .Botao_Play_Normal.Picture
        Botao_Pausa.Picture = .Botao_Pausa_Normal.Picture
        Botao_Antes.Picture = .Botao_Antes_Normal.Picture
        Botao_Seguinte.Picture = .Botao_Seguinte_Normal.Picture
        Botao_Mudo.Picture = .Som_On_Normal.Picture
        'Picture_Slide_Som.Picture = .Fundo_Slider_Volume.Picture
        Slide_Som.Picture = .Slide_Som_Normal.Picture
        Barra_Faixa.Picture = .Fundo_Barra_Faixa_Mini.Picture
        Image_Progresso.backcolor = .Cor_Slider_Music.backcolor
    End With
End Sub

Private Sub Barra_Botoes_DblClick()
    'Atalho para
    Botao_Restaurar_Click
End Sub

Private Sub Barra_Botoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Mini_Player
End Sub

Private Sub Barra_Botoes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Mini_Player
End Sub

Private Sub Barra_Botoes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Mini_Player
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    With Me
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Fundo_Barra_Player.Height)
        .Width = Screen.TwipsPerPixelX * (20 + Form_Skin.Fundo_Barra_Faixa_Mini.Width + 20)
    End With
    
    With Barra_ControlBox
        .Height = Fundo_Barra_ControlBox.Height
        .top = 0
        .Width = Me.ScaleWidth
        .left = 0
    End With
    
    With Fundo_Barra_ControlBox
        .Stretch = True
        .top = 0
        .Width = Barra_ControlBox.ScaleWidth
        .left = 0
    End With
    
    With Label_Titulo
        .top = (Barra_ControlBox.ScaleHeight - .Height) / 2
        .left = 26
    End With
    
    'Botões do controlbox
    Dim Ajustar_Botoes As String
    Ajustar_Botoes = "False" 'ReadINI("Dimensions", "Adjust_Button_ControlBox", Localizacao_Ficheiro_Skin)
    
    With Botao_Fechar
        If Ajustar_Botoes = "False" Then
            .top = (Barra_ControlBox.ScaleHeight - .Height) / 2
        Else
            .top = 0
        End If
        .left = Barra_ControlBox.ScaleWidth - .Width - 8
    End With
    
    With Botao_Restaurar
        .top = Botao_Fechar.top
        .left = Botao_Fechar.left - .Width - 8
    End With
    
    With Botao_Minimizar
        .top = Botao_Fechar.top
        .left = Botao_Restaurar.left - .Width - 8
    End With
    
    With Botao_Tray
        .top = Botao_Fechar.top
        .left = Botao_Minimizar.left - .Width - 8
    End With
    
    'Barra dos botoes e visor do player
    With Barra_Botoes
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Height = Fundo_Barra_Player.Height
        .left = 0
        .Width = Me.ScaleWidth
    End With
    
    With Fundo_Barra_Player
        .Stretch = True
        .Height = Barra_Botoes.ScaleHeight
        .top = 0
        .Width = Barra_Botoes.ScaleWidth
        .left = 0
    End With
    
    With Barra_Faixa
        .Height = Form_Skin.Fundo_Barra_Faixa_Mini.Height
        .top = 10
        .Width = Form_Skin.Fundo_Barra_Faixa_Mini.Width
        .left = 20
    End With
    
    With Label_Faixa
        .Width = Barra_Faixa.ScaleWidth - 16
        .left = 8
    End With
    
    With SliderBar
        .Height = Form_Skin.Image_Barra_Slide.Height
        .top = Barra_Faixa.ScaleHeight - .ScaleHeight - 10
        .Width = Form_Skin.Image_Barra_Slide_Mini.Width
        .left = (Barra_Faixa.ScaleWidth - .ScaleWidth) / 2
    End With
    
    With Slide
        .Height = Form_Skin.Slide_Musica_Normal.Height
        .top = (SliderBar.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Slide_Musica_Normal.Width
    End With
    
    With Image_Progresso
        .Height = Form_Skin.Slide_Musica_Normal.Height
        .top = (SliderBar.ScaleHeight - .ScaleHeight) / 2
    End With
    
    With Tempo_Estimado
        .top = Barra_Faixa.ScaleHeight - .Height - 10
        .left = 10
    End With
    
    With Label_Duracao
        .top = Tempo_Estimado.top
        .left = Barra_Faixa.ScaleWidth - .Width - 10
    End With
    
    With Label_Video
        .left = Barra_Faixa.ScaleWidth - .Width - 10
    End With
    
    With Label_Lista
        .left = Barra_Faixa.ScaleWidth - .Width - 10
    End With
    
    'Botões do player
    Dim Ajustar_Botoes_Player As String
    Ajustar_Botoes_Player = "False" 'ReadINI("Dimensions", "Adjust_Button_Player", Localizacao_Ficheiro_Skin)
    
    Dim espacamento As Integer: espacamento = Barra_Botoes.ScaleHeight - Barra_Faixa.ScaleHeight - 10
    With Botao_Antes
        .top = Barra_Faixa.top + Barra_Faixa.ScaleHeight + (espacamento - .Height) / 2
        .left = Barra_Faixa.left
    End With
    
    With Botao_Play
        .top = Botao_Antes.top
        If Ajustar_Botoes_Player = True Then
            .left = Botao_Antes.left + Botao_Antes.Width
        Else
            .left = Botao_Antes.left + Botao_Antes.Width + 6
        End If
    End With
    
    With Botao_Pausa
        .top = Botao_Antes.top
        .left = Botao_Play.left
    End With
    
    With Botao_Seguinte
        .top = Botao_Antes.top
        If Ajustar_Botoes_Player = True Then
            .left = Botao_Play.left + Botao_Play.Width
        Else
            .left = Botao_Play.left + Botao_Play.Width + 6
        End If
    End With
    
    'Barra do som
    With Botao_Mudo
        .top = Barra_Faixa.top + Barra_Faixa.ScaleHeight + (espacamento - .Height) / 2
    End With
    
    With Picture_Slide_Som
        .Height = Form_Skin.Fundo_Slider_Volume.Height
        .top = Barra_Faixa.top + Barra_Faixa.ScaleHeight + (espacamento - .Height) / 2
        .Width = Form_Skin.Fundo_Slider_Volume.Width
        .left = Botao_Mudo.left + Botao_Mudo.Width + 10
    End With
    
    With Slide_Som
        .Height = Form_Skin.Slide_Som_Normal.Height
        .top = (Picture_Slide_Som.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Slide_Som_Normal.Width
    End With
End Sub

Private Sub Form_Resize()
    Desenhar_Formulario
End Sub

Private Sub Label_Faixa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Caso o nome da faixa em reprodução seja demasiado grande que não caiba no visor de reprodução á a possibilidade de o
    'utilizador visualiza-la atraves do mouse
    If Label_Faixa.Caption <> "" Then Label_Faixa.ToolTipText = Label_Faixa.Caption
End Sub

Private Sub Label_Lista_Click()
    'Ver a lista de reprodução
    DoEvents
    If Label_Lista.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor Then
        With Form_Lista
            Me.MousePointer = 11
            .Grelha_Lista_Em_Reproducao.Clear
            .Grelha_Lista_Em_Reproducao.Rows = Form_Principal.Grelha_Reproduzida.Rows
            .Grelha_Lista_Em_Reproducao.Cols = 2 'Form_Principal.Grelha_Reproduzida.Cols
            Dim coluna As Integer: For coluna = 0 To 1 'Form_Principal.Grelha_Reproduzida.Cols - 1
                Dim Linha As Integer: For Linha = 0 To Form_Principal.Grelha_Reproduzida.Rows - 1
                    .Grelha_Lista_Em_Reproducao.TextMatrix(Linha, coluna) = Form_Principal.Grelha_Reproduzida.TextMatrix(Linha, coluna)
                    .Grelha_Lista_Em_Reproducao.ColWidth(coluna) = Form_Principal.Grelha_Reproduzida.ColWidth(coluna)
                Next
            Next
            .Grelha_Lista_Em_Reproducao.Row = Form_Principal.Grelha_Reproduzida.Row
            Me.MousePointer = 0
            .Show
        End With
        Label_Lista.ForeColor = Form_Skin.Cor_Contorno_Caixas.backcolor
        
    Else
        Form_Lista.Hide
        Label_Lista.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
    End If
End Sub

Private Sub Label_Titulo_DblClick()
    'Atalho para
    Botao_Restaurar_Click
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Mini_Player
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Mini_Player
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Mini_Player
End Sub

Private Sub Label_Video_Click()
    'Ver formulário da tela de video
    If Label_Video.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor Then
        Form_Wmp.Show
        Label_Video.ForeColor = Form_Skin.Cor_Contorno_Caixas.backcolor
        
    Else
        Form_Wmp.Hide
        Label_Video.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
    End If
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pichook é uma picture box, utilizada pelo Windows para reconhecer o ícone na barra de tarefas.
    Static rec As Boolean, Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                'Remover do sistema o icon do programa
                Form_Principal.Remover_Tray_Icon
    
            Case WM_LBUTTONDOWN:
                'Chamar o procedimento
                Form_Principal.Mostrar_Faixa_Musica_Formulario_Popup
                With Form_PopUp
                    .Show
                    .Tempo = 0
                    .Timer1.Enabled = True
                End With
                
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                'Ver o menu icon se for pressionado o botão direito
                'Me.PopupMenu Menu_Icon
        End Select
        rec = False
    End If
End Sub

Private Sub Picture_Slide_Som_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Colocar o slide som na posição pretendida
    Picture_Slide_Som.CurrentX = X
    Slide_Som.left = Picture_Slide_Som.CurrentX
    Form_Principal.Slide_Som.left = Picture_Slide_Som.CurrentX
    Form_Principal.Slide_Som_Mini.left = Picture_Slide_Som.CurrentX
    Form_Mini_Player.Slide_Som.left = Picture_Slide_Som.CurrentX
    Form_PopUp.Slide_Som.left = Picture_Slide_Som.CurrentX
    
    Form_Principal.Verificar_Volume
    Form_Principal.Text_Slide_Som.Text = Slide_Som.left
    Form_Mini_Player.Text_Slide_Som.Text = Slide_Som.left
    
    'Caso o player esteja sem som
    If Form_Principal.Mudo = True Then
        Form_Principal.Wmp.settings.mute = False
        Form_Wmp.Wmp.settings.mute = True
        Form_Principal.Mudo = False
        
        Form_Principal.Menu_Controlos(4).Caption = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_Principal.Botao_Mudo_Mini.ToolTipText = Idioma_Mudo_On
        Form_Mini_Player.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_PopUp.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_Principal.Botao_Mudo_Mini.Picture = Form_Skin.Som_On_Normal_Mini.Picture
        Form_Mini_Player.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_PopUp.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
    End If
    Slide_Som.Picture = Form_Skin.Slide_Som_Normal.Picture
    Form_Mini_Player.Slide_Som.Picture = Form_Skin.Slide_Som_Normal.Picture
    Form_PopUp.Slide_Som.Picture = Form_Skin.Slide_Som_Normal.Picture
End Sub

Private Sub Slide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DNa = True
    Txa = X
End Sub

Private Sub Slide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DNa Then
        NewLeft = Slide.left + X - Txa
        If NewLeft < Image_Barra_Slide.left + 3 Then
            NewLeft = Image_Barra_Slide.left + 3
        End If
        If NewLeft > Image_Barra_Slide.Width + Image_Barra_Slide.left - 7 - Slide.Width Then
            NewLeft = Image_Barra_Slide.Width + Image_Barra_Slide.left - 7 - Slide.Width
        End If
        Form_Principal.Slide.left = NewLeft
        Form_Principal.Slide_Mini.left = NewLeft
        Form_Mini_Player.Slide.left = NewLeft
        Form_Principal.Image_Progresso.Width = Form_Principal.Slide.left
        Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
    End If
End Sub

Private Sub Slide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Colocar o slide na posição largada
    'On Error Resume Next
    Dim offseti As Single
    DNa = False
'    offseti = (Form_Principal.Slide.Left - Form_Principal.Image_Barra_Slide.Left - 3) / (Form_Principal.Image_Barra_Slide.Width - 10 - Form_Principal.Slide.Width)
    offseti = (Slide.left - Image_Barra_Slide.left - 3) / (Image_Barra_Slide.Width - 10 - Slide.Width)
    Form_Principal.Wmp.Controls.CurrentPosition = Int(Form_Principal.Wmp.currentMedia.Duration * offseti)
    Form_Wmp.Wmp.Controls.CurrentPosition = Int(Form_Principal.Wmp.currentMedia.Duration * offseti)
    Form_Principal.Image_Progresso.Width = Form_Principal.Slide.left
    Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
End Sub

Private Function ScrollText(MyText As String) As String
    'Função para fazer scrolling na faixa em reprodução
    On Error Resume Next
    MyText = (Right$(MyText, Len(MyText) - 1)) & left$(MyText, 1)
    ScrollText = MyText
End Function

Private Sub Slide_Som_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Captar a posição do slider do som
    DNa_Som = True
    Txa_Som = X
    Slide_Som.Picture = Form_Skin.Slide_Som_Down.Picture
End Sub

Private Sub Slide_Som_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o slider do som
    If DNa_Som Then
        NewLeft_Som = Slide_Som.left + X - Txa_Som '- 6
        If NewLeft_Som < 0 Then
            NewLeft_Som = 0
        End If
        If NewLeft_Som > Picture_Slide_Som.Width - Slide_Som.Width Then
            NewLeft_Som = Picture_Slide_Som.Width - Slide_Som.Width
        End If
        Form_Principal.Slide_Som.left = NewLeft_Som
        Form_Principal.Slide_Som_Mini.left = NewLeft_Som
        Form_Mini_Player.Slide_Som.left = NewLeft_Som
        Form_PopUp.Slide_Som.left = NewLeft_Som
    End If
    Form_Principal.Verificar_Volume
End Sub

Private Sub Slide_Som_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Posicionar o slider do som na posição largada
    On Error Resume Next
'''    Dim offseti As Single
    DNa_Som = False
    Form_Principal.Verificar_Volume
    Form_Principal.Text_Slide_Som.Text = Slide_Som.left
    Form_Mini_Player.Text_Slide_Som.Text = Slide_Som.left
        
    'Caso o player esteja sem som
    If Form_Principal.Mudo = True Then
        Form_Principal.Wmp.settings.mute = False
        Form_Wmp.Wmp.settings.mute = True
        Form_Principal.Mudo = False
        
        Form_Principal.Menu_Controlos(4).Caption = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_Principal.Botao_Mudo_Mini.ToolTipText = Idioma_Mudo_On
        Form_Mini_Player.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_PopUp.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_Principal.Botao_Mudo_Mini.Picture = Form_Skin.Som_On_Normal_Mini.Picture
        Form_Mini_Player.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_PopUp.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
    End If
    Slide_Som.Picture = Form_Skin.Slide_Som_Normal.Picture
End Sub

Private Sub SliderBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Colocar o slide som na posição pretendida
    SliderBar.CurrentX = X
    Slide.left = SliderBar.CurrentX
    Form_Principal.Slide_Mini.left = SliderBar.CurrentX
    Form_Principal.Slide.left = SliderBar.CurrentX
    Image_Progresso.Width = Slide.left
    Form_Principal.Image_Progresso.Width = Form_Principal.Slide.left
    
    'Colocar o slide na posição largada
    On Error Resume Next
    Dim offseti As Single
    DNa = False
    offseti = (Slide.left - Image_Barra_Slide.left - 3) / (Image_Barra_Slide.Width - 10 - Slide.Width)
    Form_Principal.Wmp.Controls.CurrentPosition = Int(Form_Principal.Wmp.currentMedia.Duration * offseti)
    Form_Wmp.Wmp.Controls.CurrentPosition = Int(Form_Principal.Wmp.currentMedia.Duration * offseti)
    Image_Progresso.Width = Slide.left
    Form_Principal.Image_Progresso.Width = Form_Principal.Slide.left
End Sub

