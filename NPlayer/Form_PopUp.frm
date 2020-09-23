VERSION 5.00
Begin VB.Form Form_PopUp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_PopUp.frx":0000
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture_Slide_Som 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2520
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1650
      Begin VB.PictureBox Slide_Som 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         Picture         =   "Form_PopUp.frx":14CFA
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   960
   End
   Begin VB.PictureBox Pic_Capa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00101010&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   120
      Picture         =   "Form_PopUp.frx":14FA4
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   120
      Width           =   750
   End
   Begin VB.Image Botao_Mudo 
      Height          =   150
      Left            =   2280
      Picture         =   "Form_PopUp.frx":158F8
      ToolTipText     =   "Mudo"
      Top             =   1605
      Width           =   180
   End
   Begin VB.Image Botao_Ampliar 
      Height          =   210
      Left            =   4380
      Picture         =   "Form_PopUp.frx":15AA2
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Botao_Biblioteca 
      Height          =   240
      Left            =   4740
      Picture         =   "Form_PopUp.frx":15D84
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Botao_Janela 
      Height          =   240
      Left            =   5100
      Picture         =   "Form_PopUp.frx":160C6
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Botao_Pausa 
      Height          =   810
      Left            =   600
      Picture         =   "Form_PopUp.frx":16408
      ToolTipText     =   "Pausa"
      Top             =   1320
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image Botao_Seguinte 
      Height          =   585
      Left            =   1320
      Picture         =   "Form_PopUp.frx":18532
      ToolTipText     =   "Faixa seguinte"
      Top             =   1380
      Width           =   540
   End
   Begin VB.Image Botao_Antes 
      Height          =   600
      Left            =   120
      Picture         =   "Form_PopUp.frx":195E8
      ToolTipText     =   "Faixa anterior"
      Top             =   1380
      Width           =   555
   End
   Begin VB.Image Botao_Play 
      Height          =   810
      Left            =   4680
      Picture         =   "Form_PopUp.frx":1A7AA
      ToolTipText     =   "Reproduzir"
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Linha_Separador 
      BackColor       =   &H00404040&
      Height          =   15
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label_Contador 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   1020
      TabIndex        =   3
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label_Artista 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Height          =   180
      Left            =   1020
      TabIndex        =   1
      Top             =   600
      Width           =   2880
   End
   Begin VB.Label Label_Faixa 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1020
      TabIndex        =   0
      Top             =   360
      Width           =   2880
   End
End
Attribute VB_Name = "Form_PopUp"
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

'Variável para o temporador de encerramento
Public Tempo As Integer

'Variável para determinar a altura do formulário
Dim Tamanho_Standard As Boolean
Public Modo_de_Trabalho As Boolean

'Variáveis do idioma
Dim Idioma_Visualizacao_Simples As String
Dim Idioma_Visualizacao_Personalizada  As String
Dim Idioma_Ver_Biblioteca  As String
Dim Idioma_Janela_Visivel As String
Dim Idioma_Janela_Oculta  As String

'VARIÁVERIS DO SLIDER SOM
Dim TX_Som As Integer, Ty_Som As Integer, DN_Som As Boolean
Dim Txa_Som As Integer, DNa_Som As Boolean
Dim Tyb_Som, Dnb_Som As Boolean
Dim NewLeft_Som As Integer

'API's para poder arredondar os cantos do formulário
Private Declare Function CreateRoundRectRgn Lib _
        "gdi32" (ByVal X1 As Long, ByVal Y1 As _
        Long, ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
        (ByVal hwnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long
Private Declare Function GetClientRect Lib "user32" _
        (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
  left As Long
  top As Long
  Right As Long
  Bottom As Long
End Type

Sub Retangulo(m_hWnd As Long, Fator As Byte)
    'Procedimento para poder arredondar os cantos do formulário
    Dim RGN As Long
    Dim RC As RECT
    Call GetClientRect(m_hWnd, RC)
    RGN = CreateRoundRectRgn(RC.left, RC.top, RC.Right, RC.Bottom, Fator, Fator)
    SetWindowRgn m_hWnd, RGN, True
End Sub

Private Sub Botao_Ampliar_Click()
    'Ajustar o tamanho do form
    If Tamanho_Standard = True Then
        Tamanho_Standard = False
        With Me
            .Picture = Form_Skin.Fundo_PopUp_G.Picture
            .Height = Screen.TwipsPerPixelY * Form_Skin.Fundo_PopUp_G.Height
        End With
        Botao_Ampliar.ToolTipText = Idioma_Visualizacao_Simples
        Botao_Ampliar.Picture = Form_Skin.Icon_Form_Down.Picture
        
    Else
        Tamanho_Standard = True
        With Me
            .Picture = Form_Skin.Fundo_PopUp_P.Picture
            .Height = Screen.TwipsPerPixelY * Form_Skin.Fundo_PopUp_P.Height
        End With
        Botao_Ampliar.ToolTipText = Idioma_Visualizacao_Personalizada
        Botao_Ampliar.Picture = Form_Skin.Icon_Form_Up.Picture
    End If
    
    Retangulo Me.hwnd, 5
    Me.top = Screen.Height - Me.Height - 500
End Sub

Public Sub Botao_Janela_Click()
    'Des/activar o form popUp
    If Modo_de_Trabalho = False Then
        Timer1.Enabled = False
        Modo_de_Trabalho = True
        Botao_Janela.Picture = Form_Skin.Icon_Ver_Janela.Picture
        Botao_Janela.ToolTipText = Idioma_Janela_Oculta
    Else
        'Me.Hide
        Tempo = 0
        Timer1.Enabled = True
        Modo_de_Trabalho = False
        Botao_Janela.Picture = Form_Skin.Icon_Ocultar_Janela.Picture
        Botao_Janela.ToolTipText = Idioma_Janela_Visivel
    End If
End Sub

Private Sub Botao_Biblioteca_Click()
    'Chamar o procedimento
    Restaurar_Janela_do_Player
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

Private Sub Botao_Antes_Click()
    'Atalho para
    Form_Principal.Botao_Antes_Click
End Sub

Private Sub Botao_Antes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Antes.Picture = Form_Skin.Botao_Antes_Down.Picture
End Sub

Private Sub Botao_Antes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Antes.Picture = Form_Skin.Botao_Antes_Normal.Picture
End Sub

Private Sub Botao_Seguinte_Click()
    'Atalho para
    Form_Principal.Botao_Seguinte_Click
End Sub

Private Sub Botao_Seguinte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Seguinte.Picture = Form_Skin.Botao_Seguinte_Down.Picture
End Sub

Private Sub Botao_Seguinte_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Seguinte.Picture = Form_Skin.Botao_Seguinte_Normal.Picture
End Sub

Private Sub Form_DblClick()
    'Chamar o procedimento
    Restaurar_Janela_do_Player
End Sub

Private Sub Restaurar_Janela_do_Player()
    'Procedimento para restaurar a janela do player
    Form_Principal.Remover_Tray_Icon
    Timer1.Enabled = False
    Modo_de_Trabalho = False
    Botao_Janela.Picture = Form_Skin.Icon_Ocultar_Janela.Picture
    Botao_Janela.ToolTipText = Idioma_Janela_Visivel
    Form_Principal.Show
    Form_Principal.Modo_Tray = False
    Me.Hide
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    
    'Colocar o formulário por cima dos outros
    OnTop Me, True
    
    'Arredondar os cantos do formulário
    Retangulo Me.hwnd, 5

    Tempo = 0
    Tamanho_Standard = True
    Modo_de_Trabalho = False
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Idioma_Visualizacao_Simples = ReadINI("PopUp", "Simple_View", Localizacao_Ficheiro_Lingua)
    Idioma_Visualizacao_Personalizada = ReadINI("PopUp", "Custom_View", Localizacao_Ficheiro_Lingua)
    Idioma_Ver_Biblioteca = ReadINI("PopUp", "View_Library", Localizacao_Ficheiro_Lingua)
    Idioma_Janela_Visivel = ReadINI("PopUp", "State_Window_Normal", Localizacao_Ficheiro_Lingua)
    Idioma_Janela_Oculta = ReadINI("PopUp", "State_Window_Over", Localizacao_Ficheiro_Lingua)
    
    Botao_Ampliar.ToolTipText = Idioma_Visualizacao_Personalizada
    Botao_Biblioteca.ToolTipText = Idioma_Ver_Biblioteca
    Botao_Janela.ToolTipText = Idioma_Janela_Visivel
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Botao_Ampliar.Picture = .Icon_Form_Up.Picture
        Botao_Biblioteca.Picture = .Icon_Ver_Biblioteca.Picture
        Botao_Janela.Picture = .Icon_Ocultar_Janela.Picture
        Me.Picture = .Fundo_PopUp_P.Picture
        Pic_Capa.Picture = .Image_Capa.Picture
        Label_Contador.ForeColor = .Cor_Label_Contador_Popup.backcolor
        Label_Faixa.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Artista.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Botao_Play.Picture = .Botao_Play_Normal.Picture
        Botao_Pausa.Picture = .Botao_Pausa_Normal.Picture
        Botao_Antes.Picture = .Botao_Antes_Normal.Picture
        Botao_Seguinte.Picture = .Botao_Seguinte_Normal.Picture
        Botao_Mudo.Picture = .Som_On_Normal.Picture
        Picture_Slide_Som.Picture = .Fundo_Slider_Volume.Picture
        Slide_Som.Picture = .Slide_Som_Normal.Picture
    End With
End Sub

Private Sub Desenhar_Formulario()
    'Procedimento para construir o formulário
    With Me
        .Height = Screen.TwipsPerPixelY * Form_Skin.Fundo_PopUp_P.Height
        .top = Screen.Height - .Height - 600
        .Width = Screen.TwipsPerPixelX * Form_Skin.Fundo_PopUp_P.Width
        .left = Screen.Width - .Width - 50
    End With
    
    With Pic_Capa
        .Height = Form_Skin.Image_Capa.Height
        .top = ((Me.ScaleHeight - .ScaleHeight) / 2)
        .Width = Form_Skin.Image_Capa.Width
        .left = .top
    End With

    With Botao_Janela
        .Height = Form_Skin.Icon_Form_Up.Height
        .top = Pic_Capa.top
        .Width = Form_Skin.Icon_Form_Up.Width
        .left = Me.ScaleWidth - .Width - .top
    End With

    With Botao_Biblioteca
        .Height = Botao_Janela.Height
        .top = Botao_Janela.top
        .Width = Botao_Janela.Width
        .left = Botao_Janela.left - .Width - .top
    End With
    
    With Botao_Ampliar
        .Height = Botao_Janela.Height
        .top = Botao_Janela.top
        .Width = Botao_Janela.Width
        .left = Botao_Biblioteca.left - .Width - .top
    End With
    
    With Linha_Separador
        .Width = Me.ScaleWidth - (2 * Pic_Capa.left)
        .left = Pic_Capa.left
    End With
    
    With Botao_Antes
        .Height = Form_Skin.Botao_Seguinte_Normal.Height
        .Width = Form_Skin.Botao_Seguinte_Normal.Width
        .left = Pic_Capa.left
    End With
    
    With Botao_Play
        .Height = Form_Skin.Botao_Play_Normal.Height
        .Width = Form_Skin.Botao_Play_Normal.Width
        .left = Botao_Antes.left + Botao_Antes.Width
    End With
    
    With Botao_Pausa
        .Height = Botao_Play.Height
        .top = Botao_Play.top
        .Width = Botao_Play.Width
        .left = Botao_Play.left
    End With
    
    With Botao_Seguinte
        .Height = Botao_Antes.Height
        .top = Botao_Antes.top
        .Width = Botao_Antes.Width
        .left = Botao_Play.left + Botao_Play.Width
    End With
    
    With Botao_Mudo
        .Height = Form_Skin.Som_On_Normal.Height
        .Width = Form_Skin.Som_On_Normal.Width
    End With
    
    With Picture_Slide_Som
        .Height = Form_Skin.Fundo_Slider_Volume.Height
        .Width = Form_Skin.Fundo_Slider_Volume.Width
    End With
    
    With Slide_Som
        .Height = Form_Skin.Slide_Som_Normal.Height
        .top = (Picture_Slide_Som.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Slide_Som_Normal.Width
    End With
End Sub

Private Sub Label1_Click()
    'Chamar o procedimento
    Restaurar_Janela_do_Player
End Sub

Private Sub lblArtist_Click()
    'Chamar o procedimento
    Restaurar_Janela_do_Player
End Sub

Private Sub lblSong_Click()
    'Chamar o procedimento
    Restaurar_Janela_do_Player
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
    If Mudo = True Then
        Form_Principal.Wmp.settings.mute = False
        Form_Wmp.Wmp.settings.mute = True
        Mudo = False
        
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

Private Sub Timer1_Timer()
    'Ocultar o formulário
    On Error Resume Next
    If Modo_de_Trabalho = False Then
        Tempo = Tempo + 1
        If Tempo = 50 Then
            Modo_de_Trabalho = False
            Botao_Janela.Picture = Form_Skin.Icon_Ocultar_Janela.Picture
            Botao_Janela.ToolTipText = Idioma_Janela_Visivel
            Me.Hide
            Timer1.Enabled = False
        End If
    End If
End Sub
