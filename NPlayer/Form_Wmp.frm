VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form_Wmp 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
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
   Icon            =   "Form_Wmp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Frame_Video 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   3375
      Left            =   120
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   380
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5700
      Begin WMPLibCtl.WindowsMediaPlayer Wmp 
         Height          =   600
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   900
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "none"
         stretchToFit    =   -1  'True
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   1588
         _cy             =   1058
      End
   End
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6015
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   5640
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Video"
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
         Width           =   540
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.Image Fundo_Barra_Player 
      Enabled         =   0   'False
      Height          =   525
      Left            =   0
      Top             =   600
      Width           =   465
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00212121&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Wmp"
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
Option Explicit
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

Private Sub Form_Activate()
    Form_Wmp.Wmp.settings.mute = True
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Wmp
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Wmp
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Wmp
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    Me.Hide
    Form_Mini_Player.Label_Video.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
'    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    'Chamar o procedimento para contruir o formulário
    With Me
        .Height = Form_Mini_Player.Height
        .Width = Form_Mini_Player.Width
    End With
    
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, False

    Wmp.settings.Volume = 0
    Wmp.settings.mute = True
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento para contruir o formulário
    Desenhar_Formulario
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Fundo_Barra_Player.Picture = .Fundo_Mini_Player.Picture
    End With
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Wmp", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Wmp", "Button_Close", Localizacao_Ficheiro_Lingua)
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Ajustar_Formulario Form_Wmp, False, False, False, False
    
    With Fundo_Barra_Player
        .Stretch = True
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Width = Me.ScaleWidth
        .left = 0
    End With
    
    With Frame_Video
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - 2 - 5
        .Width = Barra_ControlBox.ScaleWidth - 10
        .left = 5
    End With
    
    With Wmp
        .Height = Frame_Video.ScaleHeight
        .top = 0
        .Width = Frame_Video.ScaleWidth
        .left = 0
    End With
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Wmp
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Wmp
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Wmp
End Sub

