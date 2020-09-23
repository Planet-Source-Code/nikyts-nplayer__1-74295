VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD - Copy.OCX"
Begin VB.Form Form_Lista 
   Appearance      =   0  'Flat
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Lista.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grelha_Lista_Em_Reproducao 
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   3
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   3223857
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColorFixed  =   2171169
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      OLEDropMode     =   1
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
      ScaleWidth      =   401
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6015
      Begin VB.Image Botao_Fechar 
         Height          =   135
         Left            =   5520
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Lista"
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
         TabIndex        =   1
         Top             =   120
         Width           =   465
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
      Height          =   375
      Left            =   3840
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Fundo_Barra_Player 
      Enabled         =   0   'False
      Height          =   525
      Left            =   0
      Top             =   480
      Width           =   465
   End
End
Attribute VB_Name = "Form_Lista"
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

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    Me.Hide
    Form_Mini_Player.Label_Lista.ForeColor = Form_Skin.Cor_Label_Barra_Visor.backcolor
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
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento para contruir o formulário
    Desenhar_Formulario
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Playlist", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Playlist", "Button_Close", Localizacao_Ficheiro_Lingua)
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Grelha_Lista_Em_Reproducao.backcolor = .Cor_Grid_BackColor.backcolor
        Grelha_Lista_Em_Reproducao.BackColorBkg = .Cor_Grid_BackColorBkg.backcolor
        Grelha_Lista_Em_Reproducao.BackColorFixed = .Cor_Grid_BackColorFixed.backcolor
        Grelha_Lista_Em_Reproducao.BackColorSel = .Cor_Grid_BackColorSel.backcolor
        Grelha_Lista_Em_Reproducao.ForeColor = .Cor_Grid_ForeColor.backcolor
        Grelha_Lista_Em_Reproducao.ForeColorFixed = .Cor_Grid_ForeColorFixed.backcolor
        Grelha_Lista_Em_Reproducao.ForeColorSel = .Cor_Grid_ForeColorSel.backcolor
        Grelha_Lista_Em_Reproducao.GridColor = .Cor_Grid_Color.backcolor
        Grelha_Lista_Em_Reproducao.GridColorFixed = .Cor_Grid_ColorFixed.backcolor
        Fundo_Barra_Player.Picture = .Fundo_Mini_Player.Picture
    End With
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Ajustar_Formulario Form_Lista, False, False, False, False
    
    With Shape_Contorno
        .Height = Me.ScaleHeight
        .top = 0
        .Width = Me.ScaleWidth
        .left = 0
    End With
    
    With Fundo_Barra_Player
        .Stretch = True
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Width = Me.ScaleWidth
        .left = 0
    End With
    
    With Grelha_Lista_Em_Reproducao
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - 2 - 5
        .Width = Barra_ControlBox.ScaleWidth - 10
        .left = 5
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    'Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Grelha_Lista_Em_Reproducao_Click()
    'Igual a linha á lista da biblioteca
    Form_Principal.Grelha_Reproduzida.Row = Grelha_Lista_Em_Reproducao.Row
    Form_Principal.Grelha_Reproduzida.ColSel = Form_Principal.Grelha_Reproduzida.Cols - 1
End Sub

Private Sub Grelha_Lista_Em_Reproducao_DblClick()
    'Reproduzir a música da linha selecionada
    If Grelha_Lista_Em_Reproducao.Rows <= 1 Then Exit Sub
    Musica_Linha_Pressionada = Grelha_Lista_Em_Reproducao.Row 'Form_Principal.Grelha_Reproduzida.Row
    Form_Principal.Reproduzir_Musica_da_Grelha
End Sub

Private Sub Grelha_Lista_Em_Reproducao_EnterCell()
    'Igual a linha á lista da biblioteca
    Form_Principal.Grelha_Reproduzida.Row = Grelha_Lista_Em_Reproducao.Row
    Form_Principal.Grelha_Reproduzida.ColSel = Form_Principal.Grelha_Reproduzida.Cols - 1
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Lista
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Lista
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Lista
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Lista
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Lista
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Lista
End Sub

