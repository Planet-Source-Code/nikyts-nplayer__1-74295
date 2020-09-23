VERSION 5.00
Begin VB.Form Form_Reportar_Erro 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   6375
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   1005
         ItemData        =   "Form_Reportar_Erro.frx":0000
         Left            =   0
         List            =   "Form_Reportar_Erro.frx":0002
         TabIndex        =   23
         Top             =   3240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.PictureBox Lista_Assunto 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00101010&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   1080
         ScaleHeight     =   63
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   303
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label Label_Assunto 
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Assunto"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   960
         End
         Begin VB.Label Shape_Sombra 
            BackColor       =   &H00D88316&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.PictureBox Barra_Text_Mensagem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1920
         Left            =   240
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2520
         Width           =   5475
         Begin VB.TextBox Text_Mensagem 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   1740
            Left            =   600
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   30
            Width           =   1860
         End
         Begin VB.Shape Contorno_Mensagem 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Assunto 
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5475
         Begin VB.PictureBox Seta_Assunto 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5160
            Picture         =   "Form_Reportar_Erro.frx":0004
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   11
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   165
         End
         Begin VB.TextBox Text_Assunto 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Assunto 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Email 
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
         Top             =   840
         Width           =   5475
         Begin VB.TextBox Text_Email 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   600
            TabIndex        =   0
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Email 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Image Image_Erro 
         Enabled         =   0   'False
         Height          =   210
         Left            =   300
         Picture         =   "Form_Reportar_Erro.frx":02DC
         Top             =   240
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label_Mensagem 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label Label_Texto 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label_Erro 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         BackStyle       =   0  'Transparent
         Caption         =   "Indique um endereço de email válido."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label_De 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label_Info 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(Para possivel contacto caso seja necessário)"
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
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   3480
      End
      Begin VB.Shape Shape_Erro 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H008080FF&
         Height          =   315
         Left            =   240
         Top             =   180
         Visible         =   0   'False
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
      AutoRedraw      =   -1  'True
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   393
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5400
      Width           =   5895
      Begin VB.PictureBox Botao_Ok 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2040
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   3
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
            Caption         =   "Enviar"
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   600
            TabIndex        =   10
            Top             =   45
            Width           =   540
         End
      End
      Begin VB.PictureBox Botao_Cancelar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3960
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   4
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
            TabIndex        =   9
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
      AutoRedraw      =   -1  'True
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
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
         Caption         =   "Suporte técnico"
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
         Width           =   1530
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
Attribute VB_Name = "Form_Reportar_Erro"
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
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'Variável para indicar qual a linha que está selecionada da lista linguas
Dim Linha_Selecionada As Integer

'Variável para o idioma
Dim Idioma_Reportar As String
Dim Idioma_Sugestao As String
Dim Idioma_Questao As String
Dim Idioma_Outro As String

Dim Idioma_Erro As String
Dim Idioma_Descricao As String
Dim Idioma_Erro_Execucao As String
Dim Idioma_Conectar_Servidor As String
Dim Idioma_Internet_Desligada As String
Dim Idioma_Mensagem_Enviada As String

Private Sub Barra_ControlBox_Click()
    'Ocultar frame
    Lista_Assunto.Visible = False
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

Private Sub Botao_Fechar_Click()
    'Fechar formulário
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
    If KeyCode = vbKeyRight Then Botao_Ok_LostFocus: Botao_Cancelar_GotFocus: Botao_Cancelar.SetFocus
End Sub

Private Sub Botao_Ok_LostFocus()
    'Ao perder o focus no botao
    Contorno_Ok.Visible = False
End Sub

Private Sub Form_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    'Iniciar o formulário
    Carregar_Idioma
    Carregar_Skin
    Desenhar_Formulario
    
    'Variáveis para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, True
    
    'Preencher a lista com as opções de assunto
    With List1
        .AddItem Idioma_Reportar
        .AddItem Idioma_Sugestao
        .AddItem Idioma_Questao
        .AddItem Idioma_Outro
    End With
    'Text_Assunto.Text = List1.List(0)
    
    'On Error Resume Next
    'Criar a lista consoante o nº de assuntos existentes
    Label_Assunto(0).Caption = ""
    Label_Assunto(0).Visible = True
    Dim Objecto As Integer
    For Objecto = 1 To List1.ListCount - 1
        Load Label_Assunto(Objecto)
        Label_Assunto(Objecto).Move Label_Assunto(Objecto - 1).left, Label_Assunto(Objecto - 1).top + Label_Assunto(Objecto - 1).Height
        Label_Assunto(Objecto).Visible = True
        
        Load Shape_Sombra(Objecto)
        Shape_Sombra(Objecto).Move Shape_Sombra(Objecto - 1).left, Shape_Sombra(Objecto - 1).top + Shape_Sombra(Objecto - 1).Height
        Shape_Sombra(Objecto).Visible = False
        Shape_Sombra(Objecto).ZOrder 1
    Next Objecto
        
    'Preencher as label's com o titulo do assunto
    Dim Z As Integer
    List1.ListIndex = 0
    For Z = 0 To List1.ListCount - 1
        Label_Assunto(Z).Caption = List1.List(Z)
    Next Z
    
    'Selecionar a 1ªlinha da lista assunto
    Linha_Selecionada = 0
    Shape_Sombra(0).Visible = True
    Label_Assunto(0).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
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
        Contorno_Email.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Assunto.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Mensagem.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_De.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Info.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Texto.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Mensagem.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        'Barra_Text_Email.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Email.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Email.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Email.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Email.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Assunto.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Assunto.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Assunto.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Assunto.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Assunto.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Assunto.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Assunto.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Assunto.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Mensagem.Picture = .Caixa_de_Observacoes.Picture
        Barra_Text_Mensagem.Picture = Nothing
        Barra_Text_Mensagem.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 10, 0, 0, 10, 10
        Barra_Text_Mensagem.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Mensagem.ScaleWidth, 10, 10, 0, 40, 10
        Barra_Text_Mensagem.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Mensagem.ScaleWidth - 10), 0, 10, 10, 51, 0, 10, 10
        Barra_Text_Mensagem.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 10, 10, (Barra_Text_Mensagem.ScaleHeight - 20), 0, 10, 10, 10
        Barra_Text_Mensagem.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Mensagem.ScaleWidth - 10), 10, 10, (Barra_Text_Mensagem.ScaleHeight - 20), 51, 10, 10, 10
        Barra_Text_Mensagem.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, (Barra_Text_Mensagem.ScaleHeight - 10), 10, 10, 0, 17, 10, 10
        Barra_Text_Mensagem.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, (Barra_Text_Mensagem.ScaleHeight - 10), (Barra_Text_Mensagem.ScaleWidth - 20), 10, 10, 17, 40, 10
        Barra_Text_Mensagem.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Mensagem.ScaleWidth - 10), (Barra_Text_Mensagem.ScaleHeight - 10), 10, 10, 51, 17, 10, 10
        Text_Mensagem.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Mensagem.ForeColor = .Cor_Letra_Textbox.backcolor
        Lista_Assunto.backcolor = .Cor_Fundo_Textbox.backcolor
        Shape_Sombra(0).backcolor = .Cor_Contorno_Caixas.backcolor
        Label_Assunto(0).ForeColor = .Cor_Letra_Textbox.backcolor
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Ok.Picture = .Pic_Button.Picture
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Seta_Assunto.Picture = .Seta_Combo.Picture
    End With
End Sub

Public Sub Desactivar_Objectos()
    'as textboxs
    Text_Mensagem.Enabled = False
End Sub

Public Sub Activar_Objectos()
    'as textboxs
    Text_Mensagem.Enabled = True
End Sub

Public Sub Limpa_Campos()
    'Limpa o conteudo das caixas de texto
    Text_Email.Text = ""
    Text_Assunto.Text = ""
    Text_Mensagem.Text = ""
End Sub

Private Sub Form_Resize()
    Desenhar_Formulario
End Sub

Private Sub Frame_Botoes_Click()
    'Ocultar frame
    Lista_Assunto.Visible = False
End Sub

Private Sub Frame_Centro_Click()
    'Ocultar frame
    Lista_Assunto.Visible = False
End Sub

Private Sub Label_Assunto_Click(Index As Integer)
    'Indicar a lingua selecionada pelo utilizador
    Text_Assunto.Text = Label_Assunto(Index).Caption
    
    Lista_Assunto.Visible = False
    Text_Assunto.SetFocus
End Sub

Private Sub Label_Assunto_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada = Index Then Exit Sub
    Shape_Sombra(Linha_Selecionada).Visible = False
    Label_Assunto(Linha_Selecionada).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
    Shape_Sombra(Index).Visible = True
    Label_Assunto(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    Linha_Selecionada = Index
End Sub

Private Sub Label_Cancelar_Click()
    'Cancelar operação
    Unload Me
End Sub

Private Sub Label_Ok_Click()
    'Verificar o preencimento das textboxs
    On Error GoTo Corrige_Erro
    Label_Erro.Visible = False
    Shape_Erro.Visible = False
    Image_Erro.Visible = False
    
    'Verifica se o campo email está no formato correcto
    If Not IsEmail(Text_Email.Text) Then
        Label_Erro.Visible = True
        Shape_Erro.Visible = True
        Image_Erro.Visible = True
        Text_Email.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = 11
    Botao_Ok.Enabled = False
    Label_Ok.Enabled = False
    Contorno_Ok.Visible = False
    
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/suporte/" & "enviarmensagem.asp?Email=" & Text_Email.Text & "&Assunto=" & App.ProductName & " - " & Text_Assunto.Text & "&Mensagem=" & Text_Mensagem.Text, False
    servidor.send 'envia o pedido para o servidor

    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
            Mensagem_de_Aviso "Information", Idioma_Mensagem_Enviada
            
            'Limpar os campos para se poder enviar uma nova mensagem
            Me.MousePointer = 0
            Limpa_Campos
            Botao_Ok.Enabled = True
            Label_Ok.Enabled = True
            Contorno_Ok.Visible = True
            Verificar_o_Prenchimento
            Contorno_Ok.Visible = False
            Text_Email.SetFocus
        End If
    End If
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Reportar_Erro
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Reportar_Erro
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Reportar_Erro
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Reportar_Erro
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Reportar_Erro
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Reportar_Erro
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Email.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text_Email.ScaleWidth) + (2 * Barra_Text_Email.left) + 20)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Shape_Erro.left + Shape_Erro.Height + 3 + Label_De.Height + 3 _
                + Form_Skin.Caixa_de_Texto.Height + 6 + 3 + Label_Texto.Height + 3 + Form_Skin.Caixa_de_Texto.Height + 6 + 3 _
                + Label_Mensagem.Height + 3 + Form_Skin.Caixa_de_Observacoes.Height + 6 + Fundo_Frame_Botoes.Height + (2 * Shape_Erro.left))
    End With
    
    Ajustar_Formulario Form_Reportar_Erro, False, False, True, True
    
    Ajustar_Botao Form_Reportar_Erro, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Reportar_Erro, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With

    Ajustar_Caixa_Texto Barra_Text_Email, Text_Email, Contorno_Email, False
    Ajustar_Caixa_Texto Barra_Text_Assunto, Text_Assunto, Contorno_Assunto, False
    Ajustar_Caixa_Texto Barra_Text_Mensagem, Text_Mensagem, Contorno_Mensagem, True
    
    With Shape_Sombra(0)
        .Width = Lista_Assunto.ScaleWidth
    End With
    
    With Label_Assunto(0)
        .Width = Lista_Assunto.ScaleWidth
    End With
        
    With Shape_Erro
        .top = .left
        .Width = Barra_Text_Email.ScaleWidth
    End With
    
    With Image_Erro
        .top = (Shape_Erro.top + Shape_Erro.Height) / 2
    End With
    
    With Label_Erro
        .top = Image_Erro.top
    End With
    
    With Label_De
        .top = Shape_Erro.top + Shape_Erro.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Info
        .top = Label_De.top
        .left = Label_De.left + Label_De.Width + 3
    End With
    
    With Barra_Text_Email
        .top = Label_De.top + Label_De.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Texto
        .top = Barra_Text_Email.top + Barra_Text_Email.Height + 6
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Assunto
        .top = Label_Texto.top + Label_Texto.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Mensagem
        .top = Barra_Text_Assunto.top + Barra_Text_Assunto.Height + 6
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Mensagem
        .Height = Form_Skin.Caixa_de_Observacoes.Height
        .top = Label_Mensagem.top + Label_Mensagem.Height + 3
        .left = Shape_Erro.left
        .Width = Form_Skin.Caixa_de_Observacoes.Width
    End With
    
    With Seta_Assunto
        .Height = Form_Skin.Seta_Combo.Height
        .top = (Barra_Text_Assunto.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Seta_Combo.Width
        .left = Barra_Text_Assunto.ScaleWidth - .ScaleWidth - .top
    End With

    With Lista_Assunto
        .top = Barra_Text_Assunto.top + Barra_Text_Assunto.ScaleHeight - 1
        .Width = Barra_Text_Assunto.ScaleWidth
        .left = Barra_Text_Assunto.left
    End With
    
    With Shape_Sombra(0)
        .Width = Lista_Assunto.Width
        .left = 0
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Seta_Assunto_Click()
    'Ver/ocultar lista
    If Lista_Assunto.Visible = True Then
        Lista_Assunto.Visible = False
    Else
        Lista_Assunto.Visible = True
    End If
End Sub

Private Sub Text_Assunto_Change()
    'Chamar o procedimento
    Verificar_o_Prenchimento
End Sub

Private Sub Text_Assunto_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Private Sub Text_Assunto_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Assunto.Visible = True
End Sub

Private Sub Text_Assunto_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Assunto.Visible = False
End Sub

Private Sub Text_Email_Change()
    'Chamar o procedimento
    Verificar_o_Prenchimento
End Sub

Private Sub Text_Email_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Private Sub Text_Email_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Email.Visible = True
End Sub

Private Sub Text_Email_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Email.Visible = False
End Sub

Private Sub Text_Mensagem_Change()
    'Chamar o procedimento
    Verificar_o_Prenchimento
End Sub

Private Sub Text_Mensagem_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Public Sub Verificar_o_Prenchimento()
    'Procedimento para verificar se as caixas de texto estão devidamente bem preenchidas
    If Len(Trim(Text_Email.Text)) = 0 Or Len(Trim(Text_Assunto.Text)) = 0 Or Len(Trim(Text_Mensagem.Text)) = 0 Then
        Botao_Ok.Enabled = False
        Label_Ok.Enabled = False
        
    Else
        Botao_Ok.Enabled = True
        Label_Ok.Enabled = True
    End If
End Sub

Private Sub Text_Mensagem_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Mensagem.Visible = True
End Sub

Private Sub Text_Mensagem_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Mensagem.Visible = False
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Support", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Support", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Erro.Caption = ReadINI("Support", "Label_Error", Localizacao_Ficheiro_Lingua)
    Label_Info.Caption = ReadINI("Support", "Label_Info", Localizacao_Ficheiro_Lingua)
    Label_De.Caption = ReadINI("Support", "Label_Of", Localizacao_Ficheiro_Lingua)
    Label_Texto.Caption = ReadINI("Support", "Label_Subject", Localizacao_Ficheiro_Lingua)
    Label_Mensagem.Caption = ReadINI("Support", "Label_Message", Localizacao_Ficheiro_Lingua)
    Idioma_Reportar = ReadINI("Support", "Label_Report", Localizacao_Ficheiro_Lingua)
    Idioma_Sugestao = ReadINI("Support", "Label_Suggestion", Localizacao_Ficheiro_Lingua)
    Idioma_Questao = ReadINI("Support", "Label_Question", Localizacao_Ficheiro_Lingua)
    Idioma_Outro = ReadINI("Support", "Label_Other", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Support", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Support", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    
    Idioma_Erro = ReadINI("Message", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Message", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Message", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Conectar_Servidor = ReadINI("Message", "Error_Connect", Localizacao_Ficheiro_Lingua)
    Idioma_Internet_Desligada = ReadINI("Message", "Error_Internet", Localizacao_Ficheiro_Lingua)
    Idioma_Mensagem_Enviada = ReadINI("Message", "Info_Posted", Localizacao_Ficheiro_Lingua)
End Sub
