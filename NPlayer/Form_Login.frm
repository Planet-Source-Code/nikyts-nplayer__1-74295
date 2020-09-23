VERSION 5.00
Begin VB.Form Form_Login 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
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
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   0
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Width           =   6255
      Begin VB.PictureBox Pic_Lembrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2070
         Width           =   195
      End
      Begin VB.PictureBox Barra_Text_Senha 
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5475
         Begin VB.TextBox Text_Senha 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   600
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Senha 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Utilizador 
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
         Begin VB.TextBox Text_Utilizador 
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
         Begin VB.Shape Contorno_Utilizador 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.CheckBox Check_Lembrar 
         BackColor       =   &H00313131&
         Caption         =   "Lembre-se de mim"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label Label_Criar_Conta 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         BackStyle       =   0  'Transparent
         Caption         =   "Criar conta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D88316&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   3000
         Width           =   960
      End
      Begin VB.Label Label_Erro 
         BackColor       =   &H00F5F5F5&
         BackStyle       =   0  'Transparent
         Caption         =   "Os dados de acesso estão inválidos."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   4125
      End
      Begin VB.Label Label_Esqueceu 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         BackStyle       =   0  'Transparent
         Caption         =   "Esqueceu-se dos seus dados de acesso?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D88316&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   3465
      End
      Begin VB.Label Label_Senha 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label_Utilizador 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Utilizador"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   795
      End
      Begin VB.Image Image_Erro 
         Enabled         =   0   'False
         Height          =   210
         Left            =   300
         Picture         =   "Form_Login.frx":0000
         Top             =   240
         Visible         =   0   'False
         Width           =   210
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
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4320
      Width           =   6375
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
         Begin VB.Label Label_Cancelar 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   480
            TabIndex        =   11
            Top             =   45
            Width           =   780
         End
         Begin VB.Shape Contorno_Cancelar 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
      End
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
         Begin VB.Label Label_Ok 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ok"
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   750
            TabIndex        =   10
            Top             =   45
            Width           =   240
         End
         Begin VB.Shape Contorno_Ok 
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Image Botao_Fechar 
         Height          =   135
         Left            =   5640
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar sessão"
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
         TabIndex        =   8
         Top             =   120
         Width           =   1380
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
Attribute VB_Name = "Form_Login"
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

Private Sub Botao_Cancelar_Click()
    'Fechar formulário
    Unload Me
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
    'Atalho para
    Botao_Cancelar_Click
End Sub

Private Sub Check_Lembrar_Click()
    'Des/Activar a opcção de "Lembrar"
    If Check_Lembrar.Value = 1 Then
        Pic_Lembrar.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Lembrar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, True
    
    'Verificar se exite utilizador guardado
    Dim Lembre_se_de_Mim As String
    Lembre_se_de_Mim = ReadINI("Login", "Remember", Localizacao_Ficheiro_Preferencias)
    If Lembre_se_de_Mim = "True" Then
        Check_Lembrar.Value = 1
        Pic_Lembrar.Picture = Form_Skin.Check_Over.Picture
        Text_Utilizador.Text = ReadINI("Login", "User", Localizacao_Ficheiro_Preferencias)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Login", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Login", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Utilizador.Caption = ReadINI("Login", "Label_User", Localizacao_Ficheiro_Lingua)
    Label_Senha.Caption = ReadINI("Login", "Label_Password", Localizacao_Ficheiro_Lingua)
    Check_Lembrar.Caption = ReadINI("Login", "Check_Remember", Localizacao_Ficheiro_Lingua)
    Label_Esqueceu.Caption = ReadINI("Login", "Label_Forgot", Localizacao_Ficheiro_Lingua)
    Label_Criar_Conta.Caption = ReadINI("Login", "Label_Create", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Login", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Login", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    Label_Erro.Caption = ReadINI("Login", "Label_Error", Localizacao_Ficheiro_Lingua)
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
        Contorno_Utilizador.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Senha.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Utilizador.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Senha.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        'Barra_Text_Utilizador.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Utilizador.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Utilizador.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Utilizador.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Utilizador.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Utilizador.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Utilizador.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Utilizador.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Senha.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Senha.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Senha.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Senha.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Senha.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Senha.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Senha.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Senha.ForeColor = .Cor_Letra_Textbox.backcolor
        Check_Lembrar.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Lembrar.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Pic_Lembrar.Picture = .Check_Normal.Picture
        Pic_Lembrar.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Label_Esqueceu.ForeColor = .Cor_Contorno_Caixas.backcolor
        Label_Criar_Conta.ForeColor = .Cor_Contorno_Caixas.backcolor
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Ok.Picture = .Pic_Button.Picture
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
    End With
End Sub

Private Sub Label_Cancelar_Click()
    'Atalho para
    Botao_Cancelar_Click
End Sub

Private Sub Label_Criar_Click()
    'Criar conta de acesso
    Unload Me
    Form_Criar.Show vbModal
End Sub

Private Sub Label_Criar_Conta_Click()
    'Criar conta
    Unload Me
    Form_Criar.Show 'vbModal
End Sub

Private Sub Label_Esqueceu_Click()
    'Recuperar dados de acesso
    Unload Me
    Form_Recuperar_Conta.Show vbModal
End Sub

Private Sub Label_Ok_Click()
    'Atalho para
    Botao_Ok_Click
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para construir o formulario, ajustando os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Utilizador.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text_Utilizador.ScaleWidth) + (2 * Barra_Text_Utilizador.left) + 20)
'        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Fundo_Frame_Botoes.Height + Frame_Centro.ScaleHeight)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Shape_Erro.left + Shape_Erro.Height + 3 + Label_Utilizador.Height + 3 _
                + Form_Skin.Caixa_de_Texto.Height + 6 + Label_Senha.Height + Form_Skin.Caixa_de_Texto.Height + 6 + Shape_Erro.left + Pic_Lembrar.Height _
                + Shape_Erro.left + Label_Esqueceu.Height + 3 + Label_Criar_Conta.Height + Fundo_Frame_Botoes.Height + (3 * Shape_Erro.left))
    End With
    
    Ajustar_Formulario Form_Login, False, False, True, True
    
    Ajustar_Botao Form_Login, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Login, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With
    
    Ajustar_Caixa_Texto Barra_Text_Utilizador, Text_Utilizador, Contorno_Utilizador, False
    Ajustar_Caixa_Texto Barra_Text_Senha, Text_Senha, Contorno_Senha, False
    
    Ajustar_ChecBox Pic_Lembrar, Check_Lembrar
    
    With Shape_Erro
        .top = .left
        .Width = Barra_Text_Utilizador.ScaleWidth
    End With
    
    With Image_Erro
        .top = (Shape_Erro.top + Shape_Erro.Height) / 2
    End With
    
    With Label_Erro
        .top = Image_Erro.top
    End With
    
    With Label_Utilizador
        .top = Shape_Erro.top + Shape_Erro.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Utilizador
        .top = Label_Utilizador.top + Label_Utilizador.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Senha
        .top = Barra_Text_Utilizador.top + Barra_Text_Utilizador.ScaleHeight + 6
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Senha
        .top = Label_Senha.top + Label_Senha.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Check_Lembrar
        .top = Barra_Text_Senha.top + Barra_Text_Senha.ScaleHeight + Shape_Erro.left
        .Width = Barra_Text_Senha.ScaleWidth
        .left = Shape_Erro.left
    End With
    
    With Pic_Lembrar
        .top = Check_Lembrar.top
        .left = Check_Lembrar.left
    End With
    
    With Label_Esqueceu
        .top = Check_Lembrar.top + Check_Lembrar.Height + Shape_Erro.left
        .left = Shape_Erro.left
    End With
        
    With Label_Criar_Conta
        .top = Label_Esqueceu.top + Label_Esqueceu.Height + 3
        .left = Shape_Erro.left
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Login
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Login
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Login
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Login
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Login
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Login
End Sub

Private Sub Botao_Ok_Click()
    'Verifica primeiro o prenchimento da caixas de texto
    On Error GoTo Corrige_Erro
    If Text_Utilizador.Text = "" Or Text_Senha.Text = "" Then Exit Sub
    Label_Erro.Visible = False
    Shape_Erro.Visible = False
    Image_Erro.Visible = False
    
    Me.MousePointer = 11
    Me.Enabled = False
    Botao_Ok.Enabled = False
            
    'Conclui operação
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "contas.asp?utilizador=" & Text_Utilizador.Text & "&senha=" & Text_Senha.Text, False
    servidor.setRequestHeader "content-type", "text/XML"
    
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If servidor.responseText = "false" Then
        Me.MousePointer = 0
        Me.Enabled = True
        Botao_Ok.Enabled = True
        Text_Utilizador.Text = ""
        Text_Senha.Text = ""
        Label_Erro.Visible = True
        Shape_Erro.Visible = True
        Image_Erro.Visible = True
        Utilizador_Logado = False
        Text_Utilizador.SetFocus
        Exit Sub
    
        
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            With Form_Principal
                .Label_Botao(6).Caption = ReadINI("Main", "Button_Log_Off", Localizacao_Ficheiro_Lingua)
                .Label_Botao(3).Caption = ReadINI("Main", "Button_My_Account", Localizacao_Ficheiro_Lingua)
                .Ajustar_Objectos_Na_Horizontal
                
                onHttpRequest servidor.responseText
                Utilizador_Logado = True
                Form_Principal.Carregar_Minha_Musica
                Form_Principal.Carregar_Meus_Amigos
                Form_Principal.Carregar_Minhas_Mensagens
                Form_Principal.Carregar_Meus_Contactos
                Form_Principal.Carregar_Meus_eventos
                
                Me.Hide
                If Form_Principal.Separador_Clicado = "adicionar_my_music" Then
                    Form_Principal.Label_Botao_Click 2
                End If
                
                If Form_Principal.Separador_Clicado = "my_music" Then
                    Form_Principal.Label_Barra_Drive_Click (9)
                End If

                If Form_Principal.Separador_Clicado = "download" Then
                    Form_Principal.Label_Botao_Click 1
                End If

                If Form_Principal.Separador_Clicado = "adicionar_link" Then
                    Form_Principal.Label_Botao_Click 5
                End If
                
                If Form_Principal.Separador_Clicado = "friends" Then
                    Form_Principal.Label_Barra_Drive_Click (12)
                End If
                
                If Form_Principal.Separador_Clicado = "my_contacts" Then
                    Form_Principal.Label_Barra_Drive_Click (3)
                End If
                
                If Form_Principal.Separador_Clicado = "my_events" Then
                    Form_Principal.Label_Barra_Drive_Click (4)
                End If
                    
                If Form_Principal.Separador_Clicado = "messages" Then
                    Form_Principal.Label_Barra_Drive_Click (13)
                End If
                
                'Verificar se a checkBox "Lembre-se de mim" está activada
                If Check_Lembrar.Value = 1 Then
                    Call WriteINI("Login", "Remember", "True", (Localizacao_Ficheiro_Preferencias))
                    Call WriteINI("Login", "User", Text_Utilizador.Text, (Localizacao_Ficheiro_Preferencias))
                Else
                    Call WriteINI("Login", "Remember", "False", (Localizacao_Ficheiro_Preferencias))
                End If
                .Ajustar_Objectos_Na_Horizontal
            End With
            Me.MousePointer = 0
            Unload Me
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
Me.Enabled = True
Botao_Ok.Enabled = True
End Sub

Private Sub Botao_Ok_GotFocus()
    'Colocar o focus no botao
    Contorno_Ok.Visible = True
End Sub

Private Sub Botao_Ok_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Ok_Click
    If KeyCode = vbKeyLeft Then Botao_Ok_LostFocus: Botao_Cancelar_GotFocus: Botao_Cancelar.SetFocus
End Sub

Private Sub Botao_Ok_LostFocus()
    'Remover o focus no botao
    Contorno_Ok.Visible = False
End Sub

Private Sub onHttpRequest(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    
    If xml.loadXML(responseText) Then
        Dim dados: Set dados = xml.selectSingleNode("/dados")
        Form_Perfil.Label_Utilizador.Caption = dados.selectSingleNode("utilizador").Text
        Form_Perfil.Label_Minha_Senha.Caption = Text_Senha.Text
        
        Dim Sexo As String
        'dim  Imagem as String
        Sexo = dados.selectSingleNode("genero").Text
        'Imagem = dados.selectSingleNode("foto").Text
        
        'Carregar o perfil do utilizador
        With Form_Perfil
            .Text_Nome.Text = dados.selectSingleNode("nome").Text
            .Text_Email.Text = dados.selectSingleNode("email").Text
            If Sexo = "M" Then
                .Opcao_Masculino.Value = True
            Else
                .Opcao_Feminino.Value = True
            End If
            .Text_Dia.Text = dados.selectSingleNode("dia").Text
            .Text_Mes.Text = dados.selectSingleNode("mes").Text
            .Text_Ano.Text = dados.selectSingleNode("ano").Text
            .Text_Pais.Text = dados.selectSingleNode("pais").Text
            If Imagem = "" Then
                If Sexo = "F" Then
                    .Image_Foto.Picture = Form_Skin.Foto_Feminino.Picture
                Else
                    .Image_Foto.Picture = Form_Skin.Foto_Masculino.Picture
                End If
            End If
            
            .Text_Perfil.Text = dados.selectSingleNode("perfil").Text
            .Text_Servidor.Text = dados.selectSingleNode("servidor").Text
            .Text_Usuario.Text = dados.selectSingleNode("usuario").Text
            .Text_Password.Text = dados.selectSingleNode("senha2").Text
        End With
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Private Sub Pic_Lembrar_Click()
    'Des/Activar a opcção de "Lembrar"
    If Check_Lembrar.Value = 0 Then
        Check_Lembrar.Value = 1
        Pic_Lembrar.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Lembrar.Value = 0
        Pic_Lembrar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Text_Senha_Change()
    'Chamar o procedimento
    Verificar_Preenchimento
End Sub

Private Sub Text_Senha_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Senha.Visible = True
End Sub

Private Sub Text_Senha_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Senha.Visible = False
End Sub

Private Sub Text_Senha_KeyPress(KeyAscii As Integer)
    'Atalho das teclas
    If KeyAscii = vbKeyReturn Then Botao_Ok_Click
End Sub

Private Sub Text_Utilizador_Change()
    'Chamar o procedimento
    Verificar_Preenchimento
End Sub

Private Sub Text_Utilizador_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Utilizador.Visible = True
End Sub

Private Sub Text_Utilizador_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Utilizador.Visible = False
    Text_Utilizador.Text = StrConv(Text_Utilizador.Text, vbProperCase)
End Sub

Private Sub Text_Utilizador_KeyPress(KeyAscii As Integer)
    'Atalho das teclas
    If KeyAscii = vbKeyReturn Then Botao_Ok_Click
End Sub

Public Sub Verificar_Preenchimento()
    'Verifcar o preenchimento da caixa de texto
    If Text_Utilizador.Text = "" Or Text_Senha.Text = "" Then
        Botao_Ok.Enabled = False
        Label_Ok.Enabled = False
    
    Else
        Botao_Ok.Enabled = True
        Label_Ok.Enabled = True
    End If
End Sub

