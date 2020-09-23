VERSION 5.00
Begin VB.Form Form_Criar 
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   5625
   ClientLeft      =   13095
   ClientTop       =   1725
   ClientWidth     =   6360
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
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00222222&
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
      ScaleWidth      =   449
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Criar conta"
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
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1080
      End
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   5760
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
      AutoRedraw      =   -1  'True
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
      Height          =   660
      Left            =   0
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4920
      Width           =   5775
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
         TabIndex        =   6
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
            TabIndex        =   22
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
         TabIndex        =   5
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
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   750
            TabIndex        =   9
            Top             =   45
            Width           =   240
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
      AutoRedraw      =   -1  'True
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   6135
      Begin VB.PictureBox Barra_Text_Confirmar 
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3000
         Width           =   5475
         Begin VB.TextBox Text_Confirmar 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   600
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Confirmar 
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5475
         Begin VB.TextBox Text_Email 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   600
            TabIndex        =   1
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
         TabIndex        =   19
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2280
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
            TabIndex        =   2
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
      Begin VB.PictureBox Pic_Visualizar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3630
         Width           =   195
      End
      Begin VB.CheckBox Check_Visualizar 
         BackColor       =   &H00313131&
         Caption         =   "Visualizar o conteúdo dos campos"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   3630
         Width           =   4575
      End
      Begin VB.Label Label_Utilizador 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Utilizador"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label_Email 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label_Senha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label_Confirmar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmar senha"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2760
         Width           =   1440
      End
      Begin VB.Label Label_Erro 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         BackStyle       =   0  'Transparent
         Caption         =   "Indique um endereço de email válido."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Image Image_Erro 
         Enabled         =   0   'False
         Height          =   210
         Left            =   300
         Picture         =   "Form_Criar.frx":0000
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
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00212121&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Criar"
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

'Variável para o idioma
Dim Idioma_Erro As String
Dim Idioma_Descricao As String
Dim Idioma_Erro_Execucao As String
Dim Idioma_Conectar_Servidor As String
Dim Idioma_Internet_Desligada As String

Dim Idioma_Error_Email_Invalid As String
Dim Idioma_Error_TextBox_Required As String
Dim Idioma_Error_User_Already_Exists As String
Dim Idioma_Info_Create_Account_Success As String
Dim Idioma_Error_Password_Characters As String

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

Private Sub Botao_Ok_Click()
    'Verifica primeiro o prenchimento da caixas de texto
    On Error GoTo Corrige_Erro
    If Text_Utilizador.Text = "" Or Text_Senha.Text = "" Or Text_Email.Text = "" Or Text_Confirmar.Text = "" Then Exit Sub
    Shape_Erro.Visible = False
    Label_Erro.Visible = False
    Image_Erro.Visible = False
    
    If Not IsEmail(Text_Email.Text) Then
        Label_Erro.Caption = Idioma_Error_Email_Invalid
        Shape_Erro.Visible = True
        Label_Erro.Visible = True
        Image_Erro.Visible = True
        Text_Email.SetFocus
        Exit Sub
    End If

    'Verificar o tamanho da senha
    If Len(Text_Senha.Text) < 6 Or Len(Text_Confirmar.Text) < 6 Then
        Label_Erro.Caption = Idioma_Error_Password_Characters
        Shape_Erro.Visible = True
        Label_Erro.Visible = True
        Image_Erro.Visible = True
        Text_Senha.Text = ""
        Text_Confirmar.Text = ""
        Text_Senha.SetFocus
        Exit Sub
    End If

    'Verificar se a senha é igual á confirmação da senha
    If Text_Senha.Text <> Text_Confirmar.Text Then
        Label_Erro.Caption = Idioma_Error_TextBox_Required
        Shape_Erro.Visible = True
        Label_Erro.Visible = True
        Image_Erro.Visible = True
        Text_Senha.Text = ""
        Text_Confirmar.Text = ""
        Text_Senha.SetFocus
        Exit Sub
    End If
    
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "criar.asp?Utilizador=" & Text_Utilizador.Text & "&Senha=" & Text_Senha.Text & "&Email=" & Text_Email.Text, False
    servidor.send

    'Verificar os dados acesso
    If servidor.responseText = "Existe" Then
        Text_Utilizador.Text = ""
        Text_Senha.Text = ""
        Text_Confirmar.Text = ""
        Text_Email.Text = ""
        Label_Erro.Caption = Idioma_Error_User_Already_Exists
        Shape_Erro.Visible = True
        Label_Erro.Visible = True
        Image_Erro.Visible = True
        Text_Utilizador.SetFocus
        Exit Sub

    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        With Form_Principal
        
            If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then  ' 4 - deu resposta e 200 validou
                Mensagem_de_Aviso "Information", Idioma_Info_Create_Account_Success
                Unload Me
            End If
        End With
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
    'Remover o focus no botao
    Contorno_Ok.Visible = False
End Sub

Private Sub Check_Visualizar_Click()
    'Des/Activar a opcção
    If Check_Visualizar.Value = 1 Then
        Pic_Visualizar.Picture = Form_Skin.Check_Over.Picture
        Text_Senha.PasswordChar = ""
        Text_Confirmar.PasswordChar = ""
        
    Else
        Pic_Visualizar.Picture = Form_Skin.Check_Normal.Picture
        Text_Senha.PasswordChar = "*"
        Text_Confirmar.PasswordChar = "*"
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
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Create_Account", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Create_Account", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Utilizador.Caption = ReadINI("Create_Account", "Label_User", Localizacao_Ficheiro_Lingua)
    Label_Email.Caption = ReadINI("Create_Account", "Label_Email", Localizacao_Ficheiro_Lingua)
    Label_Senha.Caption = ReadINI("Create_Account", "Label_Password", Localizacao_Ficheiro_Lingua)
    Label_Confirmar.Caption = ReadINI("Create_Account", "Label_Confirm", Localizacao_Ficheiro_Lingua)
    Check_Visualizar.Caption = ReadINI("Create_Account", "Check_View", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Create_Account", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Create_Account", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    
    Idioma_Erro = ReadINI("Message", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Message", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Message", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Conectar_Servidor = ReadINI("Message", "Error_Connect", Localizacao_Ficheiro_Lingua)
    Idioma_Internet_Desligada = ReadINI("Message", "Error_Internet", Localizacao_Ficheiro_Lingua)
    
    Idioma_Error_Email_Invalid = ReadINI("Message", "Error_Email_Invalid", Localizacao_Ficheiro_Lingua)
    Idioma_Error_TextBox_Required = ReadINI("Message", "Error_TextBox_Required", Localizacao_Ficheiro_Lingua)
    Idioma_Error_User_Already_Exists = ReadINI("Message", "Error_User_Already_Exists", Localizacao_Ficheiro_Lingua)
    Idioma_Error_Password_Characters = ReadINI("Message", "Error_Password_Characters", Localizacao_Ficheiro_Lingua)
    Idioma_Info_Create_Account_Success = ReadINI("Message", "Info_Create_Account_Success", Localizacao_Ficheiro_Lingua)
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
        Contorno_Email.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Senha.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Confirmar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Utilizador.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Email.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Senha.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Confirmar.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        'Barra_Text_Utilizador.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Utilizador.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Utilizador.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Utilizador.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Utilizador.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Utilizador.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Utilizador.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Utilizador.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Utilizador.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Utilizador.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Email.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Email.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Email.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Email.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Email.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Senha.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Senha.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Senha.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Senha.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Senha.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Senha.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Senha.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Senha.ForeColor = .Cor_Letra_Textbox.backcolor
        'Barra_Text_Confirmar.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Confirmar.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Confirmar.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Confirmar.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Confirmar.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Confirmar.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Confirmar.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Confirmar.ForeColor = .Cor_Letra_Textbox.backcolor
        Check_Visualizar.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Visualizar.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Pic_Visualizar.Picture = .Check_Normal.Picture
        Pic_Visualizar.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
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

Private Sub Label_Ok_Click()
    'Atalho para
    Botao_Ok_Click
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para construir o formulario, ajustando os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Utilizador.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text_Email.ScaleWidth) + (2 * Barra_Text_Email.left) + 20)
'        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Fundo_Frame_Botoes.Height + Frame_Centro.ScaleHeight)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Shape_Erro.left + Shape_Erro.Height + 3 + Label_Utilizador.Height + 3 _
                + Form_Skin.Caixa_de_Texto.Height + 6 + Label_Utilizador.Height + 3 + Form_Skin.Caixa_de_Texto.Height + 6 + Label_Utilizador.Height _
                + 3 + Form_Skin.Caixa_de_Texto.Height + 6 + Label_Utilizador.Height + 3 + Form_Skin.Caixa_de_Texto.Height + Shape_Erro.left _
                + Check_Visualizar.Height + Fundo_Frame_Botoes.Height + (3 * Shape_Erro.left))
    End With
    
    Ajustar_Formulario Form_Criar, False, False, True, True
    
    Ajustar_Botao Form_Criar, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Criar, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With

    Ajustar_Caixa_Texto Barra_Text_Utilizador, Text_Utilizador, Contorno_Utilizador, False
    Ajustar_Caixa_Texto Barra_Text_Email, Text_Email, Contorno_Email, False
    Ajustar_Caixa_Texto Barra_Text_Senha, Text_Senha, Contorno_Senha, False
    Ajustar_Caixa_Texto Barra_Text_Confirmar, Text_Confirmar, Contorno_Confirmar, False
    
    Ajustar_ChecBox Pic_Visualizar, Check_Visualizar
    
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
    
    With Label_Email
        .top = Barra_Text_Utilizador.top + Barra_Text_Utilizador.ScaleHeight + 6
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Email
        .top = Label_Email.top + Label_Email.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Senha
        .top = Barra_Text_Email.top + Barra_Text_Email.ScaleHeight + 6
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Senha
        .top = Label_Senha.top + Label_Senha.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Confirmar
        .top = Barra_Text_Senha.top + Barra_Text_Senha.ScaleHeight + 6
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Confirmar
        .top = Label_Confirmar.top + Label_Confirmar.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Check_Visualizar
        .top = Barra_Text_Confirmar.top + Barra_Text_Confirmar.ScaleHeight + Shape_Erro.left
        .Width = Barra_Text_Confirmar.ScaleWidth
        .left = Shape_Erro.left
    End With
    
    With Pic_Visualizar
        .top = Check_Visualizar.top
        .left = Check_Visualizar.left
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Criar
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Criar
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Criar
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Criar
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Criar
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Criar
End Sub

Private Sub Pic_Visualizar_Click()
    'Des/Activar a opcção
    If Check_Visualizar.Value = 0 Then
        Check_Visualizar.Value = 1
        Pic_Visualizar.Picture = Form_Skin.Check_Over.Picture
        Text_Senha.PasswordChar = ""
        Text_Confirmar.PasswordChar = ""
        
    Else
        Check_Visualizar.Value = 0
        Pic_Visualizar.Picture = Form_Skin.Check_Normal.Picture
        Text_Senha.PasswordChar = "*"
        Text_Confirmar.PasswordChar = "*"
    End If
End Sub

Private Sub Text_Confirmar_Change()
    'Chamar o procedimento
    Verificar_Preenchimento
End Sub

Private Sub Text_Confirmar_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Confirmar.Visible = True
End Sub

Private Sub Text_Confirmar_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Confirmar.Visible = False
End Sub

Private Sub Text_Email_Change()
    'Chamar o procedimento
    Verificar_Preenchimento
End Sub

Private Sub Text_Email_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Email.Visible = True
End Sub

Private Sub Text_Email_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Email.Visible = False
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

Private Sub Text_Utilizador_Change()
    'Chamar o procedimento
    Verificar_Preenchimento
End Sub

Public Sub Verificar_Preenchimento()
    'Verifcar o preenchimento da caixa de texto
    If Text_Utilizador.Text = "" Or Text_Email.Text = "" Or Text_Senha.Text = "" Or Text_Confirmar.Text = "" Then
        Botao_Ok.Enabled = False
        Label_Ok.Enabled = False
    
    Else
        Botao_Ok.Enabled = True
        Label_Ok.Enabled = True
    End If
End Sub

Private Sub Text_Utilizador_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Utilizador.Visible = True
End Sub

Private Sub Text_Utilizador_LostFocus()
    'Contorno da text box ao perder o focus
    Text_Utilizador.Text = StrConv(Text_Utilizador.Text, vbProperCase)
    Contorno_Utilizador.Visible = False
End Sub
