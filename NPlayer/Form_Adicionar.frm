VERSION 5.00
Begin VB.Form Form_Adicionar 
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   5085
   ClientLeft      =   13095
   ClientTop       =   1725
   ClientWidth     =   6525
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
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   6015
      Begin VB.PictureBox Barra_Text_Data 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   240
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3000
         Width           =   3915
         Begin VB.TextBox Text_Data 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   600
            TabIndex        =   3
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Data 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Artista 
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2280
         Width           =   5475
         Begin VB.TextBox Text_Artista 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   600
            TabIndex        =   2
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Artista 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Titulo 
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5475
         Begin VB.TextBox Text_Titulo 
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
         Begin VB.Shape Contorno_Titulo 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Barra_Text_Servidor 
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   840
         Width           =   5475
         Begin VB.TextBox Text_Servidor 
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
         Begin VB.Shape Contorno_Servidor 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Shape Shape_Centro 
         BorderColor     =   &H00212121&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label_Data 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data de adicionamento:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label_Artista 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Artista:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label_Musica 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label_Servidor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   735
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
      ScaleWidth      =   401
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4440
      Width           =   6015
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
         TabIndex        =   5
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
            TabIndex        =   10
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
         TabIndex        =   4
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
            TabIndex        =   9
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
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Adicionar link"
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
         Width           =   1350
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
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00212121&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Adicionar"
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

Dim Idioma_Info_Link_Add_Success As String
Dim Idioma_Info_Thank_You As String

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
    If Text_Servidor.Text = "" Or Text_Titulo.Text = "" Or Text_Artista.Text = "" Then Exit Sub
    
    Me.MousePointer = 11
    Botao_Ok.Enabled = False
    Label_Ok.Enabled = False
    
    'Enviar o pedido para o servidor
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "adicionarlink.asp?Servidor=" & Text_Servidor.Text & "&Titulo=" & Text_Titulo.Text & "&Artista=" & Text_Artista.Text & "&Adicionado=" & Form_Perfil.Label_Utilizador.Caption & "&Data=" & Text_Data.Text, False
    servidor.send

    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
            Mensagem_de_Aviso "Information", Idioma_Info_Link_Add_Success & vbNewLine & Idioma_Info_Thank_You
            
            'Limpar os campos para adicionar novo link
            Text_Servidor.Text = Empty
            Text_Titulo.Text = Empty
            Text_Artista.Text = Empty
            Botao_Ok.Enabled = False
            Label_Ok.Enabled = False
            Text_Servidor.SetFocus
        End If
    End If
    Me.MousePointer = 0
    Botao_Ok.Enabled = True
    Label_Ok.Enabled = True
    
    
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
Botao_Ok.Enabled = True
Label_Ok.Enabled = True
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
    
    Label_Titulo.Caption = ReadINI("Add_Link", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Add_Link", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Servidor.Caption = ReadINI("Add_Link", "Label_Server", Localizacao_Ficheiro_Lingua)
    Label_Musica.Caption = ReadINI("Add_Link", "Label_Title", Localizacao_Ficheiro_Lingua)
    Label_Artista.Caption = ReadINI("Add_Link", "Label_Artist", Localizacao_Ficheiro_Lingua)
    Label_Data.Caption = ReadINI("Add_Link", "Label_Date", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Add_Link", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Add_Link", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    Idioma_Info_Link_Add_Success = ReadINI("Add_Link", "Info_Link_Add_Success", Localizacao_Ficheiro_Lingua)
    Idioma_Info_Thank_You = ReadINI("Add_Link", "Info_Thank_You", Localizacao_Ficheiro_Lingua)
    
    Idioma_Erro = ReadINI("Message", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Message", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Message", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Conectar_Servidor = ReadINI("Message", "Error_Connect", Localizacao_Ficheiro_Lingua)
    Idioma_Internet_Desligada = ReadINI("Message", "Error_Internet", Localizacao_Ficheiro_Lingua)
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
        Contorno_Servidor.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Titulo.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Artista.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Data.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Servidor.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Musica.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Artista.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Data.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Text_Servidor.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Servidor.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Titulo.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Titulo.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Artista.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Artista.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Data.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Data.ForeColor = .Cor_Letra_Textbox.backcolor
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Ok.Picture = .Pic_Button.Picture
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Barra_Text_Servidor.backcolor = .Cor_Fundo_Textbox.backcolor
        'Barra_Text_Servidor.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Servidor.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Servidor.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Servidor.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Servidor.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Servidor.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Titulo.backcolor = .Cor_Fundo_Textbox.backcolor
        'Barra_Text_Titulo.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Titulo.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Titulo.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Titulo.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Titulo.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Titulo.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Artista.backcolor = .Cor_Fundo_Textbox.backcolor
        'Barra_Text_Artista.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Artista.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Artista.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Artista.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Artista.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Artista.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Data.backcolor = .Cor_Fundo_Textbox.backcolor
        'Barra_Text_Data.Picture = .TextBox_Intermediate.Picture
        Barra_Text_Data.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Data.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Data.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Data.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Data.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
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
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Servidor.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text_Servidor.ScaleWidth) + (2 * Barra_Text_Servidor.left) + 20)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Label_Servidor.left + 3 + Form_Skin.Caixa_de_Texto.Height + 6 _
        + Label_Musica.left + 3 + Form_Skin.Caixa_de_Texto.Height + 6 + Label_Artista.left + 3 + Form_Skin.Caixa_de_Texto.Height + 6 _
        + Label_Data.left + 3 + Form_Skin.Caixa_de_Texto.Height + 6 + Fundo_Frame_Botoes.Height + (2 * Label_Servidor.left))
    End With
    
    Ajustar_Formulario Form_Adicionar, False, False, True, True
    
    Ajustar_Botao Form_Adicionar, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Adicionar, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With

    Ajustar_Caixa_Texto Barra_Text_Servidor, Text_Servidor, Contorno_Servidor, False
    Ajustar_Caixa_Texto Barra_Text_Titulo, Text_Titulo, Contorno_Titulo, False
    Ajustar_Caixa_Texto Barra_Text_Artista, Text_Artista, Contorno_Artista, False
    Ajustar_Caixa_Texto_Media Barra_Text_Data, Text_Data, Contorno_Data
    
    With Label_Servidor
        .top = Label_Servidor.left
    End With
    
    With Barra_Text_Servidor
        .top = Label_Servidor.top + Label_Servidor.Height + 3
        .left = Label_Servidor.left
    End With
    
    With Label_Musica
        .top = Barra_Text_Servidor.top + Barra_Text_Servidor.ScaleHeight + 6
        .left = Label_Servidor.left
    End With
    
    With Barra_Text_Titulo
        .top = Label_Musica.top + Label_Musica.Height + 3
        .left = Label_Servidor.left
    End With
    
    With Label_Artista
        .top = Barra_Text_Titulo.top + Barra_Text_Titulo.ScaleHeight + 6
        .left = Label_Servidor.left
    End With
    
    With Barra_Text_Artista
        .top = Label_Artista.top + Label_Artista.Height + 3
        .left = Label_Servidor.left
    End With
    
    With Label_Data
        .top = Barra_Text_Artista.top + Barra_Text_Artista.ScaleHeight + 6
        .left = Label_Servidor.left
    End With
    
    With Barra_Text_Data
        .top = Label_Data.top + Label_Data.Height + 3
        .left = Label_Servidor.left
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Adicionar
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Adicionar
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Adicionar
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Adicionar
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Adicionar
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Adicionar
End Sub

Private Sub Text_Servidor_Change()
    'Chamar o procedimento
    Verificar_Preenchimento
End Sub

Private Sub Text_Servidor_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Servidor.Visible = True
End Sub

Private Sub Text_Servidor_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Servidor.Visible = False
End Sub

Private Sub Text_Titulo_Change()
    'Chamar o procedimento
    Verificar_Preenchimento
End Sub

Private Sub Text_Titulo_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Titulo.Visible = True
End Sub

Private Sub Text_Titulo_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Titulo.Visible = False
End Sub

Private Sub Text_Artista_Change()
    'Chamar o procedimento
    Verificar_Preenchimento
End Sub

Private Sub Text_Artista_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Artista.Visible = True
End Sub

Private Sub Text_Artista_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Artista.Visible = False
End Sub

Public Sub Verificar_Preenchimento()
    'Verifcar o preenchimento da caixa de texto
    If Text_Artista.Text = "" Or Text_Servidor.Text = "" Or Text_Titulo.Text = "" Then
        Botao_Ok.Enabled = False
        Label_Ok.Enabled = False
    
    Else
        Botao_Ok.Enabled = True
        Label_Ok.Enabled = True
    End If
End Sub

Private Sub Text_Data_Change()
    'Chamar o procedimento
    Verificar_Preenchimento
End Sub
