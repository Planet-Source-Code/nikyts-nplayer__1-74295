VERSION 5.00
Begin VB.Form Form_Download 
   Appearance      =   0  'Flat
   BackColor       =   &H00292929&
   BorderStyle     =   0  'None
   Caption         =   "Update"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6405
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
   Icon            =   "Form_Download.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   427
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   6255
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   5475
         Begin VB.TextBox Text_Servidor 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   600
            Locked          =   -1  'True
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
      Begin NPlayer.dl dl 
         Left            =   4680
         Top             =   2040
         _extentx        =   1799
         _extenty        =   1667
      End
      Begin VB.PictureBox Barra_Text_Pasta_Destino 
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
         Top             =   1320
         Width           =   5475
         Begin VB.PictureBox Botao_Pesquisar 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   5040
            Picture         =   "Form_Download.frx":57E2
            ScaleHeight     =   12
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   12
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Selecionar pasta"
            Top             =   120
            Width           =   180
         End
         Begin VB.TextBox Text_Pasta_Destino 
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
         Begin VB.Shape Contorno_Pasta_Destino 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Pic_Fechar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1920
         Width           =   210
      End
      Begin VB.CheckBox Check_Fechar 
         Appearance      =   0  'Flat
         BackColor       =   &H00313131&
         Caption         =   "Iniciar o programa após a actualização"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   2
         Top             =   1920
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin NPlayer.NProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Visible         =   0   'False
         Width           =   5460
         _extentx        =   9631
         _extenty        =   450
      End
      Begin VB.Label Label_Servidor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label_Estado_Transferencia 
         AutoSize        =   -1  'True
         BackColor       =   &H00292929&
         Caption         =   "Estado da transferência"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label Label_Pasta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Localização do programa:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00393939&
         Visible         =   0   'False
         X1              =   24
         X2              =   384
         Y1              =   192
         Y2              =   192
      End
      Begin VB.Shape Shape_Centro 
         BorderColor     =   &H00212121&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   375
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Image Botao_Minimizar 
         Height          =   135
         Left            =   5520
         ToolTipText     =   "Minimizar"
         Top             =   120
         Width           =   135
      End
      Begin VB.Image Icon_do_Programa 
         Enabled         =   0   'False
         Height          =   210
         Left            =   75
         Picture         =   "Form_Download.frx":5B1B
         Top             =   60
         Width           =   210
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Download"
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
         Left            =   480
         TabIndex        =   9
         Top             =   120
         Width           =   945
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
      AutoRedraw      =   -1  'True
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   393
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4680
      Width           =   5895
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
            TabIndex        =   7
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
            Caption         =   "Transferir"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   450
            TabIndex        =   6
            Top             =   45
            Width           =   840
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
         Width           =   1620
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
Attribute VB_Name = "Form_Download"
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

'Variáveis para poder mover o formulário
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'Variável para o idioma
Dim Idioma_Erro As String
Dim Idioma_Descricao As String
Dim Idioma_Erro_Execucao As String
Dim Idioma_Conectar_Servidor As String
Dim Idioma_Error_Download_Program As String

'API's para selecionar uma pasta
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'Tipo para def
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Sub Botao_Cancelar_Click()
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
    'Fechar o formulário
    dl.cancel
    ProgressBar1.Value = 0
    Unload Me
End Sub

Private Sub Botao_Minimizar_Click()
    'Minimizar formulário
    Me.WindowState = 1
End Sub

Public Sub Botao_Ok_Click()
    'atalho para
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
    'Remover o focus no botao
    Contorno_Ok.Visible = False
End Sub

Private Sub Botao_Pesquisar_Click()
    'Selecionar a pas ta do programa
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    'Personaliza a procura
    szTitle = Botao_Pesquisar.ToolTipText
    With tBrowseInfo
        .hWndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_EDITBOX
    End With
    
    'Abre a janela de procura
    'E retorna o caminho da pasta selecionada
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    'Se existir alguma pasta selecionada extrair
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        Text_Pasta_Destino.Text = sBuffer & "\"
    End If
End Sub

Private Sub Check_Fechar_Click()
    'Des/Activar a opcção "fechar automaticamente"
    If Check_Fechar.Value = 1 Then
        Pic_Fechar.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Fechar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Form_Load()
    'Iniciar o formulário
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    
    'Variáveis para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, True
    
    'Carregar a pasta de destino
    If Not ArquivoExiste(App.Path & "\Downloads", True) Then
        MkDir App.Path & "\Downloads\"
    End If
    Text_Pasta_Destino.Text = App.Path & "\Downloads\"
    
    'Nome do formulário
    Me.Caption = Label_Titulo.Caption
    
    'Alterar cores do progreesbar
    ProgressBar1.backcolor = Form_Skin.Cor_Contorno_Caixas.backcolor
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
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
        Botao_Minimizar.Picture = .Botao_Minimizar_Normal.Picture
        Contorno_Pasta_Destino.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Servidor.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Text_Servidor.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Servidor.ForeColor = .Cor_Letra_Textbox.backcolor
        Contorno_Servidor.BorderColor = .Cor_Contorno_Caixas.backcolor
        Barra_Text_Servidor.backcolor = .Cor_Fundo_Textbox.backcolor
        Barra_Text_Servidor.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Servidor.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Servidor.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Servidor.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Servidor.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Label_Pasta.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Estado_Transferencia.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Estado_Transferencia.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Fechar.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Check_Fechar.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Pic_Fechar.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Barra_Text_Pasta_Destino.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Pasta_Destino.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Pasta_Destino.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Pasta_Destino.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Pasta_Destino.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Pasta_Destino.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Pasta_Destino.ForeColor = .Cor_Letra_Textbox.backcolor
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Ok.Picture = .Pic_Button.Picture
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Botao_Pesquisar.Picture = .Botao_Pesquisar.Picture
        Pic_Fechar.Picture = Form_Skin.Check_Over.Picture
        Line1.BorderColor = .Cor_Line_Border_Frames.backcolor
    End With
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Download", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Download", "Button_Close", Localizacao_Ficheiro_Lingua)
    Botao_Minimizar.ToolTipText = ReadINI("Main", "Button_Minimize", Localizacao_Ficheiro_Lingua)
    Label_Servidor.Caption = ReadINI("Download", "Label_Server", Localizacao_Ficheiro_Lingua)
    Label_Pasta.Caption = ReadINI("Download", "Label_Path", Localizacao_Ficheiro_Lingua)
    Check_Fechar.Caption = ReadINI("Download", "Check_Close", Localizacao_Ficheiro_Lingua)
    Idioma_Error_Download_Program = ReadINI("Download", "Error_Download_Program", Localizacao_Ficheiro_Lingua)
    Botao_Pesquisar.ToolTipText = ReadINI("Download", "Select_Folder", Localizacao_Ficheiro_Lingua)
    Label_Estado_Transferencia.Caption = ReadINI("Download", "Label_State", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Download", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Download", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    
    Idioma_Erro = ReadINI("Message", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Message", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Message", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Conectar_Servidor = ReadINI("Message", "Error_Connect", Localizacao_Ficheiro_Lingua)
    Idioma_Internet_Desligada = ReadINI("Message", "Error_Internet", Localizacao_Ficheiro_Lingua)
    Idioma_Mensagem_Enviada = ReadINI("Message", "Info_Posted", Localizacao_Ficheiro_Lingua)
End Sub

Private Sub Pic_Fechar_Click()
    'Des/Activar a opcção "fechar automaticamente"
    If Check_Fechar.Value = 0 Then
        Check_Fechar.Value = 1
        Pic_Fechar.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Fechar.Value = 0
        Pic_Fechar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Me
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Me
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Me
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Me
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Me
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Me
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Pasta_Destino.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text_Pasta_Destino.ScaleWidth) + (2 * Barra_Text_Pasta_Destino.left) + 20)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Label_Servidor.left + Label_Servidor.Height + 3 _
                + Form_Skin.Caixa_de_Texto.Height + 12 + Label_Pasta.Height + 3 + Label_Pasta.Height + Form_Skin.Caixa_de_Texto.Height + 12 _
                + Check_Fechar.Height + Fundo_Frame_Botoes.Height + (2 * Label_Servidor.left) + 90)
    End With
    
    Ajustar_Formulario Me, True, False, True, True
    
    With Botao_Minimizar
        .top = Botao_Fechar.top
        .left = Botao_Fechar.left - .Width - 8
    End With
    
    Ajustar_Botao Me, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Me, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With

    Ajustar_Caixa_Texto Barra_Text_Pasta_Destino, Text_Pasta_Destino, Contorno_Pasta_Destino, False
    Ajustar_Caixa_Texto Barra_Text_Servidor, Text_Servidor, Contorno_Servidor, False
    
    With Botao_Pesquisar
        .Height = Form_Skin.Botao_Pesquisar.Height
        .top = (Barra_Text_Pasta_Destino.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Pesquisar.Width
        .left = Barra_Text_Pasta_Destino.ScaleWidth - .ScaleWidth - .top
    End With
    
    With Label_Servidor
        .left = 16
        .top = .left
    End With
    
    With Barra_Text_Servidor
        .top = Label_Servidor.top + Label_Servidor.Height + 3
        .left = Label_Servidor.left
    End With
    
    With Label_Pasta
        .top = Barra_Text_Servidor.top + Barra_Text_Servidor.ScaleHeight + 12
        .left = Label_Servidor.left
    End With
    
    With Barra_Text_Pasta_Destino
        .top = Label_Pasta.top + Label_Pasta.Height + 3
        .left = Label_Servidor.left
    End With
    
    Ajustar_ChecBox Pic_Fechar, Check_Fechar
    With Check_Fechar
        .top = Barra_Text_Pasta_Destino.top + Barra_Text_Pasta_Destino.ScaleHeight + 12
        .left = Label_Servidor.left
    End With
    
    With Pic_Fechar
        .top = Check_Fechar.top
        .left = Check_Fechar.left
    End With
End Sub

Private Sub Text_Pasta_Destino_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Pasta_Destino.Visible = True
End Sub

Private Sub Text_Pasta_Destino_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Pasta_Destino.Visible = False
End Sub

Private Sub Label_Cancelar_Click()
    'Cancelar o download
    If ProgressBar1.Visible = True Then
        dl.cancel
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False
        'Label_Percentagem.Visible = False
        Label_Estado_Transferencia.Visible = False
        Line1.Visible = False
    End If
    Unload Me
End Sub

Private Sub dl_DowloadComplete()
    'Transferência concluida
    ProgressBar1.Visible = False
    'Label_Percentagem.Visible = False
    Label_Estado_Transferencia.Visible = False
    Line1.Visible = False
    GetFileName (Text_Servidor.Text)
    ProgressBar1.Value = 0
    GetFileName (Text_Servidor.Text)
    Label_Ok.Enabled = True
    Botao_Ok.Enabled = True
    Label_Cancelar.Enabled = True
    Botao_Cancelar.Enabled = False
    
    'Verificar checkbox
    If Check_Fechar.Value = 1 Then
        dl.cancel
        Unload Me
    End If
    
    Me.MousePointer = 0
End Sub

Private Sub dl_DownloadErrors(strError As String)
    'Caso ocorra um erro durante o download
    Mensagem_de_Aviso "Error", Idioma_Error_Download_Program
    ProgressBar1.Visible = False
    'Label_Percentagem.Visible = False
    Label_Estado_Transferencia.Visible = False
    Line1.Visible = False
    Label_Ok.Enabled = True
    Botao_Ok.Enabled = True
    Me.MousePointer = 0
End Sub

Private Sub dl_DownloadEvents(strEvent As String)
    'Evento do download
    Label_Ok.Enabled = False
    Botao_Ok.Enabled = False
    Label_Cancelar.Enabled = True
    Botao_Cancelar.Enabled = True
End Sub

Private Sub dl_DownloadProgress(intPercent As String)
    'Mostrar o progresso do download
    ProgressBar1.Value = intPercent
    'Label_Percentagem.Caption = ProgressBar1.Value & " %"
    GetFileName (Text_Servidor.Text)
    Label_Cancelar.Enabled = True
End Sub

Private Sub Label_Ok_Click()
    'Verificar se a pasta dos downloads existe
    If Not ArquivoExiste(App.Path & "\Downloads", True) Then
        MkDir App.Path & "\Downloads\"
    End If
    
    'Efectuar o download da musica
    Me.MousePointer = 11
    Label_Ok.Enabled = False
    Botao_Ok.Enabled = False
    Label_Estado_Transferencia.Visible = True
    Line1.Visible = True
    ProgressBar1.Visible = True
    'Label_Percentagem.Visible = True
    dl.DownloadFile Text_Servidor.Text, Text_Pasta_Destino.Text & GetFileName(Text_Servidor.Text)
    On Error GoTo 0 'tratamento de erros.
    
Exit Sub
errHand:
End Sub

Private Sub Text_Servidor_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Servidor.Visible = True
End Sub

Private Sub Text_Servidor_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Servidor.Visible = False
End Sub

