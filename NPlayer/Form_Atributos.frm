VERSION 5.00
Begin VB.Form Form_Atributos 
   Appearance      =   0  'Flat
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   5550
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   6765
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
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Width           =   6375
      Begin VB.PictureBox Pic_Sistema 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3510
         Width           =   195
      End
      Begin VB.PictureBox Pic_Arquivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3510
         Width           =   195
      End
      Begin VB.PictureBox Pic_Ocultar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3030
         Width           =   195
      End
      Begin VB.PictureBox Pic_Ler 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3030
         Width           =   195
      End
      Begin VB.CheckBox Check_Sistema 
         Appearance      =   0  'Flat
         BackColor       =   &H00313131&
         Caption         =   "Sistema"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   3510
         Width           =   1935
      End
      Begin VB.CheckBox Check_Arquivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00313131&
         Caption         =   "Ficheiro"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   3510
         Width           =   1695
      End
      Begin VB.CheckBox Check_Ocultar 
         Appearance      =   0  'Flat
         BackColor       =   &H00313131&
         Caption         =   "Ocultar"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2760
         TabIndex        =   2
         Top             =   3030
         Width           =   1815
      End
      Begin VB.CheckBox Check_Ler 
         Appearance      =   0  'Flat
         BackColor       =   &H00313131&
         Caption         =   "Ler apenas"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   3030
         Width           =   1815
      End
      Begin VB.PictureBox Barra_Text_Ficheiro 
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
         Top             =   840
         Width           =   5475
         Begin VB.TextBox Text_Ficheiro 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Ficheiro 
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
      Begin VB.Label Label2 
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "qqqqqqqqqqqqq"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   20
         Top             =   1560
         Width           =   4365
      End
      Begin VB.Label Label2 
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "qqqqqqqqqqq"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   19
         Top             =   1920
         Width           =   4275
      End
      Begin VB.Label Label2 
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "qqqqqqqqqqqq"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   18
         Top             =   2280
         Width           =   4380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Tamanho:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   870
      End
      Begin VB.Label Label_Caminho 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ficheiro"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   660
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Propriedades do ficheiro"
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
         TabIndex        =   11
         Top             =   120
         Width           =   2415
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
      ScaleWidth      =   401
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4800
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
            TabIndex        =   9
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
         Begin VB.Label Label_Ok 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ok"
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   750
            TabIndex        =   8
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
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00212121&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Atributos"
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
'Variável para obter as propriedades dos ficheiros
Public Filepath As String, FileSize As Long, Filedate As String

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
    'Atalho para
    Botao_Cancelar_Click
End Sub

Private Sub Check_Arquivo_Click()
    'Des/Activar a opcção
    If Check_Arquivo.Value = 1 Then
        Pic_Arquivo.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Arquivo.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Ler_Click()
    'Des/Activar a opcção
    If Check_Ler.Value = 1 Then
        Pic_Ler.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Ler.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Ocultar_Click()
    'Des/Activar a opcção
    If Check_Ocultar.Value = 1 Then
        Pic_Ocultar.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Ocultar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Sistema_Click()
    'Des/Activar a opcção
    If Check_Sistema.Value = 1 Then
        Pic_Sistema.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Sistema.Picture = Form_Skin.Check_Normal.Picture
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
    
    'Limpar campos
    Label2(1).Caption = ""
    Label2(2).Caption = ""
    Label2(3).Caption = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Attributes", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Attributes", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Caminho.Caption = ReadINI("Attributes", "Label_File", Localizacao_Ficheiro_Lingua)
    Label1(1).Caption = ReadINI("Attributes", "Label_Name", Localizacao_Ficheiro_Lingua)
    Label1(2).Caption = ReadINI("Attributes", "Label_Date", Localizacao_Ficheiro_Lingua)
    Label1(3).Caption = ReadINI("Attributes", "Label_Size", Localizacao_Ficheiro_Lingua)
    Check_Ler.Caption = ReadINI("Attributes", "Check_Read_Only", Localizacao_Ficheiro_Lingua)
    Check_Ocultar.Caption = ReadINI("Attributes", "Check_Hide", Localizacao_Ficheiro_Lingua)
    Check_Arquivo.Caption = ReadINI("Attributes", "Check_File", Localizacao_Ficheiro_Lingua)
    Check_Sistema.Caption = ReadINI("Attributes", "Check_System", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Attributes", "Button_Ok", Localizacao_Ficheiro_Lingua)
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
        Contorno_Ficheiro.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Caminho.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label1(1).ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label1(2).ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label1(3).ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label2(1).ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label2(2).ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label2(3).ForeColor = .Cor_Letra_Label_Formulario.backcolor
        'Barra_Text_Ficheiro.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Ficheiro.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Ficheiro.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Ficheiro.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Ficheiro.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Ficheiro.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Ficheiro.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Ficheiro.ForeColor = .Cor_Letra_Textbox.backcolor
        Check_Ler.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Ler.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Check_Ocultar.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Ocultar.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Check_Arquivo.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Arquivo.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Check_Sistema.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Sistema.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Pic_Ler.Picture = .Check_Normal.Picture
        Pic_Ler.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Pic_Ocultar.Picture = .Check_Normal.Picture
        Pic_Ocultar.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Pic_Arquivo.Picture = .Check_Normal.Picture
        Pic_Arquivo.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Pic_Sistema.Picture = .Check_Normal.Picture
        Pic_Sistema.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
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
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Label_Ok_Click()
    'Aplicar as propriedades alteradas
    Dim userset As Long
    If Check_Arquivo.Value = 1 Then
    userset = userset + 32
    End If
    If Check_Ler.Value = 1 Then
    userset = userset + 1
    End If
    If Check_Ocultar.Value = 1 Then
    userset = userset + 2
    End If
    If Check_Sistema.Value = 1 Then
    userset = userset + 4
    End If
    SetAttr Text_Ficheiro.Text, userset
    Ver_Propriedades
    Unload Me
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Atributos
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Atributos
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Atributos
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Ficheiro.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text_Ficheiro.ScaleWidth) + (2 * Barra_Text_Ficheiro.left) + 20)
'        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Fundo_Frame_Botoes.Height + Frame_Centro.ScaleHeight)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Label_Caminho.left + Label_Caminho.Height + 3 _
                + Form_Skin.Caixa_de_Texto.Height + (2 * Label_Caminho.left) + Label1(1).Height + 6 + Label1(2).Height + 6 + Label1(3).Height _
                + Check_Ler.Height + (2 * Label_Caminho.left) + Check_Sistema.Height + Label_Caminho.left + Fundo_Frame_Botoes.Height + (2 * Label_Caminho.left))
    End With
    
    Ajustar_Formulario Form_Atributos, False, False, True, True
    
    Ajustar_Botao Form_Atributos, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Atributos, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With

    Ajustar_Caixa_Texto Barra_Text_Ficheiro, Text_Ficheiro, Contorno_Ficheiro, False

    With Label_Caminho
        .top = .left
    End With
    
    With Barra_Text_Ficheiro
        .top = Label_Caminho.top + Label_Caminho.Height + 3
        .left = Label_Caminho.left
    End With
    
    With Label1(1)
        .top = Barra_Text_Ficheiro.top + Barra_Text_Ficheiro.ScaleHeight + (2 * Label_Caminho.left)
        .left = Label_Caminho.left
    End With
    
    With Label1(2)
        .top = Label1(1).top + Label1(1).Height + 6
        .left = Label_Caminho.left
    End With
    
    With Label1(3)
        .top = Label1(2).top + Label1(2).Height + 6
        .left = Label_Caminho.left
    End With
    
    With Label2(1)
        .top = Label1(1).top
        .left = Label1(3).left + Label1(1).Width + 16
    End With
    
    With Label2(2)
        .top = Label1(2).top
        .left = Label2(1).left
    End With
    
    With Label2(3)
        .top = Label1(3).top
        .left = Label2(1).left
    End With
    
    With Check_Ler
        .top = Label1(3).top + Label1(3).Height + (2 * Label_Caminho.left)
        .left = Label_Caminho.left
    End With
    
    With Pic_Ler
        .top = Check_Ler.top
        .left = Check_Ler.left
    End With
    
    With Check_Ocultar
        .top = Check_Ler.top
    End With
    
    With Pic_Ocultar
        .top = Check_Ocultar.top
        .left = Check_Ocultar.left
    End With
    
    With Check_Arquivo
        .top = Check_Ler.top + Check_Ler.Height + Label_Caminho.left
        .left = Label_Caminho.left
    End With
    
    With Pic_Arquivo
        .top = Check_Arquivo.top
        .left = Check_Arquivo.left
    End With
    
    With Check_Sistema
        .top = Check_Arquivo.top
    End With
    
    With Pic_Sistema
        .top = Check_Sistema.top
        .left = Check_Sistema.left
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Atributos
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Atributos
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
   Largar_Formulario Form_Atributos
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

Function Ver_Propriedades()
    'Preencher as proproedades adquiridas
    If Text_Ficheiro <> "" Then
    GetInfo (Text_Ficheiro)
    Filepath = Mid(Text_Ficheiro.Text, 1, InStr(1, Text_Ficheiro.Text, Text_Ficheiro.Text, vbTextCompare) - 1)
    FileSize = FileLen(Text_Ficheiro)
    Filedate = FileDateTime(Text_Ficheiro)
'    Label2(0).Caption = Filepath
    Label2(1).Caption = Text_Ficheiro.Text
    Label2(2).Caption = Filedate
    Label2(3).Caption = FileSize & " Byte(s)"
    Text_Ficheiro.Locked = True
    Else
    Text_Ficheiro.Locked = False
    End If
End Function

Function GetInfo(filen As String)
    'Obter as propriedades do ficheiro
    Select Case GetAttr(filen)
    Case 1
    Check_Arquivo.Value = 0
    Check_Ler.Value = 1
    Check_Ocultar.Value = 0
    Check_Sistema.Value = 0
    Case 2
    Check_Arquivo.Value = 0
    Check_Ler.Value = 0
    Check_Ocultar.Value = 1
    Check_Sistema.Value = 0
    Case 3
    Check_Arquivo.Value = 0
    Check_Ler.Value = 1
    Check_Ocultar.Value = 1
    Check_Sistema.Value = 0
    Case 4
    Check_Arquivo.Value = 0
    Check_Ler.Value = 0
    Check_Ocultar.Value = 0
    Check_Sistema.Value = 1
    Case 5
    Check_Arquivo.Value = 0
    Check_Ler.Value = 1
    Check_Ocultar.Value = 0
    Check_Sistema.Value = 1
    Case 6
    Check_Arquivo.Value = 0
    Check_Ler.Value = 0
    Check_Ocultar.Value = 1
    Check_Sistema.Value = 1
    Case 7
    Check_Arquivo.Value = 0
    Check_Ler.Value = 1
    Check_Ocultar.Value = 1
    Check_Sistema.Value = 1
    Case 32
    Check_Arquivo.Value = 1
    Check_Ler.Value = 0
    Check_Ocultar.Value = 0
    Check_Sistema.Value = 0
    Case 33
    Check_Arquivo.Value = 1
    Check_Ler.Value = 1
    Check_Ocultar.Value = 0
    Check_Sistema.Value = 0
    Case 34
    Check_Arquivo.Value = 1
    Check_Ler.Value = 0
    Check_Ocultar.Value = 1
    Check_Sistema.Value = 0
    Case 35
    Check_Arquivo.Value = 1
    Check_Ler.Value = 1
    Check_Ocultar.Value = 1
    Check_Sistema.Value = 0
    Case 36
    Check_Arquivo.Value = 1
    Check_Ler.Value = 0
    Check_Ocultar.Value = 0
    Check_Sistema.Value = 1
    Case 38
    Check_Arquivo.Value = 1
    Check_Ler.Value = 0
    Check_Ocultar.Value = 1
    Check_Sistema.Value = 1
    Case 39
    Check_Arquivo.Value = 1
    Check_Ler.Value = 1
    Check_Ocultar.Value = 1
    Check_Sistema.Value = 1
    End Select
End Function

Private Sub Pic_Arquivo_Click()
    'Des/Activar a opcção
    If Check_Arquivo.Value = 0 Then
        Check_Arquivo.Value = 1
        Pic_Arquivo.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Arquivo.Value = 0
        Pic_Arquivo.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Ler_Click()
    'Des/Activar a opcção
    If Check_Ler.Value = 0 Then
        Check_Ler.Value = 1
        Pic_Ler.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Ler.Value = 0
        Pic_Ler.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Ocultar_Click()
    'Des/Activar a opcção
    If Check_Ocultar.Value = 0 Then
        Check_Ocultar.Value = 1
        Pic_Ocultar.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Ocultar.Value = 0
        Pic_Ocultar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Sistema_Click()
    'Des/Activar a opcção
    If Check_Sistema.Value = 0 Then
        Check_Sistema.Value = 1
        Pic_Sistema.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Sistema.Value = 0
        Pic_Sistema.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Text_Ficheiro_Change()
    'Verificar o preenchimento da caixa de texto
    If Text_Ficheiro.Text = Empty Then
        Botao_Ok.Enabled = False
        Label_Ok.Enabled = False
    Else
        Botao_Ok.Enabled = True
        Label_Ok.Enabled = True
    End If
End Sub

Private Sub Text_Ficheiro_GotFocus()
    'Ver contorno
    Contorno_Ficheiro.Visible = True
End Sub

Private Sub Text_Ficheiro_LostFocus()
    'Ocultar contorno
    Contorno_Ficheiro.Visible = False
End Sub
