VERSION 5.00
Begin VB.Form Form_Sobre 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   3555
   ClientLeft      =   11025
   ClientTop       =   2325
   ClientWidth     =   6570
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
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   960
      Left            =   360
      Picture         =   "Form_Sobre.frx":0000
      Top             =   360
      Width           =   960
   End
   Begin VB.Label Label_Versao 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "versão"
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
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label_Site 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.nikyts.com"
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
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   1950
   End
   Begin VB.Label Label_Web 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "Web:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   450
   End
   Begin VB.Label Label_Direitos 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "© 2011 Nikyts Software"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label_Autor 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "Desenvolvido por: Nelson do Carmo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   3120
   End
   Begin VB.Label Label_Programa 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "NPlayer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   840
   End
   Begin VB.Image Botao_Fechar 
      Height          =   195
      Left            =   6240
      ToolTipText     =   "Fechar"
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "Form_Sobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   NPlayer
'   Copyright © 2011-2012 Nikyts software ™ - Informática e tecnologia
'   www.nikyts.com / nikyts@hotmail.com
'   Desenvolvido por: Nelson do Carmo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'API para abrir web
Private Const SW_NORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Variáveis para o idioma
Dim Idioma_Versao As String
Dim Idioma_Desenvolvido As String
Dim Idioma_Informatica As String

'Variáveis da scroll frame Info
Dim Info_tx As Integer, Info_Ty As Integer, Info_DN As Boolean
Dim Info_Txa As Integer, Info_DNa As Boolean
Dim Info_Tyb, Info_DNb As Boolean
Dim Info_NewY As Integer

'Variável para indicar o número de linhas
Dim Numero_Linhas As Integer
Dim Altura_Linha As Integer

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

Private Sub Botao_Fechar_Click()
    'Fechar o formulario
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Carregar_Idioma
    
    Desenhar_Formulario
    Carregar_Skin
    
    'Definir os valores de x e y para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    'Arredondar os cantos do formulário
    Retangulo Me.hwnd, 5
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Botao_Fechar.ToolTipText = ReadINI("About", "Button_Close", Localizacao_Ficheiro_Lingua)
    Idioma_Versao = ReadINI("About", "Label_Version", Localizacao_Ficheiro_Lingua)
    Idioma_Desenvolvido = ReadINI("About", "Label_Developer", Localizacao_Ficheiro_Lingua)
    Idioma_Informatica = ReadINI("About", "Label_Informatic", Localizacao_Ficheiro_Lingua)
    
    'Propriedades iniciais do formulário
    Label_Programa.Caption = App.ProductName
    Label_Versao.Caption = Idioma_Versao & " " & App.Major & "." & App.Minor & "." & App.Revision
    Label_Direitos.Caption = "© 2011-2012 Nikyts software" & " . " & Idioma_Informatica
    Label_Autor.Caption = Idioma_Desenvolvido & " " & "Nelson do Carmo"
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Me.Picture = .BackGround_Form_About.Picture
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Label_Programa.ForeColor = .Cor_Form_About_Letter.backcolor
        Label_Versao.ForeColor = .Cor_Form_About_Letter.backcolor
        Label_Direitos.ForeColor = .Cor_Form_About_Letter.backcolor
        Label_Autor.ForeColor = .Cor_Form_About_Letter.backcolor
        Label_Web.ForeColor = .Cor_Form_About_Letter.backcolor
        Label_Site.ForeColor = .Cor_Contorno_Caixas.backcolor
    End With
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    With Me
        .Width = Screen.TwipsPerPixelX * (Label_Direitos.Width + (2 * Label_Direitos.left) + 40)
        .Height = Screen.TwipsPerPixelX * Form_Skin.BackGround_Form_About.Height
    End With
        
    With Label_Versao
        .top = Label_Programa.top + Label_Programa.Height + 3
        .left = Label_Programa.left
    End With
    
    With Label_Web
        .top = Me.ScaleHeight - .Height - 20
    End With
    
    With Label_Site
        .top = Label_Web.top
    End With
    
    Dim Ajustar_Botoes As String
    Ajustar_Botoes = "False" 'ReadINI("Dimensions", "Adjust_Button_ControlBox", Localizacao_Ficheiro_Skin)
    With Botao_Fechar
        .Height = Form_Skin.Botao_Fechar.Height
        .Width = Form_Skin.Botao_Fechar.Width
        .left = Me.ScaleWidth - .Width - 6
        If Ajustar_Botoes = "False" Then
            .top = 6
        Else
            .top = 0
        End If
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    'Shape_Contorno.left = 0
    'Shape_Contorno.Width = Me.ScaleWidth - 1
End Sub

Private Sub Form_Resize()
    Desenhar_Formulario
End Sub

Private Sub Label_Site_Click()
    'Abrir página pessoal
    Call ShellExecute(0, "open", Label_Site.Caption, vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Sobre
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Sobre
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Sobre
End Sub
