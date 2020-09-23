VERSION 5.00
Begin VB.Form Form_Codigo 
   Appearance      =   0  'Flat
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   5730
   ClientLeft      =   90
   ClientTop       =   0
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Código fonte"
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
         TabIndex        =   5
         Top             =   120
         Width           =   1230
      End
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   5880
         Picture         =   "Form_Codigo.frx":0000
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   195
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Codigo.frx":0232
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   465
      Width           =   6255
      Begin VB.PictureBox Barra_Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   240
         ScaleHeight     =   201
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   5475
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00808080&
            Height          =   2820
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   30
            Width           =   2940
         End
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
      ScaleWidth      =   376
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4920
      Width           =   5640
      Begin VB.Image Fundo_Frame_Botoes 
         Height          =   615
         Left            =   0
         Picture         =   "Form_Codigo.frx":09B8
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00404040&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Codigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   NPlayer
'   COPYRIGHT © 2011-2012 Nikyts software ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Declaração das variáveis
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

Private Sub Botao_Fechar_Click()
    'Fechar a aplicação
    Unload Me
End Sub

Private Sub Form_Load()
    'Propriedades inicais do formulário
    'Carregar_Idioma
    'Carregar_Skin
    Desenhar_Formulario
    
    Arredondar_Cantos_do_Form Me, True
End Sub

Private Sub Form_Resize()
    'Atalho para
    Desenhar_Formulario
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Codigo
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Codigo
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Codigo
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Codigo
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Codigo
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Codigo
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text1.Width = Form_Skin.Frame_Componentes.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text1.ScaleWidth) + (2 * Barra_Text1.left) + 20)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + 16 + Form_Skin.Frame_Componentes.Height + Fundo_Frame_Botoes.Height _
                + Barra_Text1.left)
    End With
    
    Ajustar_Formulario Form_Codigo, False, False, True, True
    
'    Ajustar_Botao Form_Codigo, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
'    Ajustar_Botao Form_Codigo, Botao_Ok, Label_Ok, True, Contorno_Ok
    
'    With Botao_Cancelar
'        .Left = Frame_Botoes.ScaleWidth - .ScaleWidth - .Top
'    End With
'    With Botao_Ok
'        .Left = Botao_Cancelar.Left - .ScaleWidth - .Top
'    End With
    
    With Barra_Text1
        .left = 16
        .Height = Form_Skin.Frame_Componentes.Height
        .top = .left
        .Width = Form_Skin.Frame_Componentes.Width
    End With
    
    With Text1
        .Height = Barra_Text1.ScaleHeight - 8 - 8
        .top = (Barra_Text1.ScaleHeight - .Height) / 2
        .Width = Barra_Text1.ScaleWidth - 8 - 8
        .left = 8
    End With
End Sub
