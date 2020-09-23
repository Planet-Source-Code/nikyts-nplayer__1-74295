VERSION 5.00
Begin VB.Form Form_Notificacao 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   975
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6105
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   2040
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   1080
      Top             =   600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1560
      Top             =   600
   End
   Begin VB.Image Icon_Info 
      Enabled         =   0   'False
      Height          =   480
      Left            =   120
      Picture         =   "Form_Notificacao.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label_Musica_Adicionada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A música selecionada foi adicionada com sucesso."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   4305
   End
End
Attribute VB_Name = "Form_Notificacao"
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

Option Explicit
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SPI_GETWORKAREA = 48
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Type RECT
  left As Long
  top As Long
  Right As Long
  Bottom As Long
End Type

Private i As Integer
Private NormalWindowStyle As Long
Private TaskBar As Long

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Carregar_Idioma
    
    Me.top = Screen.Height
    Me.left = Screen.Width - Me.Width - 50
    NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_ALPHA
    i = 100
    
    'Ajustar os objectos
    With Icon_Info
        .top = (Me.ScaleHeight - .Height) / 2
    End With
    
    With Label_Musica_Adicionada
        .top = (Me.ScaleHeight - .Height) / 2
    End With
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Musica_Adicionada.Caption = ReadINI("Notification", "Label_Add_Music", Localizacao_Ficheiro_Lingua)
End Sub

Private Sub Timer1_Timer()
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    TaskBar = ((Screen.Height / Screen.TwipsPerPixelX) - WindowRect.Bottom) * Screen.TwipsPerPixelX
    If (Me.top + Me.Height + TaskBar) > Screen.Height Then
        Me.top = Me.top - 30
    Else
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    Timer1.Enabled = False
    Timer3.Enabled = True
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Timer3_Timer()
    If Me.top < Screen.Height And i > 0 Then
        Me.top = Me.top + 30
        i = i - 1.5
    Else
        Unload Me
    End If
End Sub
