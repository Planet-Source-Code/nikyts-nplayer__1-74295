VERSION 5.00
Begin VB.Form Form_Skin 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FF80&
   ClientHeight    =   12630
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   21690
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
   Icon            =   "Form_Skin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   842
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1446
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic_Text_Web 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   18120
      Picture         =   "Form_Skin.frx":57E2
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   960
      Width           =   1935
   End
   Begin VB.PictureBox Pic_TextBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   16800
      Picture         =   "Form_Skin.frx":5FF0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image Menu_Check_Over 
      Height          =   240
      Left            =   12360
      Picture         =   "Form_Skin.frx":72E2
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image Icon_Link 
      Height          =   720
      Left            =   11280
      Picture         =   "Form_Skin.frx":742C
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image Image_Down_Processando 
      Height          =   240
      Left            =   9600
      Picture         =   "Form_Skin.frx":8F6E
      Top             =   4080
      Width           =   240
   End
   Begin VB.Image Image_Down_Concluido 
      Height          =   240
      Left            =   9600
      Picture         =   "Form_Skin.frx":92B0
      Top             =   4440
      Width           =   240
   End
   Begin VB.Image Image_Down_Erro 
      Height          =   240
      Left            =   9600
      Picture         =   "Form_Skin.frx":95F2
      Top             =   4800
      Width           =   240
   End
   Begin VB.Image Foto_Programa 
      Height          =   1860
      Left            =   10800
      Picture         =   "Form_Skin.frx":9934
      Top             =   3360
      Width           =   3075
   End
   Begin VB.Image Botao_Programas 
      Height          =   435
      Left            =   14040
      Picture         =   "Form_Skin.frx":BF1E
      Top             =   3960
      Width           =   1365
   End
   Begin VB.Image Botao_Linha_Normal 
      Height          =   375
      Left            =   8040
      Picture         =   "Form_Skin.frx":C478
      Top             =   3120
      Width           =   2160
   End
   Begin VB.Image Botao_Linha_2_Normal 
      Height          =   375
      Left            =   8040
      Picture         =   "Form_Skin.frx":CA0D
      Top             =   4080
      Width           =   1440
   End
   Begin VB.Image Botao_Linha_Over 
      Height          =   375
      Left            =   8040
      Picture         =   "Form_Skin.frx":E66F
      Top             =   3600
      Width           =   2160
   End
   Begin VB.Image Botao_Linha_2_Over 
      Height          =   375
      Left            =   8040
      Picture         =   "Form_Skin.frx":110E1
      Top             =   4560
      Width           =   1440
   End
   Begin VB.Image Image_Estrelas_0 
      Height          =   480
      Left            =   0
      Picture         =   "Form_Skin.frx":12D43
      Top             =   9360
      Width           =   2400
   End
   Begin VB.Image Image_Estrelas_1 
      Height          =   480
      Left            =   0
      Picture         =   "Form_Skin.frx":16985
      Top             =   9840
      Width           =   2400
   End
   Begin VB.Image Image_Estrelas_2 
      Height          =   480
      Left            =   0
      Picture         =   "Form_Skin.frx":1A5C7
      Top             =   10320
      Width           =   2400
   End
   Begin VB.Image Image_Estrelas_3 
      Height          =   480
      Left            =   0
      Picture         =   "Form_Skin.frx":1E209
      Top             =   10800
      Width           =   2400
   End
   Begin VB.Image Image_Estrelas_4 
      Height          =   480
      Left            =   0
      Picture         =   "Form_Skin.frx":21E4B
      Top             =   11280
      Width           =   2400
   End
   Begin VB.Image Image_Estrelas_5 
      Height          =   480
      Left            =   0
      Picture         =   "Form_Skin.frx":25A8D
      Top             =   11760
      Width           =   2400
   End
   Begin VB.Image Linha_Normal 
      Height          =   645
      Left            =   6480
      Picture         =   "Form_Skin.frx":296CF
      Top             =   3120
      Width           =   1485
   End
   Begin VB.Image Linha_Over 
      Height          =   1290
      Left            =   6480
      Picture         =   "Form_Skin.frx":2C975
      Top             =   3840
      Width           =   1380
   End
   Begin VB.Image Icon_Mensagem_Down 
      Height          =   180
      Left            =   15240
      Picture         =   "Form_Skin.frx":3266F
      Top             =   1440
      Width           =   225
   End
   Begin VB.Image Icon_Mensagem_Normal 
      Height          =   180
      Left            =   14880
      Picture         =   "Form_Skin.frx":328F1
      Top             =   1440
      Width           =   225
   End
   Begin VB.Image Button_Menu_Standard_Down 
      Height          =   300
      Left            =   16200
      Picture         =   "Form_Skin.frx":32B73
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Button_Menu_Standard_Normal 
      Height          =   300
      Left            =   16200
      Picture         =   "Form_Skin.frx":33515
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Icon_Invitation 
      Height          =   720
      Left            =   12120
      Picture         =   "Form_Skin.frx":33EB7
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   14040
      Picture         =   "Form_Skin.frx":348C2
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Image Bar_View_Cover1 
      Height          =   360
      Left            =   14040
      Picture         =   "Form_Skin.frx":34CC5
      Top             =   3120
      Width           =   870
   End
   Begin VB.Image Botao_Barra_Down 
      Height          =   330
      Left            =   7920
      Picture         =   "Form_Skin.frx":35D87
      Top             =   2400
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Image Botao_Barra_Normal 
      Height          =   330
      Left            =   7920
      Picture         =   "Form_Skin.frx":37949
      Top             =   2040
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Image Icon_Seta_Anterior_Down 
      Height          =   330
      Left            =   6480
      Picture         =   "Form_Skin.frx":3950B
      Top             =   2400
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Icon_Seta_Seguinte_Down 
      Height          =   330
      Left            =   6960
      Picture         =   "Form_Skin.frx":39CDD
      Top             =   2400
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Icon_Seta_Anterior_Normal 
      Height          =   330
      Left            =   6480
      Picture         =   "Form_Skin.frx":3A507
      Top             =   2040
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Icon_Seta_Seguinte_Normal 
      Height          =   330
      Left            =   6960
      Picture         =   "Form_Skin.frx":3ACD9
      Top             =   2040
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Icon_Home_Normal 
      Height          =   330
      Left            =   7440
      Picture         =   "Form_Skin.frx":3B503
      Top             =   2040
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image_Precos 
      Height          =   4560
      Left            =   12840
      Picture         =   "Form_Skin.frx":3BD2D
      Top             =   9600
      Width           =   4020
   End
   Begin VB.Image Icon_Home_Down 
      Height          =   330
      Left            =   7440
      Picture         =   "Form_Skin.frx":7782F
      Top             =   2400
      Width           =   450
   End
   Begin VB.Image Icon_Topico_Drive_Over 
      Height          =   225
      Left            =   16560
      Picture         =   "Form_Skin.frx":78059
      Top             =   2520
      Width           =   210
   End
   Begin VB.Image Icon_Topico_Drive_Normal 
      Height          =   225
      Left            =   16320
      Picture         =   "Form_Skin.frx":78279
      Top             =   2520
      Width           =   210
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu_ForeColorSel"
      Height          =   195
      Index           =   9
      Left            =   480
      TabIndex        =   95
      Top             =   9000
      Width           =   1665
   End
   Begin VB.Label Cor_Menu_ForeColorSel 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   94
      Top             =   9000
      Width           =   255
   End
   Begin VB.Label Cor_Menu_BackColor 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   93
      Top             =   8280
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu_BackColor"
      Height          =   195
      Index           =   8
      Left            =   480
      TabIndex        =   92
      Top             =   8280
      Width           =   1440
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu_ForeColor"
      Height          =   195
      Index           =   12
      Left            =   480
      TabIndex        =   91
      Top             =   8670
      Width           =   1395
   End
   Begin VB.Label Cor_Menu_ForeColor 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   90
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Cor_Menu_BackColorSel 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   89
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu_BackColorSel"
      Height          =   195
      Index           =   11
      Left            =   480
      TabIndex        =   88
      Top             =   7950
      Width           =   1710
   End
   Begin VB.Image Icon_Topico_Musica_Over 
      Enabled         =   0   'False
      Height          =   225
      Left            =   16560
      Picture         =   "Form_Skin.frx":784E8
      Top             =   2760
      Width           =   210
   End
   Begin VB.Image Icon_Topico_Musica_Normal 
      Enabled         =   0   'False
      Height          =   225
      Left            =   16320
      Picture         =   "Form_Skin.frx":787F1
      Top             =   2760
      Width           =   210
   End
   Begin VB.Image Icon_Topico_Filmes_Normal 
      Enabled         =   0   'False
      Height          =   225
      Left            =   16320
      Picture         =   "Form_Skin.frx":78B3F
      Top             =   3000
      Width           =   210
   End
   Begin VB.Image Icon_Topico_Filmes_Over 
      Enabled         =   0   'False
      Height          =   225
      Left            =   16560
      Picture         =   "Form_Skin.frx":78E98
      Top             =   3000
      Width           =   210
   End
   Begin VB.Image Icon_Topico_radio_Normal 
      Height          =   240
      Left            =   16320
      Picture         =   "Form_Skin.frx":791F1
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image Icon_Topico_radio_Over 
      Height          =   240
      Left            =   16560
      Picture         =   "Form_Skin.frx":79531
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image Icon_Topico_Programas_Normal 
      Enabled         =   0   'False
      Height          =   240
      Left            =   16320
      Picture         =   "Form_Skin.frx":79845
      Top             =   4080
      Width           =   210
   End
   Begin VB.Image Icon_Topico_Minha_Musica_Normal 
      Enabled         =   0   'False
      Height          =   225
      Left            =   16320
      Picture         =   "Form_Skin.frx":79BA5
      Top             =   3810
      Width           =   210
   End
   Begin VB.Image Icon_Topico_MusicLink_Normal 
      Enabled         =   0   'False
      Height          =   240
      Left            =   16320
      Picture         =   "Form_Skin.frx":79F12
      Top             =   3480
      Width           =   210
   End
   Begin VB.Image Icon_Topico_Programas_Over 
      Enabled         =   0   'False
      Height          =   240
      Left            =   16560
      Picture         =   "Form_Skin.frx":7A263
      Top             =   4080
      Width           =   210
   End
   Begin VB.Image Icon_Topico_Minha_Musica_Over 
      Enabled         =   0   'False
      Height          =   225
      Left            =   16560
      Picture         =   "Form_Skin.frx":7A5C6
      Top             =   3810
      Width           =   210
   End
   Begin VB.Image Icon_Topico_MusicLink_Over 
      Enabled         =   0   'False
      Height          =   240
      Left            =   16560
      Picture         =   "Form_Skin.frx":7A908
      Top             =   3480
      Width           =   210
   End
   Begin VB.Image Icon_Topico_Lista_Normal 
      Enabled         =   0   'False
      Height          =   240
      Left            =   16320
      Picture         =   "Form_Skin.frx":7AC52
      Top             =   4320
      Width           =   210
   End
   Begin VB.Image Icon_Topico_Lista_Over 
      Enabled         =   0   'False
      Height          =   240
      Left            =   16560
      Picture         =   "Form_Skin.frx":7AFA5
      Top             =   4320
      Width           =   210
   End
   Begin VB.Image Slide_Album 
      Height          =   240
      Left            =   15000
      Top             =   4560
      Width           =   390
   End
   Begin VB.Image Image_Album_Over 
      Enabled         =   0   'False
      Height          =   3240
      Left            =   12720
      Picture         =   "Form_Skin.frx":7B2F5
      Top             =   6240
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Image Image_Album 
      Enabled         =   0   'False
      Height          =   3240
      Left            =   12240
      Picture         =   "Form_Skin.frx":9CF37
      Top             =   6000
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label Cor_Line_Border_Frames 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   87
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Line_Border_Frames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_Frame_Cover"
      Height          =   195
      Index           =   11
      Left            =   3600
      TabIndex        =   86
      Top             =   360
      Width           =   1725
   End
   Begin VB.Image Select_Topic_TaskBar 
      Height          =   300
      Left            =   6480
      Picture         =   "Form_Skin.frx":BEB79
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Image Icon_Subtitles_Down 
      Height          =   210
      Left            =   16560
      Picture         =   "Form_Skin.frx":BF27A
      Top             =   5160
      Width           =   225
   End
   Begin VB.Image Icon_Subtitles_Normal 
      Height          =   210
      Left            =   16560
      Picture         =   "Form_Skin.frx":BF5D4
      Top             =   4800
      Width           =   225
   End
   Begin VB.Image Button_Arrow_Left_Down 
      Height          =   330
      Left            =   9960
      Picture         =   "Form_Skin.frx":BF928
      Top             =   2400
      Width           =   390
   End
   Begin VB.Image Button_Arrow_Right_Down 
      Height          =   330
      Left            =   9600
      Picture         =   "Form_Skin.frx":BFC99
      Top             =   2400
      Width           =   390
   End
   Begin VB.Image Button_Arrow_Left_Normal 
      Height          =   330
      Left            =   9960
      Picture         =   "Form_Skin.frx":C0038
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image Button_Arrow_Right_Normal 
      Height          =   330
      Left            =   9600
      Picture         =   "Form_Skin.frx":C0393
      Top             =   2040
      Width           =   390
   End
   Begin VB.Image Bar_Menu 
      Height          =   435
      Left            =   15600
      Picture         =   "Form_Skin.frx":C0753
      Top             =   3360
      Width           =   675
   End
   Begin VB.Image Pic_Button 
      Height          =   390
      Left            =   14760
      Picture         =   "Form_Skin.frx":C0A26
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form_About_Letter"
      Height          =   195
      Index           =   0
      Left            =   3600
      TabIndex        =   84
      Top             =   750
      Width           =   1635
   End
   Begin VB.Label Cor_Form_About_Letter 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   83
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_Frame_Cover"
      Height          =   195
      Index           =   10
      Left            =   480
      TabIndex        =   82
      Top             =   390
      Width           =   1725
   End
   Begin VB.Label Cor_Label_Frame_Cover 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   81
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Cor_BackGround_Frame_Cover 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   80
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackGround_Frame_Cover"
      Height          =   195
      Index           =   9
      Left            =   480
      TabIndex        =   79
      Top             =   750
      Width           =   2325
   End
   Begin VB.Image Slide_Som_Down 
      Height          =   75
      Left            =   8640
      Picture         =   "Form_Skin.frx":C0F94
      Top             =   1680
      Width           =   180
   End
   Begin VB.Image Botao_Seguinte_Down 
      Height          =   330
      Left            =   8160
      Picture         =   "Form_Skin.frx":C1261
      ToolTipText     =   "Faixa seguinte"
      Top             =   480
      Width           =   465
   End
   Begin VB.Image Botao_Antes_Down 
      Height          =   330
      Left            =   6480
      Picture         =   "Form_Skin.frx":C162E
      ToolTipText     =   "Faixa anterior"
      Top             =   480
      Width           =   465
   End
   Begin VB.Image Botao_Play_Down 
      Height          =   330
      Left            =   6960
      Picture         =   "Form_Skin.frx":C1A19
      ToolTipText     =   "Reproduzir"
      Top             =   480
      Width           =   555
   End
   Begin VB.Image Botao_Pausa_Down 
      Height          =   330
      Left            =   7560
      Picture         =   "Form_Skin.frx":C1E18
      ToolTipText     =   "Pausa"
      Top             =   480
      Width           =   555
   End
   Begin VB.Image Button_Folder_Down 
      Height          =   300
      Left            =   15000
      Picture         =   "Form_Skin.frx":C21D9
      Top             =   480
      Width           =   600
   End
   Begin VB.Image Button_Folder_Normal 
      Height          =   300
      Left            =   15000
      Picture         =   "Form_Skin.frx":C25F3
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Button_Menu_Down 
      Height          =   300
      Left            =   11640
      Picture         =   "Form_Skin.frx":C2A2D
      Top             =   480
      Width           =   2100
   End
   Begin VB.Image Button_Music_Repete_Over 
      Height          =   300
      Left            =   9840
      Picture         =   "Form_Skin.frx":C2F8C
      Top             =   840
      Width           =   585
   End
   Begin VB.Image Button_Music_Randomize_Over 
      Height          =   300
      Left            =   9240
      Picture         =   "Form_Skin.frx":C33FE
      Top             =   840
      Width           =   585
   End
   Begin VB.Image Button_Playlist_View_Down 
      Height          =   300
      Left            =   14400
      Picture         =   "Form_Skin.frx":C3878
      ToolTipText     =   "Ocultar capa"
      Top             =   480
      Width           =   585
   End
   Begin VB.Image Button_Playlist_View_Normal 
      Height          =   300
      Left            =   14400
      Picture         =   "Form_Skin.frx":C3CAF
      ToolTipText     =   "Ocultar capa"
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Button_Cover_Hide_Down 
      Height          =   300
      Left            =   11040
      Picture         =   "Form_Skin.frx":C40FD
      ToolTipText     =   "Ocultar capa"
      Top             =   480
      Width           =   585
   End
   Begin VB.Image Button_Cover_Hide_Normal 
      Height          =   300
      Left            =   11040
      Picture         =   "Form_Skin.frx":C4533
      ToolTipText     =   "Ocultar capa"
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Button_Playlist_Hide_Normal 
      Height          =   300
      Left            =   13800
      Picture         =   "Form_Skin.frx":C497D
      ToolTipText     =   "Ocultar playlist"
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Button_Playlist_Hide_Down 
      Height          =   300
      Left            =   13800
      Picture         =   "Form_Skin.frx":C4DCD
      Top             =   480
      Width           =   585
   End
   Begin VB.Image Button_Music_Repete_Over_Down 
      Height          =   300
      Left            =   9840
      Picture         =   "Form_Skin.frx":C5203
      Top             =   1200
      Width           =   585
   End
   Begin VB.Image Button_Music_Randomize_Over_Down 
      Height          =   300
      Left            =   9240
      Picture         =   "Form_Skin.frx":C5661
      Top             =   1200
      Width           =   585
   End
   Begin VB.Image Button_Music_Repete_Down 
      Height          =   300
      Left            =   9840
      Picture         =   "Form_Skin.frx":C5AC8
      Top             =   480
      Width           =   585
   End
   Begin VB.Image Button_Music_Randomize_Down 
      Height          =   300
      Left            =   9240
      Picture         =   "Form_Skin.frx":C5F04
      Top             =   480
      Width           =   585
   End
   Begin VB.Image Button_Cover_View_Down 
      Height          =   300
      Left            =   10440
      Picture         =   "Form_Skin.frx":C6332
      ToolTipText     =   "Ocultar capa"
      Top             =   480
      Width           =   585
   End
   Begin VB.Image Button_Cover_View_Normal 
      Height          =   300
      Left            =   10440
      Picture         =   "Form_Skin.frx":C675E
      ToolTipText     =   "Ocultar capa"
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Button_Music_Randomize_Normal 
      Height          =   300
      Left            =   9240
      Picture         =   "Form_Skin.frx":C6B9F
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Button_Music_Repete_Normal 
      Height          =   300
      Left            =   9840
      Picture         =   "Form_Skin.frx":C6FE4
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Button_New_Playlist_Normal 
      Height          =   300
      Left            =   8640
      Picture         =   "Form_Skin.frx":C7420
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Button_New_Playlist_Down 
      Height          =   300
      Left            =   8640
      Picture         =   "Form_Skin.frx":C785C
      Top             =   480
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form_BorderColor"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   78
      Top             =   5760
      Width           =   1590
   End
   Begin VB.Label Cor_Form_BorderColor 
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   77
      Top             =   5760
      Width           =   255
   End
   Begin VB.Image Close_Wmp 
      Height          =   570
      Left            =   15360
      Picture         =   "Form_Skin.frx":C7C6D
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image Icon_Visao_Down 
      Height          =   330
      Index           =   2
      Left            =   7200
      Picture         =   "Form_Skin.frx":C8C1F
      ToolTipText     =   "Album art"
      Top             =   1200
      Width           =   390
   End
   Begin VB.Image Icon_Visao_Down 
      Height          =   330
      Index           =   1
      Left            =   6840
      Picture         =   "Form_Skin.frx":C8FFD
      ToolTipText     =   "Pesquisa avançada"
      Top             =   1200
      Width           =   390
   End
   Begin VB.Image Icon_Visao_Down 
      Height          =   330
      Index           =   0
      Left            =   6480
      Picture         =   "Form_Skin.frx":C93E5
      ToolTipText     =   "Simples"
      Top             =   1200
      Width           =   405
   End
   Begin VB.Image Icon_Visao_Normal 
      Height          =   330
      Index           =   2
      Left            =   7200
      Picture         =   "Form_Skin.frx":C9783
      ToolTipText     =   "Album art"
      Top             =   840
      Width           =   390
   End
   Begin VB.Image Icon_Visao_Normal 
      Height          =   330
      Index           =   1
      Left            =   6840
      Picture         =   "Form_Skin.frx":C9B5A
      ToolTipText     =   "Pesquisa avançada"
      Top             =   840
      Width           =   390
   End
   Begin VB.Image Icon_Visao_Normal 
      Height          =   330
      Index           =   0
      Left            =   6480
      Picture         =   "Form_Skin.frx":C9F34
      ToolTipText     =   "Simples"
      Top             =   840
      Width           =   405
   End
   Begin VB.Image Slide_Mini 
      Height          =   120
      Left            =   15480
      Picture         =   "Form_Skin.frx":CA30A
      Top             =   2520
      Width           =   120
   End
   Begin VB.Image Slide_Som_Mini 
      Height          =   165
      Left            =   15240
      Picture         =   "Form_Skin.frx":CA40C
      Top             =   2520
      Width           =   165
   End
   Begin VB.Image SliderBar_Mini 
      Height          =   150
      Left            =   15720
      Top             =   600
      Width           =   6300
   End
   Begin VB.Image Som_Off_Normal_Mini 
      Height          =   180
      Left            =   12720
      Top             =   1080
      Width           =   195
   End
   Begin VB.Image Som_On_Normal_Mini 
      Height          =   180
      Left            =   12960
      Top             =   1080
      Width           =   195
   End
   Begin VB.Image Button_Menu_Normal 
      Height          =   300
      Left            =   11640
      Picture         =   "Form_Skin.frx":CA5DA
      Top             =   120
      Width           =   2100
   End
   Begin VB.Image Icon_FullScreen_Off 
      Height          =   255
      Left            =   14880
      Picture         =   "Form_Skin.frx":CABC1
      ToolTipText     =   "Tela cheia"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Icon_FullScreen_On 
      Height          =   255
      Left            =   14880
      Picture         =   "Form_Skin.frx":CAF77
      ToolTipText     =   "Tela cheia"
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image Picture_Slide_Som_Mini 
      Height          =   165
      Left            =   19080
      Top             =   360
      Width           =   1650
   End
   Begin VB.Image Fundo_Barra_Mini_Player 
      Height          =   885
      Left            =   17400
      Top             =   6720
      Width           =   6660
   End
   Begin VB.Label Cor_BackGround_Bar_Label_Button 
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   76
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackGround_Bar_Label_Buton"
      Height          =   195
      Index           =   8
      Left            =   3600
      TabIndex        =   75
      Top             =   1110
      Width           =   2610
   End
   Begin VB.Image Fundo_Barra_Botoes_Musica 
      Enabled         =   0   'False
      Height          =   360
      Left            =   17400
      Top             =   1680
      Width           =   2865
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_Button_ForeColor"
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   74
      Top             =   1110
      Width           =   2055
   End
   Begin VB.Label Cor_Label_Button_ForeColor 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Topic_Task_Bar"
      Height          =   195
      Index           =   6
      Left            =   3600
      TabIndex        =   72
      Top             =   1470
      Width           =   1365
   End
   Begin VB.Label Cor_Topic_Task_Bar 
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   71
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Cor_Form_Main 
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   70
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo form main"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   69
      Top             =   1470
      Width           =   1455
   End
   Begin VB.Image Fundo_Barra_Actualizar 
      Height          =   540
      Left            =   15600
      Picture         =   "Form_Skin.frx":CB32D
      Top             =   4320
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slider music"
      Height          =   195
      Index           =   7
      Left            =   3600
      TabIndex        =   68
      Top             =   7560
      Width           =   1050
   End
   Begin VB.Label Cor_Slider_Music 
      BackColor       =   &H00222222&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   67
      Top             =   7560
      Width           =   255
   End
   Begin VB.Image Botao_Actualizar_Programa 
      Height          =   390
      Left            =   14040
      Picture         =   "Form_Skin.frx":CC0EF
      Top             =   5040
      Width           =   2355
   End
   Begin VB.Image Fundo_Frame_Caixa_Pesquisa 
      Height          =   780
      Left            =   17400
      Picture         =   "Form_Skin.frx":CF121
      Top             =   7680
      Width           =   6780
   End
   Begin VB.Label Cor_Download_Add_Ons_Buttons 
      BackColor       =   &H00222222&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   66
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Download_Add_Ons_Buttons"
      Height          =   195
      Index           =   6
      Left            =   3600
      TabIndex        =   65
      Top             =   7200
      Width           =   2460
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor_Display"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   64
      Top             =   7560
      Width           =   1620
   End
   Begin VB.Label Cor_BackColor_Display 
      BackColor       =   &H00A2AFA7&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   63
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letter_Download_Add_Ons"
      Height          =   195
      Index           =   4
      Left            =   3600
      TabIndex        =   62
      Top             =   6840
      Width           =   2310
   End
   Begin VB.Label Cor_Letter_Download_Add_Ons 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   61
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor_Download_Add_Ons"
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   60
      Top             =   6480
      Width           =   2700
   End
   Begin VB.Label Cor_BackColor_Download_Add_Ons 
      BackColor       =   &H00212121&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   59
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image Estrela_Over 
      Height          =   165
      Left            =   8400
      Picture         =   "Form_Skin.frx":E04D3
      Top             =   1560
      Width           =   165
   End
   Begin VB.Image Estrela_Normal 
      Height          =   165
      Left            =   8160
      Picture         =   "Form_Skin.frx":E07EF
      Top             =   1560
      Width           =   165
   End
   Begin VB.Label Cor_Contorno_Frame_Centro 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   58
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contorno frame centro"
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   57
      Top             =   6120
      Width           =   1965
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   424
      X2              =   424
      Y1              =   8
      Y2              =   640
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   " Cores do programa "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   55
      Top             =   0
      Width           =   2205
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   424
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   -8
      X2              =   424
      Y1              =   640
      Y2              =   640
   End
   Begin VB.Label Cor_Letra_Bar_Info 
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   54
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra barra informação"
      Height          =   195
      Index           =   4
      Left            =   3600
      TabIndex        =   53
      Top             =   2190
      Width           =   1980
   End
   Begin VB.Image Fundo_Barra_Informacoes 
      Height          =   435
      Left            =   20160
      Picture         =   "Form_Skin.frx":E09F2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fundo task bar"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   52
      Top             =   1830
      Width           =   1245
   End
   Begin VB.Label Cor_Fundo_Task_Bar 
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Cor_Letra_Tab_Normal 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra tab normal"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   49
      Top             =   2190
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra tab over"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   48
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Label Cor_Letra_Tab_Over 
      BackColor       =   &H00B67B26&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Cor_Grid_ColorFixed 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   46
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_ColorFixed"
      Height          =   195
      Index           =   0
      Left            =   3600
      TabIndex        =   45
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Image Fundo_Barra_Setup 
      Height          =   1200
      Left            =   1080
      Top             =   9840
      Visible         =   0   'False
      Width           =   5475
   End
   Begin VB.Image Separador_Mini_Normal 
      Height          =   420
      Left            =   14160
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Image Fundo_Barra_Lateral 
      Height          =   810
      Left            =   6600
      Top             =   10200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image Fundo_Frame_Album 
      Enabled         =   0   'False
      Height          =   4830
      Left            =   0
      Top             =   9840
      Width           =   975
   End
   Begin VB.Image Menu_Check_Normal 
      Height          =   240
      Left            =   12120
      Picture         =   "Form_Skin.frx":E0D86
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image Background_Scroll_Bar 
      Height          =   165
      Left            =   13200
      Picture         =   "Form_Skin.frx":E0ED0
      Top             =   840
      Width           =   210
   End
   Begin VB.Image Scroll_Info_Slider_Barras 
      Enabled         =   0   'False
      Height          =   75
      Left            =   13260
      Picture         =   "Form_Skin.frx":E1186
      Top             =   1440
      Width           =   105
   End
   Begin VB.Image Scroll_Info_Up 
      Height          =   240
      Left            =   13200
      Picture         =   "Form_Skin.frx":E1240
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Scroll_Info_Down 
      Height          =   240
      Left            =   13200
      Picture         =   "Form_Skin.frx":E1582
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label Cor_Scroll_Bar 
      BackColor       =   &H00212121&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   44
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo scroll bar"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   43
      Top             =   1830
      Width           =   1365
   End
   Begin VB.Image Opcao_Over 
      Height          =   180
      Left            =   9360
      Picture         =   "Form_Skin.frx":E18C4
      Top             =   1620
      Width           =   180
   End
   Begin VB.Image Opcao_Normal 
      Height          =   180
      Left            =   9120
      Picture         =   "Form_Skin.frx":E1C29
      Top             =   1620
      Width           =   180
   End
   Begin VB.Image Bar_View_Cover 
      Height          =   270
      Left            =   14040
      Picture         =   "Form_Skin.frx":E1F54
      Top             =   2760
      Width           =   4440
   End
   Begin VB.Image TextBox_Intermediate 
      Height          =   390
      Left            =   9840
      Top             =   10200
      Width           =   2535
   End
   Begin VB.Image Moldura_Pic_Tag_Editor 
      Height          =   3075
      Left            =   18960
      Top             =   11880
      Width           =   3105
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra topico over"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   42
      Top             =   7230
      Width           =   1455
   End
   Begin VB.Label Cor_Letra_Topico_Over 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Cor_Fundo_Topico_Over 
      BackColor       =   &H00484947&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo topico over"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   39
      Top             =   6870
      Width           =   1530
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra topico normal"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   38
      Top             =   6510
      Width           =   1665
   End
   Begin VB.Label Cor_Letra_Topico_Normal 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo topico normal"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   36
      Top             =   6150
      Width           =   1740
   End
   Begin VB.Label Cor_Fundo_Topico_Normal 
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_Color"
      Height          =   195
      Left            =   3600
      TabIndex        =   34
      Top             =   5040
      Width           =   1350
   End
   Begin VB.Image Image_Form_Mini_Player 
      Height          =   3645
      Left            =   6600
      Top             =   11160
      Width           =   5670
   End
   Begin VB.Image Fundo_Mini_Player 
      Height          =   3285
      Left            =   13200
      Picture         =   "Form_Skin.frx":E5E06
      Top             =   1920
      Width           =   720
   End
   Begin VB.Image Caixa_de_Texto_Mini 
      Height          =   390
      Left            =   9840
      Top             =   10680
      Width           =   915
   End
   Begin VB.Image Moldura_Foto 
      Height          =   2070
      Left            =   18960
      Top             =   9720
      Width           =   2085
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Não pertence ao skin. Seve apenas como medida."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   0
      Left            =   7680
      TabIndex        =   33
      Top             =   13680
      Width           =   4965
   End
   Begin VB.Image Image_Capa 
      Height          =   750
      Left            =   13920
      Picture         =   "Form_Skin.frx":E6215
      Top             =   840
      Width           =   750
   End
   Begin VB.Image Icon_Form_Down 
      Height          =   210
      Left            =   10440
      Picture         =   "Form_Skin.frx":E66C5
      Top             =   1620
      Width           =   240
   End
   Begin VB.Image Icon_Form_Up 
      Height          =   210
      Left            =   10080
      Picture         =   "Form_Skin.frx":E69A7
      Top             =   1620
      Width           =   240
   End
   Begin VB.Image Icon_Ver_Biblioteca 
      Height          =   240
      Left            =   10800
      Picture         =   "Form_Skin.frx":E6C89
      Top             =   1620
      Width           =   240
   End
   Begin VB.Image Icon_Ocultar_Janela 
      Height          =   240
      Left            =   11160
      Picture         =   "Form_Skin.frx":E6FCB
      Top             =   1620
      Width           =   240
   End
   Begin VB.Image Fundo_Frame_Botoes 
      Height          =   615
      Left            =   13560
      Picture         =   "Form_Skin.frx":E730D
      Top             =   840
      Width           =   405
   End
   Begin VB.Image Fundo_Barra_ControlBox 
      Height          =   360
      Left            =   12840
      Picture         =   "Form_Skin.frx":E7615
      Top             =   1440
      Width           =   285
   End
   Begin VB.Image Botao_Fechar 
      Height          =   195
      Left            =   7680
      Picture         =   "Form_Skin.frx":E78BE
      Top             =   840
      Width           =   180
   End
   Begin VB.Image Icon_Ver_Janela 
      Height          =   240
      Left            =   11520
      Picture         =   "Form_Skin.frx":E7BC4
      Top             =   1620
      Width           =   240
   End
   Begin VB.Image Fundo_PopUp_P 
      Height          =   990
      Left            =   6480
      Picture         =   "Form_Skin.frx":E7F06
      Top             =   5400
      Width           =   5460
   End
   Begin VB.Image Foto_Masculino 
      Height          =   1920
      Left            =   6720
      Picture         =   "Form_Skin.frx":E835C
      Top             =   6960
      Width           =   1920
   End
   Begin VB.Image Image_Barra_Slide_Mini 
      Enabled         =   0   'False
      Height          =   150
      Left            =   15720
      Picture         =   "Form_Skin.frx":EA5FD
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3255
   End
   Begin VB.Image Fundo_Barra_Faixa_Mini 
      Height          =   1770
      Left            =   12120
      Picture         =   "Form_Skin.frx":EABFC
      Top             =   6720
      Width           =   5220
   End
   Begin VB.Image Frame_Componentes 
      Height          =   3060
      Left            =   1080
      Top             =   11160
      Width           =   5475
   End
   Begin VB.Image Caixa_de_Observacoes 
      Height          =   1905
      Left            =   12360
      Top             =   11160
      Width           =   5475
   End
   Begin VB.Image Seta_Combo 
      Height          =   315
      Left            =   8880
      Picture         =   "Form_Skin.frx":EB9B8
      Top             =   1080
      Width           =   285
   End
   Begin VB.Label Cor_Label_Barra_Titulo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo da barra de titulo"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   31
      Top             =   3270
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Label_Barra_Visor"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   30
      Top             =   5790
      Width           =   1995
   End
   Begin VB.Label Cor_Label_Barra_Visor 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo caixas de texto"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   28
      Top             =   4350
      Width           =   1875
   End
   Begin VB.Label Cor_Fundo_Textbox 
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Width           =   255
   End
   Begin VB.Image Som_On_Normal 
      Height          =   180
      Left            =   12960
      Picture         =   "Form_Skin.frx":EBCD8
      Top             =   840
      Width           =   210
   End
   Begin VB.Image Som_Off_Normal 
      Height          =   180
      Left            =   12720
      Picture         =   "Form_Skin.frx":EC00B
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Cor_Contorno_Caixas 
      BackColor       =   &H00CBB534&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor da progressbar"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   25
      Top             =   3990
      Width           =   1680
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_BackColorFixed"
      Height          =   195
      Index           =   6
      Left            =   3600
      TabIndex        =   24
      Top             =   3270
      Width           =   2220
   End
   Begin VB.Label Cor_Grid_BackColorFixed 
      BackColor       =   &H00313131&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Cor_Grid_BackColorBkg 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_BackColorBkg"
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   21
      Top             =   2910
      Width           =   2100
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Letra_Label_Formulario"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   20
      Top             =   3630
      Width           =   2430
   End
   Begin VB.Label Cor_Letra_Label_Formulario 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra do botao"
      Height          =   195
      Index           =   11
      Left            =   480
      TabIndex        =   18
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label Cor_da_Letra_do_Botao 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo dos formulários"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   16
      Top             =   2910
      Width           =   1905
   End
   Begin VB.Label Cor_do_Fundo_dos_Formularios 
      BackColor       =   &H00313131&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Cor_Grid_Color 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_ForeColor"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   13
      Top             =   3990
      Width           =   1725
   End
   Begin VB.Label Cor_Grid_ForeColor 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Cor_Grid_BackColor 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_BackColor"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   10
      Top             =   2550
      Width           =   1770
   End
   Begin VB.Label Cor_Grid_BackColorSel 
      BackColor       =   &H00B67B26&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_BackColorSel"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   8
      Top             =   3630
      Width           =   2040
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_ForeColorSel"
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   4710
      Width           =   1995
   End
   Begin VB.Label Cor_Grid_ForeColorSel 
      BackColor       =   &H00E6E6E6&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_ForeColorFixed"
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   5
      Top             =   4350
      Width           =   2175
   End
   Begin VB.Label Cor_Grid_ForeColorFixed 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   4320
      Width           =   255
   End
   Begin VB.Image Icon_Info 
      Enabled         =   0   'False
      Height          =   720
      Left            =   11880
      Picture         =   "Form_Skin.frx":EC2E9
      Top             =   840
      Width           =   720
   End
   Begin VB.Image Icon_Quest 
      Enabled         =   0   'False
      Height          =   720
      Left            =   11160
      Picture         =   "Form_Skin.frx":ED07B
      Top             =   840
      Width           =   720
   End
   Begin VB.Image Icon_Error 
      Enabled         =   0   'False
      Height          =   720
      Left            =   10440
      Picture         =   "Form_Skin.frx":EDED5
      Top             =   840
      Width           =   720
   End
   Begin VB.Image Caixa_Pesquisar_Musica 
      Height          =   390
      Left            =   17400
      Picture         =   "Form_Skin.frx":EECF9
      Top             =   2160
      Width           =   2625
   End
   Begin VB.Image Fundo_Slider_Volume 
      Height          =   240
      Left            =   6480
      Picture         =   "Form_Skin.frx":EF507
      Top             =   1560
      Width           =   1650
   End
   Begin VB.Image Slide_Som_Normal 
      Height          =   75
      Left            =   8640
      Picture         =   "Form_Skin.frx":EF928
      Top             =   1560
      Width           =   180
   End
   Begin VB.Image Slide_Musica_Normal 
      Height          =   75
      Left            =   11880
      Picture         =   "Form_Skin.frx":EFC06
      Top             =   1680
      Width           =   75
   End
   Begin VB.Image Botao_Pausa_Normal 
      Height          =   330
      Left            =   7560
      Picture         =   "Form_Skin.frx":EFEC1
      ToolTipText     =   "Pausa"
      Top             =   120
      Width           =   555
   End
   Begin VB.Image Botao_Play_Normal 
      Height          =   330
      Left            =   6960
      Picture         =   "Form_Skin.frx":F02AE
      ToolTipText     =   "Reproduzir"
      Top             =   120
      Width           =   555
   End
   Begin VB.Image Botao_Antes_Normal 
      Height          =   330
      Left            =   6480
      Picture         =   "Form_Skin.frx":F06C5
      ToolTipText     =   "Faixa anterior"
      Top             =   120
      Width           =   465
   End
   Begin VB.Image Botao_Seguinte_Normal 
      Height          =   330
      Left            =   8160
      Picture         =   "Form_Skin.frx":F0ABE
      ToolTipText     =   "Faixa seguinte"
      Top             =   120
      Width           =   465
   End
   Begin VB.Image Botao_Form 
      Height          =   390
      Left            =   6600
      Top             =   9720
      Width           =   1440
   End
   Begin VB.Image Check_Normal 
      Height          =   195
      Left            =   9600
      Picture         =   "Form_Skin.frx":F0EA0
      Top             =   1620
      Width           =   195
   End
   Begin VB.Image Check_Over 
      Height          =   195
      Left            =   9840
      Picture         =   "Form_Skin.frx":F11C4
      Top             =   1620
      Width           =   195
   End
   Begin VB.Label Cor_Letra_Textbox 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra caixas de texto"
      Height          =   195
      Index           =   13
      Left            =   480
      TabIndex        =   2
      Top             =   4710
      Width           =   1800
   End
   Begin VB.Image Botao_Minimizar_Normal 
      Height          =   195
      Left            =   8400
      Picture         =   "Form_Skin.frx":F1510
      ToolTipText     =   "Minimizar"
      Top             =   840
      Width           =   180
   End
   Begin VB.Image Botao_Restaurar_Normal 
      Height          =   195
      Left            =   8160
      Picture         =   "Form_Skin.frx":F1816
      ToolTipText     =   "Restaurar"
      Top             =   840
      Width           =   180
   End
   Begin VB.Image Botao_Maximizar_Normal 
      Height          =   195
      Left            =   7920
      Picture         =   "Form_Skin.frx":F1B1C
      ToolTipText     =   "Maximizar"
      Top             =   840
      Width           =   180
   End
   Begin VB.Image Botao_Tray_Normal 
      Height          =   195
      Left            =   8640
      Picture         =   "Form_Skin.frx":F1E22
      ToolTipText     =   "Colocar o ícone na bandeja"
      Top             =   840
      Width           =   180
   End
   Begin VB.Image Fundo_Barra_Faixa 
      Enabled         =   0   'False
      Height          =   825
      Left            =   12120
      Picture         =   "Form_Skin.frx":F2128
      Top             =   5760
      Width           =   8355
   End
   Begin VB.Image Image_Barra_Slide 
      Height          =   120
      Left            =   15720
      Picture         =   "Form_Skin.frx":F3140
      Top             =   120
      Width           =   6300
   End
   Begin VB.Image Fundo_Barra_Player 
      Enabled         =   0   'False
      Height          =   1050
      Left            =   14040
      Picture         =   "Form_Skin.frx":F373F
      Top             =   1680
      Width           =   600
   End
   Begin VB.Image Botao_Pesquisar 
      Height          =   195
      Left            =   8880
      Picture         =   "Form_Skin.frx":F3AA0
      ToolTipText     =   "Pesquisar"
      Top             =   840
      Width           =   195
   End
   Begin VB.Image Caixa_de_Texto 
      Height          =   390
      Left            =   10920
      Top             =   10680
      Width           =   5475
   End
   Begin VB.Label Cor_Label_Contador_Popup 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Label_Contador_Popup"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   0
      Top             =   5430
      Width           =   2400
   End
   Begin VB.Image Image_Sem_Capa 
      Height          =   3000
      Left            =   16920
      Picture         =   "Form_Skin.frx":F3DD0
      Top             =   2760
      Width           =   3000
   End
   Begin VB.Image Foto_Feminino 
      Height          =   1920
      Left            =   9360
      Picture         =   "Form_Skin.frx":F4C61
      Top             =   6960
      Width           =   1920
   End
   Begin VB.Image Fundo_PopUp_G 
      Height          =   2415
      Left            =   6480
      Picture         =   "Form_Skin.frx":F734A
      Top             =   6720
      Width           =   5460
   End
   Begin VB.Image BackGround_Form_About 
      Height          =   3375
      Left            =   6480
      Picture         =   "Form_Skin.frx":F79CE
      Top             =   1920
      Width           =   6600
   End
   Begin VB.Menu Menu_Icon 
      Caption         =   "Menu_Icon"
      Begin VB.Menu Menu_Fechar 
         Caption         =   "Fechar"
      End
   End
End
Attribute VB_Name = "Form_Skin"
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
Dim Localizacao_Pasta_Skin As String

Public Sub Carregar_Imagens_do_Skin()
    'On Error Resume Next
'    'Procedimento para carregar as imagens do skins escolhido
'    Localizacao_Pasta_Skin = App.Path & "\Skins\" & Form_Preferencias.Text_Skin.Text & "\Images\"
'
'    'Imagens
'    Botao_Fechar.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Close.jpg")
'    Fundo_Barra_ControlBox.Picture = LoadPicture(Localizacao_Pasta_Skin & "Bar_Control_Box.jpg")
'    Fundo_Frame_Botoes.Picture = LoadPicture(Localizacao_Pasta_Skin & "Bar_Buttons.jpg")
'    Pic_Button.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Normal.jpg")
'    Seta_Combo.Picture = LoadPicture(Localizacao_Pasta_Skin & "Arrow_ComboBox.jpg")
'    Check_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Check_Normal.jpg")
'    Check_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Check_Over.jpg")
'    Icon_Form_Up.Picture = LoadPicture(Localizacao_Pasta_Skin & "Arrow_Form_Up.jpg")
'    Icon_Form_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Arrow_Form_Down.jpg")
'    Icon_Ver_Biblioteca.Picture = LoadPicture(Localizacao_Pasta_Skin & "View_Library.jpg")
'    Icon_Ocultar_Janela.Picture = LoadPicture(Localizacao_Pasta_Skin & "Hide_Window.jpg")
'    Icon_Ver_Janela.Picture = LoadPicture(Localizacao_Pasta_Skin & "View_Window.jpg")
'    Fundo_PopUp_G.Picture = LoadPicture(Localizacao_Pasta_Skin & "Form_Popup_Custom.jpg")
'    Fundo_PopUp_P.Picture = LoadPicture(Localizacao_Pasta_Skin & "Form_Popup_Simple.jpg")
'    Image_Capa.Picture = LoadPicture(Localizacao_Pasta_Skin & "No_Cover_Mini.jpg")
'    Botao_Play_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Play_Normal.jpg")
'    Botao_Pausa_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Pause_Normal.jpg")
'    Botao_Antes_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Previous_Normal.jpg")
'    Botao_Seguinte_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Next_Normal.jpg")
'    Botao_Play_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Play_Down.jpg")
'    Botao_Pausa_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Pause_Down.jpg")
'    Botao_Antes_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Previous_Down.jpg")
'    Botao_Seguinte_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Next_Down.jpg")
'    Som_Off_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Sound_Off.jpg")
'    Som_On_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Sound_On.jpg")
'    Slide_Som_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Slide_Sound_Normal.jpg")
'    Slide_Som_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Slide_Sound_Down.jpg")
'    Fundo_Slider_Volume.Picture = LoadPicture(Localizacao_Pasta_Skin & "Bar_Slide_Sound.jpg")
'    Icon_Quest.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Quest.jpg")
'    Icon_Error.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Error.jpg")
'    Icon_Info.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Info.jpg")
'    Botao_Maximizar_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Maximize.jpg")
'    Botao_Restaurar_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Restore.jpg")
'    Botao_Minimizar_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Minimize.jpg")
'    Botao_Tray_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Tray.jpg")
'    Botao_Pesquisar.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_TextBox_Search.jpg")
'    Slide_Musica_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Slide_Music.jpg")
'    Estrela_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Star_Normal.jpg")
'    Estrela_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Star_Over.jpg")
'    Fundo_Barra_Player.Picture = LoadPicture(Localizacao_Pasta_Skin & "Background_Bar_Buttons_Player.jpg")
'    Foto_Masculino.Picture = LoadPicture(Localizacao_Pasta_Skin & "Photo_Male.jpg")
'    Foto_Feminino.Picture = LoadPicture(Localizacao_Pasta_Skin & "Photo_Female.jpg")
'    Opcao_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Option_Normal.jpg")
'    Opcao_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Option_Over.jpg")
'    Fundo_Mini_Player.Picture = LoadPicture(Localizacao_Pasta_Skin & "Background_Form_Mini_Player.jpg")
'    Bar_View_Cover.Picture = LoadPicture(Localizacao_Pasta_Skin & "Bar_View_Cover.jpg")
'    Scroll_Info_Up.Picture = LoadPicture(Localizacao_Pasta_Skin & "Scroll_Up.jpg")
'    Scroll_Info_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Scroll_Down.jpg")
'    Scroll_Info_Slider_Barras.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Center_Bar_Scroll.jpg")
'    Background_Scroll_Bar.Picture = LoadPicture(Localizacao_Pasta_Skin & "Background_Scroll_Bar.jpg")
'    Active_Menu.Picture = LoadPicture(Localizacao_Pasta_Skin & "Active_Menu.jpg")
'    Fundo_Barra_Informacoes.Picture = LoadPicture(Localizacao_Pasta_Skin & "Bar_Information.jpg")
'    Pic_TextBox.Picture = LoadPicture(Localizacao_Pasta_Skin & "TextBox.bmp")
'    Image_Sem_Capa.Picture = LoadPicture(Localizacao_Pasta_Skin & "No_Cover_Large.jpg")
'    Image_Capa.Picture = LoadPicture(Localizacao_Pasta_Skin & "No_Cover_Mini.jpg")
'    Fundo_Barra_Faixa.Picture = LoadPicture(Localizacao_Pasta_Skin & "Background_Display.jpg")
'    Image_Barra_Slide.Picture = LoadPicture(Localizacao_Pasta_Skin & "Background_Slider_Music.jpg")
'    Caixa_Pesquisar_Musica.Picture = LoadPicture(Localizacao_Pasta_Skin & "TextBox_Search.bmp")
'    Pic_Text_Web.Picture = LoadPicture(Localizacao_Pasta_Skin & "TextBox_Search.bmp")
'    Button_New_Playlist_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_New_Playlist_Normal.jpg")
'    Button_New_Playlist_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_New_Playlist_Down.jpg")
'    Button_Music_Randomize_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Music_Randomize_Normal.jpg")
'    Button_Music_Randomize_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Music_Randomize_Down.jpg")
'    Button_Music_Randomize_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Music_Randomize_Over.jpg")
'    Button_Music_Randomize_Over_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Music_Randomize_Over_Down.jpg")
'    Button_Music_Repete_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Music_Repete_Normal.jpg")
'    Button_Music_Repete_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Music_Repete_Down.jpg")
'    Button_Music_Repete_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Music_Repete_Over.jpg")
'    Button_Music_Repete_Over_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Music_Repete_Over_Down.jpg")
'    Button_Cover_View_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Cover_View_Normal.jpg")
'    Button_Cover_View_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Cover_View_Down.jpg")
'    Button_Cover_Hide_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Cover_Hide_Normal.jpg")
'    Button_Cover_Hide_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Cover_Hide_Down.jpg")
'    Button_Menu_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Menu_Normal.jpg")
'    Button_Menu_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Menu_Down.jpg")
'    Icon_Subtitles_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Subtitles_Normal.jpg")
'    Icon_Subtitles_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Subtitles_Down.jpg")
'    Button_Playlist_Hide_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Playlist_Hide_Normal.jpg")
'    Button_Playlist_Hide_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Playlist_Hide_Down.jpg")
'    Button_Playlist_View_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Playlist_View_Normal.jpg")
'    Button_Playlist_View_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Playlist_View_Down.jpg")
'    Button_Folder_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Folder_Normal.jpg")
'    Button_Folder_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Folder_Down.jpg")
'    Icon_Visao_Normal(0).Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_View_Grid_Normal.jpg")
'    Icon_Visao_Normal(1).Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_View_Advanced_Normal.jpg")
'    Icon_Visao_Normal(2).Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_View_Album_Normal.jpg")
'    Icon_Visao_Down(0).Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_View_Grid_Down.jpg")
'    Icon_Visao_Down(1).Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_View_Advanced_Down.jpg")
'    Icon_Visao_Down(2).Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_View_Album_Down.jpg")
'    BackGround_Form_About.Picture = LoadPicture(Localizacao_Pasta_Skin & "BackGround_Form_About.jpg")
'    Bar_Menu.Picture = LoadPicture(Localizacao_Pasta_Skin & "Bar_Menu.jpg")
'    Button_Arrow_Right_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Arrow_Right_Normal.jpg")
'    Button_Arrow_Right_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Arrow_Right_Down.jpg")
'    Button_Arrow_Left_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Arrow_Left_Normal.jpg")
'    Button_Arrow_Left_Down.Picture = LoadPicture(Localizacao_Pasta_Skin & "Button_Arrow_Left_Down.jpg")
'    Select_Topic_TaskBar.Picture = LoadPicture(Localizacao_Pasta_Skin & "Select_Topic_TaskBar.jpg")
'    Icon_Topico_Musica_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Music_Normal.jpg")
'    Icon_Topico_Musica_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Music_Over.jpg")
'    Icon_Topico_Filmes_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Movies_Normal.jpg")
'    Icon_Topico_Filmes_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Movies_Over.jpg")
'    Icon_Topico_radio_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Radio_Normal.jpg")
'    Icon_Topico_radio_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Radio_Over.jpg")
'    Icon_Topico_Procurar_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Find_Normal.jpg")
'    Icon_Topico_Procurar_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Find_Over.jpg")
'    Icon_Topico_Minha_Musica_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_My_Music_Normal.jpg")
'    Icon_Topico_Minha_Musica_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_My_Music_Over.jpg")
'    Icon_Topico_Resultado_Pesquisa_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Result_Normal.jpg")
'    Icon_Topico_Resultado_Pesquisa_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Result_Over.jpg")
'    Icon_Topico_Lista_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_List_Normal.jpg")
'    Icon_Topico_Lista_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_List_Over.jpg")
'    Fundo_Barra_Faixa_Mini.Picture = LoadPicture(Localizacao_Pasta_Skin & "Background_Display_Mini.jpg")
'    Icon_Topico_Drive_Normal.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Drive_Normal.jpg")
'    Icon_Topico_Drive_Over.Picture = LoadPicture(Localizacao_Pasta_Skin & "Icon_Drive_Over.jpg")
    
    'Cores
    Cor_Form_Main.BackColor = RGB(49, 49, 49)
    Cor_Form_BorderColor.BackColor = RGB(192, 192, 192)
    Cor_do_Fundo_dos_Formularios.BackColor = RGB(238, 238, 238)
    Cor_Label_Barra_Titulo.BackColor = RGB(255, 255, 255)
    Cor_Letra_Label_Formulario.BackColor = RGB(0, 0, 0)
    Cor_Label_Contador_Popup.BackColor = RGB(153, 153, 153)
    Cor_Contorno_Caixas.BackColor = RGB(52, 181, 203)
    Cor_Fundo_Textbox.BackColor = RGB(255, 255, 255)
    Cor_Letra_Textbox.BackColor = RGB(0, 0, 0)
    Cor_da_Letra_do_Botao.BackColor = RGB(0, 0, 0)
    Cor_Label_Barra_Visor.BackColor = RGB(255, 255, 255)
    Cor_Grid_BackColor.BackColor = RGB(255, 255, 255)
    Cor_Grid_BackColorBkg.BackColor = RGB(255, 255, 255)
    Cor_Grid_BackColorFixed.BackColor = RGB(199, 199, 201)
    Cor_Grid_BackColorSel.BackColor = RGB(52, 181, 203)
    Cor_Grid_ForeColor.BackColor = RGB(0, 0, 0)
    Cor_Grid_ForeColorFixed.BackColor = RGB(46, 53, 69)
    Cor_Grid_ForeColorSel.BackColor = RGB(255, 255, 255)
    Cor_Grid_Color.BackColor = RGB(223, 223, 223)
    Cor_Grid_ColorFixed.BackColor = RGB(230, 230, 230)
    Cor_Fundo_Topico_Normal.BackColor = RGB(218, 222, 231)
    Cor_Letra_Topico_Normal.BackColor = RGB(0, 0, 0)
    Cor_Fundo_Topico_Over.BackColor = RGB(84, 84, 84)
    Cor_Letra_Topico_Over.BackColor = RGB(255, 255, 255)
    Cor_Scroll_Bar.BackColor = RGB(32, 34, 31)
    Cor_Letra_Tab_Normal.BackColor = RGB(255, 255, 255)
    Cor_Letra_Tab_Over.BackColor = RGB(52, 181, 203)
    Cor_Fundo_Task_Bar.BackColor = RGB(218, 222, 231)
    Cor_Letra_Bar_Info.BackColor = RGB(255, 255, 255)
    Cor_Contorno_Frame_Centro.BackColor = RGB(238, 238, 238)
    Cor_BackColor_Download_Add_Ons.BackColor = RGB(33, 33, 33)
    Cor_Letter_Download_Add_Ons.BackColor = RGB(255, 255, 255)
    Cor_BackColor_Display.BackColor = RGB(40, 40, 40)
    Cor_Download_Add_Ons_Buttons.BackColor = RGB(49, 49, 49)
    Cor_Slider_Music.BackColor = RGB(123, 123, 123)
    Cor_Topic_Task_Bar.BackColor = RGB(121, 123, 148)
    Cor_Label_Button_ForeColor.BackColor = RGB(49, 49, 49)
    Cor_BackGround_Bar_Label_Button.BackColor = RGB(255, 255, 255)
    Cor_BackGround_Frame_Cover.BackColor = RGB(255, 255, 255)
    Cor_Label_Frame_Cover.BackColor = RGB(128, 128, 128)
    Cor_Form_About_Letter.BackColor = RGB(255, 255, 255)
    Cor_Line_Border_Frames.BackColor = RGB(192, 192, 192)
'    Cor_Menu_BackColorSel.backcolor = RGB(218, 222, 231)
'    Cor_Menu_BackColor.backcolor = RGB(22, 22, 22)
'    Cor_Menu_ForeColor.backcolor = RGB(255, 255, 255)
'    Cor_Menu_ForeColorSel.backcolor = RGB(0, 0, 0)

'    Possibilitar vários skins
'    Cor_Form_Main.backcolor = Ler_RGB("Colors", "Form_Main_Background")
'    Cor_Form_BorderColor.backcolor = Ler_RGB("Colors", "Form_BorderColor")
'    Cor_do_Fundo_dos_Formularios.backcolor = Ler_RGB("Colors", "Form_Background")
'    Cor_Label_Barra_Titulo.backcolor = Ler_RGB("Colors", "Form_Title")
'    Cor_Letra_Label_Formulario.backcolor = Ler_RGB("Colors", "Form_Letter")
'    Cor_Label_Contador_Popup.backcolor = Ler_RGB("Colors", "Form_Letter_2")
'    Cor_Contorno_Caixas.backcolor = Ler_RGB("Colors", "Outline_Objects")
'    Cor_Fundo_Textbox.backcolor = Ler_RGB("Colors", "TextBox_Background")
'    Cor_Letra_Textbox.backcolor = Ler_RGB("Colors", "TextBox_Letter")
'    Cor_da_Letra_do_Botao.backcolor = Ler_RGB("Colors", "Button_Letter")
'    Cor_Label_Barra_Visor.backcolor = Ler_RGB("Colors", "Display_Letter")
'    Cor_Grid_BackColor.backcolor = Ler_RGB("Colors", "Grid_BackColor")
'    Cor_Grid_BackColorBkg.backcolor = Ler_RGB("Colors", "Grid_BackColorBkg")
'    Cor_Grid_BackColorFixed.backcolor = Ler_RGB("Colors", "Grid_BackColorFixed")
'    Cor_Grid_BackColorSel.backcolor = Ler_RGB("Colors", "Grid_BackColorSel")
'    Cor_Grid_ForeColor.backcolor = Ler_RGB("Colors", "Grid_ForeColor")
'    Cor_Grid_ForeColorFixed.backcolor = Ler_RGB("Colors", "Grid_ForeColorFixed")
'    Cor_Grid_ForeColorSel.backcolor = Ler_RGB("Colors", "Grid_ForeColorSel")
'    Cor_Grid_Color.backcolor = Ler_RGB("Colors", "Grid_Color")
'    Cor_Grid_ColorFixed.backcolor = Ler_RGB("Colors", "Grid_ColorFixed")
'    Cor_Fundo_Topico_Normal.backcolor = Ler_RGB("Colors", "Topic_BackColor")
'    Cor_Letra_Topico_Normal.backcolor = Ler_RGB("Colors", "Topic_ForeColor")
'    Cor_Fundo_Topico_Over.backcolor = Ler_RGB("Colors", "Topic_BackColorSel")
'    Cor_Letra_Topico_Over.backcolor = Ler_RGB("Colors", "Topic_ForeColorSel")
'    Cor_Scroll_Bar.backcolor = Ler_RGB("Colors", "Scroll_BackColor")
'    Cor_Letra_Tab_Normal.backcolor = Ler_RGB("Colors", "Tab_Letter_Normal")
'    Cor_Letra_Tab_Over.backcolor = Ler_RGB("Colors", "Tab_Letter_Over")
'    Cor_Fundo_Task_Bar.backcolor = Ler_RGB("Colors", "Task_Bar_Background")
'    Cor_Letra_Bar_Info.backcolor = Ler_RGB("Colors", "Bar_Infotmation_Letter")
'    Cor_Contorno_Frame_Centro.backcolor = Ler_RGB("Colors", "Shape_Frame_Center")
'    Cor_BackColor_Download_Add_Ons.backcolor = Ler_RGB("Colors", "Download_Add_Ons_BackColor")
'    Cor_Letter_Download_Add_Ons.backcolor = Ler_RGB("Colors", "Download_Add_Ons_Letter")
'    Cor_BackColor_Display.backcolor = Ler_RGB("Colors", "Display_BackColor")
'    Cor_Download_Add_Ons_Buttons.backcolor = Ler_RGB("Colors", "Download_Add_Ons_Buttons")
'    Cor_Slider_Music.backcolor = Ler_RGB("Colors", "Slider_Progress")
'    Cor_Topic_Task_Bar.backcolor = Ler_RGB("Colors", "Topic_Task_Bar")
'    Cor_Label_Button_ForeColor.backcolor = Ler_RGB("Colors", "Label_Button_ForeColor")
'    Cor_BackGround_Bar_Label_Button.backcolor = Ler_RGB("Colors", "BackGround_Bar_Label_Button")
'    Cor_BackGround_Frame_Cover.backcolor = Ler_RGB("Colors", "BackGround_Frame_Cover")
'    Cor_Label_Frame_Cover.backcolor = Ler_RGB("Colors", "Label_Frame_Cover")
'    Cor_Form_About_Letter.backcolor = Ler_RGB("Colors", "Form_About_Letter")
'    Cor_Line_Border_Frames.backcolor = Ler_RGB("Colors", "Lines_Border_Frames")
'    Cor_Menu_BackColorSel.backcolor = Ler_RGB("Colors", "Menu_BackColorSel")
'    Cor_Menu_BackColor.backcolor = Ler_RGB("Colors", "Menu_BackColor")
'    Cor_Menu_ForeColor.backcolor = Ler_RGB("Colors", "Menu_ForeColor")
'    Cor_Menu_ForeColorSel.backcolor = Ler_RGB("Colors", "Menu_ForeColorSel")
End Sub

'Public Function Ler_RGB(Pai As String, Filho As String) As String
'    'Procedimento para ler a cor em RGB
'    On Error Resume Next
'    Dim CorRGB() As String
'    Dim Cor_Objecto As String: Cor_Objecto = ReadINI(Pai, Filho, App.Path & "\Skins\" & Form_Preferencias.Text_Skin.Text & "\Style.ini")
'    CorRGB = Split(Cor_Objecto, ",")
'    Ler_RGB = RGB(CorRGB(0), CorRGB(1), CorRGB(2))
'End Function

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Carregar_Imagens_do_Skin
End Sub

Private Sub Menu_Fechar_Click()
    'Chamar o procedimento
    Form_Principal.Botao_Fechar_Click
End Sub
