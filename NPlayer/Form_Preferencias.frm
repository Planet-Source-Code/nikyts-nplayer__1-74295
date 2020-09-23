VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD - Copy.OCX"
Begin VB.Form Form_Preferencias 
   Appearance      =   0  'Flat
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14985
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
   ScaleHeight     =   605
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   999
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   993
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Width           =   14895
      Begin VB.PictureBox Frame_Skin 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   9000
         ScaleHeight     =   153
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   401
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   6015
         Begin VB.ListBox List_Skins 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            Height          =   420
            Left            =   120
            TabIndex        =   46
            Top             =   5880
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.DirListBox Dir_Skins 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H00000000&
            Height          =   540
            Left            =   120
            TabIndex        =   45
            Top             =   5280
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.PictureBox Lista_Skins 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00101010&
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   960
            ScaleHeight     =   95
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   303
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   5280
            Visible         =   0   'False
            Width           =   4575
            Begin VB.Label Label_Skin 
               BackColor       =   &H00EEEEEE&
               BackStyle       =   0  'Transparent
               Caption         =   "Skin"
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   0
               Left            =   120
               TabIndex        =   43
               Top             =   0
               Width           =   960
            End
            Begin VB.Label Shape_Sombra_Skin 
               BackColor       =   &H00D88316&
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   3975
            End
         End
         Begin VB.PictureBox Barra_Text_Wallpaper 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   0
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   240
            Width           =   5475
            Begin VB.PictureBox Botao_Pesquisar 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   5160
               ScaleHeight     =   12
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   12
               TabIndex        =   60
               TabStop         =   0   'False
               ToolTipText     =   "Selecionar pasta"
               Top             =   120
               Width           =   180
            End
            Begin VB.TextBox Text_Wallpaper 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   600
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   0
               Width           =   1380
            End
            Begin VB.Shape Contorno_Wallpaper 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.PictureBox Pic_Wallpaper 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   720
            Width           =   195
         End
         Begin VB.PictureBox Barra_Text_Skin 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   120
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   4800
            Visible         =   0   'False
            Width           =   5475
            Begin VB.PictureBox Seta_Skin 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   5160
               ScaleHeight     =   19
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   11
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox Text_Skin 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Skin 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.CheckBox Check_Wallpaper 
            BackColor       =   &H00313131&
            Caption         =   "Não mostar o papel de parede"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   59
            Top             =   720
            Width           =   5415
         End
         Begin VB.Label Label_Wallpaper 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Papel de parede do MusicLink"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label Label_Skin_Programa 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Skin do programa"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   4560
            Visible         =   0   'False
            Width           =   1545
         End
      End
      Begin VB.PictureBox Frame_Complementos 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   8880
         ScaleHeight     =   265
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   441
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   6615
         Begin VB.PictureBox Frame_Download 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3735
            Left            =   480
            ScaleHeight     =   249
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   369
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   0
            Width           =   5535
            Begin NPlayer.NProgressBar ProgressBar 
               Height          =   150
               Index           =   0
               Left            =   120
               TabIndex        =   54
               Top             =   720
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   265
            End
            Begin VB.Label Label_Remover 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Remover"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   3480
               TabIndex        =   49
               Top             =   120
               Width           =   780
            End
            Begin VB.Label Label_Instalar 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Instalar"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   4680
               TabIndex        =   48
               Top             =   120
               Width           =   660
            End
            Begin VB.Label Label_Download_Transferindo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transferindo..."
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   34
               Top             =   360
               Width           =   1260
            End
            Begin VB.Label Label_Download_Titulo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Graphite"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   33
               Top             =   120
               Width           =   840
            End
            Begin VB.Shape Shape_Remover 
               BackColor       =   &H00313131&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00404040&
               BorderStyle     =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   3360
               Shape           =   4  'Rounded Rectangle
               Top             =   60
               Width           =   1095
            End
            Begin VB.Shape Shape_Instalar 
               BackColor       =   &H00313131&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00404040&
               BorderStyle     =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   4560
               Shape           =   4  'Rounded Rectangle
               Top             =   60
               Width           =   1095
            End
            Begin VB.Label Label_Fundo_Download 
               BackColor       =   &H00B67B26&
               Height          =   1095
               Index           =   0
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Width           =   5535
            End
         End
      End
      Begin VB.PictureBox Frame_Visualizacao 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   360
         ScaleHeight     =   185
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   377
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   5655
         Begin VB.PictureBox Pic_Ver_Capa 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   360
            Width           =   195
         End
         Begin VB.PictureBox Pic_Ver_Playlist 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   0
            Width           =   195
         End
         Begin VB.CheckBox Check_Ver_Playlist 
            Appearance      =   0  'Flat
            BackColor       =   &H00313131&
            Caption         =   "Ver a lista de reprodução"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   4095
         End
         Begin VB.CheckBox Check_Ver_Capa 
            Appearance      =   0  'Flat
            BackColor       =   &H00313131&
            Caption         =   "Ver a capa do album"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   3
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.PictureBox Frame_Idioma 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   9240
         ScaleHeight     =   281
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   457
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   6855
         Begin VB.PictureBox Lista_Linguas 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00101010&
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   0
            ScaleHeight     =   95
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   303
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   600
            Visible         =   0   'False
            Width           =   4575
            Begin VB.Label Label_Lingua 
               BackColor       =   &H00EEEEEE&
               BackStyle       =   0  'Transparent
               Caption         =   "Idioma"
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   0
               Left            =   120
               TabIndex        =   38
               Top             =   0
               Width           =   960
            End
            Begin VB.Label Shape_Sombra_Lingua 
               BackColor       =   &H00D88316&
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Width           =   3975
            End
         End
         Begin VB.FileListBox File_Lingua 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            Height          =   420
            Left            =   5760
            Pattern         =   "*.lng"
            TabIndex        =   41
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.DirListBox Dir_Lingua 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            Height          =   540
            Left            =   5760
            TabIndex        =   40
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.PictureBox Barra_Text_Lingua 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   0
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   5475
            Begin VB.PictureBox Seta_Lingua 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   5160
               ScaleHeight     =   19
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   11
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox Text_Lingua 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   4
               Top             =   30
               Width           =   1500
            End
            Begin VB.Shape Contorno_Lingua 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.Label Label_Idioma_Programa 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Idioma do programa"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   1770
         End
      End
      Begin VB.TextBox Text_Tela_Cheia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10560
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox Frame_Geral 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   2400
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   441
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Width           =   6615
         Begin VB.PictureBox Pic_Tray 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   2280
            Width           =   195
         End
         Begin VB.PictureBox Pic_Downloads 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   720
            Width           =   195
         End
         Begin VB.PictureBox Pic_Actualizacoes 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1920
            Width           =   195
         End
         Begin VB.PictureBox Barra_Text_Downloads 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   0
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Width           =   5475
            Begin VB.PictureBox Botao_Selecionar_Pasta 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   5040
               ScaleHeight     =   12
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   12
               TabIndex        =   63
               TabStop         =   0   'False
               ToolTipText     =   "Selecionar pasta"
               Top             =   120
               Width           =   180
            End
            Begin VB.TextBox Text_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   600
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   0
               Width           =   1380
            End
            Begin VB.Shape Contorno_Downloads 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.PictureBox Pic_Guardar_Lista 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1560
            Width           =   195
         End
         Begin VB.CheckBox Check_Guardar_Lista 
            BackColor       =   &H00313131&
            Caption         =   "Guardar a lista de reprodução ao fechar o programa"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   1
            Top             =   1560
            Width           =   5415
         End
         Begin VB.CheckBox Check_Actualizacoes 
            BackColor       =   &H00313131&
            Caption         =   "Verificar actualizações automaticamente"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   53
            Top             =   1920
            Width           =   5415
         End
         Begin VB.CheckBox Check_Downloads 
            BackColor       =   &H00313131&
            Caption         =   "Pasta predefinida pelo programa"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   62
            Top             =   720
            Width           =   5415
         End
         Begin VB.CheckBox Check_Tray 
            BackColor       =   &H00313131&
            Caption         =   "Minimizar o programa no systray (ao lado do clock)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   68
            Top             =   2280
            Width           =   5415
         End
         Begin VB.Label Label_Downloads 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guardar as músicas transferidas em"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   3120
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Lista_Opcoes 
         Height          =   2775
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   4895
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   15790578
         ForeColorFixed  =   0
         BackColorSel    =   14200408
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         Redraw          =   -1  'True
         FocusRect       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   0
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Linha_Horizontal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   8040
         TabIndex        =   66
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label_Topico_Selecionado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tópico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   8160
         TabIndex        =   65
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Linha_Vertical 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   7920
         TabIndex        =   64
         Top             =   120
         Width           =   15
      End
      Begin VB.Label Label_Close 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5520
         TabIndex        =   51
         ToolTipText     =   "Ocultar"
         Top             =   270
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape_Centro 
         BorderColor     =   &H00212121&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label_Info 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         BackStyle       =   0  'Transparent
         Caption         =   "Preferências actualizadas com sucesso."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   50
         Top             =   300
         Visible         =   0   'False
         Width           =   3390
      End
      Begin VB.Shape Shape_Info 
         BackColor       =   &H00F0E7D7&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B67B26&
         Height          =   315
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   5475
      End
   End
   Begin VB.PictureBox Frame_Botoes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   561
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8160
      Width           =   8415
      Begin VB.PictureBox Botao_Actualizar 
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
         TabIndex        =   6
         Top             =   120
         Width           =   1740
         Begin VB.Shape Contorno_Actualizar 
            BorderColor     =   &H00D88316&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label_Actualizar 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actualizar"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   450
            TabIndex        =   14
            Top             =   45
            Width           =   840
         End
      End
      Begin VB.PictureBox Botao_Cancelar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   5880
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   8
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
            TabIndex        =   13
            Top             =   45
            Width           =   780
         End
      End
      Begin VB.PictureBox Botao_Ok 
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
         TabIndex        =   7
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
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   750
            TabIndex        =   12
            Top             =   45
            Width           =   240
         End
      End
      Begin VB.Image Fundo_Frame_Botoes 
         Height          =   615
         Left            =   0
         Picture         =   "Form_Preferencias.frx":0000
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Preferências"
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
         TabIndex        =   10
         Top             =   120
         Width           =   1245
      End
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   5880
         Picture         =   "Form_Preferencias.frx":0A82
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   195
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Preferencias.frx":0CB4
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
Attribute VB_Name = "Form_Preferencias"
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
Dim Linha_Selecionada_Lingua As Integer
Dim Linha_Selecionada_Skin As Integer

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

Private Sub Botao_Actualizar_Click()
    'Atalho para
    Label_Actualizar_Click
End Sub

Private Sub Botao_Actualizar_GotFocus()
    'Colocar o focus no botao
    Contorno_Actualizar.Visible = True
End Sub

Private Sub Botao_Actualizar_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Actualizar_Click
    If KeyCode = vbKeyRight Then Botao_Actualizar_LostFocus: Botao_Ok_GotFocus: Botao_Ok.SetFocus
End Sub

Private Sub Botao_Actualizar_LostFocus()
    'Remover o focus no botao
    Contorno_Actualizar.Visible = False
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
    If KeyCode = vbKeyLeft Then Botao_Ok_LostFocus: Botao_Actualizar_GotFocus: Botao_Actualizar.SetFocus
    If KeyCode = vbKeyRight Then Botao_Ok_LostFocus: Botao_Cancelar_GotFocus: Botao_Cancelar.SetFocus
End Sub

Private Sub Botao_Ok_LostFocus()
    'Remover o focus no botao
    Contorno_Ok.Visible = False
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    Label_Cancelar_Click
End Sub

Private Sub Botao_Pesquisar_Click()
    'Selecionar um papel de parede do meu computador
    Dim Explorador As New Class_Dialog
    
    With Explorador
        .Filter = ("*.bmp,*.jpg")
        
        ' Decide a pasta inicial
        If Text_Wallpaper <> "" Then
            If Dir(Text_Wallpaper, vbDirectory) <> "" Then
                ' Se for uma pasta, é ela mesma
                .Path = Text_Wallpaper
            Else
                ' Se for um arquivo, extraia só o caminho
                .Path = left(Text_Wallpaper, InStrRev(Text_Wallpaper, "\"))
            End If
        End If
        
        .FileFlags = PATHMUSTEXIST
        .FileFlags = .FileFlags + EXPLORER
        
        ' Mostra o diálogo
        .DialogFile OpenFile
        If .cancel = True Then Exit Sub
        Text_Wallpaper.Text = .FullPath
    End With
    
    Set Explorador = Nothing
End Sub

Private Sub Botao_Selecionar_Pasta_Click()
    'Selecionar a pasta para salvar os downloads
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    'Personaliza a procura
    szTitle = Botao_Pesquisar.ToolTipText
    With tBrowseInfo
        .hWndOwner = Me.hWnd
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
        Text_Downloads.Text = sBuffer & "\"
    End If

End Sub

Private Sub Check_Actualizacoes_Click()
    'Des/Activar a opcção
    If Check_Actualizacoes.Value = 1 Then
        Pic_Actualizacoes.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Actualizacoes.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Downloads_Click()
    'Des/Activar a opcção
    If Check_Downloads.Value = 1 Then
        Pic_Downloads.Picture = Form_Skin.Check_Over.Picture
        Text_Downloads.Enabled = False
    Else
        Pic_Downloads.Picture = Form_Skin.Check_Normal.Picture
        Text_Downloads.Enabled = True
    End If
End Sub

Private Sub Check_Tray_Click()
    'Des/Activar a opcção
    If Check_Tray_.Value = 1 Then
        Pic_Tray_.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Tray_.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Wallpaper_Click()
    'Des/Activar a opcção
    If Check_Wallpaper.Value = 1 Then
        Pic_Wallpaper.Picture = Form_Skin.Check_Over.Picture
        Text_Wallpaper.Enabled = True
        Botao_Pesquisar.Enabled = True
    Else
        Pic_Wallpaper.Picture = Form_Skin.Check_Normal.Picture
        Text_Wallpaper.Enabled = False
        Botao_Pesquisar.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    'Desenhar o formulário
    Desenhar_Formulario
End Sub

Private Sub Label_Close_Click()
    'Ocultar a frame de informação
    Shape_Info.Visible = False
    Label_Info.Visible = False
    Label_Close.Visible = False
    Frame_Geral.top = Shape_Info.top
    Frame_Geral.Height = Frame_Geral.ScaleHeight + Shape_Info.Height + 16
    Ajustar_Frames
End Sub

Private Sub Label_Lingua_Click(Index As Integer)
    'Indicar a lingua selecionada pelo utilizador
    Text_Lingua.Text = Label_Lingua(Index).Caption
    
    'Chamar o procedimento
    'Carregar_Idioma
    
    Lista_Linguas.Visible = False
    Text_Lingua.SetFocus
End Sub

Private Sub Label_Lingua_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Lingua = Index Then Exit Sub
    Shape_Sombra_Lingua(Linha_Selecionada_Lingua).Visible = False
    Label_Lingua(Linha_Selecionada_Lingua).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
    Shape_Sombra_Lingua(Index).Visible = True
    Label_Lingua(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    Linha_Selecionada_Lingua = Index
End Sub

Private Sub Label_Skin_Click(Index As Integer)
    'Indicar a Skin selecionada pelo utilizador
    Text_Skin.Text = Label_Skin(Index).Caption
    
    'Chamar o procedimento
    'Carregar_Skin
    
    Lista_Skins.Visible = False
    Text_Skin.SetFocus
End Sub

Private Sub Label_Skin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Skin = Index Then Exit Sub
    Shape_Sombra_Skin(Linha_Selecionada_Skin).Visible = False
    Label_Skin(Linha_Selecionada_Skin).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
    Shape_Sombra_Skin(Index).Visible = True
    Label_Skin(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    Linha_Selecionada_Skin = Index
End Sub

Private Sub Lista_Opcoes_Click()
    'Selecionar linha
    Ocultar_Objectos
    Ocultar_Frames
    
    Label_Topico_Selecionado.Caption = Lista_Opcoes.TextMatrix(Lista_Opcoes.Row, 1)
    Select Case Lista_Opcoes.Row
        Case 1
            Frame_Geral.Visible = True
        Case 2
            Frame_Visualizacao.Visible = True
        Case 3
            Frame_Idioma.Visible = True
        Case 4
            Frame_Skin.Visible = True
'        Case 5
'            Frame_Complementos.Visible = True
    End Select
End Sub

Private Sub Lista_Opcoes_SelChange()
    'Atalho para
    Lista_Opcoes_Click
End Sub

Private Sub Pic_Actualizacoes_Click()
    'Des/Activar a opcção
    If Check_Actualizacoes.Value = 0 Then
        Check_Actualizacoes.Value = 1
        Pic_Actualizacoes.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Actualizacoes.Value = 0
        Pic_Actualizacoes.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Downloads_Click()
    'Des/Activar a opcção
    If Check_Downloads.Value = 0 Then
        Check_Downloads.Value = 1
        Pic_Downloads.Picture = Form_Skin.Check_Over.Picture
        Text_Downloads.Enabled = False
        
    Else
        Check_Downloads.Value = 0
        Pic_Downloads.Picture = Form_Skin.Check_Normal.Picture
        Text_Downloads.Enabled = True
    End If
End Sub

Private Sub Pic_Tray_Click()
    'Des/Activar a opcção
    If Check_Tray.Value = 0 Then
        Check_Tray.Value = 1
        Pic_Tray.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Tray.Value = 0
        Pic_Tray.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Ver_Capa_Click()
    'Des/Activar a opcção
    If Check_Ver_Capa.Value = 0 Then
        Check_Ver_Capa.Value = 1
        Pic_Ver_Capa.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Ver_Capa.Value = 0
        Pic_Ver_Capa.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Ver_Capa_Click()
    'Des/Activar a opcção
    If Check_Ver_Capa.Value = 1 Then
        Pic_Ver_Capa.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Ver_Capa.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Ver_Playlist_Click()
    'Des/Activar a opcção
    If Check_Ver_Playlist.Value = 0 Then
        Check_Ver_Playlist.Value = 1
        Pic_Ver_Playlist.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Ver_Playlist.Value = 0
        Pic_Ver_Playlist.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Ver_Playlist_Click()
    'Des/Activar a opcção
    If Check_Ver_Playlist.Value = 1 Then
        Pic_Ver_Playlist.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Ver_Playlist.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Botao_Cancelar_Click
End Sub

Private Sub Form_Load()
    'Carregar opções do programa
    Verificar_Opcoes_Programa
    
    'Formatar a lista de opções
    With Lista_Opcoes
        .RowHeight(0) = 0
        .ColWidth(0) = 0
        .ColWidth(1) = 6000
        .Rows = 5
    End With
    
    'Chamar o procedimento
    Carregar_Idioma
    Carregar_Skin
    Desenhar_Formulario
    
    'Propriedades iniciais do formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, True

    'Chamar procedimentos
    Carregar_Idiomas_Existentes
    'Carregar_Skins_Existentes
    
    'Selecionar a 1ªlinha da lista linguas
    Linha_Selecionada_Lingua = 0
    Shape_Sombra_Lingua(0).Visible = True
    Label_Lingua(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    
    'Selecionar a 1ªlinha da lista skins
    Linha_Selecionada_Skin = 0
    Shape_Sombra_Skin(0).Visible = True
    Label_Skin(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    
    'Alterar cores do progreesbar
    ProgressBar(0).backcolor = Form_Skin.Cor_Contorno_Caixas.backcolor
    
    'Tópico activo inicialmente
    Label_Topico_Selecionado.Caption = Lista_Opcoes.TextMatrix(Lista_Opcoes.Row, 1)
End Sub

Public Sub Verificar_Opcoes_Programa()
    'Procedimento para carregar os valores guardados das preferências
    Text_Tela_Cheia.Text = ReadINI("Settings", "FullScreen", Localizacao_Ficheiro_Preferencias)
    Text_Lingua.Text = ReadINI("Settings", "Language", Localizacao_Ficheiro_Preferencias)
    'Text_Skin.Text = ReadINI("Preferences", "Skin_Of_Program", Localizacao_Ficheiro_Preferencias)
    
    Dim opcao_download As String
    opcao_download = ReadINI("Preferences", "Check_Downloads", Localizacao_Ficheiro_Preferencias)
    If opcao_download = "1" Then
        Check_Downloads.Value = 1
        Pic_Downloads.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Downloads.Value = 0
        Pic_Downloads.Picture = Form_Skin.Check_Normal.Picture
    End If
    Text_Downloads.Text = ReadINI("Preferences", "Path_Downloads", Localizacao_Ficheiro_Preferencias)
        
    Dim Lista As String
    Lista = ReadINI("Visualization", "Frame_Playlist", Localizacao_Ficheiro_Preferencias)
    If Lista = "1" Then
        Check_Ver_Playlist.Value = 1
        Pic_Ver_Playlist.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Ver_Playlist.Value = 0
        Pic_Ver_Playlist.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    Dim Capa As String
    Capa = ReadINI("Visualization", "Frame_Cover", Localizacao_Ficheiro_Preferencias)
    If Capa = "1" Then
        Check_Ver_Capa.Value = 1
        Pic_Ver_Capa.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Ver_Capa.Value = 0
        Pic_Ver_Capa.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    Dim Guardar As String
    Guardar = ReadINI("Preferences", "Check_Save_Playlist", Localizacao_Ficheiro_Preferencias)
    If Guardar = "1" Then
        Check_Guardar_Lista.Value = 1
        Pic_Guardar_Lista.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Guardar_Lista.Value = 0
        Pic_Guardar_Lista.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    Dim actualizacoes As String
    actualizacoes = ReadINI("Preferences", "Check_Update", Localizacao_Ficheiro_Preferencias)
    If actualizacoes = "1" Then
        Check_Actualizacoes.Value = 1
        Pic_Actualizacoes.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Actualizacoes.Value = 0
        Pic_Actualizacoes.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    Dim minimizar As String
    minimizar = ReadINI("Preferences", "Check_Tray", Localizacao_Ficheiro_Preferencias)
    If minimizar = "1" Then
        Check_Tray.Value = 1
        Pic_Tray.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Tray.Value = 0
        Pic_Tray.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    Dim papel As String
    papel = ReadINI("Preferences", "Check_Wallpaper", Localizacao_Ficheiro_Preferencias)
    If papel = "1" Then
        Check_Wallpaper.Value = 1
        Pic_Wallpaper.Picture = Form_Skin.Check_Over.Picture
        Text_Wallpaper.Enabled = True
        Botao_Pesquisar.Enabled = True
    Else
        Check_Wallpaper.Value = 0
        Pic_Wallpaper.Picture = Form_Skin.Check_Normal.Picture
        Text_Wallpaper.Enabled = False
        Botao_Pesquisar.Enabled = False
    End If
    Text_Wallpaper.Text = ReadINI("Preferences", "Text_Wallpaper", Localizacao_Ficheiro_Preferencias)
End Sub

'Public Sub Carregar_Skins_Existentes()
'    'Procedimento para carregar os skins existentes
'    Dir_Skins.Path = App.Path & "\Skins\"
'    Dim Temas As Integer: For Temas = 0 To Dir_Skins.ListCount - 1
'        List_Skins.AddItem Dir(Dir_Skins.List(Temas), vbDirectory)
'    Next Temas
'
'    'Criar a lista consoante o nº de skins disponiveis
'    If List_Skins.ListCount > 0 Then
'        Label_Skin(0).Caption = ""
'        Label_Skin(0).Visible = True
'        Dim Objecto As Integer
'        For Objecto = 1 To List_Skins.ListCount - 1
'            Load Label_Skin(Objecto)
'            Label_Skin(Objecto).Move Label_Skin(Objecto - 1).left, Label_Skin(Objecto - 1).top + Label_Skin(Objecto - 1).Height
'            Label_Skin(Objecto).Visible = True
'
'            Load Shape_Sombra_Skin(Objecto)
'            Shape_Sombra_Skin(Objecto).Move Shape_Sombra_Skin(Objecto - 1).left, Shape_Sombra_Skin(Objecto - 1).top + Shape_Sombra_Skin(Objecto - 1).Height
'            Shape_Sombra_Skin(Objecto).Visible = False
'            Shape_Sombra_Skin(Objecto).ZOrder 1
'        Next Objecto
'
'        'Preencher as label's com as Skins disponiveis
'        Dim Z As Integer
'        List_Skins.ListIndex = 0
'        For Z = 0 To List_Skins.ListCount - 1
'            Label_Skin(Z).Caption = List_Skins.List(Z)
'        Next Z
'    End If
'End Sub

Public Sub Carregar_Idiomas_Existentes()
    'Procedimento para carregar os idiomas do programa
    'Carregar idiomas disponiveis
    Dir_Lingua.Path = App.Path & "\Languages\"
    File_Lingua.Path = Dir_Lingua.Path
    File_Lingua.Pattern = "*.lng"
    
    'Criar a lista consoante o nº de idiomas disponiveis
    Label_Lingua(0).Caption = ""
    Label_Lingua(0).Visible = True
    Dim Objecto As Integer
    For Objecto = 1 To File_Lingua.ListCount - 1
        Load Label_Lingua(Objecto)
        Label_Lingua(Objecto).Move Label_Lingua(Objecto - 1).left, Label_Lingua(Objecto - 1).top + Label_Lingua(Objecto - 1).Height
        Label_Lingua(Objecto).Visible = True
        
        Load Shape_Sombra_Lingua(Objecto)
        Shape_Sombra_Lingua(Objecto).Move Shape_Sombra_Lingua(Objecto - 1).left, Shape_Sombra_Lingua(Objecto - 1).top + Shape_Sombra_Lingua(Objecto - 1).Height
        Shape_Sombra_Lingua(Objecto).Visible = False
        Shape_Sombra_Lingua(Objecto).ZOrder 1
    Next Objecto
        
    'Preencher as label's com as linguas disponiveis
    Dim Z As Integer
    File_Lingua.ListIndex = 0
    For Z = 0 To File_Lingua.ListCount - 1
        Label_Lingua(Z).Caption = left$(File_Lingua.List(Z), InStr(File_Lingua.List(Z), ".") - (1)) 'Retirar a extensão do ficheiro ".lng"
    Next Z
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Frame_Centro.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Frame_Geral.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Frame_Visualizacao.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Frame_Complementos.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Frame_Skin.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Frame_Idioma.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Centro.BorderColor = .Cor_Contorno_Frame_Centro.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Ok.Picture = .Pic_Button.Picture
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Actualizar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Actualizar.Picture = .Pic_Button.Picture
        Contorno_Actualizar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Downloads.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Barra_Text_Downloads.backcolor = .Cor_Fundo_Textbox.backcolor
        Barra_Text_Downloads.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Downloads.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Downloads.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Downloads.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Downloads.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Contorno_Downloads.BorderColor = .Cor_Contorno_Caixas.backcolor
        Text_Downloads.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Downloads.ForeColor = .Cor_Letra_Textbox.backcolor
        Pic_Guardar_Lista.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Guardar_Lista.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Guardar_Lista.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Pic_Ver_Playlist.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Pic_Ver_Capa.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Ver_Playlist.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Ver_Playlist.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Check_Ver_Capa.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Ver_Capa.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Skin_Programa.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Seta_Skin.Picture = .Seta_Combo.Picture
        Barra_Text_Skin.backcolor = .Cor_Fundo_Textbox.backcolor
        Barra_Text_Skin.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Skin.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Skin.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Skin.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Skin.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Contorno_Skin.BorderColor = .Cor_Contorno_Caixas.backcolor
        Text_Skin.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Skin.ForeColor = .Cor_Letra_Textbox.backcolor
        Lista_Skins.backcolor = .Cor_Fundo_Textbox.backcolor
        Shape_Sombra_Skin(0).backcolor = .Cor_Contorno_Caixas.backcolor
        Label_Skin(0).ForeColor = .Cor_Letra_Textbox.backcolor
        Label_Idioma_Programa.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Seta_Lingua.Picture = .Seta_Combo.Picture
        Barra_Text_Lingua.backcolor = .Cor_Fundo_Textbox.backcolor
        Barra_Text_Lingua.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Lingua.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Lingua.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Lingua.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Lingua.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Contorno_Lingua.BorderColor = .Cor_Contorno_Caixas.backcolor
        Text_Lingua.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Lingua.ForeColor = .Cor_Letra_Textbox.backcolor
        Lista_Linguas.backcolor = .Cor_Fundo_Textbox.backcolor
        Shape_Sombra_Lingua(0).backcolor = .Cor_Contorno_Caixas.backcolor
        Label_Lingua(0).ForeColor = .Cor_Letra_Textbox.backcolor
        Frame_Download.backcolor = .Cor_Fundo_Textbox.backcolor
        ProgressBar(0).backcolor = .Cor_Contorno_Caixas.backcolor
        Label_Fundo_Download(0).backcolor = .Cor_BackColor_Download_Add_Ons.backcolor
        Label_Download_Titulo(0).ForeColor = .Cor_Letter_Download_Add_Ons.backcolor
        Label_Download_Transferindo(0).ForeColor = .Cor_Letter_Download_Add_Ons.backcolor
        Label_Remover(0).ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Instalar(0).ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Shape_Remover(0).backcolor = .Cor_Download_Add_Ons_Buttons.backcolor
        Shape_Instalar(0).backcolor = .Cor_Download_Add_Ons_Buttons.backcolor
        Pic_Actualizacoes.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Actualizacoes.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Actualizacoes.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Pic_Tray.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Tray.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Tray.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Check_Downloads.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Downloads.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Pic_Downloads.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Label_Wallpaper.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Barra_Text_Wallpaper.backcolor = .Cor_Fundo_Textbox.backcolor
        Barra_Text_Wallpaper.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Wallpaper.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Wallpaper.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Wallpaper.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Wallpaper.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Contorno_Wallpaper.BorderColor = .Cor_Contorno_Caixas.backcolor
        Text_Wallpaper.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Wallpaper.ForeColor = .Cor_Letra_Textbox.backcolor
        Check_Wallpaper.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Wallpaper.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Botao_Pesquisar.Picture = .Botao_Pesquisar.Picture
        Botao_Selecionar_Pasta.Picture = .Botao_Pesquisar.Picture
        Lista_Opcoes.backcolor = .Cor_Grid_BackColor.backcolor
        Lista_Opcoes.BackColorBkg = .Cor_Grid_BackColorBkg.backcolor
        Lista_Opcoes.BackColorFixed = .Cor_Grid_BackColorFixed.backcolor
        Lista_Opcoes.BackColorSel = .Cor_Grid_BackColorSel.backcolor
        Lista_Opcoes.ForeColor = .Cor_Grid_ForeColor.backcolor
        Lista_Opcoes.ForeColorFixed = .Cor_Grid_ForeColorFixed.backcolor
        Lista_Opcoes.ForeColorSel = .Cor_Grid_ForeColorSel.backcolor
        Lista_Opcoes.GridColor = .Cor_Grid_Color.backcolor
        Lista_Opcoes.GridColorFixed = .Cor_Grid_BackColor.backcolor '.Cor_Grid_ColorFixed.backcolor
        Label_Topico_Selecionado.ForeColor = .Cor_Line_Border_Frames.backcolor
        Linha_Vertical.backcolor = .Cor_Line_Border_Frames.backcolor
        Linha_Horizontal.backcolor = .Cor_Line_Border_Frames.backcolor
    End With
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Preferences", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Preferences", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Actualizar.Caption = ReadINI("Preferences", "Button_Update", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Preferences", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Preferences", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    With Lista_Opcoes
        .TextMatrix(1, 1) = ReadINI("Preferences", "Label_Main", Localizacao_Ficheiro_Lingua)
        .TextMatrix(2, 1) = ReadINI("Preferences", "Label_Preview", Localizacao_Ficheiro_Lingua)
        .TextMatrix(3, 1) = ReadINI("Preferences", "Label_Language", Localizacao_Ficheiro_Lingua)
        .TextMatrix(4, 1) = ReadINI("Preferences", "Label_Skin", Localizacao_Ficheiro_Lingua)
'        .TextMatrix(5, 1) = ReadINI("Preferences", "Label_Add", Localizacao_Ficheiro_Lingua)
    End With
    Label_Topico_Selecionado.Caption = Lista_Opcoes.TextMatrix(Lista_Opcoes.Row, 1)
    Label_Downloads.Caption = ReadINI("Preferences", "Label_Downloads", Localizacao_Ficheiro_Lingua)
    Check_Guardar_Lista.Caption = ReadINI("Preferences", "Check_Save_Playlist", Localizacao_Ficheiro_Lingua)
    Check_Ver_Playlist.Caption = ReadINI("Preferences", "Check_View_Playlist", Localizacao_Ficheiro_Lingua)
    Check_Ver_Capa.Caption = ReadINI("Preferences", "Check_View_Cover", Localizacao_Ficheiro_Lingua)
    Label_Idioma_Programa.Caption = ReadINI("Preferences", "Check_Language_Program", Localizacao_Ficheiro_Lingua)
    Label_Skin_Programa.Caption = ReadINI("Preferences", "Check_Skin_Program", Localizacao_Ficheiro_Lingua)
    Label_Info.Caption = ReadINI("Preferences", "Label_Info", Localizacao_Ficheiro_Lingua)
    Label_Close.ToolTipText = ReadINI("Preferences", "Label_Close", Localizacao_Ficheiro_Lingua)
    Label_Remover(0).Caption = ReadINI("Preferences", "Label_Remover", Localizacao_Ficheiro_Lingua)
    Label_Instalar(0).Caption = ReadINI("Preferences", "Label_Install", Localizacao_Ficheiro_Lingua)
    Check_Actualizacoes.Caption = ReadINI("Preferences", "Check_Updated", Localizacao_Ficheiro_Lingua)
    Check_Tray.Caption = ReadINI("Preferences", "Check_Tray", Localizacao_Ficheiro_Lingua)
    Label_Wallpaper.Caption = ReadINI("Preferences", "Label_Wallpaper", Localizacao_Ficheiro_Lingua)
    Check_Wallpaper.Caption = ReadINI("Preferences", "Check_Wallpaper", Localizacao_Ficheiro_Lingua)
    Botao_Pesquisar.ToolTipText = ReadINI("Preferences", "Button_Find", Localizacao_Ficheiro_Lingua)
    Botao_Selecionar_Pasta.ToolTipText = ReadINI("Preferences", "Button_Select_Folder", Localizacao_Ficheiro_Lingua)
End Sub

Private Sub Frame_Centro_Click()
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Frame_Complementos_Click()
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Frame_Geral_Click()
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Frame_Idioma_Click()
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Frame_Skin_Click()
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Frame_Visualizacao_Click()
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Label_Actualizar_Click()
    'Ver as alterações feitas automaticamente
    If Shape_Info.Visible = False Then
        Shape_Info.Visible = True
        Label_Info.Visible = True
        Label_Close.Visible = True
        Frame_Geral.top = Shape_Info.top + Shape_Info.Height + 16
        Frame_Geral.Height = Frame_Geral.ScaleHeight - Shape_Info.Height - 16
        Ajustar_Frames
    End If

'    Form_Principal.Verificar_Opcoes_do_Programa
'    Salvar_Valores

    'Idioma do programa
    If Frame_Idioma.Visible = True Then
        Me.MousePointer = 11
        Call WriteINI("Settings", "Language", Text_Lingua.Text, (Localizacao_Ficheiro_Preferencias))
        Call Aplicar_Idioma_do_Programa
        Me.MousePointer = 0
    End If
    
    If Frame_Skin.Visible = True Then
        'Call WriteINI("Preferences", "Skin_Of_Program", Text_Skin.Text, (Localizacao_Ficheiro_Preferencias))
        Call WriteINI("Preferences", "Text_Wallpaper", Text_Wallpaper.Text, (Localizacao_Ficheiro_Preferencias))
        Call WriteINI("Preferences", "Check_Wallpaper", Check_Wallpaper.Value, (Localizacao_Ficheiro_Preferencias))
        'Verificar se é para carregar algum wallpaper
        If Check_Wallpaper.Value = 0 Then
            Form_Principal.Fundo_Frame_Music_Link.Visible = False
        Else
            If ArquivoExiste(Form_Preferencias.Text_Wallpaper.Text, False) Then 'Verificar se existe o wallpaper indicado
                Form_Principal.Fundo_Frame_Music_Link.Picture = LoadPicture(Text_Wallpaper.Text)
                Form_Principal.Fundo_Frame_Music_Link.Visible = True
            Else
                Form_Principal.Fundo_Frame_Music_Link.Visible = False
            End If
        End If
    End If
End Sub

Public Sub Ajustar_Frames()
    'Procedimento para ajustar o top das frames
    Frame_Visualizacao.top = Frame_Geral.top
    Frame_Complementos.top = Frame_Geral.top
    Frame_Skin.top = Frame_Geral.top
    Frame_Idioma.top = Frame_Geral.top
    
    Frame_Visualizacao.Height = Frame_Geral.Height
    Frame_Complementos.Height = Frame_Geral.Height
    Frame_Skin.Height = Frame_Geral.Height
    Frame_Idioma.Height = Frame_Geral.Height
End Sub

Public Sub Salvar_Valores()
    'Actualizar as preferências do programa
    If Text_Downloads.Text <> Empty Then
        Call WriteINI("Preferences", "Path_Downloads", Text_Downloads.Text, (Localizacao_Ficheiro_Preferencias))
    End If
    Call WriteINI("Preferences", "Check_Downloads", Check_Downloads.Value, (Localizacao_Ficheiro_Preferencias))

    Call WriteINI("Visualization", "Frame_Playlist", Check_Ver_Playlist.Value, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Visualization", "Frame_Cover", Check_Ver_Capa.Value, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Preferences", "Check_Save_Playlist", Check_Guardar_Lista.Value, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Preferences", "Check_Updates", Check_Actualizacoes.Value, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Preferences", "Check_Wallpaper", Check_Wallpaper.Value, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Preferences", "Check_Tray", Check_Tray.Value, (Localizacao_Ficheiro_Preferencias))
    
    If Text_Wallpaper.Text <> Empty Then
        Call WriteINI("Preferences", "Text_Wallpaper", Text_Wallpaper, (Localizacao_Ficheiro_Preferencias))
    End If
    
    'Call WriteINI("Preferences", "Skin_Of_Program", Text_Skin.Text, (Localizacao_Ficheiro_Preferencias))
End Sub

Private Sub Label_Cancelar_Click()
    'Fechar o formulárion
    'Unload Me
    Label_Close_Click
    Me.Hide
End Sub

Private Sub Ocultar_Frames()
    'Procedimento para repors as imagens originais dos separadores
    Frame_Geral.Visible = False
    Frame_Visualizacao.Visible = False
    Frame_Idioma.Visible = False
    Frame_Skin.Visible = False
    Frame_Complementos.Visible = False
End Sub

Private Sub Label_Ok_Click()
    'Concluir
    'Unload Me
    Label_Close_Click
    Me.Hide
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Preferencias
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Preferencias
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Preferencias
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Preferencias
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Preferencias
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Preferencias
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Downloads.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * (Lista_Opcoes.Width + 1 + 16 + Barra_Text_Downloads.ScaleWidth + 16 + 3)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Fundo_Frame_Botoes.Height + Frame_Centro.ScaleHeight)
    End With
    
    'Ajustar_Formulario_com_Menu Form_Preferencias, False, False, True, True
'-----------------------------------------------
    With Shape_Contorno
        .Height = Me.ScaleHeight
        .top = 0
        .Width = Me.ScaleWidth
        .left = 0
    End With
    
    With Barra_ControlBox
        .Height = Form_Skin.Fundo_Barra_ControlBox.Height
        .top = 0 ' 1
        .Width = Me.ScaleWidth '- 2
        .left = 0 ' 1
    End With
    
    With Fundo_Barra_ControlBox
        .Stretch = True
        .top = 0
        .Width = Me.Barra_ControlBox.ScaleWidth
        .left = 0
    End With

    With Label_Titulo
        .top = (Me.Barra_ControlBox.ScaleHeight - .Height) / 2
        If Icon_Visivel = False Then .left = 10 Else: .left = 26
    End With
    
    'Botões do controlbox
    Dim Ajustar_Botoes As String
    Ajustar_Botoes = "False" 'ReadINI("Dimensions", "Adjust_Button_ControlBox", Localizacao_Ficheiro_Skin)
    
    With Botao_Fechar
        .Height = Form_Skin.Botao_Fechar.Height
        If Ajustar_Botoes = "False" Then
            .top = (Me.Barra_ControlBox.ScaleHeight - .Height) / 2
        Else
            .top = 0
        End If
        .Width = Form_Skin.Botao_Fechar.Width
        .left = Me.Barra_ControlBox.Width - .Width - 6
    End With
    
    With Frame_Botoes
        .Height = Form_Skin.Fundo_Frame_Botoes.Height
        .top = Me.ScaleHeight - .ScaleHeight - 1
        .Width = Me.ScaleWidth - 2
        .left = 1
    End With
    
    With Fundo_Frame_Botoes
        .Stretch = True
        .top = 0
        .Width = Me.Frame_Botoes.ScaleWidth
        .left = 0
    End With
    
    With Frame_Centro
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - Frame_Botoes.ScaleHeight - 2
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Width = Me.ScaleWidth - 3
        .left = 1
    End With
    
    With Shape_Centro
        .top = 0
        .Height = Me.Frame_Centro.Height
        .left = 0
        .Width = Me.Frame_Centro.Width
        .Visible = False
    End With
'-----------------------------------------------
    
    Ajustar_Botao Form_Preferencias, Botao_Actualizar, Label_Actualizar, True, Contorno_Actualizar
    Ajustar_Botao Form_Preferencias, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Preferencias, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With
    With Botao_Actualizar
        .left = Botao_Ok.left - .ScaleWidth - .top
    End With
    
    With Lista_Opcoes
        .Height = Frame_Centro.ScaleHeight
        .top = 0
        .left = 0
    End With
    
    With Linha_Vertical
        .Height = Lista_Opcoes.Height
        .top = Lista_Opcoes.top
        .left = Lista_Opcoes.left + Lista_Opcoes.Width
    End With
    
    With Label_Topico_Selecionado
        .top = 16
        .left = Lista_Opcoes.left + Lista_Opcoes.Width + 16
    End With
    
    With Linha_Horizontal
        .top = Label_Topico_Selecionado.top + Label_Topico_Selecionado.Height + 3
        .left = Label_Topico_Selecionado.left
        .Width = Form_Skin.Caixa_de_Texto.Width
    End With
    
    With Shape_Info
        .top = Label_Topico_Selecionado.top + Label_Topico_Selecionado.Height + 26
        .Width = Barra_Text_Downloads.ScaleWidth
        .left = Lista_Opcoes.left + Lista_Opcoes.Width + 16
    End With
    
    With Label_Info
        .top = Shape_Info.top + ((Shape_Info.Height - .Height) / 2)
        .left = Shape_Info.left + 10
    End With
    
    With Label_Close
        .top = Label_Info.top
        .left = Shape_Info.left + Shape_Info.Width - .Width - 10
    End With
    
    With Frame_Geral
        .top = Shape_Info.top
        .Width = Frame_Centro.ScaleWidth - 32
        .left = Shape_Info.left
    End With
    
    With Frame_Visualizacao
        .top = Frame_Geral.top
        .Height = Frame_Geral.ScaleHeight
        .left = Frame_Geral.left
        .Width = Frame_Geral.ScaleWidth
    End With
    
    With Frame_Idioma
        .top = Frame_Geral.top
        .Height = Frame_Geral.ScaleHeight
        .left = Frame_Geral.left
        .Width = Frame_Geral.ScaleWidth
    End With
    
    With Frame_Skin
        .top = Frame_Geral.top
        .Height = Frame_Geral.ScaleHeight
        .left = Frame_Geral.left
        .Width = Frame_Geral.ScaleWidth
    End With
    
    With Frame_Complementos
        .top = Frame_Geral.top
        .Height = Frame_Geral.ScaleHeight
        .left = Frame_Geral.left
        .Width = Frame_Geral.ScaleWidth
    End With
    
    Ajustar_Caixa_Texto Barra_Text_Downloads, Text_Downloads, Contorno_Downloads, False
    Ajustar_Caixa_Texto Barra_Text_Lingua, Text_Lingua, Contorno_Lingua, False
    Ajustar_Caixa_Texto Barra_Text_Skin, Text_Skin, Contorno_Skin, False
    Ajustar_Caixa_Texto Barra_Text_Wallpaper, Text_Wallpaper, Contorno_Wallpaper, False
    
    With Check_Downloads
        .left = Barra_Text_Downloads.left
        .top = Barra_Text_Downloads.top + Barra_Text_Downloads.ScaleHeight + 6
    End With

    With Pic_Downloads
        .top = Check_Downloads.top
        .left = Check_Downloads.left
    End With
    
    With Seta_Lingua
        .Height = Form_Skin.Seta_Combo.Height
        .top = (Barra_Text_Lingua.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Seta_Combo.Width
        .left = Barra_Text_Lingua.ScaleWidth - .ScaleWidth - .top
    End With
    
    With Lista_Linguas
        .top = Barra_Text_Lingua.top + Barra_Text_Lingua.ScaleHeight
        .Width = Barra_Text_Lingua.ScaleWidth
        .left = Barra_Text_Lingua.left
    End With
    
    With Shape_Sombra_Lingua(0)
        .Width = Lista_Linguas.ScaleWidth
    End With
    
    With Label_Lingua(0)
        .Width = Lista_Linguas.ScaleWidth
    End With
    
    With Seta_Skin
        .Height = Form_Skin.Seta_Combo.Height
        .top = (Barra_Text_Skin.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Seta_Combo.Width
        .left = Barra_Text_Skin.ScaleWidth - .ScaleWidth - .top
    End With
    
    With Lista_Skins
        .top = Barra_Text_Skin.top + Barra_Text_Skin.ScaleHeight
        .Width = Barra_Text_Skin.ScaleWidth
        .left = Barra_Text_Skin.left
    End With
    
    With Shape_Sombra_Skin(0)
        .Width = Lista_Skins.ScaleWidth
    End With
    
    With Label_Skin(0)
        .Width = Lista_Skins.ScaleWidth
    End With
    
    With Frame_Download
        .Height = Frame_Complementos.ScaleHeight
        .top = 0
        .Width = Frame_Complementos.ScaleWidth
        .left = 0
    End With
    
    With Label_Fundo_Download(0)
        .top = 0
        .Width = Frame_Download.Width
        .left = 0
    End With
    
    With Label_Download_Titulo(0)
        .left = 6
        .top = .left
    End With
    
    With Label_Download_Transferindo(0)
        .top = Label_Download_Titulo(0).top + Label_Download_Titulo(0).Height + 3
        .left = Label_Download_Titulo(0).left
    End With
    
    With ProgressBar(0)
        .Width = Frame_Download.Width - 12
        .left = 6
    End With
    
    With Shape_Instalar(0)
        .Height = Label_Instalar(0).Height + 10
        .top = Label_Download_Titulo(0).top
        .Width = Label_Instalar(0).Width + 10
        .left = Frame_Download.ScaleWidth - .Width - .top
    End With
    
    With Shape_Remover(0)
        .Height = Shape_Instalar(0).Height
        .top = Shape_Instalar(0).top
        .Width = Label_Remover(0).Width + 10
        .left = Shape_Instalar(0).left - .Width - .top
    End With
    
    With Label_Instalar(0)
        .top = Shape_Instalar(0).top + 5
        .left = Shape_Instalar(0).left + 5
    End With
    
    With Label_Remover(0)
        .top = Label_Instalar(0).top
        .left = Shape_Remover(0).left + 5
    End With
    
    Ajustar_ChecBox Pic_Guardar_Lista, Check_Guardar_Lista
    Ajustar_ChecBox Pic_Actualizacoes, Check_Actualizacoes
    Ajustar_ChecBox Pic_Ver_Playlist, Check_Ver_Playlist
    Ajustar_ChecBox Pic_Ver_Capa, Check_Ver_Capa
    Ajustar_ChecBox Pic_Wallpaper, Check_Wallpaper
    Ajustar_ChecBox Pic_Downloads, Check_Downloads
    Ajustar_ChecBox Pic_Tray, Check_Tray
    
    With Botao_Selecionar_Pasta
        .Height = Form_Skin.Botao_Pesquisar.Height
        .top = (Barra_Text_Downloads.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Pesquisar.Width
        .left = Barra_Text_Downloads.ScaleWidth - .ScaleWidth - .top
    End With
    
    With Text_Downloads
        .Width = Barra_Text_Downloads.ScaleWidth - 8 - 8 - Botao_Selecionar_Pasta.ScaleWidth - 8
        .left = 8
    End With
    
    With Botao_Pesquisar
        .Height = Form_Skin.Botao_Pesquisar.Height
        .top = (Barra_Text_Wallpaper.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Pesquisar.Width
        .left = Barra_Text_Wallpaper.ScaleWidth - .ScaleWidth - .top
    End With
    
    With Text_Wallpaper
        .Width = Barra_Text_Wallpaper.ScaleWidth - 8 - 8 - Botao_Pesquisar.ScaleWidth - 8
        .left = 8
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Pic_Wallpaper_Click()
    'Des/Activar a opcção
    If Check_Wallpaper.Value = 0 Then
        Check_Wallpaper.Value = 1
        Pic_Wallpaper.Picture = Form_Skin.Check_Over.Picture
        Text_Wallpaper.Enabled = True
        Botao_Pesquisar.Enabled = True
        
    Else
        Check_Wallpaper.Value = 0
        Pic_Wallpaper.Picture = Form_Skin.Check_Normal.Picture
        Text_Wallpaper.Enabled = False
        Botao_Pesquisar.Enabled = False
    End If
End Sub

Private Sub Pic_Guardar_Lista_Click()
    'Des/Activar a opcção
    If Check_Guardar_Lista.Value = 0 Then
        Check_Guardar_Lista.Value = 1
        Pic_Guardar_Lista.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Guardar_Lista.Value = 0
        Pic_Guardar_Lista.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Guardar_Lista_Click()
    'Des/Activar a opcção
    If Check_Guardar_Lista.Value = 1 Then
        Pic_Guardar_Lista.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Guardar_Lista.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Seta_Lingua_Click()
    'Ver/ocultar lista
    If Lista_Linguas.Visible = True Then
        Lista_Linguas.Visible = False
    Else
        Lista_Linguas.Visible = True
    End If
End Sub

Private Sub Seta_Skin_Click()
    'Ver/ocultar lista
    If Lista_Skins.Visible = True Then
        Lista_Skins.Visible = False
    Else
        Lista_Skins.Visible = True
    End If
End Sub

Private Sub Shape_Sombra_Lingua_Click(Index As Integer)
    'Atalho para
    Label_Lingua_Click (Index)
End Sub

Private Sub Shape_Sombra_Lingua_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Lingua = Index Then Exit Sub
    Shape_Sombra_Lingua(Linha_Selecionada_Lingua).Visible = False
    Label_Lingua(Linha_Selecionada_Lingua).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
    Shape_Sombra_Lingua(Index).Visible = True
    Label_Lingua(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    Linha_Selecionada_Lingua = Index
End Sub

Private Sub Shape_Sombra_Skin_Click(Index As Integer)
    'Atalho para
    Label_Skin_Click (Index)
End Sub

Private Sub Shape_Sombra_Skin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Skin = Index Then Exit Sub
    Shape_Sombra_Skin(Linha_Selecionada_Skin).Visible = False
    Label_Skin(Linha_Selecionada_Skin).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
    Shape_Sombra_Skin(Index).Visible = True
    Label_Skin(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    Linha_Selecionada_Skin = Index
End Sub

Private Sub Text_Lingua_Click()
    'Ocultar lista
    Lista_Linguas.Visible = False
End Sub

Private Sub Text_Lingua_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Lingua.Visible = True
End Sub

Private Sub Text_Lingua_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Lingua.Visible = False
End Sub

Private Sub Ocultar_Objectos()
    'Procedimento para ocultar objectos (ex. as listas das comboboxs)
    Lista_Linguas.Visible = False
    Lista_Skins.Visible = False
End Sub

Private Sub Text_Downloads_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Downloads.Visible = True
End Sub

Private Sub Text_Downloads_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Downloads.Visible = False
End Sub

Private Sub Text_Skin_Click()
    'Ocultar frame
    Lista_Skins.Visible = False
End Sub

Private Sub Text_Skin_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Skin.Visible = True
End Sub

Private Sub Text_Skin_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Skin.Visible = False
End Sub

Private Sub Text_Wallpaper_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Wallpaper.Visible = True
End Sub

Private Sub Text_Wallpaper_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Wallpaper.Visible = False
End Sub

