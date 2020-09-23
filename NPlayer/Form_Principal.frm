VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD - Copy.OCX"
Begin VB.Form Form_Principal 
   Appearance      =   0  'Flat
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   Caption         =   "NPlayer"
   ClientHeight    =   28005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Principal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1867
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1533
   Begin VB.PictureBox Frame_Evento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   10800
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   227
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   3750
      Begin VB.Label Label_Evento 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   2
         Left            =   240
         TabIndex        =   231
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label_Evento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " x "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   3
         Left            =   3480
         TabIndex        =   230
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label_Evento 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do evento"
         ForeColor       =   &H00808080&
         Height          =   555
         Index           =   1
         Left            =   1200
         TabIndex        =   229
         Top             =   480
         Width           =   2400
      End
      Begin VB.Label Label_Evento 
         BackStyle       =   0  'Transparent
         Caption         =   "Evento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00212121&
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   228
         Top             =   240
         Width           =   1800
      End
      Begin VB.Shape Shape_Evento 
         BorderColor     =   &H00C0C0C0&
         Height          =   3000
         Left            =   240
         Top             =   -120
         Width           =   3750
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   660
         Left            =   240
         Picture         =   "Form_Principal.frx":57E2
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox Frame_Menu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00101010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   14640
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ListBox List_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   225
         Index           =   0
         ItemData        =   "Form_Principal.frx":70E4
         Left            =   0
         List            =   "Form_Principal.frx":70E6
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Linha_Ficheiro 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   0
         TabIndex        =   52
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Menu_Ficheiro 
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu ficheiro"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   51
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label Sombra_Ficheiro 
         BackColor       =   &H00838F89&
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   3975
      End
      Begin VB.Shape Shape_Vertical 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   0
         Left            =   1200
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Frame_Menu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00101010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   1
      Left            =   16440
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ListBox List_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   225
         Index           =   1
         ItemData        =   "Form_Principal.frx":70E8
         Left            =   0
         List            =   "Form_Principal.frx":70EA
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Linha_Editar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   0
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Menu_Editar 
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu editar"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   46
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label Sombra_Editar 
         BackColor       =   &H00838F89&
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   3975
      End
      Begin VB.Shape Shape_Vertical 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   1
         Left            =   1200
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Frame_Menu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00101010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   2
      Left            =   14640
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ListBox List_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   225
         Index           =   2
         ItemData        =   "Form_Principal.frx":70EC
         Left            =   0
         List            =   "Form_Principal.frx":70EE
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image Menu_Check 
         Height          =   240
         Index           =   4
         Left            =   1440
         Picture         =   "Form_Principal.frx":70F0
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Menu_Check 
         Height          =   240
         Index           =   3
         Left            =   1200
         Picture         =   "Form_Principal.frx":723A
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Menu_Check 
         Height          =   240
         Index           =   2
         Left            =   960
         Picture         =   "Form_Principal.frx":7384
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Menu_Check 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "Form_Principal.frx":74CE
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Menu_Check 
         Height          =   240
         Index           =   0
         Left            =   480
         Picture         =   "Form_Principal.frx":7618
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Linha_Ver 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   0
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Menu_Ver 
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu ver"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   41
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label Sombra_Ver 
         BackColor       =   &H00404040&
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   3975
      End
      Begin VB.Shape Shape_Vertical 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   2
         Left            =   1200
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Frame_Menu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00101010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   3
      Left            =   16440
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ListBox List_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   225
         Index           =   3
         ItemData        =   "Form_Principal.frx":7762
         Left            =   0
         List            =   "Form_Principal.frx":7764
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Linha_Controlos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   0
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Menu_Controlos 
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu controlos"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   36
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label Sombra_Controlos 
         BackColor       =   &H00838F89&
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   3975
      End
      Begin VB.Shape Shape_Vertical 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   3
         Left            =   1200
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Frame_Menu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00101010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   5
      Left            =   16440
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ListBox List_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   225
         Index           =   5
         ItemData        =   "Form_Principal.frx":7766
         Left            =   0
         List            =   "Form_Principal.frx":7768
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Linha_Ajuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Menu_Ajuda 
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu ajuda"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   31
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label Sombra_Ajuda 
         BackColor       =   &H00838F89&
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   3975
      End
      Begin VB.Shape Shape_Vertical 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   5
         Left            =   1200
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Frame_Menu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00101010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   4
      Left            =   14640
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ListBox List_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   225
         Index           =   4
         ItemData        =   "Form_Principal.frx":776A
         Left            =   0
         List            =   "Form_Principal.frx":776C
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Linha_Ferramentas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Menu_Ferramentas 
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu ferramentas"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   25
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label Sombra_Ferramentas 
         BackColor       =   &H00838F89&
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   3975
      End
      Begin VB.Shape Shape_Vertical 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   4
         Left            =   1200
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Barra_Actualizar 
      Appearance      =   0  'Flat
      BackColor       =   &H008DE7F2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   3120
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   728
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   10920
      Begin VB.PictureBox Botao_Actualizar_Programa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3840
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   159
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   2385
         Begin VB.Label Label_Actualizar_Programa 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Actualizar programa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00B67B26&
            Height          =   195
            Left            =   255
            TabIndex        =   99
            Top             =   90
            Width           =   2025
         End
      End
      Begin VB.Label Linha_Barra_Actualizar 
         BackColor       =   &H00808080&
         Caption         =   "Label1"
         Enabled         =   0   'False
         Height          =   15
         Left            =   0
         TabIndex        =   124
         Top             =   0
         Width           =   7215
      End
      Begin VB.Label Label_Nova_Versao 
         AutoSize        =   -1  'True
         BackColor       =   &H008DE7F2&
         BackStyle       =   0  'Transparent
         Caption         =   "De momento não existem actualizações"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   100
         Top             =   180
         Width           =   3420
      End
      Begin VB.Image Close_Barra_Actualizar 
         Height          =   210
         Left            =   10560
         Picture         =   "Form_Principal.frx":776E
         ToolTipText     =   "Ocultar"
         Top             =   120
         Width           =   210
      End
   End
   Begin VB.PictureBox Close_Wmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   9840
      Picture         =   "Form_Principal.frx":7A18
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Barra_Mini_Player 
      Appearance      =   0  'Flat
      BackColor       =   &H00212121&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   3120
      Picture         =   "Form_Principal.frx":89CA
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   444
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   5640
      Width           =   6660
      Begin VB.PictureBox SliderBar_Mini 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   180
         Picture         =   "Form_Principal.frx":1BD08
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   420
         TabIndex        =   119
         Top             =   600
         Width           =   6300
         Begin VB.PictureBox Slide_Mini 
            Appearance      =   0  'Flat
            BackColor       =   &H00212121&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   0
            Picture         =   "Form_Principal.frx":1EE82
            ScaleHeight     =   8
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Image Image_Barra_Slide_Mini 
            Height          =   150
            Left            =   0
            Picture         =   "Form_Principal.frx":1EF84
            Top             =   0
            Visible         =   0   'False
            Width           =   6300
         End
      End
      Begin VB.PictureBox Picture_Slide_Som_Mini 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   480
         Picture         =   "Form_Principal.frx":220FE
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   110
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   240
         Width           =   1650
         Begin VB.PictureBox Slide_Som_Mini 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   0
            Picture         =   "Form_Principal.frx":22F84
            ScaleHeight     =   11
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   11
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   0
            Width           =   165
         End
      End
      Begin VB.Image Botao_Player_Mini 
         Height          =   360
         Index           =   2
         Left            =   3120
         Picture         =   "Form_Principal.frx":23152
         Top             =   165
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image Botao_Mudo_Mini 
         Height          =   180
         Left            =   180
         Picture         =   "Form_Principal.frx":23854
         Top             =   240
         Width           =   195
      End
      Begin VB.Image Botao_Player_Mini 
         Height          =   225
         Index           =   3
         Left            =   3720
         Picture         =   "Form_Principal.frx":23A76
         Top             =   225
         Width           =   405
      End
      Begin VB.Image Botao_Player_Mini 
         Height          =   225
         Index           =   0
         Left            =   2400
         Picture         =   "Form_Principal.frx":23FA4
         Top             =   240
         Width           =   405
      End
      Begin VB.Image Botao_Player_Mini 
         Height          =   360
         Index           =   1
         Left            =   3120
         Picture         =   "Form_Principal.frx":244D2
         Top             =   165
         Width           =   360
      End
      Begin VB.Image Botao_Player_Mini 
         Height          =   255
         Index           =   4
         Left            =   6120
         Picture         =   "Form_Principal.frx":24BD4
         ToolTipText     =   "Tela cheia"
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.PictureBox Frame_Wmp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   16440
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      Begin WMPLibCtl.WindowsMediaPlayer Wmp 
         Height          =   600
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   900
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "none"
         stretchToFit    =   -1  'True
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   1588
         _cy             =   1058
      End
   End
   Begin VB.PictureBox Barra_Informacoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00212121&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   3120
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5280
      Width           =   10935
      Begin VB.PictureBox Botao_Mensagens 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8400
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   225
         Top             =   30
         Visible         =   0   'False
         Width           =   660
         Begin VB.Label Label_Mensagens 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00CBB534&
            Height          =   195
            Left            =   345
            TabIndex        =   226
            Top             =   30
            Width           =   105
         End
         Begin VB.Image Icon_Mensagens 
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            Top             =   30
            Width           =   255
         End
      End
      Begin VB.PictureBox Botao_Legendas 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   140
         TabIndex        =   121
         Top             =   60
         Width           =   2100
         Begin VB.Image Icon_Legendas 
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            Top             =   30
            Width           =   255
         End
         Begin VB.Label Label_Legendas 
            BackStyle       =   0  'Transparent
            Caption         =   "Legendas on-line"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   360
            TabIndex        =   122
            Top             =   30
            Width           =   1470
         End
      End
      Begin VB.Image Botao_Redimensionar 
         Height          =   150
         Left            =   9360
         Picture         =   "Form_Principal.frx":24F8A
         Top             =   120
         Width           =   150
      End
      Begin VB.Image Icon_Barra_Informacoes 
         Height          =   300
         Index           =   4
         Left            =   9600
         Top             =   60
         Width           =   600
      End
      Begin VB.Image Icon_Barra_Informacoes 
         Height          =   300
         Index           =   2
         Left            =   2040
         Top             =   60
         Width           =   585
      End
      Begin VB.Image Icon_Barra_Informacoes 
         Height          =   300
         Index           =   1
         Left            =   1440
         Top             =   60
         Width           =   585
      End
      Begin VB.Image Icon_Barra_Informacoes 
         Height          =   300
         Index           =   3
         Left            =   2640
         ToolTipText     =   "Ocultar capa"
         Top             =   60
         Width           =   585
      End
      Begin VB.Image Icon_Barra_Informacoes 
         Height          =   300
         Index           =   5
         Left            =   10200
         ToolTipText     =   "Ocultar playlist"
         Top             =   60
         Width           =   585
      End
      Begin VB.Image Icon_Barra_Informacoes 
         Height          =   300
         Index           =   0
         Left            =   840
         Top             =   60
         Width           =   600
      End
      Begin VB.Label Label_Contador 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   420
      End
      Begin VB.Image Fundo_Barra_Informacoes 
         Enabled         =   0   'False
         Height          =   435
         Left            =   0
         Picture         =   "Form_Principal.frx":2510C
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox Barra_Botoes_Musica 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   3120
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   3960
      Width           =   10935
      Begin VB.Image Imagem_Votar 
         Height          =   285
         Left            =   6480
         Picture         =   "Form_Principal.frx":25EE6
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eu gosto!"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   20
         Left            =   5520
         TabIndex        =   474
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eliminar"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   19
         Left            =   4680
         TabIndex        =   224
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Editar"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   18
         Left            =   3840
         TabIndex        =   223
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Novo evento"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   17
         Left            =   2520
         TabIndex        =   222
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enviar email"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   16
         Left            =   1200
         TabIndex        =   221
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remover"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   220
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Editar"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   14
         Left            =   9720
         TabIndex        =   219
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Novo contacto"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   13
         Left            =   8040
         TabIndex        =   218
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enviar mensagem"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   12
         Left            =   6240
         TabIndex        =   203
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O que ando a ouvir"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   11
         Left            =   4320
         TabIndex        =   200
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enviar convite"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   10
         Left            =   2760
         TabIndex        =   199
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ver o perfil"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   9
         Left            =   1440
         TabIndex        =   198
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guardar lista"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   144
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nova lista"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   7
         Left            =   9240
         TabIndex        =   126
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Linha_Barra_Botoes_Musica 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   15
         Left            =   0
         TabIndex        =   113
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ A minha música"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   7440
         TabIndex        =   112
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transferir"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   6480
         TabIndex        =   111
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adicionar link"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   5
         Left            =   5160
         TabIndex        =   110
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Criar conta"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   3960
         TabIndex        =   109
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar sessão"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   6
         Left            =   2520
         TabIndex        =   108
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remover"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   4
         Left            =   1560
         TabIndex        =   107
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label_Botao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nova biblioteca"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   106
         Top             =   120
         Width           =   1320
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
      ScaleWidth      =   905
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   13575
      Begin VB.PictureBox pichook 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3840
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   13200
         ToolTipText     =   "Fechar"
         Top             =   0
         Width           =   195
      End
      Begin VB.Image Botao_Minimizar 
         Height          =   135
         Left            =   12360
         ToolTipText     =   "Minimizar"
         Top             =   0
         Width           =   135
      End
      Begin VB.Image Botao_Restaurar 
         Height          =   135
         Left            =   12840
         ToolTipText     =   "Restaurar"
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Botao_Maximizar 
         Height          =   135
         Left            =   11880
         ToolTipText     =   "Maximizar"
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label_Menu 
         AutoSize        =   -1  'True
         BackColor       =   &H00101010&
         BackStyle       =   0  'Transparent
         Caption         =   "Ficheiro"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4560
         TabIndex        =   19
         Top             =   150
         Width           =   660
      End
      Begin VB.Label Label_Menu 
         AutoSize        =   -1  'True
         BackColor       =   &H00101010&
         BackStyle       =   0  'Transparent
         Caption         =   "Editar"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   5520
         TabIndex        =   18
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label_Menu 
         AutoSize        =   -1  'True
         BackColor       =   &H00101010&
         BackStyle       =   0  'Transparent
         Caption         =   "Ver"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   6240
         TabIndex        =   17
         Top             =   150
         Width           =   300
      End
      Begin VB.Label Label_Menu 
         AutoSize        =   -1  'True
         BackColor       =   &H00101010&
         BackStyle       =   0  'Transparent
         Caption         =   "Ferramentas"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   7800
         TabIndex        =   16
         Top             =   150
         Width           =   1080
      End
      Begin VB.Label Label_Menu 
         AutoSize        =   -1  'True
         BackColor       =   &H00101010&
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuda"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   9120
         TabIndex        =   15
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label_Menu 
         AutoSize        =   -1  'True
         BackColor       =   &H00101010&
         BackStyle       =   0  'Transparent
         Caption         =   "Controlos"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   6720
         TabIndex        =   14
         Top             =   150
         Width           =   825
      End
      Begin VB.Image Icon_do_Programa 
         Enabled         =   0   'False
         Height          =   210
         Left            =   75
         Picture         =   "Form_Principal.frx":26480
         Top             =   60
         Width           =   210
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "NPlayer (beta) - Nikyts software"
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
         Left            =   465
         TabIndex        =   13
         Top             =   120
         Width           =   3180
      End
      Begin VB.Label Shape_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00838F89&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   9000
         TabIndex        =   59
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Shape_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00838F89&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   7680
         TabIndex        =   58
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Shape_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00838F89&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   6720
         TabIndex        =   57
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Shape_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00838F89&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   6120
         TabIndex        =   56
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Shape_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00838F89&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   55
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Shape_Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H00838F89&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   54
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox Barra_Player 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1089
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   16335
      Begin VB.PictureBox Barra_Botoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00212121&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   0
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   306
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   4590
         Begin VB.PictureBox Picture_Slide_Som 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2700
            ScaleHeight     =   16
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   110
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   300
            Width           =   1650
            Begin VB.TextBox Text_Slide_Som 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   600
               TabIndex        =   101
               Top             =   15
               Visible         =   0   'False
               Width           =   210
            End
            Begin VB.PictureBox Slide_Som 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   0
               ScaleHeight     =   14
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   14
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   0
               Width           =   210
            End
         End
         Begin VB.Image Botao_Play 
            Height          =   810
            Left            =   720
            ToolTipText     =   "Reproduzir"
            Top             =   360
            Width           =   765
         End
         Begin VB.Image Botao_Antes 
            Height          =   600
            Left            =   240
            ToolTipText     =   "Faixa anterior"
            Top             =   180
            Width           =   555
         End
         Begin VB.Image Botao_Seguinte 
            Height          =   585
            Left            =   1440
            ToolTipText     =   "Faixa seguinte"
            Top             =   180
            Width           =   540
         End
         Begin VB.Image Botao_Mudo 
            Height          =   150
            Left            =   2400
            ToolTipText     =   "Mudo"
            Top             =   330
            Width           =   180
         End
         Begin VB.Image Botao_Pausa 
            Height          =   810
            Left            =   720
            ToolTipText     =   "Pausa"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Image Fundo_Barra_Botoes 
            Enabled         =   0   'False
            Height          =   1050
            Left            =   0
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.TextBox Text_Visualizacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   12360
         TabIndex        =   143
         Text            =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Barra_Faixa 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1BEB6&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   4590
         ScaleHeight     =   55
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   513
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   7695
         Begin VB.TextBox Text_Classificacao 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   6240
            TabIndex        =   103
            Top             =   120
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.Timer Timer_Slider_Video 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   0
            Top             =   0
         End
         Begin VB.PictureBox SliderBar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00BFCCC4&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   150
            Left            =   720
            ScaleHeight     =   10
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   420
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   525
            Width           =   6300
            Begin VB.PictureBox Image_Progresso 
               Appearance      =   0  'Flat
               BackColor       =   &H00212121&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   135
               Left            =   0
               ScaleHeight     =   9
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1
               TabIndex        =   96
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   15
            End
            Begin VB.PictureBox Slide 
               Appearance      =   0  'Flat
               BackColor       =   &H00212121&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   150
               Left            =   0
               ScaleHeight     =   10
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   10
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Image Image_Barra_Slide 
               Height          =   150
               Left            =   0
               Picture         =   "Form_Principal.frx":2672A
               Top             =   0
               Visible         =   0   'False
               Width           =   6300
            End
         End
         Begin VB.Image Estrela 
            Height          =   165
            Index           =   4
            Left            =   7470
            Top             =   120
            Width           =   165
         End
         Begin VB.Image Estrela 
            Height          =   165
            Index           =   3
            Left            =   7200
            Top             =   120
            Width           =   165
         End
         Begin VB.Image Estrela 
            Height          =   165
            Index           =   2
            Left            =   6960
            Top             =   120
            Width           =   165
         End
         Begin VB.Image Estrela 
            Height          =   165
            Index           =   1
            Left            =   6735
            Top             =   120
            Width           =   165
         End
         Begin VB.Image Estrela 
            Height          =   165
            Index           =   0
            Left            =   6480
            Top             =   120
            Width           =   165
         End
         Begin VB.Label Label_Faixa 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label Label_Duracao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00A2AFA7&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   6720
            TabIndex        =   21
            Top             =   525
            Width           =   900
         End
         Begin VB.Label Tempo_Estimado 
            BackColor       =   &H00B1BEB6&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   0
            TabIndex        =   7
            Top             =   525
            Width           =   900
         End
      End
      Begin VB.PictureBox Barra_Caixa_Pesquisar_Musica 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   13920
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   240
         Width           =   2610
         Begin VB.TextBox Text_Pesquisar_Musica 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00808080&
            Height          =   300
            Left            =   360
            TabIndex        =   128
            Top             =   30
            Width           =   1260
         End
         Begin VB.Image Image_Lupa 
            Enabled         =   0   'False
            Height          =   255
            Left            =   1920
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Image Icon_Visao 
         Height          =   330
         Index           =   2
         Left            =   13440
         ToolTipText     =   "Album art"
         Top             =   240
         Width           =   390
      End
      Begin VB.Image Icon_Visao 
         Height          =   330
         Index           =   1
         Left            =   13080
         ToolTipText     =   "Pesquisa avançada"
         Top             =   240
         Width           =   390
      End
      Begin VB.Image Icon_Visao 
         Height          =   330
         Index           =   0
         Left            =   12720
         ToolTipText     =   "Simples"
         Top             =   240
         Width           =   405
      End
      Begin VB.Image Fundo_Barra_Player 
         Enabled         =   0   'False
         Height          =   1050
         Left            =   0
         Picture         =   "Form_Principal.frx":26B3C
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Barra_Lateral 
      Appearance      =   0  'Flat
      BackColor       =   &H00101010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8505
      Left            =   0
      ScaleHeight     =   567
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1515
      Width           =   3135
      Begin VB.PictureBox Barra_Conexao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   0
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   209
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   8040
         Width           =   3135
         Begin VB.Label Linha_Barra_Conexao 
            BackColor       =   &H00808080&
            Enabled         =   0   'False
            Height          =   15
            Left            =   0
            TabIndex        =   115
            Top             =   0
            Width           =   3495
         End
      End
      Begin VB.PictureBox Frame_Separadores 
         Appearance      =   0  'Flat
         BackColor       =   &H00101010&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   0
         ScaleHeight     =   321
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   0
         Width           =   3015
         Begin VB.PictureBox Frame_Separador_Barra_Lateral 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1545
            Index           =   1
            Left            =   120
            ScaleHeight     =   103
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   1560
            Width           =   2535
            Begin VB.Image Icon_Topico 
               Enabled         =   0   'False
               Height          =   240
               Index           =   5
               Left            =   120
               Picture         =   "Form_Principal.frx":28C4E
               Top             =   1080
               Width           =   210
            End
            Begin VB.Label Label_Topico_Programas 
               BackStyle       =   0  'Transparent
               Caption         =   "App library"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   360
               TabIndex        =   232
               Top             =   1125
               Width           =   1920
            End
            Begin VB.Label Label_Topico_MusicLink 
               BackStyle       =   0  'Transparent
               Caption         =   "Music link"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   360
               TabIndex        =   197
               Top             =   405
               Width           =   1215
            End
            Begin VB.Image Icon_Topico 
               Enabled         =   0   'False
               Height          =   240
               Index           =   3
               Left            =   120
               Picture         =   "Form_Principal.frx":28F50
               Top             =   360
               Width           =   210
            End
            Begin VB.Label Label_Topico_Drive 
               BackStyle       =   0  'Transparent
               Caption         =   "My other drive"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   360
               TabIndex        =   145
               Top             =   765
               Width           =   1320
            End
            Begin VB.Image Icon_Topico 
               Enabled         =   0   'False
               Height          =   225
               Index           =   4
               Left            =   120
               Picture         =   "Form_Principal.frx":29252
               Top             =   720
               Width           =   210
            End
            Begin VB.Image Icon_Topico 
               Enabled         =   0   'False
               Height          =   240
               Index           =   2
               Left            =   120
               Picture         =   "Form_Principal.frx":294C1
               Top             =   0
               Width           =   240
            End
            Begin VB.Label Label_Topico_Radio 
               BackStyle       =   0  'Transparent
               Caption         =   "Rádio"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   360
               TabIndex        =   92
               Top             =   45
               Width           =   480
            End
            Begin VB.Image Shape_Topico 
               Enabled         =   0   'False
               Height          =   300
               Index           =   2
               Left            =   0
               Top             =   0
               Width           =   300
            End
            Begin VB.Image Shape_Topico 
               Enabled         =   0   'False
               Height          =   300
               Index           =   4
               Left            =   0
               Top             =   720
               Width           =   300
            End
            Begin VB.Image Shape_Topico 
               Enabled         =   0   'False
               Height          =   300
               Index           =   3
               Left            =   0
               Top             =   360
               Width           =   300
            End
            Begin VB.Image Shape_Topico 
               Enabled         =   0   'False
               Height          =   300
               Index           =   5
               Left            =   0
               Top             =   1080
               Width           =   300
            End
         End
         Begin VB.PictureBox Separador_Barra_Lateral 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   2
            Left            =   120
            ScaleHeight     =   20
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   3360
            Width           =   2535
            Begin VB.Label Label_Topico_Barra_Lateral 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Listas"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   90
               Top             =   30
               Width           =   570
            End
         End
         Begin VB.PictureBox Separador_Barra_Lateral 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   120
            ScaleHeight     =   20
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2535
            Begin VB.Label Label_Topico_Barra_Lateral 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Serviços"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   88
               Top             =   30
               Width           =   840
            End
         End
         Begin VB.PictureBox Separador_Barra_Lateral 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   120
            ScaleHeight     =   20
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   0
            Width           =   2535
            Begin VB.Label Label_Topico_Barra_Lateral 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Biblioteca"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   83
               Top             =   30
               Width           =   960
            End
         End
         Begin VB.PictureBox Frame_Separador_Barra_Lateral 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   2
            Left            =   120
            ScaleHeight     =   65
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   3600
            Width           =   2535
            Begin VB.TextBox Text_Nome_Lista 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1200
               TabIndex        =   130
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.DirListBox Dir_Lista 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Height          =   315
               Left            =   1800
               TabIndex        =   81
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.FileListBox File_Lista 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Height          =   420
               Left            =   1800
               Pattern         =   "*.ini"
               TabIndex        =   80
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label_Topico_Lista 
               BackStyle       =   0  'Transparent
               Caption         =   "Lista 1"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   129
               Top             =   0
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.Image Icon_Topico_Lista 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   120
               Picture         =   "Form_Principal.frx":29803
               Top             =   0
               Visible         =   0   'False
               Width           =   210
            End
            Begin VB.Image Shape_Topico_Lista 
               Enabled         =   0   'False
               Height          =   300
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   300
            End
         End
         Begin VB.PictureBox Frame_Separador_Barra_Lateral 
            Appearance      =   0  'Flat
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   660
            Index           =   0
            Left            =   120
            ScaleHeight     =   44
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   360
            Width           =   2535
            Begin VB.Image Icon_Topico 
               Enabled         =   0   'False
               Height          =   225
               Index           =   1
               Left            =   240
               Picture         =   "Form_Principal.frx":29B05
               Top             =   300
               Width           =   210
            End
            Begin VB.Image Icon_Topico 
               Enabled         =   0   'False
               Height          =   225
               Index           =   0
               Left            =   240
               Picture         =   "Form_Principal.frx":29DDB
               Top             =   0
               Width           =   195
            End
            Begin VB.Label Label_Topico_Musica 
               BackStyle       =   0  'Transparent
               Caption         =   "Música"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   600
               TabIndex        =   86
               Top             =   45
               Width           =   570
            End
            Begin VB.Label Label_Topico_Filmes 
               BackStyle       =   0  'Transparent
               Caption         =   "Filmes"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   600
               TabIndex        =   85
               Top             =   285
               Width           =   540
            End
            Begin VB.Image Shape_Topico 
               Enabled         =   0   'False
               Height          =   300
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   300
            End
            Begin VB.Image Shape_Topico 
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   0
               Top             =   240
               Width           =   300
            End
         End
      End
      Begin VB.PictureBox Frame_Capa 
         Appearance      =   0  'Flat
         BackColor       =   &H00101010&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3345
         Left            =   0
         ScaleHeight     =   223
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   209
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4560
         Visible         =   0   'False
         Width           =   3135
         Begin VB.PictureBox Separador_Barra_Lateral 
            Appearance      =   0  'Flat
            BackColor       =   &H00313131&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   3
            Left            =   0
            ScaleHeight     =   20
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   209
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   3135
            Begin VB.Label Label_Topico_Barra_Lateral 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Capa do album"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   125
               Top             =   0
               Width           =   1770
            End
         End
         Begin VB.PictureBox Pic_Capa_Album 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00101010&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   3000
            Left            =   0
            ScaleHeight     =   200
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   180
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   360
            Width           =   2700
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Musica 
      Height          =   735
      Left            =   14640
      TabIndex        =   70
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Filmes 
      Height          =   735
      Left            =   15240
      TabIndex        =   71
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Radio 
      Height          =   735
      Left            =   16440
      TabIndex        =   72
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Loja 
      Height          =   735
      Left            =   15840
      TabIndex        =   73
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Minha_Musica 
      Height          =   735
      Left            =   17040
      TabIndex        =   74
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Listas 
      Height          =   735
      Left            =   17640
      TabIndex        =   95
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      OLEDropMode     =   1
   End
   Begin VB.PictureBox Frame_Album 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   3120
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2160
      Width           =   10935
      Begin VB.PictureBox Frame_Slide_Album 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   273
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   4095
         Begin VB.FileListBox File_Ficheiros 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            Height          =   225
            Left            =   1920
            Pattern         =   "*.mp3"
            TabIndex        =   142
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.ListBox Lista_Pastas 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H00000000&
            Height          =   225
            ItemData        =   "Form_Principal.frx":2A075
            Left            =   1920
            List            =   "Form_Principal.frx":2A077
            Sorted          =   -1  'True
            TabIndex        =   141
            Top             =   0
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox Text_Caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1920
            TabIndex        =   140
            Top             =   720
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.PictureBox Frame_Slide 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   0
            ScaleHeight     =   105
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   129
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   0
            Width           =   1935
            Begin VB.TextBox Label_Nome_Album 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00020202&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   0
               Left            =   0
               TabIndex        =   139
               Text            =   "Nome do album"
               Top             =   1080
               Width           =   1455
            End
            Begin VB.PictureBox Image_Album 
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   495
               Index           =   0
               Left            =   0
               ScaleHeight     =   33
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   33
               TabIndex        =   138
               Top             =   0
               Width           =   495
            End
            Begin VB.Timer Timer_Mover 
               Enabled         =   0   'False
               Interval        =   1
               Left            =   1080
               Top             =   0
            End
            Begin VB.Timer Timer_Album 
               Enabled         =   0   'False
               Interval        =   1
               Left            =   720
               Top             =   0
            End
            Begin VB.Label Label_Directorio_Album 
               Alignment       =   2  'Center
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   136
               Top             =   600
               Visible         =   0   'False
               Width           =   1380
            End
         End
      End
      Begin VB.PictureBox Frame_Grelhas_Pesquisa 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   7560
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   3375
         Begin MSFlexGridLib.MSFlexGrid Grelha_Artista 
            Height          =   615
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1085
            _Version        =   393216
            Rows            =   1
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   3223857
            ForeColorFixed  =   12632256
            BackColorSel    =   13870394
            ForeColorSel    =   16777215
            BackColorBkg    =   11254195
            GridColor       =   14737632
            GridColorFixed  =   2763306
            Redraw          =   -1  'True
            FocusRect       =   0
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            BorderStyle     =   0
            Appearance      =   0
            OLEDropMode     =   1
         End
         Begin MSFlexGridLib.MSFlexGrid Grelha_Album 
            Height          =   615
            Left            =   2160
            TabIndex        =   64
            Top             =   0
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1085
            _Version        =   393216
            Rows            =   1
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   3223857
            ForeColorFixed  =   12632256
            BackColorSel    =   13870394
            ForeColorSel    =   16777215
            BackColorBkg    =   11254195
            GridColor       =   14737632
            GridColorFixed  =   2763306
            Redraw          =   -1  'True
            FocusRect       =   0
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            BorderStyle     =   0
            Appearance      =   0
            OLEDropMode     =   1
         End
         Begin MSFlexGridLib.MSFlexGrid Grelha_Genero 
            Height          =   615
            Left            =   1080
            TabIndex        =   65
            Top             =   0
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1085
            _Version        =   393216
            Rows            =   1
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   3223857
            ForeColorFixed  =   12632256
            BackColorSel    =   13870394
            ForeColorSel    =   16777215
            BackColorBkg    =   11254195
            GridColor       =   14737632
            GridColorFixed  =   2763306
            Redraw          =   -1  'True
            FocusRect       =   0
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            BorderStyle     =   0
            Appearance      =   0
            OLEDropMode     =   1
         End
      End
      Begin VB.PictureBox Barra_Slider_Album 
         Appearance      =   0  'Flat
         BackColor       =   &H00212121&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4320
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   169
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   2535
         Begin VB.PictureBox Barra_Slider_Album_Center 
            Appearance      =   0  'Flat
            BackColor       =   &H00313131&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   89
            TabIndex        =   132
            Top             =   0
            Width           =   1335
            Begin VB.PictureBox Slide_Album 
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   0
               Picture         =   "Form_Principal.frx":2A079
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   26
               TabIndex        =   137
               Top             =   15
               Width           =   390
            End
            Begin VB.Image Fundo_Barra_Slider_Album_Center 
               Enabled         =   0   'False
               Height          =   240
               Left            =   0
               Picture         =   "Form_Principal.frx":2A56B
               Top             =   0
               Width           =   675
            End
         End
         Begin VB.Image Fundo_Barra_Slider_Album_Dir 
            Height          =   240
            Left            =   2160
            Picture         =   "Form_Principal.frx":2AE2D
            Top             =   0
            Width           =   315
         End
         Begin VB.Image Fundo_Barra_Slider_Album_Esq 
            Height          =   240
            Left            =   0
            Picture         =   "Form_Principal.frx":2B26F
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.Label Label_Album 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Eminem"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   4320
         TabIndex        =   133
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Linha_Frame_Album 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   30
         Left            =   4320
         TabIndex        =   131
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label_Nenhum_Album 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nenhum album disponivel"
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
         Left            =   4320
         TabIndex        =   76
         Top             =   0
         Width           =   3180
      End
   End
   Begin VB.PictureBox Barra_Playlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00ABB9B3&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   14640
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
      Begin VB.TextBox Text_Lista_Actual 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   94
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid Grelha_Lista_Em_Reproducao 
         Height          =   735
         Left            =   120
         TabIndex        =   93
         Top             =   120
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1296
         _Version        =   393216
         Cols            =   16
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   3223857
         ForeColorFixed  =   12632256
         BackColorSel    =   13870394
         ForeColorSel    =   16777215
         BackColorBkg    =   11254195
         GridColor       =   14737632
         GridColorFixed  =   2763306
         Redraw          =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         OLEDropMode     =   1
      End
      Begin VB.Label Linha_Barra_Playlist 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   0
         TabIndex        =   105
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Label_Carregar_Favoritos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar favoritos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   1080
         Visible         =   0   'False
         Width           =   1605
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Contactos 
      Height          =   735
      Left            =   15840
      TabIndex        =   146
      Top             =   4200
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Comunidade 
      Height          =   735
      Left            =   17040
      TabIndex        =   189
      Top             =   5040
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Ficheiros 
      Height          =   735
      Left            =   17640
      TabIndex        =   190
      Top             =   4200
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Favoritos 
      Height          =   735
      Left            =   16440
      TabIndex        =   191
      Top             =   5040
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Recentes 
      Height          =   735
      Left            =   15840
      TabIndex        =   192
      Top             =   5040
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Eventos 
      Height          =   735
      Left            =   16440
      TabIndex        =   193
      Top             =   4200
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Mensagens 
      Height          =   735
      Left            =   17040
      TabIndex        =   194
      Top             =   4200
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      OLEDropMode     =   1
   End
   Begin VB.PictureBox Frame_My_Drive 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12615
      Left            =   3120
      ScaleHeight     =   841
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   985
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   17160
      Visible         =   0   'False
      Width           =   14775
      Begin VB.PictureBox Frame_Home 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   11895
         Left            =   480
         ScaleHeight     =   793
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   956
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   0
         Width           =   14340
         Begin VB.PictureBox Picture_Tabela 
            Appearance      =   0  'Flat
            BackColor       =   &H00212121&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4560
            Index           =   2
            Left            =   9120
            ScaleHeight     =   304
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   268
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   5640
            Width           =   4020
            Begin VB.Label Label_Ficheiros 
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Gerenciamento de ficheiros: Ilimitado"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   179
               Top             =   3960
               Width           =   4095
            End
            Begin VB.Label Label_Dados 
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Armazenamento de dados: Ilimitado"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   178
               Top             =   3720
               Width           =   4095
            End
            Begin VB.Label Label_Funcionalidades 
               AutoSize        =   -1  'True
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Funcionalidades?"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   177
               Top             =   3480
               Width           =   1440
            End
            Begin VB.Label Label_Mensalidade 
               AutoSize        =   -1  'True
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Mensalidade p/ mês"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   170
               Top             =   2475
               Width           =   1710
            End
            Begin VB.Label Label_Preco 
               AutoSize        =   -1  'True
               BackColor       =   &H00C4AD2F&
               BackStyle       =   0  'Transparent
               Caption         =   "30€"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   48
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1125
               Index           =   2
               Left            =   600
               TabIndex        =   167
               Top             =   1320
               Width           =   1305
            End
            Begin VB.Label Label_Plano 
               AutoSize        =   -1  'True
               BackColor       =   &H00C4AD2F&
               BackStyle       =   0  'Transparent
               Caption         =   "PROFISSIONAL"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   2
               Left            =   120
               TabIndex        =   164
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.PictureBox Picture_Tabela 
            Appearance      =   0  'Flat
            BackColor       =   &H00212121&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4560
            Index           =   1
            Left            =   4560
            ScaleHeight     =   304
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   268
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   5640
            Width           =   4020
            Begin VB.Label Label_Popular 
               AutoSize        =   -1  'True
               BackColor       =   &H00C4AD2F&
               BackStyle       =   0  'Transparent
               Caption         =   "Mais popular"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   2580
               TabIndex        =   180
               Top             =   195
               Width           =   1080
            End
            Begin VB.Shape Shape_Popular 
               BackColor       =   &H0080FFFF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00404040&
               BorderStyle     =   0  'Transparent
               Height          =   375
               Left            =   2400
               Shape           =   4  'Rounded Rectangle
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label_Ficheiros 
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Gerenciamento de ficheiros: Limitado"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   176
               Top             =   3960
               Width           =   4095
            End
            Begin VB.Label Label_Dados 
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Armazenamento de dados: Ilimitado"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   175
               Top             =   3720
               Width           =   4095
            End
            Begin VB.Label Label_Funcionalidades 
               AutoSize        =   -1  'True
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Funcionalidades?"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   174
               Top             =   3480
               Width           =   1440
            End
            Begin VB.Label Label_Mensalidade 
               AutoSize        =   -1  'True
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Mensalidade p/ mês"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   169
               Top             =   2475
               Width           =   1710
            End
            Begin VB.Label Label_Preco 
               AutoSize        =   -1  'True
               BackColor       =   &H00C4AD2F&
               BackStyle       =   0  'Transparent
               Caption         =   "10€"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   48
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1125
               Index           =   1
               Left            =   600
               TabIndex        =   166
               Top             =   1320
               Width           =   1305
            End
            Begin VB.Label Label_Plano 
               AutoSize        =   -1  'True
               BackColor       =   &H00C4AD2F&
               BackStyle       =   0  'Transparent
               Caption         =   "AVANÇADO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   1
               Left            =   120
               TabIndex        =   163
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture_Tabela 
            Appearance      =   0  'Flat
            BackColor       =   &H00212121&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4560
            Index           =   0
            Left            =   0
            ScaleHeight     =   304
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   268
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   5640
            Width           =   4020
            Begin VB.Label Label_Ficheiros 
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Gerenciamento ficheiros: Não disponivel"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   173
               Top             =   3960
               Width           =   4095
            End
            Begin VB.Label Label_Dados 
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Armazenamento de dados: Limitado"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   172
               Top             =   3720
               Width           =   4095
            End
            Begin VB.Label Label_Funcionalidades 
               AutoSize        =   -1  'True
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Funcionalidades?"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   171
               Top             =   3480
               Width           =   1440
            End
            Begin VB.Label Label_Mensalidade 
               AutoSize        =   -1  'True
               BackColor       =   &H00CBB534&
               BackStyle       =   0  'Transparent
               Caption         =   "Mensalidade p/ mês"
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   168
               Top             =   2475
               Width           =   1710
            End
            Begin VB.Label Label_Preco 
               AutoSize        =   -1  'True
               BackColor       =   &H00C4AD2F&
               BackStyle       =   0  'Transparent
               Caption         =   "0€"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   48
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1125
               Index           =   0
               Left            =   600
               TabIndex        =   165
               Top             =   1320
               Width           =   870
            End
            Begin VB.Label Label_Plano 
               AutoSize        =   -1  'True
               BackColor       =   &H00C4AD2F&
               BackStyle       =   0  'Transparent
               Caption         =   "PADRÃO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   0
               Left            =   120
               TabIndex        =   162
               Top             =   240
               Width           =   795
            End
         End
         Begin VB.Label Label_Texto 
            AutoSize        =   -1  'True
            BackColor       =   &H00CBB534&
            BackStyle       =   0  'Transparent
            Caption         =   "PLANOS QUE ATENDAM AS SUAS NECESSIDADES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Index           =   3
            Left            =   600
            TabIndex        =   158
            Top             =   4680
            Width           =   7005
         End
         Begin VB.Label Label_Texto 
            AutoSize        =   -1  'True
            BackColor       =   &H00C4AD2F&
            BackStyle       =   0  'Transparent
            Caption         =   "My other drive"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   570
            Index           =   0
            Left            =   6840
            TabIndex        =   157
            Top             =   360
            Width           =   3915
         End
         Begin VB.Label Label_Texto 
            AutoSize        =   -1  'True
            BackColor       =   &H00C4AD2F&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTALMENTE GRÁTIS"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   7320
            TabIndex        =   156
            Top             =   1920
            Width           =   3435
         End
         Begin VB.Label Label_Texto 
            AutoSize        =   -1  'True
            BackColor       =   &H00CBB534&
            BackStyle       =   0  'Transparent
            Caption         =   "SERVIDORES E INFRA-ESTRUTURAS DE DEMANDA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   6840
            TabIndex        =   155
            Top             =   960
            Width           =   4560
         End
         Begin VB.Label Label_Aderir_Agora 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Aderir agora"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00CBB534&
            Height          =   270
            Left            =   7320
            TabIndex        =   154
            Top             =   2760
            Width           =   3360
         End
         Begin VB.Image Imagem_Nuvens 
            Enabled         =   0   'False
            Height          =   4200
            Left            =   600
            Picture         =   "Form_Principal.frx":2B6B1
            Top             =   0
            Width           =   5085
         End
         Begin VB.Image Botao_Aderir_Agora 
            Height          =   720
            Left            =   7320
            Picture         =   "Form_Principal.frx":33B52
            Top             =   2520
            Width           =   3360
         End
         Begin VB.Shape Shape_Sky 
            BackColor       =   &H00CBB534&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00CBB534&
            Height          =   4200
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   13980
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C4AD2F&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   10680
         TabIndex        =   152
         Top             =   2040
         Width           =   60
      End
      Begin VB.Shape Shape_Sky 
         BackColor       =   &H00CBB534&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00CBB534&
         Height          =   4200
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Frame_Music_Link 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00CBB534&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   3120
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   181
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   6855
      Begin VB.PictureBox Frame_Caixa_Pesquisa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   0
         Picture         =   "Form_Principal.frx":3B994
         ScaleHeight     =   52
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   452
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   840
         Width           =   6780
         Begin VB.TextBox Text_Pesquisar 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   240
            TabIndex        =   185
            Text            =   "Pesquisar música"
            Top             =   264
            Width           =   5325
         End
         Begin VB.Image Botao_Pesquisar 
            Height          =   600
            Left            =   6090
            Picture         =   "Form_Principal.frx":4CD46
            Top             =   90
            Width           =   600
         End
      End
      Begin VB.Label Label_Texto 
         AutoSize        =   -1  'True
         BackColor       =   &H00CBB534&
         BackStyle       =   0  'Transparent
         Caption         =   "PESQUISE, OUÇA E TRANSFIRA MÚSICAS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   1560
         TabIndex        =   182
         Top             =   600
         Width           =   3795
      End
      Begin VB.Label Label_Texto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C4AD2F&
         BackStyle       =   0  'Transparent
         Caption         =   "Music Link"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   570
         Index           =   4
         Left            =   1560
         TabIndex        =   183
         Top             =   0
         Width           =   2820
      End
      Begin VB.Image Fundo_Frame_Music_Link 
         Enabled         =   0   'False
         Height          =   1260
         Left            =   0
         Top             =   0
         Width           =   1260
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grelha_Amigos 
      Height          =   735
      Left            =   17640
      TabIndex        =   202
      Top             =   5040
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   12632256
      BackColorSel    =   13870394
      ForeColorSel    =   16777215
      BackColorBkg    =   11254195
      GridColor       =   14737632
      GridColorFixed  =   2763306
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.PictureBox Frame_Programas 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBB534&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   3120
      ScaleHeight     =   585
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   825
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   12375
      Begin VB.PictureBox Frame_Informacoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   0
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   816
         TabIndex        =   451
         Top             =   4080
         Visible         =   0   'False
         Width           =   12240
         Begin VB.PictureBox Botao_Frame_Informacoes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   0
            Left            =   8520
            Picture         =   "Form_Principal.frx":4E048
            ScaleHeight     =   29
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   91
            TabIndex        =   463
            TabStop         =   0   'False
            Top             =   1770
            Width           =   1365
            Begin VB.Label Label_Botao_Frame_Informacoes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Transferir"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   464
               Top             =   120
               Width           =   1365
            End
         End
         Begin VB.TextBox Text_Servidor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   2160
            TabIndex        =   462
            Top             =   1800
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtZip 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   4200
            TabIndex        =   461
            Top             =   1800
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.PictureBox Frame_Avaliacao 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1005
            Left            =   8040
            Picture         =   "Form_Principal.frx":4E594
            ScaleHeight     =   67
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   173
            TabIndex        =   458
            TabStop         =   0   'False
            Top             =   360
            Width           =   2595
            Begin VB.Label Label_Votos 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1680
               TabIndex        =   460
               Top             =   600
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label_Frame_Informacoes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "0 avaliações"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   459
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.TextBox Text_Informacao 
            BackColor       =   &H00F9F9F9&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   457
            TabStop         =   0   'False
            Top             =   3600
            Width           =   7575
         End
         Begin VB.PictureBox Botao_Frame_Informacoes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   1
            Left            =   4560
            Picture         =   "Form_Principal.frx":521D6
            ScaleHeight     =   29
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   91
            TabIndex        =   454
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1365
            Begin VB.Label Label_Botao_Frame_Informacoes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Cancelar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   455
               Top             =   120
               Width           =   1365
            End
         End
         Begin VB.PictureBox Botao_Frame_Informacoes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   2
            Left            =   6120
            Picture         =   "Form_Principal.frx":52722
            ScaleHeight     =   29
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   91
            TabIndex        =   452
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1365
            Begin VB.Label Label_Botao_Frame_Informacoes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Executar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   453
               Top             =   120
               Width           =   1365
            End
         End
         Begin NPlayer.NProgressBar ProgressBar1 
            Height          =   375
            Left            =   6960
            TabIndex        =   456
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
         End
         Begin NPlayer.dl dl 
            Left            =   3960
            Top             =   1440
            _ExtentX        =   1799
            _ExtentY        =   1667
         End
         Begin VB.Image Image_Download 
            Height          =   240
            Left            =   600
            Top             =   2970
            Width           =   240
         End
         Begin VB.Image Image_Tela 
            Enabled         =   0   'False
            Height          =   1860
            Left            =   8760
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   3075
         End
         Begin VB.Label Label_Frame_Informacoes 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ficheiro.zip"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   3
            Left            =   600
            TabIndex        =   473
            Top             =   1860
            Width           =   1110
         End
         Begin VB.Label Label_Id_Programa 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2400
            TabIndex        =   472
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label_Transferencias 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6120
            TabIndex        =   471
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label_Site_Programa 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   470
            Top             =   2520
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Label Label_Frame_Informacoes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Site oficial do programa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   600
            TabIndex        =   469
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label Label_Frame_Informacoes 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição do programa"
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   468
            Top             =   960
            Width           =   2010
         End
         Begin VB.Label Label_Frame_Informacoes 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do programa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   345
            Index           =   0
            Left            =   1920
            TabIndex        =   467
            Top             =   600
            Width           =   3075
         End
         Begin VB.Image Image_Logo 
            Enabled         =   0   'False
            Height          =   960
            Left            =   720
            Top             =   480
            Width           =   960
         End
         Begin VB.Shape Shape_Foto 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   1920
            Left            =   8520
            Top             =   2520
            Width           =   3135
         End
         Begin VB.Label Label_Frame_Informacoes 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "(Downloads: 0)"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   5
            Left            =   5160
            TabIndex        =   466
            Top             =   720
            Width           =   1320
         End
         Begin VB.Label Label_Frame_Informacoes 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   960
            TabIndex        =   465
            Top             =   2970
            Width           =   60
         End
         Begin VB.Shape Shape_Transferir 
            BackColor       =   &H00EEF0EF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   600
            Left            =   600
            Top             =   1680
            Width           =   9615
         End
         Begin VB.Shape Shape_Estado 
            BackColor       =   &H00EEF0EF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   600
            Left            =   600
            Top             =   2880
            Width           =   7095
         End
      End
      Begin VB.PictureBox Frame_Lista 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   817
         TabIndex        =   240
         TabStop         =   0   'False
         Top             =   2760
         Visible         =   0   'False
         Width           =   12255
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   20
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   646
            TabStop         =   0   'False
            Top             =   10800
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   20
               Left            =   3480
               Picture         =   "Form_Principal.frx":52C6E
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   651
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   20
                  Left            =   210
                  TabIndex        =   652
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   20
               Left            =   720
               Picture         =   "Form_Principal.frx":53203
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   649
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   20
                  Left            =   360
                  TabIndex        =   650
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   20
               Left            =   6480
               Picture         =   "Form_Principal.frx":53798
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   647
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   20
                  Left            =   360
                  TabIndex        =   648
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   20
               Left            =   9000
               TabIndex        =   653
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   0
               TabIndex        =   664
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   1440
               TabIndex        =   663
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   20
               Left            =   720
               TabIndex        =   662
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   20
               Left            =   720
               TabIndex        =   661
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   20
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   2880
               TabIndex        =   660
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   4320
               TabIndex        =   659
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   5760
               TabIndex        =   658
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   7200
               TabIndex        =   657
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   20
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   20
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   10320
               TabIndex        =   656
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   0
               TabIndex        =   655
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   5760
               TabIndex        =   654
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   19
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   627
            TabStop         =   0   'False
            Top             =   10320
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   19
               Left            =   3480
               Picture         =   "Form_Principal.frx":53C62
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   632
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   19
                  Left            =   210
                  TabIndex        =   633
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   19
               Left            =   720
               Picture         =   "Form_Principal.frx":541F7
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   630
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   19
                  Left            =   360
                  TabIndex        =   631
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   19
               Left            =   6480
               Picture         =   "Form_Principal.frx":5478C
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   628
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   19
                  Left            =   360
                  TabIndex        =   629
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   19
               Left            =   9000
               TabIndex        =   634
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   19
               Left            =   720
               TabIndex        =   643
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   19
               Left            =   720
               TabIndex        =   642
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   19
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   2880
               TabIndex        =   641
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   4320
               TabIndex        =   640
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   5760
               TabIndex        =   639
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   7200
               TabIndex        =   638
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   19
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   19
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   10320
               TabIndex        =   637
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   0
               TabIndex        =   636
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   5760
               TabIndex        =   635
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   0
               TabIndex        =   645
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   1440
               TabIndex        =   644
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   18
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   608
            TabStop         =   0   'False
            Top             =   9720
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   18
               Left            =   3480
               Picture         =   "Form_Principal.frx":54C56
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   613
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   18
                  Left            =   210
                  TabIndex        =   614
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   18
               Left            =   720
               Picture         =   "Form_Principal.frx":551EB
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   611
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   18
                  Left            =   360
                  TabIndex        =   612
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   18
               Left            =   6480
               Picture         =   "Form_Principal.frx":55780
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   609
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   18
                  Left            =   360
                  TabIndex        =   610
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   18
               Left            =   9000
               TabIndex        =   615
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   18
               Left            =   720
               TabIndex        =   624
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   18
               Left            =   720
               TabIndex        =   623
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   18
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   2880
               TabIndex        =   622
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   4320
               TabIndex        =   621
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   5760
               TabIndex        =   620
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   7200
               TabIndex        =   619
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   18
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   18
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   10320
               TabIndex        =   618
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   0
               TabIndex        =   617
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   5760
               TabIndex        =   616
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   0
               TabIndex        =   626
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   1440
               TabIndex        =   625
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   17
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   589
            TabStop         =   0   'False
            Top             =   9120
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   17
               Left            =   3480
               Picture         =   "Form_Principal.frx":55C4A
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   594
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   17
                  Left            =   210
                  TabIndex        =   595
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   17
               Left            =   720
               Picture         =   "Form_Principal.frx":561DF
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   592
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   17
                  Left            =   360
                  TabIndex        =   593
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   17
               Left            =   6480
               Picture         =   "Form_Principal.frx":56774
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   590
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   17
                  Left            =   360
                  TabIndex        =   591
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   17
               Left            =   9000
               TabIndex        =   596
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   17
               Left            =   720
               TabIndex        =   605
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   17
               Left            =   720
               TabIndex        =   604
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   17
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   2880
               TabIndex        =   603
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   4320
               TabIndex        =   602
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   5760
               TabIndex        =   601
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   7200
               TabIndex        =   600
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   17
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   17
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   10320
               TabIndex        =   599
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   0
               TabIndex        =   598
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   5760
               TabIndex        =   597
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   0
               TabIndex        =   607
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   1440
               TabIndex        =   606
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   16
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   570
            TabStop         =   0   'False
            Top             =   8760
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   16
               Left            =   3480
               Picture         =   "Form_Principal.frx":56C3E
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   575
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   16
                  Left            =   210
                  TabIndex        =   576
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   16
               Left            =   720
               Picture         =   "Form_Principal.frx":571D3
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   573
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   16
                  Left            =   360
                  TabIndex        =   574
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   16
               Left            =   6480
               Picture         =   "Form_Principal.frx":57768
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   571
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   16
                  Left            =   360
                  TabIndex        =   572
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   16
               Left            =   9000
               TabIndex        =   577
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   16
               Left            =   720
               TabIndex        =   586
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   16
               Left            =   720
               TabIndex        =   585
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   16
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   2880
               TabIndex        =   584
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   4320
               TabIndex        =   583
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   5760
               TabIndex        =   582
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   7200
               TabIndex        =   581
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   16
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   16
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   10320
               TabIndex        =   580
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   0
               TabIndex        =   579
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   5760
               TabIndex        =   578
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   0
               TabIndex        =   588
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   1440
               TabIndex        =   587
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   15
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   551
            TabStop         =   0   'False
            Top             =   8280
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   15
               Left            =   3480
               Picture         =   "Form_Principal.frx":57C32
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   556
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   15
                  Left            =   210
                  TabIndex        =   557
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   15
               Left            =   720
               Picture         =   "Form_Principal.frx":581C7
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   554
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   15
                  Left            =   360
                  TabIndex        =   555
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   15
               Left            =   6480
               Picture         =   "Form_Principal.frx":5875C
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   552
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   15
                  Left            =   360
                  TabIndex        =   553
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   15
               Left            =   9000
               TabIndex        =   558
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   15
               Left            =   720
               TabIndex        =   567
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   15
               Left            =   720
               TabIndex        =   566
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   15
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   2880
               TabIndex        =   565
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   4320
               TabIndex        =   564
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   5760
               TabIndex        =   563
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   7200
               TabIndex        =   562
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   15
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   15
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   10320
               TabIndex        =   561
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   560
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   5760
               TabIndex        =   559
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   569
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   1440
               TabIndex        =   568
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   14
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   532
            TabStop         =   0   'False
            Top             =   7800
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   14
               Left            =   3480
               Picture         =   "Form_Principal.frx":58C26
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   537
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   14
                  Left            =   210
                  TabIndex        =   538
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   14
               Left            =   720
               Picture         =   "Form_Principal.frx":591BB
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   535
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   14
                  Left            =   360
                  TabIndex        =   536
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   14
               Left            =   6480
               Picture         =   "Form_Principal.frx":59750
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   533
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   14
                  Left            =   360
                  TabIndex        =   534
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   14
               Left            =   9000
               TabIndex        =   539
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   14
               Left            =   720
               TabIndex        =   548
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   14
               Left            =   720
               TabIndex        =   547
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   14
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   2880
               TabIndex        =   546
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   4320
               TabIndex        =   545
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   5760
               TabIndex        =   544
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   7200
               TabIndex        =   543
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   14
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   14
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   10320
               TabIndex        =   542
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   0
               TabIndex        =   541
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   5760
               TabIndex        =   540
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   0
               TabIndex        =   550
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   1440
               TabIndex        =   549
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   13
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   513
            TabStop         =   0   'False
            Top             =   7440
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   13
               Left            =   3480
               Picture         =   "Form_Principal.frx":59C1A
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   518
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   13
                  Left            =   210
                  TabIndex        =   519
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   13
               Left            =   720
               Picture         =   "Form_Principal.frx":5A1AF
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   516
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   13
                  Left            =   360
                  TabIndex        =   517
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   13
               Left            =   6480
               Picture         =   "Form_Principal.frx":5A744
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   514
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   13
                  Left            =   360
                  TabIndex        =   515
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   13
               Left            =   9000
               TabIndex        =   520
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   13
               Left            =   720
               TabIndex        =   529
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   13
               Left            =   720
               TabIndex        =   528
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   13
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   2880
               TabIndex        =   527
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   4320
               TabIndex        =   526
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   5760
               TabIndex        =   525
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   7200
               TabIndex        =   524
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   13
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   13
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   10320
               TabIndex        =   523
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   0
               TabIndex        =   522
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   5760
               TabIndex        =   521
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   0
               TabIndex        =   531
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   1440
               TabIndex        =   530
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   12
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   494
            TabStop         =   0   'False
            Top             =   6960
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   12
               Left            =   3480
               Picture         =   "Form_Principal.frx":5AC0E
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   499
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   12
                  Left            =   210
                  TabIndex        =   500
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   12
               Left            =   720
               Picture         =   "Form_Principal.frx":5B1A3
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   497
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   12
                  Left            =   360
                  TabIndex        =   498
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   12
               Left            =   6480
               Picture         =   "Form_Principal.frx":5B738
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   495
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   12
                  Left            =   360
                  TabIndex        =   496
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   12
               Left            =   9000
               TabIndex        =   501
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   12
               Left            =   720
               TabIndex        =   510
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   12
               Left            =   720
               TabIndex        =   509
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   12
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   2880
               TabIndex        =   508
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   4320
               TabIndex        =   507
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   5760
               TabIndex        =   506
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   7200
               TabIndex        =   505
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   12
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   12
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   10320
               TabIndex        =   504
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   0
               TabIndex        =   503
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   5760
               TabIndex        =   502
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   0
               TabIndex        =   512
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   1440
               TabIndex        =   511
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   11
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   475
            TabStop         =   0   'False
            Top             =   6480
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   11
               Left            =   3480
               Picture         =   "Form_Principal.frx":5BC02
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   480
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   11
                  Left            =   210
                  TabIndex        =   481
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   11
               Left            =   720
               Picture         =   "Form_Principal.frx":5C197
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   478
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   11
                  Left            =   360
                  TabIndex        =   479
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   11
               Left            =   6480
               Picture         =   "Form_Principal.frx":5C72C
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   476
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   11
                  Left            =   360
                  TabIndex        =   477
                  Top             =   90
                  Width           =   750
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   11
               Left            =   9000
               TabIndex        =   482
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   11
               Left            =   720
               TabIndex        =   491
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   11
               Left            =   720
               TabIndex        =   490
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   11
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   2880
               TabIndex        =   489
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   4320
               TabIndex        =   488
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   5760
               TabIndex        =   487
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   7200
               TabIndex        =   486
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   11
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   11
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   10320
               TabIndex        =   485
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   484
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   5760
               TabIndex        =   483
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   493
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   1440
               TabIndex        =   492
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   0
            Left            =   0
            ScaleHeight     =   81
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   421
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   11175
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   0
               Left            =   8880
               TabIndex        =   440
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin NPlayer.dl Download_Programa 
               Left            =   7440
               Top             =   0
               _ExtentX        =   1799
               _ExtentY        =   1667
            End
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   0
               Left            =   6480
               Picture         =   "Form_Principal.frx":5CBF6
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   426
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   0
                  Left            =   360
                  TabIndex        =   427
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   0
               Left            =   720
               Picture         =   "Form_Principal.frx":5D0C0
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   424
               TabStop         =   0   'False
               Top             =   630
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   0
                  Left            =   360
                  TabIndex        =   425
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   0
               Left            =   3480
               Picture         =   "Form_Principal.frx":5D655
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   422
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   0
                  Left            =   300
                  TabIndex        =   423
                  Top             =   90
                  Width           =   780
               End
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   5760
               TabIndex        =   438
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   437
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   10320
               TabIndex        =   436
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   0
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   0
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   7200
               TabIndex        =   435
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   5760
               TabIndex        =   434
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   4320
               TabIndex        =   433
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   2880
               TabIndex        =   432
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   720
               TabIndex        =   431
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   720
               TabIndex        =   430
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   429
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   428
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   9
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   403
            TabStop         =   0   'False
            Top             =   5400
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   9
               Left            =   6480
               Picture         =   "Form_Principal.frx":5DB1F
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   408
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   9
                  Left            =   360
                  TabIndex        =   409
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   9
               Left            =   720
               Picture         =   "Form_Principal.frx":5DFE9
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   406
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   9
                  Left            =   360
                  TabIndex        =   407
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   9
               Left            =   3480
               Picture         =   "Form_Principal.frx":5E57E
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   404
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   9
                  Left            =   210
                  TabIndex        =   405
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   9
               Left            =   9000
               TabIndex        =   449
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   5760
               TabIndex        =   420
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   419
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   10320
               TabIndex        =   418
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   9
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   9
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   7200
               TabIndex        =   417
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   5760
               TabIndex        =   416
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   4320
               TabIndex        =   415
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   2880
               TabIndex        =   414
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   9
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   9
               Left            =   720
               TabIndex        =   413
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   9
               Left            =   720
               TabIndex        =   412
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   1440
               TabIndex        =   411
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   410
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   8
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   385
            TabStop         =   0   'False
            Top             =   4800
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   8
               Left            =   6480
               Picture         =   "Form_Principal.frx":5EB13
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   390
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   8
                  Left            =   360
                  TabIndex        =   391
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   8
               Left            =   720
               Picture         =   "Form_Principal.frx":5EFDD
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   388
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   8
                  Left            =   360
                  TabIndex        =   389
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   8
               Left            =   3480
               Picture         =   "Form_Principal.frx":5F572
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   386
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   8
                  Left            =   210
                  TabIndex        =   387
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   8
               Left            =   9000
               TabIndex        =   448
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   5760
               TabIndex        =   402
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   401
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   10320
               TabIndex        =   400
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   8
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   8
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   7200
               TabIndex        =   399
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   5760
               TabIndex        =   398
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   4320
               TabIndex        =   397
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   2880
               TabIndex        =   396
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   8
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   8
               Left            =   720
               TabIndex        =   395
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   8
               Left            =   720
               TabIndex        =   394
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   1440
               TabIndex        =   393
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   392
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   7
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   367
            TabStop         =   0   'False
            Top             =   4200
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   7
               Left            =   6480
               Picture         =   "Form_Principal.frx":5FB07
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   372
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   7
                  Left            =   360
                  TabIndex        =   373
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   7
               Left            =   720
               Picture         =   "Form_Principal.frx":5FFD1
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   370
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   7
                  Left            =   360
                  TabIndex        =   371
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   7
               Left            =   3480
               Picture         =   "Form_Principal.frx":60566
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   368
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   7
                  Left            =   210
                  TabIndex        =   369
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   7
               Left            =   9000
               TabIndex        =   447
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   5760
               TabIndex        =   384
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   383
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   10320
               TabIndex        =   382
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   7
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   7
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   7200
               TabIndex        =   381
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   5760
               TabIndex        =   380
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   4320
               TabIndex        =   379
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   2880
               TabIndex        =   378
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   7
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   7
               Left            =   720
               TabIndex        =   377
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   7
               Left            =   720
               TabIndex        =   376
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   1440
               TabIndex        =   375
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   374
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   6
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   349
            TabStop         =   0   'False
            Top             =   3600
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   6
               Left            =   6480
               Picture         =   "Form_Principal.frx":60AFB
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   354
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   6
                  Left            =   360
                  TabIndex        =   355
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   6
               Left            =   720
               Picture         =   "Form_Principal.frx":60FC5
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   352
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   6
                  Left            =   360
                  TabIndex        =   353
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   6
               Left            =   3480
               Picture         =   "Form_Principal.frx":6155A
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   350
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   6
                  Left            =   210
                  TabIndex        =   351
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   6
               Left            =   9000
               TabIndex        =   446
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   5760
               TabIndex        =   366
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   365
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   10320
               TabIndex        =   364
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   6
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   6
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   7200
               TabIndex        =   363
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   5760
               TabIndex        =   362
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   4320
               TabIndex        =   361
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   2880
               TabIndex        =   360
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   6
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   6
               Left            =   720
               TabIndex        =   359
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   6
               Left            =   720
               TabIndex        =   358
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   1440
               TabIndex        =   357
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   356
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   5
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   331
            TabStop         =   0   'False
            Top             =   3000
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   5
               Left            =   6480
               Picture         =   "Form_Principal.frx":61AEF
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   336
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   5
                  Left            =   360
                  TabIndex        =   337
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   5
               Left            =   720
               Picture         =   "Form_Principal.frx":61FB9
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   334
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   5
                  Left            =   360
                  TabIndex        =   335
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   5
               Left            =   3480
               Picture         =   "Form_Principal.frx":6254E
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   332
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   5
                  Left            =   210
                  TabIndex        =   333
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   5
               Left            =   9000
               TabIndex        =   445
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   5760
               TabIndex        =   348
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   347
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   10320
               TabIndex        =   346
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   5
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   5
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   7200
               TabIndex        =   345
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   5760
               TabIndex        =   344
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   4320
               TabIndex        =   343
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   2880
               TabIndex        =   342
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   5
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   5
               Left            =   720
               TabIndex        =   341
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   5
               Left            =   720
               TabIndex        =   340
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   1440
               TabIndex        =   339
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   338
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   4
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   313
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   4
               Left            =   6480
               Picture         =   "Form_Principal.frx":62AE3
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   318
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   4
                  Left            =   360
                  TabIndex        =   319
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   4
               Left            =   720
               Picture         =   "Form_Principal.frx":62FAD
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   316
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   4
                  Left            =   360
                  TabIndex        =   317
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   4
               Left            =   3480
               Picture         =   "Form_Principal.frx":63542
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   314
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   4
                  Left            =   210
                  TabIndex        =   315
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   4
               Left            =   9000
               TabIndex        =   444
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   5760
               TabIndex        =   330
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   329
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   10320
               TabIndex        =   328
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   4
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   4
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   7200
               TabIndex        =   327
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   5760
               TabIndex        =   326
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   4320
               TabIndex        =   325
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   2880
               TabIndex        =   324
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   4
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   4
               Left            =   720
               TabIndex        =   323
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   4
               Left            =   720
               TabIndex        =   322
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   1440
               TabIndex        =   321
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   320
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   3
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   295
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   3
               Left            =   6480
               Picture         =   "Form_Principal.frx":63AD7
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   300
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   3
                  Left            =   360
                  TabIndex        =   301
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   3
               Left            =   720
               Picture         =   "Form_Principal.frx":63FA1
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   298
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   3
                  Left            =   360
                  TabIndex        =   299
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   3
               Left            =   3480
               Picture         =   "Form_Principal.frx":64536
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   296
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   3
                  Left            =   210
                  TabIndex        =   297
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   3
               Left            =   9000
               TabIndex        =   443
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   5760
               TabIndex        =   312
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   311
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   10320
               TabIndex        =   310
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   3
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   3
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   7200
               TabIndex        =   309
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   5760
               TabIndex        =   308
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   4320
               TabIndex        =   307
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   2880
               TabIndex        =   306
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   3
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   3
               Left            =   720
               TabIndex        =   305
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   3
               Left            =   720
               TabIndex        =   304
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   1440
               TabIndex        =   303
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   302
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   2
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   277
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   2
               Left            =   6480
               Picture         =   "Form_Principal.frx":64ACB
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   282
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   2
                  Left            =   360
                  TabIndex        =   283
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   2
               Left            =   720
               Picture         =   "Form_Principal.frx":64F95
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   280
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   2
                  Left            =   360
                  TabIndex        =   281
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   2
               Left            =   3480
               Picture         =   "Form_Principal.frx":6552A
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   278
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   2
                  Left            =   210
                  TabIndex        =   279
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   2
               Left            =   8880
               TabIndex        =   442
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   5760
               TabIndex        =   294
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   293
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   10320
               TabIndex        =   292
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   2
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   2
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   7200
               TabIndex        =   291
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   5760
               TabIndex        =   290
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   4320
               TabIndex        =   289
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   2880
               TabIndex        =   288
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   2
               Left            =   720
               TabIndex        =   287
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   2
               Left            =   720
               TabIndex        =   286
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   1440
               TabIndex        =   285
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   284
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   1
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   259
            TabStop         =   0   'False
            Top             =   600
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   1
               Left            =   6480
               Picture         =   "Form_Principal.frx":65ABF
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   264
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   1
                  Left            =   360
                  TabIndex        =   265
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   1
               Left            =   3480
               Picture         =   "Form_Principal.frx":65F89
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   262
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   1
                  Left            =   210
                  TabIndex        =   263
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   1
               Left            =   720
               Picture         =   "Form_Principal.frx":6651E
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   260
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   1
                  Left            =   360
                  TabIndex        =   261
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   1
               Left            =   8880
               TabIndex        =   441
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   5760
               TabIndex        =   276
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   275
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   10320
               TabIndex        =   274
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   1
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   1
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   7200
               TabIndex        =   273
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   5760
               TabIndex        =   272
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   4320
               TabIndex        =   271
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   270
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   720
               TabIndex        =   269
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   1
               Left            =   720
               TabIndex        =   268
               Top             =   120
               Width           =   1665
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   267
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   266
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pic_Linha 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   10
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   745
            TabIndex        =   241
            TabStop         =   0   'False
            Top             =   6000
            Visible         =   0   'False
            Width           =   11175
            Begin VB.PictureBox Botao_Executar_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   10
               Left            =   6480
               Picture         =   "Form_Principal.frx":66AB3
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   96
               TabIndex        =   246
               TabStop         =   0   'False
               Top             =   630
               Width           =   1440
               Begin VB.Label Label_Executar_Programa 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   10
                  Left            =   360
                  TabIndex        =   247
                  Top             =   90
                  Width           =   750
               End
            End
            Begin VB.PictureBox Botao_Mais_Informacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   10
               Left            =   720
               Picture         =   "Form_Principal.frx":66F7D
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   244
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Mais_Informacoes 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mais informações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   10
                  Left            =   360
                  TabIndex        =   245
                  Top             =   90
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Botao_Remover_Transferencia 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   10
               Left            =   3480
               Picture         =   "Form_Principal.frx":67512
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   144
               TabIndex        =   242
               TabStop         =   0   'False
               Top             =   615
               Width           =   2160
               Begin VB.Label Label_Remover_Transferencia 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remover"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   10
                  Left            =   210
                  TabIndex        =   243
                  Top             =   90
                  Width           =   1740
               End
            End
            Begin NPlayer.NProgressBar Progresso 
               Height          =   375
               Index           =   10
               Left            =   9000
               TabIndex        =   450
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
            End
            Begin VB.Label Label_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   5760
               TabIndex        =   258
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Id 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   0
               TabIndex        =   257
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   10320
               TabIndex        =   256
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Tela_Programa 
               Height          =   255
               Index           =   10
               Left            =   9600
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Image Logotipo_Programa 
               Height          =   255
               Index           =   10
               Left            =   8760
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Tela 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   7200
               TabIndex        =   255
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Logotipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   5760
               TabIndex        =   254
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Icon 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   4320
               TabIndex        =   253
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Observacoes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   2880
               TabIndex        =   252
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Icon_Programa 
               Enabled         =   0   'False
               Height          =   375
               Index           =   10
               Left            =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label_Nome 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "VbMovieManager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   10
               Left            =   720
               TabIndex        =   251
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label_Descricao 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Programa para gerenciar os filmes do seu computador."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   10
               Left            =   720
               TabIndex        =   250
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label Label_Downloads 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   1440
               TabIndex        =   249
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   0
               TabIndex        =   248
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.Label Label_Nenum_Resultado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nenhum programa foi encontrado com esta categoria."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7500
            TabIndex        =   439
            Top             =   0
            Visible         =   0   'False
            Width           =   4650
         End
      End
      Begin VB.PictureBox Frame_Programas_Home 
         Appearance      =   0  'Flat
         BackColor       =   &H00CBB534&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   0
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   697
         TabIndex        =   236
         TabStop         =   0   'False
         Top             =   360
         Width           =   10455
         Begin VB.Label Label_Titulo_Frame_Programas 
            Alignment       =   2  'Center
            BackColor       =   &H00CBB534&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   239
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label Label_Titulo_Frame_Programas 
            AutoSize        =   -1  'True
            BackColor       =   &H00CBB534&
            BackStyle       =   0  'Transparent
            Caption         =   "SELECIONE A CATEGORIA DO PROGRAMA QUE PRETENDE PESQUISAR"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   238
            Top             =   600
            Width           =   6480
         End
         Begin VB.Label Label_Titulo_Frame_Programas 
            AutoSize        =   -1  'True
            BackColor       =   &H00C4AD2F&
            BackStyle       =   0  'Transparent
            Caption         =   "APP LIBRARY"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   570
            Index           =   0
            Left            =   0
            TabIndex        =   237
            Top             =   0
            Width           =   3630
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   3
            Left            =   6360
            Picture         =   "Form_Principal.frx":67AA7
            Top             =   840
            Width           =   1920
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   4
            Left            =   8520
            Picture         =   "Form_Principal.frx":71569
            Top             =   840
            Width           =   1920
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   0
            Left            =   0
            Picture         =   "Form_Principal.frx":7B02B
            Top             =   870
            Width           =   1920
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   1
            Left            =   2160
            Picture         =   "Form_Principal.frx":84AED
            Top             =   840
            Width           =   1920
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   2
            Left            =   4200
            Picture         =   "Form_Principal.frx":8E5AF
            Top             =   840
            Width           =   1920
         End
      End
      Begin VB.Label Label_Frame_Programas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do programa"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   235
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label_Frame_Programas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   234
         Top             =   120
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Image Separador_Frame_Programas 
         Height          =   375
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label_Frame_Programas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instalados"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   233
         Top             =   120
         Width           =   885
      End
      Begin VB.Image Separador_Frame_Programas 
         Height          =   330
         Index           =   1
         Left            =   360
         Top             =   0
         Width           =   1395
      End
      Begin VB.Image Separador_Frame_Programas 
         Height          =   330
         Index           =   3
         Left            =   3480
         Top             =   0
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Image Separador_Frame_Programas 
         Height          =   330
         Index           =   2
         Left            =   1920
         Top             =   0
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Image Barra_Top_Frame_Programas 
         Enabled         =   0   'False
         Height          =   330
         Left            =   0
         Picture         =   "Form_Principal.frx":98071
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5490
      End
   End
   Begin VB.PictureBox Frame_Perfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   15720
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   204
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   8295
      Begin VB.Line Linha_Frame_Perfil 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   184
         X2              =   528
         Y1              =   56
         Y2              =   56
      End
      Begin VB.Label Label_Nickname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Utilizador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   2760
         TabIndex        =   217
         Top             =   360
         Width           =   1680
      End
      Begin VB.Image Imagem_Foto 
         Enabled         =   0   'False
         Height          =   1920
         Left            =   240
         Picture         =   "Form_Principal.frx":98CBB
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label_Frame_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   214
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label_Frame_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   213
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label_Frame_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Género:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   212
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label Label_Frame_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data de nascimento:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   211
         Top             =   3240
         Width           =   1995
      End
      Begin VB.Label Label_Frame_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "País:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   210
         Top             =   3960
         Width           =   465
      End
      Begin VB.Label Label_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   209
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   208
         Top             =   4200
         Width           =   480
      End
      Begin VB.Label Label_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   207
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label Label_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   206
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label Label_Perfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   205
         Top             =   2040
         Width           =   480
      End
      Begin VB.Line Linha_Frame_Perfil 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         Visible         =   0   'False
         X1              =   16
         X2              =   136
         Y1              =   168
         Y2              =   168
      End
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagens"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   14760
      TabIndex        =   216
      Top             =   8040
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver perfil"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   16200
      TabIndex        =   215
      Top             =   8040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amigos"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   17640
      TabIndex        =   201
      Top             =   7680
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comunidade"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   16200
      TabIndex        =   196
      Top             =   7680
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   14760
      TabIndex        =   195
      Top             =   7680
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Favoritos"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   16200
      TabIndex        =   188
      Top             =   7200
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minha música"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   17640
      TabIndex        =   187
      Top             =   7200
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recentes"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   14760
      TabIndex        =   186
      Top             =   7200
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compartilhados"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   17640
      TabIndex        =   150
      Top             =   6720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ficheiros"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   16200
      TabIndex        =   149
      Top             =   6720
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agenda"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   14880
      TabIndex        =   148
      Top             =   6720
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label_Barra_Drive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contactos"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   17040
      TabIndex        =   147
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   6
      Left            =   17520
      Top             =   6600
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   4
      Left            =   14640
      Top             =   6600
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   5
      Left            =   16080
      Top             =   6600
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   3
      Left            =   16920
      Top             =   6120
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   2
      Left            =   16320
      Top             =   6120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   1
      Left            =   15840
      Top             =   6120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   0
      Left            =   15360
      Top             =   6120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Linha_Vertical 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   14160
      TabIndex        =   104
      Top             =   480
      Width           =   15
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00101010&
      Height          =   375
      Left            =   13680
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   8
      Left            =   16080
      Top             =   7080
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   9
      Left            =   17520
      Top             =   7080
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   7
      Left            =   14640
      Top             =   7080
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   11
      Left            =   16080
      Top             =   7560
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   10
      Left            =   14640
      Top             =   7560
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   12
      Left            =   17520
      Top             =   7560
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   14
      Left            =   16080
      Top             =   7920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Botao_Barra_Drive 
      Height          =   330
      Index           =   13
      Left            =   14640
      Top             =   7920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Barra_Drive 
      Enabled         =   0   'False
      Height          =   330
      Left            =   14640
      Picture         =   "Form_Principal.frx":99ECF
      Top             =   6120
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "Form_Principal"
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

'Option Explicit
'Declaração de variáveis
'Variáveis da scroll frame radio
Dim Radio_tx As Integer, Radio_Ty As Integer, Radio_DN As Boolean
Dim Radio_Txa As Integer, Radio_DNa As Boolean
Dim Radio_Tyb, Radio_DNb As Boolean
Dim Radio_NewY As Integer

'VARIÁVERIS DO SLIDER VIDEO
Dim tx As Integer, Ty As Integer, DN As Boolean
Dim Txa As Integer, DNa As Boolean
Dim Tyb, DNb As Boolean
Dim NewLeft As Integer

'VARIÁVERIS DO SLIDER VIDEO MASCARA
Dim Tx_2 As Integer, Ty_2 As Integer, DN_2 As Boolean
Dim Txa_2 As Integer, DNa_2 As Boolean
Dim Tyb_2, DNb_2 As Boolean
Dim NewLeft_2 As Integer

'VARIÁVERIS DO SLIDER SOM
Dim TX_Som As Integer, Ty_Som As Integer, DN_Som As Boolean
Dim Txa_Som As Integer, DNa_Som As Boolean
Dim Tyb_Som, Dnb_Som As Boolean
Dim NewLeft_Som As Integer

'Faixa em reproducao
Public Faixa_em_Reproducao As String

'Com/ Sem som
Public Mudo As Boolean

'Variavel para ver a duracao do ficheiro a reproduzir
Public VideoDuration As Double

'play ou pause
Public Musica_Play As Boolean

'tray icon
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_FINDSTRING = &H18F

'PROGRESS BAR EDITADO
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_USER = &H400
Const CCM_FIRST = &H2000&
Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Const PBM_SETBARCOLOR = (WM_USER + 9)

'Variável para verificar se está em modo hide (stray icon)
Public Modo_Tray As Boolean

'Ajusta o Form para sempre exibir a barra de tarefas do windows, full screen
Private Const SPI_GETWORKAREA = 48
Private Type RECT
  left As Long
  top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Variavel para verificar a janela do formulário
Dim Tela_Cheia As Boolean

'Variáveis para a base de dados
Public Cnn_Biblioteca As New ADODB.Connection
Public Rs_Musica As New ADODB.Recordset
Public Rs_Filmes As New ADODB.Recordset

'Variável para efectuar a pesquisa na base de dados
Dim Criterio As String

'Indicar qual a grelha activa, visivel
Public Grelha_Reproduzida As MSFlexGrid
Public Grelha_Visivel As MSFlexGrid

'Variáveis para visualizar as capas dos albums
Private MP3Path    As String
Private TPic       As IPictureDisp
Private CurIndex   As Long
Private MaxIndex   As Long

'Variável para verificar se o player esta a tocar de momento
Public Tocar As Boolean

'Variável para verificar se algum menu esta activo
Dim Menu_Activo As Boolean

'Cahamar a classe para ler as tags dos ficheiros
Dim cFile As New Classe_Tag_Editor

'Variável para criar um novo common dialog
Dim Explorador As New Class_Dialog

'Variáveis das linhas das grelhas personalizadas
Public Musica_Linha_Selecionada As String
Public Musica_Linha_Pressionada As String

'Variáveis para o idioma
Dim Idioma_Mudo_On As String
Dim Idioma_Mudo_Off As String
Dim Idioma_Desenvolvido As String
Dim Idioma_Nova_Versao   As String
Dim Idioma_Nao_Existe_Actualizacoes As String
Dim Idioma_Topico_Musica As String
Dim Idioma_Topico_Filmes As String
Dim Idioma_Topico_Loja As String
Dim Idioma_Topico_Radio As String
Dim Idioma_Total_Musicas As String
Dim Idioma_Total_Contactos As String
Dim Idioma_Total_Messagens As String
Dim Idioma_Total_Utilizadores As String
Dim Idioma_Total_Amigos As String
Dim Idioma_Total_Filmes As String
Dim Idioma_Total_Estacoes_Radio As String
Dim Idioma_Total_Ficheiros_Online As String
Dim Idioma_Total_Eventos As String
Public Idioma_Grid_Music_Col_1 As String
Public Idioma_Grid_Music_Col_2 As String
Public Idioma_Grid_Music_Col_3 As String
Public Idioma_Grid_Music_Col_4 As String
Public Idioma_Grid_Music_Col_5 As String
Public Idioma_Grid_Music_Col_6 As String
Public Idioma_Grid_Music_Col_7 As String
Public Idioma_Grid_Music_Col_8 As String
Public Idioma_Grid_Movies_Col_1 As String
Public Idioma_Grid_Movies_Col_2 As String
Public Idioma_Grid_Movies_Col_3 As String
Public Idioma_Grid_Movies_Col_4 As String
Public Idioma_Grid_Movies_Col_5 As String
Public Idioma_Grid_Movies_Col_6 As String
Public Idioma_Grid_Playlist_Col_1 As String
Public Idioma_Grid_Radio_Col_1 As String
Public Idioma_Grid_Loja_Col_1 As String
Public Idioma_Grid_Loja_Col_2 As String
Public Idioma_Grid_Loja_Col_3 As String
Public Idioma_Grid_Loja_Col_4 As String
Dim Idioma_Topico_Procurar  As String
Dim Idioma_Topico_Minha_Musica  As String
Dim Idioma_Topico_Resultado_Pesquisa  As String
Dim Idioma_Label_Topico_Barra_Lateral(2) As String
Dim Idioma_Label_Topico_Drive As String
Dim Idioma_Ver_Capa As String
Dim Idioma_Ocultar_Capa As String
Dim Idioma_Ver_Lista As String
Dim Idioma_Ocultar_Lista As String
Dim Idioma_Conectando As String
Dim Idioma_Reproduzindo As String
Dim Idioma_Pesquisa_Musica As String
Dim Idioma_Erro As String
Dim Idioma_Descricao As String
Dim Idioma_Erro_Execucao As String
Dim Idioma_Conectar_Servidor As String
Dim Idioma_Internet_Desligada As String
Dim Idioma_Mensagem_Enviada As String
Dim Idioma_Button_Fullscreen_On As String
Dim Idioma_Button_Fullscreen_Off As String
Dim Idioma_Label_Result_Search As String
Dim Idioma_Label_File_Found_0 As String
Dim Idioma_Label_File_Found_1 As String
Dim Idioma_Label_File_Found_2 As String
Dim Idioma_Name_Of_New_Playlist As String
Dim Idioma_Grid_Community_Col_1 As String
Dim Idioma_Grid_Community_Col_2 As String
Dim Idioma_Grid_Community_Col_3 As String
Dim Idioma_Grid_Community_Col_4 As String
Dim Idioma_Grid_Community_Col_5 As String
Dim Idioma_Grid_Community_Col_6 As String
Dim Idioma_Grid_Community_Col_7 As String
Dim Idioma_Grid_Community_Col_8 As String
Dim Idioma_Button_Transfer_Program As String
Dim Idioma_Button_Execute_Program As String
Dim Idioma_Button_Remove_Program As String
Dim Idioma_Button_Cancel_Program As String
Dim Idioma_Label_Rate As String

'Indicar qual foi o separador clicado antes de fazer login
Public Separador_Clicado As String

'Variáveis do idioma
Dim Idioma_Janela_Oculta  As String

'Variável para indicar qual a linha que está selecionada da frame listas
Dim Linha_Selecionada As Integer

'Actualizações
Dim Existe_Nova_Versao As Boolean

'Variável para indicar qual a linha que está selecionada no menu corrspondente
Dim Linha_Selecionada_Ficheiro As Integer
Dim Linha_Selecionada_Editar As Integer
Dim Linha_Selecionada_Ver As Integer
Dim Linha_Selecionada_Controlos As Integer
Dim Linha_Selecionada_Ferramentas As Integer
Dim Linha_Selecionada_Ajuda As Integer

'Variável para verificar qual é o album selecionado
Dim Index_Album As Integer

'APi's e variáveis para fazer scroll nas grids
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim X As Integer
Dim topRow As Integer
Dim ctl As Control
Dim lngResult As Long
Dim mygrid As Object

'Api's e variáveis para ver a capa online
Dim AID As Long
Dim albumID As Integer
Dim Artist$
Dim Author$
Dim Label$
Dim Title$
Dim adType$
Dim oldTitle$
Dim album$
Dim SmallCover$
Dim MedCover$
Dim LargeCover$
Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
Dim Capt$
Dim sCapt$
Dim LeftChar
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Const LB_ITEMFROMPOINT = &H1A9
Dim i As Integer
Dim eTitle$
Dim EMess$
Dim mError As Long

'Variável para saber qual a posição do player
Dim Posicao_do_Player As Integer

'Variável para saber se a barra mini player está activa ou não
Dim mini_player_activo As Boolean
Dim tela_video_fullscreen As Boolean

'Variável para preencher o espaço das labels dos topicos
Dim Espaco As String

'Variável para identificar qual é a visão da biblioteca
Dim visao_actual_da_biblioteca As String

'Variáveis de reprodução das músicas
Dim musica_aleatoria As Boolean
Dim musica_recomecar As Boolean

'Variáveis para mover a barra album
Dim nova_posicao As Long
Dim direcao_do_movimento As String
Dim movimento As String
Dim posicao_intermedia As Integer
Dim album_activo As Integer

'VAriáveis para a scrollBar dos albuns
Dim tx_album As Integer, Ty_album As Integer, DN_album As Boolean
Dim Txa_album As Integer, DNa_album As Boolean
Dim Tyb_album, DNb_album As Boolean
Dim NewLeft_album As Integer

'Class para verificar se exste alguma instância aberta do programa para carregar o parâmetro command$
Private WithEvents f_cPI As Class_Instancias
Attribute f_cPI.VB_VarHelpID = -1

'Redimensionar formulário
'Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const WM_NCLBUTTONDOWN = &HA1
Const HTBOTTOMRIGHT = 17
Const HTCAPTION = 2

'Variável para guardar as dimensões do formulário
Dim ALtura_Formulario As Long
Dim Largura_Formulario As Long

'Variável para saber qual é o nome da lista que está a ser editada
Dim index_lista_editada As Integer

'Variável para saber qual a lista que foi selecioanda
Dim index_lista_selecionada As Integer

'Variável para indicar qual será o nome da nova lista
Dim nome_nova_lista As String
Dim numero_lista As Integer

'Variáveis para saber qual é a coluna ordenada e qual a sua ordem
Private m_SortColumn As Integer
Private m_SortOrder As SortSettings

'Variável para saber qual é o servico que foi selecionado da barra lateral
Dim Servico_Activo As String

'Cor utilizada pelo programa
Const Azul = &HCBB534

'Variável para identificar que categoria do programa deve ser procurada
Dim Categoria_a_ser_Pesquisada As String

'Variável para identificar qual foi a linha selecionada da lista de programas
Dim Linha_Programa_Selecionado As Integer

'API para abrir web
Private Const SW_NORMAL = 1
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Api's para actualizar a date e hora dos programas após o seu download
Private Type FILETIME
    LowDateTime As Long
    HighDateTime As Long
End Type
Private Type SYSTEMTIME
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

'Variável para saber qual é o progressbar activo da lista de programas
Dim progress_activo As Integer

Function FileSetDate(ByVal sFileName As String, ByVal dFileDate As Date, Optional bSetCreationTime As Boolean = False, Optional bSetLastAccessedTime As Boolean = False, Optional bSetLastModified As Boolean = False) As Boolean
    'Função para actualizar a data e hora dos programas após o seu download
    Const GENERIC_WRITE = &H40000000, OPEN_EXISTING = 3
    Const FILE_SHARE_READ = &H1, FILE_SHARE_WRITE = &H2
    
    Dim lhwndFile As Long
    Dim tSystemTime As SYSTEMTIME
    Dim tLocalTime As FILETIME, tFileTime As FILETIME
    
    tSystemTime.Year = Year(dFileDate)
    tSystemTime.Month = Month(dFileDate)
    tSystemTime.Day = Day(dFileDate)
    tSystemTime.DayOfWeek = Weekday(dFileDate) - 1
    tSystemTime.Hour = Hour(dFileDate)
    tSystemTime.Second = Second(dFileDate)
    tSystemTime.Milliseconds = 0

    'Open the file to get the filehandle
    lhwndFile = CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If lhwndFile Then
        'File opened
        'Convert system time to local time
        SystemTimeToFileTime tSystemTime, tLocalTime
        'Convert local time to GMT
        LocalFileTimeToFileTime tLocalTime, tFileTime
'-------Change date/time property of the file
        FileSetDate = True
        If bSetCreationTime Then
            FileSetDate = FileSetDate And CBool(SetFileTime(lhwndFile, tFileTime, 0&, 0&))
        End If
        If bSetLastAccessedTime Then
            FileSetDate = FileSetDate And CBool(SetFileTime(lhwndFile, 0&, tFileTime, 0&))
        End If
        If bSetLastModified Then
            FileSetDate = FileSetDate And CBool(SetFileTime(lhwndFile, 0&, 0&, tFileTime))
        End If
        'Close the file handle
        Call CloseHandle(lhwndFile)
    End If
End Function

Public Sub ScrollUp()
    'Scroll up..
    If topRow > 1 Then
        topRow = topRow - 1
        mygrid.topRow = topRow
    End If
End Sub

Public Sub ScrollDown()
    'Scroll down..
    If topRow < mygrid.Rows - 1 Then
        topRow = topRow + 1
        mygrid.topRow = topRow
    End If
End Sub

Public Function PosFormRelativeTaskBar(F As Form)
    'Função para ao maximizar o form seja visivel a barra do windows iniciar
    'Colocar o WindowsState=0 normal
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    SetWindowPos hwnd, 0, WindowRect.left, WindowRect.top, WindowRect.Right - WindowRect.left, WindowRect.Bottom - WindowRect.top, 0
    F.top = WindowRect.Bottom * Screen.TwipsPerPixelY - F.Height
    F.left = WindowRect.Right * Screen.TwipsPerPixelX - F.Width
End Function

Public Sub Verifica_Rs_Musica()
    'VERIFICAR CONEXÃO A BASE DE DADOS
   If Rs_Musica.State = 1 Then Rs_Musica.Close
End Sub

Public Sub Verifica_Rs_Filmes()
    'VERIFICAR CONEXÃO A BASE DE DADOS
   If Rs_Filmes.State = 1 Then Rs_Filmes.Close
End Sub

Private Sub Barra_Actualizar_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Barra_Botoes_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Barra_Botoes_Lista_em_Reproducao_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Barra_Botoes_DblClick()
    'Atalho para
    Label_Titulo_DblClick
End Sub

Private Sub Barra_Botoes_Musica_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Repor_Objectos
End Sub

Private Sub Barra_Faixa_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Barra_Informacoes_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Barra_Informacoes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ocultar a barra do player/video
    Barra_Mini_Player.Visible = False
    
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Barra_Lateral_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Barra_Lateral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Barra_Mini_Player_Click()
    Ocultar_menus
End Sub

Private Sub Barra_Player_DblClick()
    'Atalho para
    Label_Titulo_DblClick
End Sub

Private Sub Barra_Player_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Barra_Mini_Player.Visible = False
End Sub

Private Sub Barra_Playlist_Click()
    Ocultar_menus
End Sub

Private Sub Barra_Slider_Album_Center_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Barra_Slider_Album_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Botao_Actualizar_Programa_Click()
    'Atalho para
    Label_Actualizar_Programa_Click
End Sub

Private Sub Botao_Click(Index As Integer)
    'Atalho das labels
    Label_Botao_Click Index
End Sub

Private Sub Botao_Antes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Antes.Picture = Form_Skin.Botao_Antes_Down.Picture
End Sub

Private Sub Botao_Antes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Antes.Picture = Form_Skin.Botao_Antes_Normal.Picture
End Sub

Private Sub Botao_Barra_Drive_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    Ocultar_menus
    
    Select Case Botao_Barra_Drive(Index).Index
        Case 0 'Ver antes
            Botao_Barra_Drive(0).Picture = Form_Skin.Icon_Seta_Anterior_Down.Picture

        Case 1 'Ver seguinte
            Botao_Barra_Drive(1).Picture = Form_Skin.Icon_Seta_Seguinte_Down.Picture
            
        Case 2 'Home
            Botao_Barra_Drive(2).Picture = Form_Skin.Icon_Home_Down.Picture
            
        Case 3 'Contacto
            Botao_Barra_Drive(3).Picture = Form_Skin.Botao_Barra_Down.Picture
        
        Case 4 'Agenda
            Botao_Barra_Drive(4).Picture = Form_Skin.Botao_Barra_Down.Picture
        
        Case 5 'Mensagens
            Botao_Barra_Drive(5).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 6 'Ficheiros
            Botao_Barra_Drive(6).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 7 'Recomendo
            Botao_Barra_Drive(7).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 8 'Favoritos
            Botao_Barra_Drive(8).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 9 'A minha música
            Botao_Barra_Drive(9).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 10 'Resultado da pesquisa
            Botao_Barra_Drive(10).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 11 'Comunidade
            Botao_Barra_Drive(11).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 12 'Os meus amigos
            Botao_Barra_Drive(12).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 13 'Mensagens
            Botao_Barra_Drive(13).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 14 'Ver perfil
            Botao_Barra_Drive(14).Picture = Form_Skin.Botao_Barra_Down.Picture
    End Select
End Sub

Private Sub Botao_Barra_Drive_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Repor as imagens originais após animar o botão
    Ocultar_menus
    
    Select Case Botao_Barra_Drive(Index).Index
        Case 0 'Ver antes
            Botao_Barra_Drive(0).Picture = Form_Skin.Icon_Seta_Anterior_Normal.Picture

        Case 1 'Ver seguinte
            Botao_Barra_Drive(1).Picture = Form_Skin.Icon_Seta_Seguinte_Normal.Picture
            
        Case 2 'Home
            Botao_Barra_Drive(2).Picture = Form_Skin.Icon_Home_Normal.Picture
                
        Case 3 'Contacto
            Botao_Barra_Drive(3).Picture = Form_Skin.Botao_Barra_Normal.Picture
        
        Case 4 'Agenda
            Botao_Barra_Drive(4).Picture = Form_Skin.Botao_Barra_Normal.Picture
        
        Case 5 'Mensagens
            Botao_Barra_Drive(5).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 6 'Ficheiros
            Botao_Barra_Drive(6).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 7 'Recomendo
            Botao_Barra_Drive(7).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 8 'Favoritos
            Botao_Barra_Drive(8).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 9 'A minha música
            Botao_Barra_Drive(9).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 10 'Resultado da pesquisa
            Botao_Barra_Drive(10).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 11 'Comunidade
            Botao_Barra_Drive(11).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 12 'Os meus amigos
            Botao_Barra_Drive(12).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 13 'Mensagens
            Botao_Barra_Drive(13).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 14 'Ver perfil
            Botao_Barra_Drive(14).Picture = Form_Skin.Botao_Barra_Normal.Picture
    End Select
End Sub

Private Sub Botao_Executar_Programa_Click(Index As Integer)
    'Atalho para
    Label_Executar_Programa_Click Index
End Sub

Private Sub Botao_Frame_Informacoes_Click(Index As Integer)
    'Efectuar operações
    Ocultar_menus
    Select Case Botao_Frame_Informacoes(Index).Index
        Case 0 'Transferir
            Label_Botao_Frame_Informacoes_Click (0)
            
        Case 1 'Cancelar
            Label_Botao_Frame_Informacoes_Click (1)
            
        Case 2 'Executar
            Label_Botao_Frame_Informacoes_Click (2)
    End Select
End Sub

Private Sub Botao_Legendas_Click()
    'Legendas on-line
    Procurar_Legendas
End Sub

Private Sub Botao_Legendas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mostar a imagem down
    Botao_Legendas.Picture = Form_Skin.Button_Menu_Down.Picture
    Icon_Legendas.Picture = Form_Skin.Icon_Subtitles_Down.Picture
End Sub

Private Sub Botao_Legendas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Legendas on-line
    Botao_Legendas.Picture = Form_Skin.Button_Menu_Normal.Picture
    Icon_Legendas.Picture = Form_Skin.Icon_Subtitles_Normal.Picture
End Sub

Private Sub Botao_Mais_Informacoes_Click(Index As Integer)
    'Atalho para
    Label_Mais_Informacoes_Click Index
End Sub

Private Sub Botao_Mensagens_Click()
    'Atalho para ver as mensagens
    Label_Mensagens_Click
End Sub

Private Sub Botao_Mensagens_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mostar a imagem down
    Botao_Mensagens.Picture = Form_Skin.Button_Menu_Standard_Down.Picture
    Icon_Mensagens.Picture = Form_Skin.Icon_Mensagem_Down.Picture
End Sub

Private Sub Botao_Mensagens_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Legendas on-line
    Botao_Mensagens.Picture = Form_Skin.Button_Menu_Standard_Normal.Picture
    Icon_Mensagens.Picture = Form_Skin.Icon_Mensagem_Normal.Picture
End Sub

Private Sub Botao_Mudo_Mini_Click()
    'Colocar o media player como mudo ou ouvir
    Botao_Mudo_Click
End Sub

Private Sub Botao_Pausa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Pausa.Picture = Form_Skin.Botao_Pausa_Down.Picture
End Sub

Private Sub Botao_Pausa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Pausa.Picture = Form_Skin.Botao_Pausa_Normal.Picture
End Sub

Private Sub Botao_Pesquisar_Click()
    'Efectuar pesquisa na base de dados consuante os dados introduzidos
    'On Error GoTo Corrige_Erro
    If Len(Trim(Text_Pesquisar.Text)) = 0 Then Exit Sub
    If Text_Pesquisar.Text = Idioma_Pesquisa_Musica Then Exit Sub
    
    Me.MousePointer = 11
    
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "pesquisarmusica.asp?Recebe_Pesquisa=" & Text_Pesquisar.Text, False
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If servidor.responseText = "false" Then
        Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Grelha_Loja.Clear
            Formatar_Grelha Grelha_Loja
            Carregar_Loja_Online servidor.responseText
            Me.MousePointer = 0

            Text_Pesquisar.Text = Idioma_Pesquisa_Musica
            Label_Barra_Drive_Click (10)
        End If
    End If
    Set servidor = Nothing
    
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

Private Sub Carregar_Loja_Online(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/pesquisa/resultado")
        Dim i As Integer: i = Grelha_Loja.Rows
        
        For Each node In nodeList
            DoEvents
            Grelha_Loja.Rows = Grelha_Loja.Rows + 1

            If Not IsEmpty(node.selectSingleNode("servidor")) Then Grelha_Loja.TextMatrix(i, 0) = node.selectSingleNode("servidor").Text
            If Not IsEmpty(node.selectSingleNode("titulo")) Then Grelha_Loja.TextMatrix(i, 1) = node.selectSingleNode("titulo").Text
            If Not IsEmpty(node.selectSingleNode("artista")) Then Grelha_Loja.TextMatrix(i, 2) = node.selectSingleNode("artista").Text
            If Not IsEmpty(node.selectSingleNode("data")) Then Grelha_Loja.TextMatrix(i, 3) = node.selectSingleNode("data").Text
            If Not IsEmpty(node.selectSingleNode("adicionado")) Then Grelha_Loja.TextMatrix(i, 4) = node.selectSingleNode("adicionado").Text
            If Not IsEmpty(node.selectSingleNode("id")) Then Grelha_Loja.TextMatrix(i, 5) = node.selectSingleNode("id").Text
            i = i + 1
        Next
                
    Else
        'Caso nenhum encontre num ficheiro referente á pesquisa efectuada
        Grelha_Loja.Rows = 1
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Private Sub Barra_Player_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Public Sub Botao_Fechar_Click()
    'Fechar a conexão á base de dados
    Me.Hide
    Rs_Musica.Close
    Rs_Filmes.Close

'    Form_Preferencias.Salvar_Valores
    
    'Salvar a grelha lista em reprodução
    If Form_Preferencias.Check_Guardar_Lista.Value = 1 Then
        If Grelha_Lista_Em_Reproducao.Rows > 1 Then
            Personalizar_Grid Grelha_Lista_Em_Reproducao
            Dim cFlexSettings As clsFlexSettings
            Set cFlexSettings = New clsFlexSettings
            Set cFlexSettings.FlexGrid = Grelha_Lista_Em_Reproducao
            cFlexSettings.SaveSettings App.Path & "\Library\Standard.ini", True, True, True, True
            Set clsFlexSettings = Nothing
        End If
    End If
    
    'Cancelar media player
    Wmp.Controls.stop: Form_Wmp.Wmp.Controls.stop
    Timer_Slider_Video.Enabled = False
    Botao_Play.Visible = True: Form_Mini_Player.Botao_Play.Visible = True: Form_PopUp.Botao_Play.Visible = True: Botao_Player_Mini(1).Visible = True
    Botao_Pausa.Visible = False: Form_Mini_Player.Botao_Pausa.Visible = False: Form_PopUp.Botao_Pausa.Visible = False: Botao_Player_Mini(2).Visible = False

    Unload Form_Wmp
    Unload Form_Lista
    Unload Form_PopUp
    
    'Remover do sistema o icon do programa
    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t  'Remove o ícone da barra de tarefas.
    
    'Actualizar as opções do programa
    Call WriteINI("Settings", "FullScreen", Form_Preferencias.Text_Tela_Cheia.Text, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Settings", "Sound", Text_Slide_Som.Text, (Localizacao_Ficheiro_Preferencias))
    If Grelha_Musica.Rows > 1 Then Call WriteINI("Library", "Playing_Track", Musica_Linha_Pressionada, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Library", "Visualization", Text_Visualizacao.Text, (Localizacao_Ficheiro_Preferencias))
    
    'Encerrar a class das instâncias do programa
    Set f_cPI = Nothing
    
    'Fechar o programa
    Unload Me
    End
End Sub

Private Sub Botao_Play_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Play.Picture = Form_Skin.Botao_Play_Down.Picture
End Sub

Private Sub Botao_Play_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Play.Picture = Form_Skin.Botao_Play_Normal.Picture
End Sub

Private Sub Botao_Player_Mini_Click(Index As Integer)
    'Botões do mini player
    Select Case Botao_Player_Mini(Index).Index
        Case 0
            Botao_Antes_Click
        Case 1
            Botao_Play_Click
        Case 2
            Botao_Pausa_Click
        Case 3
            Botao_Seguinte_Click
        Case 4
            Video_FullScreen
    End Select
End Sub

Private Sub Botao_Redimensionar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Redimensionar o formulário conforme as dimensões pretendidas
    If Button = vbLeftButton Then
        If Tela_Cheia = False Then
            ReleaseCapture
            SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
            
            'Verificar se não exedeu os limites
            If Me.Height < 8616 Then
                Me.Height = "8616"
            End If
        
            If Me.Width < 14385 Then
                Me.Width = "14385"
            End If
            ALtura_Formulario = Me.Height
            Largura_Formulario = Me.Width
            Desenhar_Formulario
        End If
    End If
End Sub

Private Sub Botao_Redimensionar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Alterar o mousepointer
    If Tela_Cheia = True Then
        Botao_Redimensionar.MousePointer = vbDefault
    Else
        Botao_Redimensionar.MousePointer = 8
    End If
End Sub

Private Sub Botao_Remover_Transferencia_Click(Index As Integer)
    'Atalho para
    Label_Remover_Transferencia_Click Index
End Sub

Private Sub Botao_Seguinte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Seguinte.Picture = Form_Skin.Botao_Seguinte_Down.Picture
End Sub

Private Sub Botao_Seguinte_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botao_Seguinte.Picture = Form_Skin.Botao_Seguinte_Normal.Picture
End Sub

Private Sub Botao_Tray_Click()
    'Mensagem no icon do projecto/ coloca-lo ao lado do clock
    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
    t.szTip = "NPlayer" & Chr$(10) 'Texto a ser exibido no icon
    Shell_NotifyIcon NIM_ADD, t
    App.TaskVisible = False
    
    'Colocar o icon do formulário ao lado do clock do windows
    Me.Hide
        
    Modo_Tray = True
    
    'Chamar o procedimento
    Mostrar_Faixa_Musica_Formulario_Popup
    Form_PopUp.Hide
End Sub

Private Sub ResizePic(nome_da_pic As PictureBox)
    'Ajustar a capa do album
    On Error Resume Next 'GoTo Corrige_Erro
    Dim nWidth  As Long
    Dim nHeight As Long

    nWidth = ScaleX(TPic.Width, vbHimetric, vbPixels)
    nHeight = ScaleY(TPic.Height, vbHimetric, vbPixels)
    With nome_da_pic
        If .ScaleWidth < (nWidth * (.ScaleHeight / nHeight)) Then
            nHeight = nHeight * (.ScaleWidth / nWidth)
            nWidth = .ScaleWidth
        Else
            nWidth = nWidth * (.ScaleHeight / nHeight)
            nHeight = .ScaleHeight
        End If
        TPic.Render .hdc, (.ScaleWidth - CLng(nWidth)) / 2, (.ScaleHeight - CLng(nHeight)) / 2, CLng(nWidth), CLng(nHeight), 0, TPic.Height, TPic.Width, -TPic.Height, ByVal 0&
    End With
    On Error GoTo 0
End Sub

Public Sub Apagar_Album_Musica()
    'Procedimento para apagar os objectos do album
    Lista_Pastas.Clear
    'Lista_Ficheiros.Clear
    Label_Album.Caption = ""
    Label_Directorio_Album(0).Caption = ""
    Label_Nome_Album(0).Text = ""
    
    Frame_Slide_Album.Visible = False
    Label_Nenhum_Album.Visible = True
    Image_Album(0).Visible = False
    
    If Image_Album.count > 1 Then
        Dim album_existente As Integer: For album_existente = 1 To Image_Album.count - 1
            Unload Image_Album(album_existente)
            Unload Label_Directorio_Album(album_existente)
            Unload Label_Nome_Album(album_existente)
        Next album_existente
    End If
End Sub

Public Sub Criar_Album_Musica()
    'Listar pastas existentes no caminho indicado
    DoEvents
    Lista_Pastas.Clear
    'Lista_Ficheiros.Clear
    ListSubDirs Text_Caminho.Text
    Frame_Slide_Album.Visible = False
    Label_Nenhum_Album.Visible = True
    
    If Lista_Pastas.ListCount <= 0 Then Exit Sub

    With Lista_Pastas
        If .ListCount > 1 Then
            If Text_Visualizacao.Text = "2" Then Frame_Slide_Album.Visible = True
            Frame_Slide.Width = Form_Skin.Image_Capa.Width + 10
            Image_Album(0).ToolTipText = ReadINI("Main", "Image_Album", Localizacao_Ficheiro_Lingua)
            Image_Album(0).Visible = True
            Image_Album(0).Picture = Form_Skin.Image_Album.Picture
            Label_Nenhum_Album.Visible = False
            Barra_Slider_Album.Visible = True
            Label_Directorio_Album(0).Caption = .List(0)
            Label_Nome_Album(0).Text = Dir(Label_Directorio_Album(0).Caption, vbDirectory)
            Label_Nome_Album(0).Visible = True
            Label_Nome_Album(0).Width = Image_Album(0).Width - 20
            Label_Nome_Album(0).left = Image_Album(0).left + 10
            Label_Nome_Album(0).top = Image_Album(0).top + Image_Album(0).Height - Label_Nome_Album(0).Height - 10
                            
            Dim novo_album As Integer
            For novo_album = 1 To .ListCount - 1
                'Criar janelas consoante o número de linhas da list pastas
                Load Image_Album(novo_album)
                Image_Album(novo_album).Visible = True
                Image_Album(novo_album).Move Image_Album(novo_album - 1).left + Image_Album(0).Width + 10, Image_Album(0).top
                Image_Album(novo_album).ToolTipText = ReadINI("Main", "Image_Album", Localizacao_Ficheiro_Lingua)
                
                Load Label_Directorio_Album(novo_album)
                Label_Directorio_Album(novo_album).Visible = False
                Label_Directorio_Album(novo_album).Width = Image_Album(0).Width - 20
                Label_Directorio_Album(novo_album).left = Image_Album(novo_album).left + 10
                Label_Directorio_Album(novo_album).Caption = .List(novo_album)
                
                Load Label_Nome_Album(novo_album)
                Label_Nome_Album(novo_album).Visible = True
                Label_Nome_Album(novo_album).Width = Image_Album(0).Width - 20
                Label_Nome_Album(novo_album).left = Image_Album(novo_album).left + 10
                Label_Nome_Album(novo_album).top = Label_Nome_Album(0).top
                Label_Nome_Album(novo_album).Text = Dir(Label_Directorio_Album(novo_album).Caption, vbDirectory)
                Label_Nome_Album(novo_album).ZOrder 0
            Next novo_album
            
            'Ajustar o tamanho da barra consoante o nº de albuns encontrados
            Frame_Slide.Width = Image_Album.count * (Image_Album(0).Width + 10)
            
            'Activa o 1º album
            album_activo = 0
            Image_Album(album_activo).Picture = Form_Skin.Image_Album_Over.Picture
            Label_Album.Caption = Label_Topico_Barra_Lateral(0).Caption 'Label_Nome_Album(album_activo).Text

        Else 'Caso não existem músicas na grelha de reprorução
            Label_Nenhum_Album.Visible = True
            Frame_Slide_Album.Visible = False
            Barra_Slider_Album.Visible = False
        End If
    End With
End Sub

Private Sub Procurar_Legendas()
    'Procedimento para procurar legendas online
    On Error GoTo Corrige_Erro
    If Grelha_Filmes.Rows <= 1 Then Exit Sub
    'Me.MousePointer = 11
    
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "pesquisarlegenda.asp?Recebe_Pesquisa=" & Grelha_Filmes.TextMatrix(Grelha_Filmes.Row, 1), False
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If servidor.responseText = "false" Then
        Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Form_Legendas.Grelha_Legendas.Rows = 1
            Carregar_Legendas servidor.responseText
            Me.MousePointer = 0
        End If
    End If
    Set servidor = Nothing
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
End Sub

Private Sub Carregar_Legendas(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    On Error GoTo Corrige_Erro
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    
    With Form_Legendas
        If xml.loadXML(responseText) Then
            Dim node As IXMLDOMNode
            Dim nodeList As IXMLDOMNodeList
            Set nodeList = xml.selectNodes("/pesquisa/resultado")
            Dim i As Integer: i = .Grelha_Legendas.Rows
            
            For Each node In nodeList
                DoEvents
                .Grelha_Legendas.Rows = .Grelha_Legendas.Rows + 1
                If Not IsEmpty(node.selectSingleNode("localizacao")) Then .Grelha_Legendas.TextMatrix(i, 0) = node.selectSingleNode("localizacao").Text
                If Not IsEmpty(node.selectSingleNode("ficheiro")) Then .Grelha_Legendas.TextMatrix(i, 1) = node.selectSingleNode("ficheiro").Text
                If Not IsEmpty(node.selectSingleNode("idioma")) Then .Grelha_Legendas.TextMatrix(i, 2) = node.selectSingleNode("idioma").Text
                If Not IsEmpty(node.selectSingleNode("formato")) Then .Grelha_Legendas.TextMatrix(i, 3) = node.selectSingleNode("formato").Text
                If Not IsEmpty(node.selectSingleNode("titulo")) Then .Grelha_Legendas.TextMatrix(i, 4) = node.selectSingleNode("titulo").Text
                i = i + 1
            Next
            .Label_Download.Enabled = True
            .Botao_Download.Enabled = True
            .Grelha_Legendas.HighLight = flexHighlightAlways
        Else
            'Caso nenhum encontre num ficheiro referente á pesquisa efectuada
            .Label_Download.Enabled = False
            .Botao_Download.Enabled = False
            .Grelha_Legendas.HighLight = flexHighlightNever
        End If
        .Show vbModal
        Set xml = Nothing
        Set nodeList = Nothing
    End With
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
End Sub

Private Sub Close_Barra_Actualizar_Click()
    'Ocultar a barra actualizar
    Barra_Actualizar.Visible = False
    Ajustar_Objectos_Na_Vertical
End Sub

Private Sub Close_Wmp_Click()
    'Fechar a tela de video
    On Error GoTo Corrige_Erro
    Ocultar_menus
    Frame_Wmp.Visible = False
    Close_Wmp.Visible = False
    Barra_Mini_Player.Visible = False
    Grelha_Reproduzida.Visible = True
    If Grelha_Reproduzida = Grelha_Filmes Then Botao_Pausa_Click
    
Exit Sub
Corrige_Erro:
Label_Topico_Musica_Click
End Sub

Private Sub Close_Wmp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ocultar a barra mini player
    Barra_Mini_Player.Visible = False
End Sub

Private Sub dl_DowloadComplete()
    'Transferência concluida
    GetFileName (Text_Servidor.Text)
    ProgressBar1.Value = 0
    GetFileName (Text_Servidor.Text)
    
    Botao_Frame_Informacoes(0).Visible = True
    ProgressBar1.Visible = False
    Me.MousePointer = 0
    Label_Frame_Informacoes(6).Caption = ReadINI("Main", "Label_Download_Complete", Localizacao_Ficheiro_Lingua)
    Botao_Frame_Informacoes(2).Enabled = True
    Label_Botao_Frame_Informacoes(2).Enabled = True
    Botao_Frame_Informacoes(1).Enabled = False
    Label_Botao_Frame_Informacoes(1).Enabled = False
    
    'Actualiza no servidor nº de downloads do programa
    Verificar_Downloads
    
    'Iniciar a decompactação do programa zipado
    DesCompacta App.Path & "\Programs\" & Label_Frame_Informacoes(3).Caption, "*.*", App.Path & "\Programs\", True
    Kill App.Path & "\Programs\" & Label_Frame_Informacoes(3).Caption
    
    'Ao terminar a transferência do ficheiro a Idioma_Button_Transfer_Program passa a ser Idioma_Button_Remove_Program
    Label_Botao_Frame_Informacoes(0).Caption = Idioma_Button_Remove_Program
    Label_Frame_Informacoes(3).Caption = ReadINI("Main", "Label_Installed_In", Localizacao_Ficheiro_Lingua) & ": " & Date & " " & Time
    
    'Actualizar a data e hora de criação do programa
    Dim Ficheiro_Para_Actualizar As String
    Ficheiro_Para_Actualizar = App.Path & "\Programs\" & Label_Frame_Informacoes(0).Caption & "\" & Label_Frame_Informacoes(0).Caption & ".exe"
    
    'Set the creation time
    FileSetDate Ficheiro_Para_Actualizar, Now, True
    'Set the last accessed time
    FileSetDate Ficheiro_Para_Actualizar, Now, , True
    'Set the last write time
    FileSetDate Ficheiro_Para_Actualizar, Now, , , True
    
    Image_Download.Picture = Form_Skin.Image_Down_Concluido.Picture
    Me.MousePointer = 0
End Sub

Private Sub dl_DownloadErrors(strError As String)
    'Caso ocorra um erro durante o download
    Label_Frame_Informacoes(6).Caption = ReadINI("Main", "Error_Transfer_Program", Localizacao_Ficheiro_Lingua)
    Label_Botao_Frame_Informacoes(0).Caption = Idioma_Button_Transfer_Program
    Botao_Frame_Informacoes(0).Visible = True
    ProgressBar1.Visible = False
    Image_Download.Picture = Form_Skin.Image_Down_Erro.Picture
    Me.MousePointer = 0
End Sub

Private Sub dl_DownloadProgress(intPercent As String)
    'Mostrar o progresso do download
    ProgressBar1.Value = intPercent
    GetFileName (Text_Servidor.Text)
    Text_Servidor.Text = ""
    Label_Frame_Informacoes(6).Caption = ReadINI("Main", "Label_Transferring_File", Localizacao_Ficheiro_Lingua)
    
    Image_Download.Picture = Form_Skin.Image_Down_Processando.Picture
End Sub

Private Sub Estrela_Click(Index As Integer)
    'Classificar o ficheiro selecionado
    Select Case Estrela(Index).Index
        Case 0
            Classificacao True, False, False, False, False
            Text_Classificacao.Text = "1"
        Case 1
            Classificacao True, True, False, False, False
            Text_Classificacao.Text = "2"
        Case 2
            Classificacao True, True, True, False, False
            Text_Classificacao.Text = "3"
        Case 3
            Classificacao True, True, True, True, False
            Text_Classificacao.Text = "4"
        Case 4
            Classificacao True, True, True, True, True
            Text_Classificacao.Text = "5"
    End Select
    
    Actualiza_Dados_da_Tabela
End Sub

Private Sub Form_Activate()
    Form_Wmp.Wmp.settings.mute = True
End Sub

Private Sub Form_Click()
    Ocultar_menus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then
        If Close_Wmp.Visible = True Then Close_Wmp_Click
    End If
End Sub

Private Sub Form_Load()
    'On Error GoTo Corrige_Erro
    'Verificar se já existe alguma instância do programa aberta, ao carregar os ficheiros através do command$
'    Set f_cPI = New Class_Instancias
'    If f_cPI.PrevInstance Then
'        Unload Me
'        End
'    Else
        
    '    'Permite que se possa fazer scroll nas grids
    '    lpFormObj = ObjPtr(Me)
    '    SetProp Me.hwnd, "PrevWndProc", SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WndProc)
    '    If GetSystemMetrics(SM_MOUSEWHEELPRESENT) Then
    '       ' MsgBox "A simple call to GetSystemMetrics tells you whether or not the mouse has a wheel. The mouse connected to this computer does have a wheel.", vbInformation + vbOKOnly, App.Title
    '       Debug.Print "Yes Wheel"
    '    Else
    '      ' MsgBox "A simple call to GetSystemMetrics tells you whether or not the mouse has a wheel. The mouse connected to this computer doesn't have a wheel.", vbInformation + vbOKOnly, App.Title
    '        Debug.Print "No Wheel"
    '    End If
    '    topRow = 1
    
        'Informações iniciais
        Dim dia_hoje As String: dia_hoje = Day(Now)
        If Len(dia_hoje) < 2 Then dia_hoje = "0" & dia_hoje: Label_Evento(2).Caption = dia_hoje
        
        Label_Contador.Caption = ""
        File_Ficheiros.Pattern = "*.mp3;*.wav;*.wma"
        mover_album = False
        Frame_Slide.left = 0
        nova_posicao = Frame_Slide.left + Form_Skin.Image_Album.Width + 10
        musica_aleatoria = False
        musica_recomecar = False
        Espaco = "         "
        Image_Barra_Slide.left = 0
        tela_video_fullscreen = False
        mini_player_activo = True
        Label_Album.Caption = ""
        Carregar_Skin
        Carregar_Idioma
        Desenhar_Formulario
        
        'Variáveis para poder mover o formulário
        iTPPX& = Screen.TwipsPerPixelX
        iTPPY& = Screen.TwipsPerPixelY
        
        'Guardar a localização do programa para depois se poder actualizar
        Call WriteINI("Path", "Location_Of_Program", App.Path & "\", (Localizacao_Ficheiro_Preferencias))
    
        Verificar_Opcoes_do_Programa
        Set Grelha_Reproduzida = Grelha_Musica
        Set Grelha_Visivel = Grelha_Musica
        Formatar_Grelha_Musica Grelha_Lista_Em_Reproducao
        Formatar_Grelha_Musica Form_Lista.Grelha_Lista_Em_Reproducao
        Formatar_Grelha_Contactos
        Formatar_Grelha Grelha_Loja
        Formatar_Grelha_Musica Grelha_Listas
        Formatar_Grelha Grelha_Minha_Musica
        Formatar_Grelha Grelha_Recentes
        Formatar_Grelha Grelha_Favoritos
        Formatar_Grelha_Comunidade
        Formatar_Grelha_Amigos
        Formatar_Grelha_Mensagens
        Formatar_Grelha_Eventos
        Formatar_Grelha_Ficheiros
        
        Conectar_a_Base_de_Dados
        Carregar_Grelhas
        
        'Verificar as definições do programa
        With Form_Preferencias
            If .Text_Tela_Cheia.Text <> Empty Then
                If .Text_Tela_Cheia.Text = "True" Then
                    Botao_Maximizar_Click
                Else
                    Botao_Restaurar_Click
                End If
            End If
        End With
        
        'Volume do player
        Text_Slide_Som.Text = ReadINI("Settings", "Sound", Localizacao_Ficheiro_Preferencias)
        Slide_Som.left = Text_Slide_Som.Text
        Slide_Som_Mini.left = Text_Slide_Som.Text
        Form_Mini_Player.Slide_Som.left = Text_Slide_Som.Text
        Form_PopUp.Slide_Som.left = Text_Slide_Som.Text
        Verificar_Volume
        
        'Valores iniciais das variáveis
        Mudo = False
        Musica_Play = False
        Modo_Tray = False
        VideoDuration = 0
        Posicao_do_Player = 0
        Verificar_Volume
        Menu_Activo = False
        Utilizador_Logado = False
        Existe_Nova_Versao = False
        
        'Carregar as estações de rádio existentes
        Carregar_Estacoes_Radio
        
        Label_Topico_Musica_Click
        If Grelha_Reproduzida.Rows > 1 Then
            'Definir qual é a linha a ser reproduzida
            Dim Ultima_Musica_Ouvida As String
            Ultima_Musica_Ouvida = ReadINI("Library", "Playing_Track", Localizacao_Ficheiro_Preferencias)
            If Ultima_Musica_Ouvida <> "" Then
                If Grelha_Musica.Rows >= Ultima_Musica_Ouvida + 1 Then '+1 corresponde ao cabeçalho da grelha
                    Musica_Linha_Pressionada = Ultima_Musica_Ouvida
                    Musica_Linha_Selecionada = Ultima_Musica_Ouvida
                Else
                    Musica_Linha_Pressionada = 1
                    Musica_Linha_Selecionada = 1
                End If
                
            Else
                Musica_Linha_Pressionada = 1
                Musica_Linha_Selecionada = 1
            End If
            
            Faixa_em_Reproducao = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 0)
            Wmp.URL = Faixa_em_Reproducao: Wmp.Controls.stop
            Form_Wmp.Wmp.URL = Faixa_em_Reproducao: Form_Wmp.Wmp.Controls.stop
            Label_Faixa.Caption = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 1)
            Form_Mini_Player.Label_Faixa.Caption = Label_Faixa.Caption
            Verificar_Classificacao
            
            'Seleciona a linha inteira da grelha
            With Grelha_Reproduzida
                .Row = Musica_Linha_Pressionada
                .Col = 0
                .ColSel = .Cols - 1
            End With
        End If
        
        'Selecionar a 1ªlinha da list das playlists existentes
        Linha_Selecionada = 0
        Carregar_Listas
        nome_nova_lista = Idioma_Name_Of_New_Playlist & ".ini"
        numero_lista = 0
        
        'Verificar se o programa já foi instalado
        Dim Programa_Instalado As String: Programa_Instalado = ReadINI("Settings", "Installed_Program", Localizacao_Ficheiro_Preferencias)
        If Programa_Instalado = "False" Then
            'Indicar que o programa já foi instalado
            Call WriteINI("Settings", "Installed_Program", "True", (Localizacao_Ficheiro_Preferencias))
        End If
        Label_Topico_Musica_Click
        If Form_Preferencias.Check_Actualizacoes.Value = 1 Then Verificar_Actualizacoes
        
        'Criar os menus do programa
        Carregar_Menu_Ficheiro
        Carregar_Menu_Editar
        Carregar_Menu_Ver
        Carregar_Menu_Controlos
        Carregar_Menu_Ferramentas
        Carregar_Menu_Ajuda
        Ajustar_Menus
        
        'Carregar albuns de música
        If importar_media = True Then
            If Text_Caminho.Text = Empty Then
                Text_Caminho.Text = ReadINI("Library", "Location_Of_Albuns", Localizacao_Ficheiro_Preferencias)
                Criar_Album_Musica
            End If
        End If
        
        'Formatar e carregar as listas de pesquisa personalizada
        Formatar_Grelha_Artista
        Formatar_Grelha_Genero
        Formatar_Grelha_Album
        Carregar_Grelha_Artista
        Carregar_Grelha_Genero
        Carregar_Grelha_Album
    
        'Verificar qual é a visão da biblioteca
        Text_Visualizacao = ReadINI("Library", "Visualization", Localizacao_Ficheiro_Preferencias)
        Icon_Visao_Click Text_Visualizacao.Text
        
        'Carregar o ficheiro recebido pelo parâmetro command$
'        f_cPI.Ready = True
'    End If
    
    'Carregar os valores da altura e largura do formulário quando fica redimensionado
    ALtura_Formulario = "8616"
    Largura_Formulario = "14385"

    'Alterar cores do progreesbar
    ProgressBar1.backcolor = RGB(235, 113, 64) 'laranja
    Dim xpto As Integer: For xpto = 0 To Pic_Linha.count - 1
        Progresso(xpto).backcolor = RGB(235, 113, 64)
    Next
    


Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case Else
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub f_cPI_PrevInstance(ByVal sCommand As String, ByVal bReady As Boolean, ByVal Files As Collection, ByVal Folders As Collection, ByVal Parameters As Collection)
    'Procedimento para trazer a janela do programa para a frente, ao carregar um novo ficheiro através do parâmetro command$
    f_cPI.ShowForm Me
    If bReady Then
        Dim vItem As Variant
        'Carregar ficheiros atrave´s do command e reproduzir automaticamente
        For Each vItem In Files 'Se for um ficheiro---------------------------------------------------------------------------------------
            Dim ficheiro As String: ficheiro = vItem 'Replace(Command, Chr(34), "") 'Chr(34) Corresponde ás " "
            If ficheiro <> "" Then
                Dim sFile As String, sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
                Dim nova_linha As Integer
                cFile.FileName = True
                cFile.FileName = ficheiro
                sTitle = Replace(cFile.Title, "'", " ", , , vbTextCompare)
                sArtist = Replace(cFile.Artist, "'", " ", , , vbTextCompare)
                sAlbum = Replace(cFile.album, "'", " ", , , vbTextCompare)
                sYear = Replace(cFile.Year, "'", " ", , , vbTextCompare)
                sGenre = Replace(cFile.Genre, "'", " ", , , vbTextCompare)
                sComment = Replace(cFile.Comments, "'", " ", , , vbTextCompare)
    
                'Caso o ficheiro não tenha tags então o titulo será o nome do ficheiro, o qual é obtido atrvés do directório
                Dim nome_ficheiro As String: nome_ficheiro = Dir(ficheiro, vbArchive)
    
                If sTitle = "" Then sTitle = Mid(nome_ficheiro, 1, InStrRev(nome_ficheiro, ".") - 1)
                If sArtist = "" Then sArtist = ""
                If sAlbum = "" Then sAlbum = ""
                If sYear = "" Then sYear = ""
                If sGenre = "" Then sGenre = ""
                If sComment = "" Then sComment = ""
    
                'Adicionar as músicas na playlist
                With Grelha_Lista_Em_Reproducao
                    nova_linha = .Rows
                    .Rows = .Rows + 1
                    .TextMatrix(nova_linha, 0) = ficheiro
                    .TextMatrix(nova_linha, 1) = sTitle
                    .TextMatrix(nova_linha, 2) = sArtist
                    .TextMatrix(nova_linha, 3) = sAlbum
                    .TextMatrix(nova_linha, 4) = sYear
                    .TextMatrix(nova_linha, 5) = sGenre
                    .TextMatrix(nova_linha, 6) = sComment
        '            .TextMatrix(nova_linha, 7) = Dir(ficheiro, vbDirectory)
                    .TextMatrix(nova_linha, 8) = "0"
                    .Row = nova_linha
    
                    'Definir que a grelha a ser reproduzida será a da playlist
                    Musica_Linha_Pressionada = nova_linha
                    Musica_Linha_Selecionada = nova_linha
                    Set Grelha_Reproduzida = Grelha_Lista_Em_Reproducao
                    Faixa_em_Reproducao = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 0)
                    Wmp.URL = Faixa_em_Reproducao: Wmp.Controls.stop
                    Form_Wmp.Wmp.URL = Faixa_em_Reproducao: Form_Wmp.Wmp.Controls.stop
                    Label_Faixa.Caption = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 1)
                    Form_Mini_Player.Label_Faixa.Caption = Label_Faixa.Caption
                    Reproduzir_Musica_da_Grelha
                    
                    'Activar o tópico música
                    Repor_a_Cor_Dos_Topicos
                    Shape_Topico(0).Visible = True
                    Label_Topico_Musica.ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
                    Icon_Topico(0).Picture = Form_Skin.Icon_Topico_Musica_Over.Picture
                    Ocultar_Objectos
                    Set Grelha_Visivel = Grelha_Musica
                    Grelha_Musica.Visible = True
                    Barra_Playlist.Visible = True
                    Dim star As Integer: For star = 0 To Estrela.count - 1
                        Estrela(star).Visible = True
                    Next star
                    If Grelha_Musica.Rows > 1 Then Text_Classificacao.Text = Grelha_Musica.TextMatrix(Grelha_Musica.Row, 8)
                    Verificar_Classificacao
                    Verificar_Contador
                    Label_Botao(0).Visible = True
                    Text_Pesquisar_Musica.Text = Empty
                    Label_Botao(4).Visible = True
                    Barra_Botoes_Musica.Visible = True
                    Barra_Lateral.Visible = True
                    Frame_Album.Visible = True
                    Ajustar_Objectos_Na_Horizontal
                                                            
                    'Ver a barra da lista de reprodução caso esta esteja oculta
                    Barra_Playlist.Visible = True
                    Icon_Barra_Informacoes(5).Picture = Form_Skin.Button_Playlist_Hide_Normal.Picture
                    Icon_Barra_Informacoes(5).ToolTipText = Idioma_Ocultar_Lista
                    Form_Preferencias.Check_Ver_Playlist.Value = 1
                    Form_Preferencias.Salvar_Valores
                    Desenhar_Formulario
                End With
            End If
        Next
    
'        'Avriguar mais tarde
'        For Each vItem In Folders 'Se for uma pasta-----------------------------------------------------------------
'            List1.AddItem "Folder: " & vItem
'        Next
'
'        For Each vItem In Parameters 'Se for um parâmetro------------------------------------------------------------
'            List1.AddItem "Parameter: " & vItem
'        Next
        Debug.Print sCommand
    End If
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    On Error Resume Next
    With Form_Skin
        Me.backcolor = vbWhite '.Cor_Form_Main.backcolor
        'Barra_Faixa.Picture = .Fundo_Barra_Faixa.Picture
        Barra_Faixa.Picture = .Fundo_Barra_Faixa.Picture
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Dim menu_form As Integer: For menu_form = 0 To Label_Menu.count - 1
            Label_Menu(menu_form).ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Next menu_form
        Fundo_Barra_Player.Picture = .Fundo_Barra_Player.Picture
        Fundo_Barra_Botoes.Picture = .Fundo_Barra_Player.Picture
        Botao_Antes.Picture = .Botao_Antes_Normal.Picture
        Botao_Pausa.Picture = .Botao_Pausa_Normal.Picture
        Botao_Play.Picture = .Botao_Play_Normal.Picture
        Botao_Seguinte.Picture = .Botao_Seguinte_Normal.Picture
        Picture_Slide_Som.Picture = .Fundo_Slider_Volume.Picture
        Slide_Som.Picture = .Slide_Som_Normal.Picture
        Botao_Mudo.Picture = .Som_On_Normal.Picture
        Label_Faixa.ForeColor = .Cor_Label_Barra_Visor.backcolor
        Tempo_Estimado.ForeColor = .Cor_Label_Barra_Visor.backcolor
        Label_Duracao.ForeColor = .Cor_Label_Barra_Visor.backcolor
        SliderBar.Picture = .Image_Barra_Slide.Picture
        Image_Barra_Slide.Picture = .Image_Barra_Slide.Picture
        Slide.Picture = .Slide_Musica_Normal.Picture
        Barra_Lateral.backcolor = .Cor_Fundo_Task_Bar.backcolor
        Separador_Barra_Lateral(0).backcolor = .Cor_Fundo_Task_Bar.backcolor
        Frame_Separador_Barra_Lateral(0).backcolor = .Cor_Fundo_Task_Bar.backcolor
        Separador_Barra_Lateral(1).backcolor = .Cor_Fundo_Task_Bar.backcolor
        Frame_Separador_Barra_Lateral(1).backcolor = .Cor_Fundo_Task_Bar.backcolor
        Separador_Barra_Lateral(2).backcolor = .Cor_Fundo_Task_Bar.backcolor
        Frame_Separador_Barra_Lateral(2).backcolor = .Cor_Fundo_Task_Bar.backcolor
        Frame_Separadores.backcolor = .Cor_Fundo_Task_Bar.backcolor
        Frame_Capa.backcolor = .Cor_BackGround_Frame_Cover.backcolor
        Pic_Capa_Album.backcolor = .Cor_BackGround_Frame_Cover.backcolor
        Label_Topico_Barra_Lateral(0).ForeColor = .Cor_Topic_Task_Bar.backcolor
        Label_Topico_Barra_Lateral(1).ForeColor = .Cor_Topic_Task_Bar.backcolor
        Label_Topico_Barra_Lateral(2).ForeColor = .Cor_Topic_Task_Bar.backcolor
        Label_Topico_Musica.ForeColor = .Cor_Letra_Topico_Over.backcolor
        Label_Topico_Filmes.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Label_Topico_Radio.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Label_Topico_Drive.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Label_Topico_MusicLink.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Label_Topico_Programas.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Shape_Topico(0).Picture = .Select_Topic_TaskBar.Picture
        Shape_Topico(1).Picture = .Select_Topic_TaskBar.Picture
        Shape_Topico(2).Picture = .Select_Topic_TaskBar.Picture
        Shape_Topico(4).Picture = .Select_Topic_TaskBar.Picture
        Shape_Topico(3).Picture = .Select_Topic_TaskBar.Picture
        Shape_Topico(5).Picture = .Select_Topic_TaskBar.Picture
        Dim xpto As Integer: For xpto = 0 To Label_Topico_Lista.count - 1
            Label_Topico_Lista(xpto).ForeColor = .Cor_Letra_Topico_Normal.backcolor
            Label_Topico_Lista(xpto).backcolor = .Cor_Fundo_Topico_Normal.backcolor
            Shape_Topico_Lista(xpto).Picture = .Select_Topic_TaskBar.Picture
        Next xpto
        Pic_Capa_Album.Picture = .Image_Sem_Capa.Picture
        Separador_Barra_Lateral(3).Picture = .Bar_View_Cover.Picture
        Personalizar_Grid Grelha_Musica
        Personalizar_Grid Grelha_Filmes
        Personalizar_Grid Grelha_Radio
        Personalizar_Grid Grelha_Contactos
        Personalizar_Grid Grelha_Minha_Musica
        Personalizar_Grid Grelha_Loja
        Personalizar_Grid Grelha_Lista_Em_Reproducao
        Personalizar_Grid Grelha_Artista
        Personalizar_Grid Grelha_Album
        Personalizar_Grid Grelha_Genero
        Personalizar_Grid Grelha_Listas
        Personalizar_Grid Grelha_Eventos
        Personalizar_Grid Grelha_Mensagens
        Personalizar_Grid Grelha_Ficheiros
        Personalizar_Grid Grelha_Recentes
        Personalizar_Grid Grelha_Favoritos
        Personalizar_Grid Grelha_Comunidade
        Personalizar_Grid Grelha_Amigos
        
        Fundo_Barra_Informacoes.Picture = .Fundo_Barra_Informacoes.Picture
        Label_Contador.ForeColor = .Cor_Letra_Bar_Info.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Botao_Restaurar.Picture = .Botao_Restaurar_Normal.Picture
        Botao_Minimizar.Picture = .Botao_Minimizar_Normal.Picture
        Botao_Maximizar.Picture = .Botao_Maximizar_Normal.Picture
        Scroll_Info.backcolor = .Cor_Scroll_Bar.backcolor
        Scroll_Info_Center.backcolor = .Cor_Scroll_Bar.backcolor
        Scroll_Info_Up.Picture = .Scroll_Info_Up.Picture
        Scroll_Info_Down.Picture = .Scroll_Info_Down.Picture
        Scroll_Info_Slider_Barras.Picture = .Scroll_Info_Slider_Barras.Picture
        Slider_Info.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Slider_Info.Picture = .Background_Scroll_Bar.Picture
        Barra_Botoes_Musica.backcolor = .Cor_BackGround_Bar_Label_Button.backcolor
        Dim Fundo_Botoes As Integer: For Fundo_Botoes = 0 To Label_Botao.count - 1
            Label_Botao(Fundo_Botoes).ForeColor = .Cor_Label_Button_ForeColor.backcolor
        Next Fundo_Botoes
        Barra_Caixa_Pesquisar_Musica.backcolor = .Cor_Fundo_Textbox.backcolor
        Barra_Caixa_Pesquisar_Musica.Picture = .Caixa_Pesquisar_Musica.Picture
        Text_Pesquisar_Musica.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Pesquisar_Musica.ForeColor = .Cor_Letra_Textbox.backcolor
    
        Dim star As Integer: For star = 0 To Estrela.count - 1
            Estrela(star).Picture = Form_Skin.Estrela_Normal.Picture
        Next star
        Tempo_Estimado.backcolor = .Cor_BackColor_Display.backcolor
        Label_Duracao.backcolor = .Cor_BackColor_Display.backcolor
        Botao_Actualizar_Programa.Picture = .Botao_Actualizar_Programa.Picture
        Image_Progresso.backcolor = .Cor_Slider_Music.backcolor
        Barra_Conexao.backcolor = .Cor_BackGround_Bar_Label_Button.backcolor
        Label_Conexao.ForeColor = .Cor_Label_Button_ForeColor.backcolor
        Icon_Topico(0).Picture = .Icon_Topico_Musica_Over.Picture
        Icon_Topico(1).Picture = .Icon_Topico_Filmes_Normal.Picture
        Icon_Topico(2).Picture = .Icon_Topico_radio_Normal.Picture
        Icon_Topico(4).Picture = .Icon_Topico_Drive_Normal.Picture
        Icon_Topico(3).Picture = .Icon_Topico_MusicLink_Normal.Picture
        Dim imagem_da_lista As Integer: For imagem_da_lista = 0 To Icon_Topico_Lista.count - 1
            Icon_Topico_Lista(imagem_da_lista).Picture = .Icon_Topico_Lista_Normal.Picture
        Next
        Label_Topico_Barra_Lateral(3).ForeColor = .Cor_Label_Frame_Cover.backcolor
        
        'Papel de parede da MusicLink
        Frame_Music_Link.backcolor = .Cor_Contorno_Caixas.backcolor
        If Form_Preferencias.Check_Wallpaper.Value = 0 Then
            'Fundo_Frame_Music_Link.Picture = .Fundo_Frame_Music_Link.Picture
            Fundo_Frame_Music_Link.Visible = False
        Else
            If ArquivoExiste(Form_Preferencias.Text_Wallpaper.Text, False) Then 'Verificar se existe o wallpaper indicado
                Fundo_Frame_Music_Link.Picture = LoadPicture(Form_Preferencias.Text_Wallpaper.Text)
                Fundo_Frame_Music_Link.Visible = True
            Else
                'Fundo_Frame_Music_Link.Picture = .Fundo_Frame_Music_Link.Picture
                Fundo_Frame_Music_Link.Visible = False
            End If
        End If
        
        Botao_Legendas.Picture = Form_Skin.Button_Menu_Normal.Picture
        Icon_Legendas.Picture = Form_Skin.Icon_Subtitles_Normal.Picture
        Icon_Barra_Informacoes(0).Picture = .Button_New_Playlist_Normal.Picture
        Botao_Mensagens.Picture = Form_Skin.Button_Menu_Standard_Normal.Picture
        Icon_Mensagens.Picture = Form_Skin.Icon_Mensagem_Normal.Picture
        
        '################### concluir mais tarde pois o icons poderão ser outros consoante o que está selecionado
        Icon_Barra_Informacoes(1).Picture = .Button_Music_Randomize_Normal.Picture
        Icon_Barra_Informacoes(2).Picture = .Button_Music_Repete_Normal.Picture
        If Frame_Capa.Visible = True Then
            Icon_Barra_Informacoes(3).Picture = .Button_Cover_Hide_Normal.Picture
        Else
            Icon_Barra_Informacoes(3).Picture = .Button_Cover_View_Normal.Picture
        End If
        Icon_Barra_Informacoes(4).Picture = .Button_Folder_Normal.Picture
        If Barra_Playlist.Visible = True Then
            Icon_Barra_Informacoes(5).Picture = .Button_Playlist_Hide_Normal.Picture
        Else
            Icon_Barra_Informacoes(5).Picture = .Button_Playlist_View_Normal.Picture
        End If
        Linha_Frame_Album.backcolor = .Cor_Line_Border_Frames.backcolor
        Linha_Vertical.backcolor = .Cor_Line_Border_Frames.backcolor
        Linha_Barra_Botoes_Musica.backcolor = .Cor_Line_Border_Frames.backcolor
        Linha_Barra_Conexao.backcolor = .Cor_Line_Border_Frames.backcolor
        Linha_Barra_Playlist.backcolor = .Cor_Line_Border_Frames.backcolor
        Frame_Grelhas_Pesquisa.backcolor = .Cor_Line_Border_Frames.backcolor
        Dim imagem_album As Integer: For imagem_abum = 0 To Image_Album.count - 1
            'Image_Album(imagem_album).backcolor = vbBlack
            Image_Album(imagem_album).Picture = .Image_Album.Picture
        Next
        Frame_Slide_Album.backcolor = vbBlack
        Frame_Slide.backcolor = vbBlack
        Frame_Album.backcolor = vbBlack
        
        Dim a As Integer: For n = 0 To Shape_Menu.count - 1
            Shape_Menu(n).backcolor = .Cor_Menu_BackColorSel.backcolor
        Next n
        
        Dim b As Integer: For b = 0 To Sombra_Ficheiro.count - 1
            Sombra_Ficheiro(b).backcolor = .Cor_Menu_BackColorSel.backcolor
            Menu_Ficheiro(b).ForeColor = .Cor_Menu_ForeColor.backcolor
        Next b
        
        Dim c As Integer: For c = 0 To Sombra_Editar.count - 1
            Sombra_Editar(c).backcolor = .Cor_Menu_BackColorSel.backcolor
            Menu_Editar(c).ForeColor = .Cor_Menu_ForeColor.backcolor
        Next c
        
        Dim d As Integer: For d = 0 To Sombra_Ver.count - 1
            Sombra_Ver(d).backcolor = .Cor_Menu_BackColorSel.backcolor
            Menu_Ver(d).ForeColor = .Cor_Menu_ForeColor.backcolor
        Next d
        
        Dim F As Integer: For F = 0 To Sombra_Controlos.count - 1
            Sombra_Controlos(F).backcolor = .Cor_Menu_BackColorSel.backcolor
            Menu_Controlos(F).ForeColor = .Cor_Menu_ForeColor.backcolor
        Next F
        
        Dim g As Integer: For g = 0 To Sombra_Ferramentas.count - 1
            Sombra_Ferramentas(g).backcolor = .Cor_Menu_BackColorSel.backcolor
            Menu_Ferramentas(g).ForeColor = .Cor_Menu_ForeColor.backcolor
        Next g
        
        Dim h As Integer: For h = 0 To Sombra_Ajuda.count - 1
            Sombra_Ajuda(h).backcolor = .Cor_Menu_BackColorSel.backcolor
            Menu_Ajuda(h).ForeColor = .Cor_Menu_ForeColor.backcolor
        Next h
        
        Dim j As Integer: For j = 0 To Frame_Menu.count - 1
            Frame_Menu(j).backcolor = .Cor_Menu_BackColor.backcolor
        Next j
        
        'Plano de preços/ My other drive
        Dim tabela_precos As Integer: For tabela_precos = 0 To Picture_Tabela.count - 1
            Picture_Tabela(tabela_precos).Picture = .Image_Precos.Picture
        Next
        
        Botao_Barra_Drive(0).Picture = Form_Skin.Icon_Seta_Anterior_Normal.Picture
        Botao_Barra_Drive(1).Picture = Form_Skin.Icon_Seta_Seguinte_Normal.Picture
        Botao_Barra_Drive(2).Picture = Form_Skin.Icon_Home_Normal.Picture
        Dim ctd As Integer: For ctd = 3 To 14
            Botao_Barra_Drive(ctd).Picture = Form_Skin.Botao_Barra_Normal.Picture
        Next
        
        Separador_Frame_Programas(0).Picture = Form_Skin.Icon_Home_Normal.Picture
        Dim tab_programas As Integer: For tab_programas = 1 To Separador_Frame_Programas.count - 1
            Separador_Frame_Programas(tab_programas).Picture = Form_Skin.Botao_Barra_Normal.Picture
        Next
    End With
End Sub

Public Sub Personalizar_Grid(nome_grelha As MSFlexGrid)
    'Procedimento para carregar o skin das grelhas
    With Form_Skin
        nome_grelha.backcolor = .Cor_Grid_BackColor.backcolor
        nome_grelha.BackColorBkg = .Cor_Grid_BackColorBkg.backcolor
        nome_grelha.BackColorFixed = .Cor_Grid_BackColorFixed.backcolor
        nome_grelha.BackColorSel = .Cor_Grid_BackColorSel.backcolor
        nome_grelha.ForeColor = .Cor_Grid_ForeColor.backcolor
        nome_grelha.ForeColorFixed = .Cor_Grid_ForeColorFixed.backcolor
        nome_grelha.ForeColorSel = .Cor_Grid_ForeColorSel.backcolor
        nome_grelha.GridColor = .Cor_Grid_Color.backcolor
        nome_grelha.GridColorFixed = .Cor_Grid_ColorFixed.backcolor
    End With
End Sub

Private Sub Carregar_Grelha_Albuns()
    'Carregar na grelha musica o albuns selecionado
    On Error GoTo Corrige_Erro
    If File_Ficheiros.ListCount = 0 Then Exit Sub
    With Grelha_Musica
        .Clear
        .Rows = 1
        Formatar_Grelha_Musica Grelha_Musica
        
        Dim cFile As New Classe_Tag_Editor
        Dim sFile As String, sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
        Dim j As Integer
        
        Me.MousePointer = 11
        File_Ficheiros.ListIndex = 0
        Dim i As Integer: i = 1
        
        For j = 0 To File_Ficheiros.ListCount - 1
            DoEvents
            cFile.FileName = Lista_Pastas.List(Index_Album) & "\" & File_Ficheiros.List(j)
            
            sTitle = Replace(cFile.Title, "'", " ", , , vbTextCompare)
            sArtist = Replace(cFile.Artist, "'", " ", , , vbTextCompare)
            sAlbum = Replace(cFile.album, "'", " ", , , vbTextCompare)
            sYear = Replace(cFile.Year, "'", " ", , , vbTextCompare)
            sGenre = Replace(cFile.Genre, "'", " ", , , vbTextCompare)
            sComment = Replace(cFile.Comments, "'", " ", , , vbTextCompare)
            
            'Caso o ficheiro não tenha tags então o titulo será o nome do ficheiro, o qual é obtido atrvés do directório
            Dim Arquivo() As String
            Dim DiretorioArq As String
            DiretorioArq = Lista_Pastas.List(Index_Album) & "\" & File_Ficheiros.List(j)
            Arquivo = Split(DiretorioArq, "\")
                       
            Dim nome_ficheiro As String: nome_ficheiro = Dir(Lista_Pastas.List(Index_Album) & "\" & File_Ficheiros.List(j), vbArchive)
            If sTitle = "" Then sTitle = Mid(nome_ficheiro, 1, InStrRev(nome_ficheiro, ".") - 1)
            If sArtist = "" Then sArtist = ""
            If sAlbum = "" Then sAlbum = ""
            If sYear = "" Then sYear = ""
            If sGenre = "" Then sGenre = ""
            If sComment = "" Then sComment = ""
                    
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = DiretorioArq
            .TextMatrix(i, 1) = sTitle
            .TextMatrix(i, 2) = sArtist
            .TextMatrix(i, 3) = sAlbum
            .TextMatrix(i, 4) = sYear
            .TextMatrix(i, 5) = sGenre
            .TextMatrix(i, 6) = sComment
            .TextMatrix(i, 7) = DiretorioArq
            .TextMatrix(i, 8) = "0"
            .TextMatrix(i, 9) = j
            i = i + 1
        Next j
            
        Me.MousePointer = 0
    End With
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
End Sub

Public Sub Carregar_Menu_Ficheiro()
    'Criar os menus consoante o nº de linhas
    If List_Menu(0).ListCount = 0 Then Exit Sub
    Dim Linha, menu As Integer
    Dim ultimo_menu As Integer: ultimo_menu = 0
    List_Menu(0).ListIndex = 0
    Menu_Ficheiro(0).Caption = List_Menu(0).List(0)
    Sombra_Ficheiro(0).top = 2
    Sombra_Ficheiro(0).left = 2
    Sombra_Ficheiro(0).Height = 22
    Menu_Ficheiro(0).top = Sombra_Ficheiro(0).top + ((Sombra_Ficheiro(0).Height - Menu_Ficheiro(0).Height) / 2)
    
    For Linha = 1 To List_Menu(0).ListCount - 1
        If List_Menu(0).List(Linha) <> "-" Then
            Load Sombra_Ficheiro(Linha)
            Sombra_Ficheiro(Linha).Move Sombra_Ficheiro(0).left, Sombra_Ficheiro(ultimo_menu).top + Sombra_Ficheiro(0).Height
            Sombra_Ficheiro(Linha).Visible = False
            
            Load Menu_Ficheiro(Linha)
            Menu_Ficheiro(Linha).Move Menu_Ficheiro(0).left, Sombra_Ficheiro(Linha).top + ((Sombra_Ficheiro(Linha).Height - Menu_Ficheiro(Linha).Height) / 2)
            Menu_Ficheiro(Linha).Visible = True
            Menu_Ficheiro(Linha).ZOrder 0
            Menu_Ficheiro(Linha).Caption = List_Menu(0).List(Linha)
            ultimo_menu = Linha
            
        Else
            If Linha_Ficheiro(0).Visible = False Then
                Linha_Ficheiro(0).Visible = True
                Linha_Ficheiro(0).top = Sombra_Ficheiro(ultimo_menu).top + Sombra_Ficheiro(0).Height
            Else
                Load Linha_Ficheiro(Linha)
                Linha_Ficheiro(Linha).Move Menu_Ficheiro(0).left, Sombra_Ficheiro(ultimo_menu).top + Sombra_Ficheiro(0).Height
                Linha_Ficheiro(Linha).Visible = True
            End If
        End If
    Next Linha
    
    'Selecionar a 1ªlinha do menu
    Linha_Selecionada_Ficheiro = 0
    Sombra_Ficheiro(0).Visible = True
    Menu_Ficheiro(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
End Sub

Public Sub Carregar_Menu_Editar()
    'Criar os menus consoante o nº de linhas
    If List_Menu(1).ListCount = 0 Then Exit Sub
    Dim Linha, menu As Integer
    Dim ultimo_menu As Integer: ultimo_menu = 0
    List_Menu(1).ListIndex = 0
    Menu_Editar(0).Caption = List_Menu(1).List(0)
    Sombra_Editar(0).top = 2
    Sombra_Editar(0).left = 2
    Sombra_Editar(0).Height = 22
    Menu_Editar(0).top = Sombra_Editar(0).top + ((Sombra_Editar(0).Height - Menu_Editar(0).Height) / 2)
    
    For Linha = 1 To List_Menu(1).ListCount - 1
        If List_Menu(1).List(Linha) <> "-" Then
            Load Sombra_Editar(Linha)
            Sombra_Editar(Linha).Move Sombra_Editar(0).left, Sombra_Editar(ultimo_menu).top + Sombra_Editar(0).Height
            Sombra_Editar(Linha).Visible = False
            
            Load Menu_Editar(Linha)
            Menu_Editar(Linha).Move Menu_Editar(0).left, Sombra_Editar(Linha).top + ((Sombra_Editar(Linha).Height - Menu_Editar(Linha).Height) / 2)
            Menu_Editar(Linha).Visible = True
            Menu_Editar(Linha).ZOrder 0
            Menu_Editar(Linha).Caption = List_Menu(1).List(Linha)
            ultimo_menu = Linha
            
        Else
            If Linha_Editar(0).Visible = False Then
                Linha_Editar(0).Visible = True
                Linha_Editar(0).top = Sombra_Editar(ultimo_menu).top + Sombra_Editar(0).Height
            Else
                Load Linha_Editar(Linha)
                Linha_Editar(Linha).Move Menu_Editar(0).left, Sombra_Editar(ultimo_menu).top + Sombra_Editar(0).Height
                Linha_Editar(Linha).Visible = True
            End If
        End If
    Next Linha
    
    'Selecionar a 1ªlinha do menu
    Linha_Selecionada_Editar = 0
    Sombra_Editar(0).Visible = True
    Menu_Editar(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
End Sub

Public Sub Carregar_Menu_Ver()
    'Criar os menus consoante o nº de linhas
    If List_Menu(2).ListCount = 0 Then Exit Sub
    Dim Linha, menu As Integer
    Dim ultimo_menu As Integer: ultimo_menu = 0
    List_Menu(2).ListIndex = 0
    Menu_Ver(0).Caption = List_Menu(2).List(0)
    Sombra_Ver(0).top = 2
    Sombra_Ver(0).left = 2
    Sombra_Ver(0).Height = 22
    Menu_Ver(0).top = Sombra_Ver(0).top + ((Sombra_Ver(0).Height - Menu_Ver(0).Height) / 2)
    
    For Linha = 1 To List_Menu(2).ListCount - 1
        If List_Menu(2).List(Linha) <> "-" Then
            Load Sombra_Ver(Linha)
            Sombra_Ver(Linha).Move Sombra_Ver(0).left, Sombra_Ver(ultimo_menu).top + Sombra_Ver(0).Height
            Sombra_Ver(Linha).Visible = False
            
            Load Menu_Ver(Linha)
            Menu_Ver(Linha).Move Menu_Ver(0).left, Sombra_Ver(Linha).top + ((Sombra_Ver(Linha).Height - Menu_Ver(Linha).Height) / 2)
            Menu_Ver(Linha).Visible = True
            Menu_Ver(Linha).ZOrder 0
            Menu_Ver(Linha).Caption = List_Menu(2).List(Linha)
            ultimo_menu = Linha
            
        Else
            If Linha_Ver(0).Visible = False Then
                Linha_Ver(0).Visible = True
                Linha_Ver(0).top = Sombra_Ver(ultimo_menu).top + Sombra_Ver(0).Height
            Else
                Load Linha_Ver(Linha)
                Linha_Ver(Linha).Move Menu_Ver(0).left, Sombra_Ver(ultimo_menu).top + Sombra_Ver(0).Height
                Linha_Ver(Linha).Visible = True
            End If
        End If
    Next Linha
    
    'Selecionar a 1ªlinha do menu
    Linha_Selecionada_Ver = 0
    Sombra_Ver(0).Visible = True
    Menu_Ver(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
End Sub

Public Sub Carregar_Menu_Controlos()
    'Criar os menus consoante o nº de linhas
    If List_Menu(3).ListCount = 0 Then Exit Sub
    Dim Linha, menu As Integer
    Dim ultimo_menu As Integer: ultimo_menu = 0
    List_Menu(3).ListIndex = 0
    Menu_Controlos(0).Caption = List_Menu(3).List(0)
    Sombra_Controlos(0).top = 2
    Sombra_Controlos(0).left = 2
    Sombra_Controlos(0).Height = 22
    Menu_Controlos(0).top = Sombra_Controlos(0).top + ((Sombra_Controlos(0).Height - Menu_Controlos(0).Height) / 2)
    
    For Linha = 1 To List_Menu(3).ListCount - 1
        If List_Menu(3).List(Linha) <> "-" Then
            Load Sombra_Controlos(Linha)
            Sombra_Controlos(Linha).Move Sombra_Controlos(0).left, Sombra_Controlos(ultimo_menu).top + Sombra_Controlos(0).Height
            Sombra_Controlos(Linha).Visible = False
            
            Load Menu_Controlos(Linha)
            Menu_Controlos(Linha).Move Menu_Controlos(0).left, Sombra_Controlos(Linha).top + ((Sombra_Controlos(Linha).Height - Menu_Controlos(Linha).Height) / 2)
            Menu_Controlos(Linha).Visible = True
            Menu_Controlos(Linha).ZOrder 0
            Menu_Controlos(Linha).Caption = List_Menu(3).List(Linha)
            ultimo_menu = Linha
            
        Else
            If Linha_Controlos(0).Visible = False Then
                Linha_Controlos(0).Visible = True
                Linha_Controlos(0).top = Sombra_Controlos(ultimo_menu).top + Sombra_Controlos(0).Height
            Else
                Load Linha_Controlos(Linha)
                Linha_Controlos(Linha).Move Menu_Controlos(0).left, Sombra_Controlos(ultimo_menu).top + Sombra_Controlos(0).Height
                Linha_Controlos(Linha).Visible = True
            End If
        End If
    Next Linha
    
    'Selecionar a 1ªlinha do menu
    Linha_Selecionada_Controlos = 0
    Sombra_Controlos(0).Visible = True
    Menu_Controlos(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
End Sub

Public Sub Carregar_Menu_Ferramentas()
    'Criar os menus consoante o nº de linhas
    If List_Menu(4).ListCount = 0 Then Exit Sub
    Dim Linha, menu As Integer
    Dim ultimo_menu As Integer: ultimo_menu = 0
    List_Menu(4).ListIndex = 0
    Menu_Ferramentas(0).Caption = List_Menu(4).List(0)
    Sombra_Ferramentas(0).top = 2
    Sombra_Ferramentas(0).left = 2
    Sombra_Ferramentas(0).Height = 22
    Menu_Ferramentas(0).top = Sombra_Ferramentas(0).top + ((Sombra_Ferramentas(0).Height - Menu_Ferramentas(0).Height) / 2)
    
    For Linha = 1 To List_Menu(4).ListCount - 1
        If List_Menu(4).List(Linha) <> "-" Then
            Load Sombra_Ferramentas(Linha)
            Sombra_Ferramentas(Linha).Move Sombra_Ferramentas(0).left, Sombra_Ferramentas(ultimo_menu).top + Sombra_Ferramentas(0).Height
            Sombra_Ferramentas(Linha).Visible = False
            
            Load Menu_Ferramentas(Linha)
            Menu_Ferramentas(Linha).Move Menu_Ferramentas(0).left, Sombra_Ferramentas(Linha).top + ((Sombra_Ferramentas(Linha).Height - Menu_Ferramentas(Linha).Height) / 2)
            Menu_Ferramentas(Linha).Visible = True
            Menu_Ferramentas(Linha).ZOrder 0
            Menu_Ferramentas(Linha).Caption = List_Menu(4).List(Linha)
            ultimo_menu = Linha
            
        Else
            If Linha_Ferramentas(0).Visible = False Then
                Linha_Ferramentas(0).Visible = True
                Linha_Ferramentas(0).top = Sombra_Ferramentas(ultimo_menu).top + Sombra_Ferramentas(0).Height
            Else
                Load Linha_Ferramentas(Linha)
                Linha_Ferramentas(Linha).Move Menu_Ferramentas(0).left, Sombra_Ferramentas(ultimo_menu).top + Sombra_Ferramentas(0).Height
                Linha_Ferramentas(Linha).Visible = True
            End If
        End If
    Next Linha
    
    'Selecionar a 1ªlinha do menu
    Linha_Selecionada_Ferramentas = 0
    Sombra_Ferramentas(0).Visible = True
    Menu_Ferramentas(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
End Sub

Public Sub Carregar_Menu_Ajuda()
    'Criar os menus consoante o nº de linhas
    If List_Menu(5).ListCount = 0 Then Exit Sub
    Dim Linha, menu As Integer
    Dim ultimo_menu As Integer: ultimo_menu = 0
    List_Menu(5).ListIndex = 0
    Menu_Ajuda(0).Caption = List_Menu(5).List(0)
    Sombra_Ajuda(0).top = 2
    Sombra_Ajuda(0).left = 2
    Sombra_Ajuda(0).Height = 22
    Menu_Ajuda(0).top = Sombra_Ajuda(0).top + ((Sombra_Ajuda(0).Height - Menu_Ajuda(0).Height) / 2)
    
    For Linha = 1 To List_Menu(5).ListCount - 1
        If List_Menu(5).List(Linha) <> "-" Then
            Load Sombra_Ajuda(Linha)
            Sombra_Ajuda(Linha).Move Sombra_Ajuda(0).left, Sombra_Ajuda(ultimo_menu).top + Sombra_Ajuda(0).Height
            Sombra_Ajuda(Linha).Visible = False
            
            Load Menu_Ajuda(Linha)
            Menu_Ajuda(Linha).Move Menu_Ajuda(0).left, Sombra_Ajuda(Linha).top + ((Sombra_Ajuda(Linha).Height - Menu_Ajuda(Linha).Height) / 2)
            Menu_Ajuda(Linha).Visible = True
            Menu_Ajuda(Linha).ZOrder 0
            Menu_Ajuda(Linha).Caption = List_Menu(5).List(Linha)
            ultimo_menu = Linha
            
        Else
            If Linha_Ajuda(0).Visible = False Then
                Linha_Ajuda(0).Visible = True
                Linha_Ajuda(0).top = Sombra_Ajuda(ultimo_menu).top + Sombra_Ajuda(0).Height
            Else
                Load Linha_Ajuda(Linha)
                Linha_Ajuda(Linha).Move Menu_Ajuda(0).left, Sombra_Ajuda(ultimo_menu).top + Sombra_Ajuda(0).Height
                Linha_Ajuda(Linha).Visible = True
            End If
        End If
    Next Linha
    
    'Selecionar a 1ªlinha do menu
    Linha_Selecionada_Ajuda = 0
    Sombra_Ajuda(0).Visible = True
    Menu_Ajuda(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
End Sub

Public Sub Carregar_Listas()
    'Procedimento para carregar as listas existentes na pasta "\playlist\"
    Dir_Lista.Path = App.Path & "\Library\Playlist\"
    File_Lista.Path = Dir_Lista.Path
    File_Lista.Pattern = "*.ini"
    
    If File_Lista.ListCount = 0 Then Exit Sub
    If File_Lista.ListCount > 0 Then
        'Criar a lista consoante o nº de idiomas disponiveis
        Label_Topico_Lista(0).Caption = ""
        Label_Topico_Lista(0).Visible = True
        Shape_Topico_Lista(0).Visible = False
        Icon_Topico_Lista(0).Visible = True
        
        Repor_a_Cor_Dos_Topicos
        
        Dim Objecto As Integer
        For Objecto = 1 To File_Lista.ListCount - 1
            Load Shape_Topico_Lista(Objecto)
            Shape_Topico_Lista(Objecto).Move Shape_Topico_Lista(Objecto - 1).left, Shape_Topico_Lista(Objecto - 1).top + Shape_Topico_Lista(Objecto - 1).Height
            Shape_Topico_Lista(Objecto).Visible = False
            'Shape_Topico_Lista(Objecto).ZOrder 1
            
            Load Label_Topico_Lista(Objecto)
            Label_Topico_Lista(Objecto).Move Label_Topico_Lista(Objecto - 1).left, Shape_Topico_Lista(Objecto).top + ((Shape_Topico_Lista(Objecto).Height - Label_Topico_Lista(Objecto).Height) / 2)
            Label_Topico_Lista(Objecto).Visible = True
            
            Load Icon_Topico_Lista(Objecto)
            Icon_Topico_Lista(Objecto).Move Icon_Topico_Lista(Objecto - 1).left, Shape_Topico_Lista(Objecto).top + ((Shape_Topico_Lista(Objecto).Height - Icon_Topico_Lista(Objecto).Height) / 2)
            Icon_Topico_Lista(Objecto).Visible = True
            
            Shape_Topico_Lista(Objecto).ZOrder 1
        Next Objecto
        
        'Preencher as label's com as Listas disponiveis
        Dim Z As Integer
        If File_Lista.ListCount > 0 Then
            File_Lista.ListIndex = 0
            For Z = 0 To File_Lista.ListCount - 1
                Label_Topico_Lista(Z).Caption = Espaco & left$(File_Lista.List(Z), InStr(File_Lista.List(Z), ".") - (1)) 'Retirar a extensão do ficheiro ".lng"
            Next Z
        End If
        
        'Ajustar o tamanho da Frame_Separador_Barra_Lateral(2)
        Frame_Separador_Barra_Lateral(2).Height = (Shape_Topico_Lista.count * Separador_Barra_Lateral(0).ScaleHeight) + Separador_Barra_Lateral(0).ScaleHeight
    End If
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Menu(0).Caption = ReadINI("Main", "Label_File", Localizacao_Ficheiro_Lingua)
    Label_Menu(1).Caption = ReadINI("Main", "Label_Edit", Localizacao_Ficheiro_Lingua)
    Label_Menu(2).Caption = ReadINI("Main", "Label_View", Localizacao_Ficheiro_Lingua)
    Label_Menu(3).Caption = ReadINI("Main", "Label_Controls", Localizacao_Ficheiro_Lingua)
    Label_Menu(4).Caption = ReadINI("Main", "Label_Tools", Localizacao_Ficheiro_Lingua)
    Label_Menu(5).Caption = ReadINI("Main", "Label_Help", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Main", "Button_Close", Localizacao_Ficheiro_Lingua)
    Botao_Restaurar.ToolTipText = ReadINI("Main", "Button_Restore", Localizacao_Ficheiro_Lingua)
    Botao_Minimizar.ToolTipText = ReadINI("Main", "Button_Minimize", Localizacao_Ficheiro_Lingua)
    Botao_Maximizar.ToolTipText = ReadINI("Main", "Button_Maximize", Localizacao_Ficheiro_Lingua)
    Botao_Antes.ToolTipText = ReadINI("Main", "Button_Previous_Track", Localizacao_Ficheiro_Lingua)
    Botao_Player_Mini(0).ToolTipText = ReadINI("Main", "Button_Previous_Track", Localizacao_Ficheiro_Lingua)
    Botao_Play.ToolTipText = ReadINI("Main", "Button_Play", Localizacao_Ficheiro_Lingua)
    Botao_Player_Mini(1).ToolTipText = ReadINI("Main", "Button_Play", Localizacao_Ficheiro_Lingua)
    Botao_Pausa.ToolTipText = ReadINI("Main", "Button_Pause", Localizacao_Ficheiro_Lingua)
    Botao_Player_Mini(2).ToolTipText = ReadINI("Main", "Button_Pause", Localizacao_Ficheiro_Lingua)
    Botao_Seguinte.ToolTipText = ReadINI("Main", "Button_Next_Track", Localizacao_Ficheiro_Lingua)
    Botao_Player_Mini(3).ToolTipText = ReadINI("Main", "Button_Next_Track", Localizacao_Ficheiro_Lingua)
    Form_Mini_Player.Botao_Play.ToolTipText = ReadINI("Main", "Button_Play", Localizacao_Ficheiro_Lingua)
    Form_Mini_Player.Botao_Pausa.ToolTipText = ReadINI("Main", "Button_Pause", Localizacao_Ficheiro_Lingua)
    Form_PopUp.Botao_Play.ToolTipText = ReadINI("Main", "Button_Play", Localizacao_Ficheiro_Lingua)
    Form_PopUp.Botao_Pausa.ToolTipText = ReadINI("Main", "Button_Pause", Localizacao_Ficheiro_Lingua)
    Label_Topico_Barra_Lateral(0).Caption = ReadINI("Main", "Topic_Library", Localizacao_Ficheiro_Lingua)
    Label_Topico_Musica.Caption = Espaco & ReadINI("Main", "Topic_Music", Localizacao_Ficheiro_Lingua)
    Label_Topico_Filmes.Caption = Espaco & ReadINI("Main", "Topic_Movies", Localizacao_Ficheiro_Lingua)
    Label_Topico_MusicLink.Caption = Espaco & "Music link"
    Label_Topico_Programas.Caption = Espaco & "App library"
    Label_Topico_Barra_Lateral(1).Caption = ReadINI("Main", "Topic_Services", Localizacao_Ficheiro_Lingua)
    Label_Topico_Radio.Caption = Espaco & ReadINI("Main", "Topic_Radio", Localizacao_Ficheiro_Lingua)
    Label_Topico_Drive.Caption = Espaco & "My other drive"
    Label_Topico_Barra_Lateral(2).Caption = ReadINI("Main", "Topic_Playlist", Localizacao_Ficheiro_Lingua)
    Idioma_Mudo_On = ReadINI("Main", "Button_Mute_On", Localizacao_Ficheiro_Lingua)
    Idioma_Mudo_Off = ReadINI("Main", "Button_Mute_Off", Localizacao_Ficheiro_Lingua)
    Label_Botao(0).Caption = ReadINI("Main", "Button_New_Library", Localizacao_Ficheiro_Lingua)
    Label_Botao(1).Caption = ReadINI("Main", "Button_Download", Localizacao_Ficheiro_Lingua)
    Label_Botao(2).Caption = ReadINI("Main", "Button_Add_My_Music", Localizacao_Ficheiro_Lingua)
    Label_Botao(4).Caption = ReadINI("Main", "Button_Remove", Localizacao_Ficheiro_Lingua)
    Label_Botao(6).Caption = ReadINI("Main", "Button_Login", Localizacao_Ficheiro_Lingua)
    Label_Botao(3).Caption = ReadINI("Main", "Button_Create_Account", Localizacao_Ficheiro_Lingua)
    Label_Botao(5).Caption = ReadINI("Main", "Button_Add_Link", Localizacao_Ficheiro_Lingua)
    Label_Botao(7).Caption = ReadINI("Main", "Button_New_Playlist", Localizacao_Ficheiro_Lingua)
    Label_Botao(8).Caption = ReadINI("Main", "Button_Save_Playlist", Localizacao_Ficheiro_Lingua)
    Label_Botao(9).Caption = ReadINI("Main", "Button_View_Profile", Localizacao_Ficheiro_Lingua)
    Label_Botao(10).Caption = ReadINI("Main", "Button_Send_Invitation", Localizacao_Ficheiro_Lingua)
    Label_Botao(11).Caption = ReadINI("Main", "Button_What_listening", Localizacao_Ficheiro_Lingua)
    Label_Botao(12).Caption = ReadINI("Main", "Button_Send_Message", Localizacao_Ficheiro_Lingua)
    Label_Botao(13).Caption = ReadINI("Main", "Button_New_Contact", Localizacao_Ficheiro_Lingua)
    Label_Botao(14).Caption = ReadINI("Main", "Button_Edit_Contact", Localizacao_Ficheiro_Lingua)
    Label_Botao(15).Caption = ReadINI("Main", "Button_Delete_Contact", Localizacao_Ficheiro_Lingua)
    Label_Botao(16).Caption = ReadINI("Main", "Button_Send_Mail", Localizacao_Ficheiro_Lingua)
    Label_Botao(17).Caption = ReadINI("Main", "Button_New_Event", Localizacao_Ficheiro_Lingua)
    Label_Botao(18).Caption = ReadINI("Main", "Button_Edit_Event", Localizacao_Ficheiro_Lingua)
    Label_Botao(19).Caption = ReadINI("Main", "Button_Delete_Event", Localizacao_Ficheiro_Lingua)
    Label_Botao(20).Caption = ReadINI("Main", "Button_If_You_Like", Localizacao_Ficheiro_Lingua)
    Imagem_Votar.ToolTipText = ReadINI("Main", "Button_I_Like", Localizacao_Ficheiro_Lingua)
    
    Idioma_Name_Of_New_Playlist = ReadINI("Main", "Name_Of_New_Playlist", Localizacao_Ficheiro_Lingua)
    Label_Legendas.Caption = ReadINI("Main", "Button_Subtitles", Localizacao_Ficheiro_Lingua)
    Botao_Mudo.ToolTipText = Idioma_Mudo_On
    Botao_Mudo_Mini.ToolTipText = Idioma_Mudo_On
    Label_Nenhum_Album.Caption = ReadINI("Main", "Label_No_Album_Available", Localizacao_Ficheiro_Lingua)
    Label_Actualizar_Programa.Caption = ReadINI("Main", "Button_Update_Program", Localizacao_Ficheiro_Lingua)
    
    Idioma_Desenvolvido = ReadINI("Main", "Info_Developer", Localizacao_Ficheiro_Lingua)
    Idioma_Nova_Versao = ReadINI("Main", "Info_New_Version", Localizacao_Ficheiro_Lingua)
    Idioma_Nao_Existe_Actualizacoes = ReadINI("Main", "Info_Updates_Unavailable", Localizacao_Ficheiro_Lingua)
    Idioma_Topico_Musica = ReadINI("Main", "Topic_Music", Localizacao_Ficheiro_Lingua)
    Idioma_Topico_Filmes = ReadINI("Main", "Topic_Movies", Localizacao_Ficheiro_Lingua)
    Idioma_Topico_Loja = ReadINI("Main", "Topic_Shop_Online", Localizacao_Ficheiro_Lingua)
    Idioma_Topico_Radio = ReadINI("Main", "Topic_Radio", Localizacao_Ficheiro_Lingua)
    Idioma_Total_Musicas = ReadINI("Main", "Label_Total_Music", Localizacao_Ficheiro_Lingua)
    Idioma_Total_Filmes = ReadINI("Main", "Label_Total_Movies", Localizacao_Ficheiro_Lingua)
    Idioma_Total_Ficheiros_Online = ReadINI("Main", "Label_Total_Files_Online", Localizacao_Ficheiro_Lingua)
    Idioma_Total_Estacoes_Radio = ReadINI("Main", "Label_Total_Radio_Stations", Localizacao_Ficheiro_Lingua)
    Idioma_Total_Contactos = ReadINI("Main", "Label_Total_Contacts", Localizacao_Ficheiro_Lingua)
    Idioma_Total_Messagens = ReadINI("Main", "Label_Total_Messages", Localizacao_Ficheiro_Lingua)
    Idioma_Total_Utilizadores = ReadINI("Main", "Label_Total_Users", Localizacao_Ficheiro_Lingua)
    Idioma_Total_Amigos = ReadINI("Main", "Label_Total_Friends", Localizacao_Ficheiro_Lingua)
    Idioma_Total_Eventos = ReadINI("Main", "Label_Total_Events", Localizacao_Ficheiro_Lingua)
    
    Idioma_Grid_Music_Col_1 = ReadINI("Main", "Grid_Music_Col_1", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Music_Col_2 = ReadINI("Main", "Grid_Music_Col_2", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Music_Col_3 = ReadINI("Main", "Grid_Music_Col_3", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Music_Col_4 = ReadINI("Main", "Grid_Music_Col_4", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Music_Col_5 = ReadINI("Main", "Grid_Music_Col_5", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Music_Col_6 = ReadINI("Main", "Grid_Music_Col_6", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Music_Col_7 = ReadINI("Main", "Grid_Music_Col_7", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Music_Col_8 = ReadINI("Main", "Grid_Music_Col_8", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Movies_Col_1 = ReadINI("Main", "Grid_Movies_Col_1", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Movies_Col_2 = ReadINI("Main", "Grid_Movies_Col_2", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Movies_Col_3 = ReadINI("Main", "Grid_Movies_Col_3", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Movies_Col_4 = ReadINI("Main", "Grid_Movies_Col_4", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Movies_Col_5 = ReadINI("Main", "Grid_Movies_Col_5", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Movies_Col_6 = ReadINI("Main", "Grid_Movies_Col_6", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Playlist_Col_1 = ReadINI("Main", "Grid_Playlist_Col_1", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Radio_Col_1 = ReadINI("Main", "Grid_Radio_Col_1", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Loja_Col_1 = ReadINI("Main", "Grid_Shop_Col_1", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Loja_Col_2 = ReadINI("Main", "Grid_Shop_Col_2", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Loja_Col_3 = ReadINI("Main", "Grid_Shop_Col_3", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Loja_Col_4 = ReadINI("Main", "Grid_Shop_Col_4", Localizacao_Ficheiro_Lingua)
    Idioma_Topico_Procurar = ReadINI("Main", "Topic_Find", Localizacao_Ficheiro_Lingua)
    Idioma_Topico_Minha_Musica = ReadINI("Main", "Topic_My_Music", Localizacao_Ficheiro_Lingua)
    Idioma_Topico_Resultado_Pesquisa = ReadINI("Main", "Topic_Search", Localizacao_Ficheiro_Lingua)
    Idioma_Label_Topico_Barra_Lateral(2) = ReadINI("Main", "Topic_Playlist", Localizacao_Ficheiro_Lingua)
    Idioma_Label_Topico_Drive = "My other drive"
    Idioma_Pesquisa_Musica = ReadINI("Main", "Text_Search_Music", Localizacao_Ficheiro_Lingua)
    Idioma_Button_Fullscreen_On = ReadINI("Main", "Button_Fullscreen_On", Localizacao_Ficheiro_Lingua)
    Idioma_Button_Fullscreen_Off = ReadINI("Main", "Button_Fullscreen_Off", Localizacao_Ficheiro_Lingua)
    Botao_Player_Mini(4).ToolTipText = Idioma_Button_Fullscreen_On
    
    With List_Menu(0)
        .Clear
        .AddItem "      " & ReadINI("Main", "Menu_File_New_Library", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Menu_File_Update_Library", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_File_Add_Media", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Menu_File_New_Playlist", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Menu_File_Save_Playlist", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_File_Open_Location", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_File_Close", Localizacao_Ficheiro_Lingua)
    End With
    
    With List_Menu(1)
        .Clear
        .AddItem "      " & ReadINI("Main", "Menu_Edit_Add_Playlist", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Menu_Edit_Remove_Library", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_Edit_Copie_Url", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_Edit_Clear_Playlist", Localizacao_Ficheiro_Lingua)
    End With

    With List_Menu(2)
        .Clear
        .AddItem "      " & ReadINI("Main", "Menu_View_Compact_Mode", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_View_Cover", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Menu_View_Playlist", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Icon_Simple", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Icon_Advanced_Search", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Icon_Album_Art", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_View_Video_Secreen", Localizacao_Ficheiro_Lingua)
    End With
    
    With List_Menu(3)
        .Clear
        .AddItem "      " & ReadINI("Main", "Menu_Controls_Play", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Menu_Controls_Previous_Track", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Menu_Controls_Next_Track", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_Controls_Mute", Localizacao_Ficheiro_Lingua)
    End With
    
    With List_Menu(4)
        .Clear
        .AddItem "      " & ReadINI("Main", "Icon_Properties", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Icon_Tag", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & "Media manager"
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_Tools_Options", Localizacao_Ficheiro_Lingua)
    End With

    With List_Menu(5)
        .Clear
        .AddItem "      " & ReadINI("Main", "Menu_Official_Website", Localizacao_Ficheiro_Lingua)
        .AddItem "      " & ReadINI("Main", "Menu_Technical_Support", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_Check_Updates", Localizacao_Ficheiro_Lingua)
        .AddItem "-"
        .AddItem "      " & ReadINI("Main", "Menu_About", Localizacao_Ficheiro_Lingua)
    End With
    
    Label_Texto(4).Caption = "MUSIC LINK"
    Label_Texto(5).Caption = ReadINI("Main", "Label_MusicLink_2", Localizacao_Ficheiro_Lingua)
    Label_Nova_Versao.Caption = Idioma_Nao_Existe_Actualizacoes
    Close_Barra_Actualizar.ToolTipText = ReadINI("Main", "Close_Bar_Updated", Localizacao_Ficheiro_Lingua)
    Label_Texto(0).Caption = "MY OTHER DRIVE"
    Label_Texto(1).Caption = ReadINI("Main", "Label_Servers_Demand", Localizacao_Ficheiro_Lingua)
    Label_Texto(2).Caption = ReadINI("Main", "Label_Free", Localizacao_Ficheiro_Lingua)
    Label_Aderir_Agora.Caption = ReadINI("Main", "Label_Join_Now", Localizacao_Ficheiro_Lingua)
    Label_Texto(3).Caption = ReadINI("Main", "Label_Plans_To_Meet", Localizacao_Ficheiro_Lingua)
    Label_Plano(0).Caption = ReadINI("Main", "Label_Standard", Localizacao_Ficheiro_Lingua)
    Label_Plano(1).Caption = ReadINI("Main", "Label_Advanced", Localizacao_Ficheiro_Lingua)
    Label_Plano(2).Caption = ReadINI("Main", "Label_Professional", Localizacao_Ficheiro_Lingua)
    Label_Popular.Caption = ReadINI("Main", "Label_Most_Popular", Localizacao_Ficheiro_Lingua)
    Label_Mensalidade(0).Caption = ReadINI("Main", "Label_Payment", Localizacao_Ficheiro_Lingua)
    Label_Mensalidade(1).Caption = ReadINI("Main", "Label_Payment", Localizacao_Ficheiro_Lingua)
    Label_Mensalidade(2).Caption = ReadINI("Main", "Label_Payment", Localizacao_Ficheiro_Lingua)
    Label_Funcionalidades(0).Caption = ReadINI("Main", "Label_Features", Localizacao_Ficheiro_Lingua)
    Label_Funcionalidades(1).Caption = ReadINI("Main", "Label_Features", Localizacao_Ficheiro_Lingua)
    Label_Funcionalidades(2).Caption = ReadINI("Main", "Label_Features", Localizacao_Ficheiro_Lingua)
    Label_Dados(0).Caption = ReadINI("Main", "Label_Storage0", Localizacao_Ficheiro_Lingua)
    Label_Dados(1).Caption = ReadINI("Main", "Label_Storage1", Localizacao_Ficheiro_Lingua)
    Label_Dados(2).Caption = ReadINI("Main", "Label_Storage2", Localizacao_Ficheiro_Lingua)
    Label_Ficheiros(0).Caption = ReadINI("Main", "Label_Management0", Localizacao_Ficheiro_Lingua)
    Label_Ficheiros(1).Caption = ReadINI("Main", "Label_Management1", Localizacao_Ficheiro_Lingua)
    Label_Ficheiros(2).Caption = ReadINI("Main", "Label_Management2", Localizacao_Ficheiro_Lingua)
    
    Idioma_Ver_Capa = ReadINI("Main", "Icon_Cover_View", Localizacao_Ficheiro_Lingua)
    Idioma_Ocultar_Capa = ReadINI("Main", "Icon_Cover_Hide", Localizacao_Ficheiro_Lingua)
    Idioma_Ver_Lista = ReadINI("Main", "Icon_Playlist_View", Localizacao_Ficheiro_Lingua)
    Idioma_Ocultar_Lista = ReadINI("Main", "Icon_Playlist_Hide", Localizacao_Ficheiro_Lingua)
    Icon_Visao(0).ToolTipText = ReadINI("Main", "Icon_Simple", Localizacao_Ficheiro_Lingua)
    Icon_Visao(1).ToolTipText = ReadINI("Main", "Icon_Advanced_Search", Localizacao_Ficheiro_Lingua)
    Icon_Visao(2).ToolTipText = ReadINI("Main", "Icon_Album_Art", Localizacao_Ficheiro_Lingua)
    Idioma_Conectando = ReadINI("Main", "Label_Connecting", Localizacao_Ficheiro_Lingua)
    Idioma_Reproduzindo = ReadINI("Main", "Label_Playing", Localizacao_Ficheiro_Lingua)
    Idioma_Erro = ReadINI("Message", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Message", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Message", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Conectar_Servidor = ReadINI("Message", "Error_Connect", Localizacao_Ficheiro_Lingua)
    Idioma_Internet_Desligada = ReadINI("Message", "Error_Internet", Localizacao_Ficheiro_Lingua)
    Idioma_Mensagem_Enviada = ReadINI("Message", "Info_Posted", Localizacao_Ficheiro_Lingua)
    Icon_Barra_Informacoes(0).ToolTipText = ReadINI("Main", "New_Playlist", Localizacao_Ficheiro_Lingua)
    Icon_Barra_Informacoes(1).ToolTipText = ReadINI("Main", "Music_Randomize", Localizacao_Ficheiro_Lingua)
    Icon_Barra_Informacoes(2).ToolTipText = ReadINI("Main", "Music_Repete", Localizacao_Ficheiro_Lingua)
    Icon_Barra_Informacoes(4).ToolTipText = ReadINI("Main", "Icon_Open", Localizacao_Ficheiro_Lingua)
    Close_Wmp.ToolTipText = ReadINI("Main", "Close_Wmp", Localizacao_Ficheiro_Lingua)
    Label_Topico_Barra_Lateral(3).Caption = ReadINI("Main", "Label_Frame_Cover", Localizacao_Ficheiro_Lingua)
    
    Label_Barra_Drive(3).Caption = ReadINI("Main", "Label_Bar_Contacts", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(4).Caption = ReadINI("Main", "Label_Bar_Events", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(5).Caption = ReadINI("Main", "Label_Bar_Files", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(6).Caption = ReadINI("Main", "Label_Bar_Files", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(7).Caption = ReadINI("Main", "Label_Bar_Recent", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(8).Caption = ReadINI("Main", "Label_Bar_Favorites", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(9).Caption = ReadINI("Main", "Topic_My_Music", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(10).Caption = ReadINI("Main", "Topic_Search", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(11).Caption = ReadINI("Main", "Label_Bar_Community", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(12).Caption = ReadINI("Main", "Label_Bar_My_Friends", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(13).Caption = ReadINI("Main", "Label_Bar_Messages", Localizacao_Ficheiro_Lingua)
    Label_Barra_Drive(14).Caption = ReadINI("Main", "Button_View_Profile", Localizacao_Ficheiro_Lingua)
    
    Idioma_Grid_Community_Col_1 = ReadINI("Main", "Grid_Community_Col_1", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Community_Col_2 = ReadINI("Main", "Grid_Community_Col_2", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Community_Col_3 = ReadINI("Main", "Grid_Community_Col_3", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Community_Col_4 = ReadINI("Main", "Grid_Community_Col_4", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Community_Col_5 = ReadINI("Main", "Grid_Community_Col_5", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Community_Col_6 = ReadINI("Main", "Grid_Community_Col_6", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Community_Col_7 = ReadINI("Main", "Grid_Community_Col_7", Localizacao_Ficheiro_Lingua)
    Idioma_Grid_Community_Col_8 = ReadINI("Main", "Grid_Community_Col_8", Localizacao_Ficheiro_Lingua)
    Label_Mensagens.ToolTipText = ReadINI("Main", "Tolltiptext_Enread_Messages", Localizacao_Ficheiro_Lingua)
    Botao_Mensagens.ToolTipText = ReadINI("Main", "Tolltiptext_Enread_Messages", Localizacao_Ficheiro_Lingua)
    Label_Frame_Programas(1).Caption = ReadINI("Main", "Label_Programs_Installed", Localizacao_Ficheiro_Lingua)
    Label_Frame_Programas(2).Caption = ReadINI("Main", "Label_Programs_Category", Localizacao_Ficheiro_Lingua)
    Idioma_Button_Transfer_Program = ReadINI("Main", "Button_Transfer_Program", Localizacao_Ficheiro_Lingua)
    Idioma_Button_Execute_Program = ReadINI("Main", "Button_Execute_Program", Localizacao_Ficheiro_Lingua)
    Idioma_Button_Remove_Program = ReadINI("Main", "Button_Remove_Program", Localizacao_Ficheiro_Lingua)
    Idioma_Button_Cancel_Program = ReadINI("Main", "Button_Cancel_Program", Localizacao_Ficheiro_Lingua)
    
    Dim botoes_lista As Integer: For botoes_lista = 0 To Botao_Mais_Informacoes.count - 1
        Label_Remover_Transferencia(botao_lista).Caption = Idioma_Button_Remove_Program
        Label_Executar_Programa(botao_lista).Caption = Idioma_Button_Execute_Program
        Label_Mais_Informacoes(botao_lista).Caption = ReadINI("Main", "Button_More_Information", Localizacao_Ficheiro_Lingua)
    Next
    Label_Botao_Frame_Informacoes(0).Caption = Idioma_Button_Transfer_Program
    Label_Botao_Frame_Informacoes(1).Caption = Idioma_Button_Cancel_Program
    Label_Botao_Frame_Informacoes(2).Caption = Idioma_Button_Execute_Program
    Label_Frame_Informacoes(4).Caption = ReadINI("Main", "Label_Program_Official_Website", Localizacao_Ficheiro_Lingua)
    Idioma_Label_Rate = ReadINI("Main", "Label_Rating", Localizacao_Ficheiro_Lingua)
    Label_Titulo_Frame_Programas(1).Caption = ReadINI("Main", "Label_Select_Category", Localizacao_Ficheiro_Lingua)
    
    Ajustar_Objectos_Na_Horizontal
End Sub

Private Sub Verificar_Actualizacoes()
    On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "verificarversao.asp?", False
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            'Ler os dados do xml referente aos dados do perfil do utilizador
            Dim versao_actual, nova_versao As String
            versao_actual = App.Major & App.Minor & App.Revision
            nova_versao = servidor.responseText 'CInt(responseText)
            
            'Verificar se existem versões novas
            If versao_actual < nova_versao Then
                'Indica que há uma nova versão caso a minha versao seja diferente á versão que está no servidor
                Barra_Actualizar.Visible = True
                Desenhar_Formulario

                Botao_Actualizar_Programa.Visible = True
                Label_Nova_Versao.Caption = Idioma_Nova_Versao
                Existe_Nova_Versao = True

            Else
                Label_Nova_Versao.Caption = Idioma_Nao_Existe_Actualizacoes
                Botao_Actualizar_Programa.Visible = False
            End If

        Else
            Label_Nova_Versao.Caption = Idioma_Nao_Existe_Actualizacoes
            Botao_Actualizar_Programa.Visible = False
        End If
    End If
    Set servidor = Nothing
    
    'Ajustar labels welcoome
    Me.MousePointer = 0
    Botao_Actualizar_Programa.left = Label_Nova_Versao.left + Label_Nova_Versao.Width + 20
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        'Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada

    Case -2147417848
        'exit sub

    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    'On Error GoTo Corrige_Erro
    If Me.WindowState = 1 Then Exit Sub
    Ajustar_Formulario Form_Principal, True, True, False, False
    
    'Player------------------------------------------------------------------------------------------------------------
    With Barra_Player
        .Height = Form_Skin.Fundo_Barra_Player.Height
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Width = Me.ScaleWidth - 2
        .left = 1
    End With
    
    With Fundo_Barra_Player
        .Stretch = True
        .top = 0
        .Width = Barra_Player.ScaleWidth
        .left = 0
    End With
    
    With Barra_Botoes
        .Height = Form_Skin.Fundo_Barra_Player.Height
        .top = 0
        .left = 0
    End With
    
    With Fundo_Barra_Botoes
        .Stretch = True
        .top = 0
        .Width = Barra_Player.ScaleWidth
        .left = 0
    End With
    
    Dim Ajustar_Botoes_Player As String
    Ajustar_Botoes_Player = "False" 'ReadINI("Dimensions", "Adjust_Button_Player", Localizacao_Ficheiro_Skin)
    With Botao_Antes
        .top = (Barra_Botoes.ScaleHeight - .Height) / 2
        .left = 16
    End With
    
    With Botao_Play
        .top = (Barra_Botoes.ScaleHeight - .Height) / 2
        If Ajustar_Botoes_Player = True Then
            .left = Botao_Antes.left + Botao_Antes.Width
        Else
            .left = Botao_Antes.left + Botao_Antes.Width + 6
        End If
    End With
    
    With Botao_Pausa
        .top = Botao_Play.top
        .left = Botao_Play.left
    End With
    
    With Botao_Seguinte
        .top = (Barra_Botoes.ScaleHeight - .Height) / 2
        If Ajustar_Botoes_Player = True Then
            .left = Botao_Play.left + Botao_Play.Width
        Else
            .left = Botao_Play.left + Botao_Play.Width + 6
        End If
    End With
    
    With Barra_Faixa
        .Height = Form_Skin.Fundo_Barra_Faixa.Height
        .top = (Barra_Player.ScaleHeight - .Height) / 2
        .Width = Form_Skin.Fundo_Barra_Faixa.Width
        .left = (Barra_Player.ScaleWidth - .ScaleWidth) / 2
        If .left < (Barra_Botoes.left + Barra_Botoes.ScaleWidth) Then
            .left = Barra_Botoes.left + Barra_Botoes.ScaleWidth
        End If
    End With
    
    'Verificar a largura da barra player e ocultar a caixa de pesquisa se assim for necessário
    Dim largura_total As Integer: largura_total = Barra_Player.ScaleWidth - Barra_Botoes.ScaleWidth - Barra_Faixa.ScaleWidth
    Dim largura_objectos_musica As Integer: largura_objectos_musica = Barra_Caixa_Pesquisar_Musica.ScaleWidth + (3 * Icon_Visao(0).Width) + 22
    If largura_total >= largura_objectos_musica Then 'Tem tamnho suficiente para se ver a caixa de pesquisa
        Barra_Caixa_Pesquisar_Musica.Visible = True
        Icon_Visao(0).Visible = True
        Icon_Visao(1).Visible = True
        Icon_Visao(2).Visible = True
    Else
        Barra_Caixa_Pesquisar_Musica.Visible = False 'se não tiver oculta-a
        Icon_Visao(0).Visible = False
        Icon_Visao(1).Visible = False
        Icon_Visao(2).Visible = False
    End If
    
    With Label_Faixa
        .top = 8
        .Width = Barra_Faixa.Width - 150
        .left = 10
    End With
    
    With Picture_Slide_Som
        .Height = Form_Skin.Fundo_Slider_Volume.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Fundo_Slider_Volume.Width
    End With
    
    With Slide_Som
        .Height = Form_Skin.Slide_Som_Normal.Height
        .top = (Picture_Slide_Som.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Slide_Som_Normal.Width
    End With
    
    With Botao_Mudo
        .top = (Barra_Botoes.ScaleHeight - .Height) / 2
    End With
    
    With SliderBar
        .Height = Form_Skin.Image_Barra_Slide.Height
        .top = Barra_Faixa.ScaleHeight - .ScaleHeight - 10
        .Width = Form_Skin.Image_Barra_Slide.Width
        .left = (Barra_Faixa.ScaleWidth - .ScaleWidth) / 2
    End With
    
    With Slide
        .Height = Form_Skin.Slide_Musica_Normal.Height
        .top = (SliderBar.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Slide_Musica_Normal.Width
    End With
    
    With Image_Progresso
        .Height = Form_Skin.Slide_Musica_Normal.Height
        .top = (SliderBar.ScaleHeight - .ScaleHeight) / 2
    End With
    
    With Tempo_Estimado
        .top = Barra_Faixa.ScaleHeight - .Height - 10
        .left = 10
    End With
    
    With Label_Duracao
        .top = Tempo_Estimado.top
        .left = Barra_Faixa.ScaleWidth - .Width - 10
    End With
    
    '--------------------------------------------------------------------------------------------------------
    
    'Barra lateral-------------------------------------------------------------------------------------------
    With Barra_Lateral
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - Barra_Player.ScaleHeight - Barra_Informacoes.ScaleHeight
        .top = Barra_Player.top + Barra_Player.ScaleHeight
        .Width = Form_Skin.Select_Topic_TaskBar.Width 'Fundo_Barra_Lateral.Width
        .left = 1
    End With
    
    With Linha_Vertical
        .top = Barra_Lateral.top
        .Height = Barra_Lateral.Height
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth
    End With
    
    With Frame_Separadores
        .Height = Barra_Lateral.ScaleHeight - Frame_Capa.ScaleHeight - Barra_Conexao.ScaleHeight
        .top = 0
        .Width = Barra_Lateral.ScaleWidth
        .left = 0
    End With
    
    With Frame_Capa
        .Height = Form_Skin.Bar_View_Cover.Height + Form_Skin.Image_Sem_Capa.Height
        .top = Barra_Lateral.ScaleHeight - .ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height
        .Width = Barra_Lateral.ScaleWidth
        .left = 0
    End With
    
    'Chamar procedimentos
    Ajustar_Separadores_Barra_Lateral
    Ajustar_Objectos_Na_Horizontal
    Ajustar_Objectos_Na_Vertical
    
    With Barra_Caixa_Pesquisar_Musica
        .Height = Form_Skin.Caixa_Pesquisar_Musica.Height
        .top = (Barra_Player.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Caixa_Pesquisar_Musica.Width
        .left = Barra_Player.ScaleWidth - .ScaleWidth - 10
    End With
    
    With Text_Pesquisar_Musica
        .Height = Barra_Caixa_Pesquisar_Musica.ScaleHeight - 8 - 8
        .top = (Barra_Caixa_Pesquisar_Musica.ScaleHeight - .Height) / 2
        .Width = Barra_Caixa_Pesquisar_Musica.ScaleWidth - 8 - 8 - 20
        .left = 8
    End With
    
    With Icon_Visao(2)
        .top = (Barra_Player.ScaleHeight - .Height) / 2
        .left = Barra_Caixa_Pesquisar_Musica.left - .Width - 16
    End With
    
    With Icon_Visao(1)
        .top = Icon_Visao(2).top
        .left = Icon_Visao(2).left - .Width
    End With
    
    With Icon_Visao(0)
        .top = Icon_Visao(2).top
        .left = Icon_Visao(1).left - .Width
    End With
    
    'Barra de actualizações---------------------------------------------------------------------------------------------------------
    With Barra_Actualizar
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Linha_Barra_Actualizar
        .Width = Barra_Actualizar.ScaleWidth
        .left = 0
    End With
    
    With Close_Barra_Actualizar
        .top = (Barra_Actualizar.ScaleHeight - .Height) / 2
        .left = Barra_Actualizar.ScaleWidth - .Width - .top
    End With
    
    With Label_Nova_Versao
        .top = (Barra_Actualizar.ScaleHeight - .Height) / 2
        .left = .top
    End With
    
    With Botao_Actualizar_Programa
        .Height = Form_Skin.Botao_Actualizar_Programa.Height
        .top = (Barra_Actualizar.ScaleHeight - .Height) / 2
        .Width = Form_Skin.Botao_Actualizar_Programa.Width
        .left = Label_Nova_Versao.left + Label_Nova_Versao.Width + 20
    End With
    
    With Label_Actualizar_Programa
        .top = (Botao_Actualizar_Programa.ScaleHeight - .Height) / 2
        .left = 0
        .Width = Botao_Actualizar_Programa.ScaleWidth
    End With
    
    'Album---------------------------------------------------------------------------------------------------------
    With Frame_Album
        If Form_Preferencias.Check_Ver_Playlist.Value = 1 Then
            .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3 - Form_Skin.Fundo_Barra_Lateral.Width
        Else
            .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        End If
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Label_Album
        .top = Frame_Slide_Album.top + Frame_Slide_Album.ScaleHeight + 20
        .Width = Frame_Album.ScaleWidth
        .left = 0
    End With
    
    With Barra_Slider_Album
        .top = Frame_Album.ScaleHeight - .ScaleHeight - 10
        .Height = Fundo_Barra_Slider_Album_Esq.Height
        .left = 10
        .Width = Frame_Album.ScaleWidth - 20
    End With
    
    With Fundo_Barra_Slider_Album_Esq
        .top = 0
        .left = 0
    End With
    
    With Fundo_Barra_Slider_Album_Dir
        .top = 0
        .left = Barra_Slider_Album.ScaleWidth - .Width
    End With
    
    With Barra_Slider_Album_Center
        .Height = Barra_Slider_Album.Height
        .top = 0
        .Width = Barra_Slider_Album.ScaleWidth - Fundo_Barra_Slider_Album_Esq.Width - Fundo_Barra_Slider_Album_Dir.Width
        .left = Fundo_Barra_Slider_Album_Esq.left + Fundo_Barra_Slider_Album_Esq.Width
    End With
    
    With Fundo_Barra_Slider_Album_Center
        .Stretch = True
        .top = 0
        .Width = Barra_Slider_Album_Center.ScaleWidth
        .left = 0
    End With
    
    With Slide_Album
        .Height = Form_Skin.Slide_Album.Height
        .top = 1
        .Width = Form_Skin.Slide_Album.Width
    End With
    
    With Frame_Slide_Album
        .top = 20
        .Height = Form_Skin.Image_Album.Height
        .left = 10
        .Width = Frame_Album.ScaleWidth - 20
    End With
    
    With Frame_Slide
        .top = 0
        .Height = Form_Skin.Image_Album.Height
    End With
    
    With Image_Album(0)
        .top = 0
        .Height = Form_Skin.Image_Album.Height
        .left = 0
        .Width = Form_Skin.Image_Album.Width
    End With
    
    With Linha_Frame_Album
        .top = Frame_Album.ScaleHeight - .Height
        .left = 0
        .Width = Frame_Album.ScaleWidth
    End With
    
    With Label_Nenhum_Album
        .top = (Frame_Album.ScaleHeight - .Height) / 2
        .left = (Frame_Album.ScaleWidth - .Width) / 2
    End With
    
    '------------------------------------------------------------------------------------------------
    With Pic_Capa_Album
        .Height = Form_Skin.Image_Sem_Capa.Height
        .top = Separador_Barra_Lateral(3).top + Separador_Barra_Lateral(3).ScaleHeight
        .Width = Frame_Capa.ScaleWidth
        .left = 0
    End With
    
    'Frame lista de reprodução-------------------------------------------------------------------------------------------------
    With Barra_Playlist
        .Width = Form_Skin.Fundo_Barra_Lateral.Width
        .left = Frame_Album.left + Frame_Album.ScaleWidth
    End With
    
    With Label_Carregar_Favoritos
        .top = ((Barra_Playlist.ScaleHeight - .Height) / 2)
        .left = ((Barra_Playlist.ScaleWidth - .Width) / 2)
    End With
    
    With Grelha_Lista_Em_Reproducao
        .Width = Barra_Playlist.Width - 1
        .left = 1
    End With
    
    With Linha_Barra_Playlist
        .left = 0
    End With
    
    '----------------------------------------------------------------------------------------------
    With Frame_Wmp
        .Height = Barra_Lateral.ScaleHeight
        .top = Barra_Lateral.top
        .Width = Me.ScaleWidth - 2
        .left = 1
    End With
    
    With Wmp
        .Height = Frame_Wmp.ScaleHeight
        .top = 0
        .Width = Frame_Wmp.ScaleWidth
        .left = 0
    End With
    
    With Close_Wmp
        .Height = Form_Skin.Close_Wmp.Height
        .top = Frame_Wmp.top '+ 10
        .Width = Form_Skin.Close_Wmp.Width
        .left = Me.ScaleWidth - .ScaleWidth - 10
    End With
    
    With Barra_Mini_Player
        .Height = Form_Skin.Fundo_Barra_Mini_Player.Height
        .top = Frame_Wmp.top + Frame_Wmp.ScaleHeight - .ScaleHeight - 50
        .Width = Form_Skin.Fundo_Barra_Mini_Player.Width
        .left = (Frame_Wmp.ScaleWidth - .ScaleWidth) / 2
    End With
    
    With Picture_Slide_Som_Mini
        .Height = Form_Skin.Picture_Slide_Som_Mini.Height
        .Width = Form_Skin.Picture_Slide_Som_Mini.Width
    End With
    
    With Slide_Som_Mini
        .Height = Form_Skin.Slide_Som_Mini.Height
        .Width = Form_Skin.Slide_Som_Mini.Width
    End With
    
    With SliderBar_Mini
        .Height = Form_Skin.SliderBar_Mini.Height
        .Width = Form_Skin.SliderBar_Mini.Width
    End With
    
    With Slide_Mini
        .Height = Form_Skin.Slide_Mini.Height
        .Width = Form_Skin.Slide_Mini.Width
    End With
    
    'Grelhas do centro-------------------------------------------------------------------------------
    With Grelha_Musica
        If Form_Preferencias.Check_Ver_Playlist.Value = 1 Then
            .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3 - Form_Skin.Fundo_Barra_Lateral.Width
        Else
            .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        End If
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Grelha_Filmes
        .Width = Grelha_Musica.Width
        .left = Grelha_Musica.left
    End With
    
    With Grelha_Radio
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    'Seviços-----------------------------------------------------------------------------
    With Barra_Drive
        .Stretch = True
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Botao_Barra_Drive(0)
        .top = Barra_Drive.top
        .left = Barra_Drive.left
    End With
    
    With Botao_Barra_Drive(1)
        .top = Barra_Drive.top
        .left = Botao_Barra_Drive(0).left + Botao_Barra_Drive(0).Width
    End With
    
    With Botao_Barra_Drive(2)
        .top = Barra_Drive.top
        .left = Barra_Drive.left 'Botao_Barra_Drive(1).left + Botao_Barra_Drive(1).Width + 80
    End With
    
    Dim h As Integer: For h = 3 To 6
        Botao_Barra_Drive(h).Stretch = True
        Botao_Barra_Drive(h).Width = Label_Barra_Drive(h).Width + 40
        Botao_Barra_Drive(h).top = Barra_Drive.top
        Botao_Barra_Drive(h).left = Botao_Barra_Drive(h - 1).left + Botao_Barra_Drive(h - 1).Width
    Next
    
    Dim k As Integer: For k = 7 To 14
        Botao_Barra_Drive(k).Stretch = True
        Botao_Barra_Drive(k).Width = Label_Barra_Drive(k).Width + 40
        Botao_Barra_Drive(k).top = Barra_Drive.top
    Next
    Botao_Barra_Drive(7).left = Botao_Barra_Drive(2).left + Botao_Barra_Drive(2).Width
    Botao_Barra_Drive(8).left = Botao_Barra_Drive(7).left + Botao_Barra_Drive(7).Width
    Botao_Barra_Drive(9).left = Botao_Barra_Drive(8).left + Botao_Barra_Drive(8).Width
    Botao_Barra_Drive(10).left = Botao_Barra_Drive(9).left + Botao_Barra_Drive(9).Width
    Botao_Barra_Drive(11).left = Botao_Barra_Drive(10).left + Botao_Barra_Drive(10).Width
    Botao_Barra_Drive(12).left = Botao_Barra_Drive(11).left + Botao_Barra_Drive(11).Width
    Botao_Barra_Drive(13).left = Botao_Barra_Drive(12).left + Botao_Barra_Drive(12).Width
    Botao_Barra_Drive(14).left = Botao_Barra_Drive(13).left + Botao_Barra_Drive(13).Width
    
    Dim g As Integer: For g = 3 To 6
        Label_Barra_Drive(g).top = Barra_Drive.top + ((Barra_Drive.Height - Label_Barra_Drive(g).Height) / 2)
        Label_Barra_Drive(g).left = Botao_Barra_Drive(g).left + 20
    Next
    
    Dim m As Integer: For m = 7 To 14
        Label_Barra_Drive(m).top = Barra_Drive.top + ((Barra_Drive.Height - Label_Barra_Drive(m).Height) / 2)
        Label_Barra_Drive(m).left = Botao_Barra_Drive(m).left + 20
    Next
    
    With Frame_My_Drive
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Shape_Sky(0)
        .top = 0
        .Height = Imagem_Nuvens.Height
        .left = 0
        .Width = Frame_My_Drive.ScaleWidth
    End With
    
    With Shape_Sky(1)
        .top = 0
        .Height = Imagem_Nuvens.Height
        .left = 0
        .Width = Frame_My_Drive.ScaleWidth
    End With
    
    With Frame_Home
        .top = 0
        .Width = (3 * Form_Skin.Image_Precos.Width) + (4 * 30)
        If Frame_My_Drive.Width >= .Width Then
            .left = (Frame_My_Drive.ScaleWidth - .ScaleWidth) / 2
        Else
            .left = 0
        End If
    End With
    
    Dim tabela_precos As Integer: For tabela_precos = 0 To Picture_Tabela.count - 1
        Picture_Tabela(tabela_precos).Height = Form_Skin.Image_Precos.Height
        Picture_Tabela(tabela_precos).Width = Form_Skin.Image_Precos.Width
        
        Label_Plano(tabela_precos).left = 10
        
        Label_Preco(tabela_precos).left = (Picture_Tabela(tabela_precos).ScaleWidth - Label_Preco(tabela_precos).Width) / 2
        
        Label_Mensalidade(tabela_precos).left = (Picture_Tabela(tabela_precos).ScaleWidth - Label_Mensalidade(tabela_precos).Width) / 2
    Next
    Picture_Tabela(0).left = 30
    Picture_Tabela(1).left = Picture_Tabela(0).left + Picture_Tabela(0).ScaleWidth + 30
    Picture_Tabela(2).left = Picture_Tabela(1).left + Picture_Tabela(1).ScaleWidth + 30
    
    With Label_Texto(3)
        .left = (Frame_Home.ScaleWidth - .Width) / 2
    End With
    
    With Grelha_Contactos
        If Barra_Actualizar.Visible = True Then
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Actualizar.ScaleHeight - Barra_Drive.Height
        Else
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Drive.Height
        End If
        .top = Barra_Drive.top + Barra_Drive.Height
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Grelha_Eventos
        .top = Grelha_Contactos.top
        .Height = Grelha_Contactos.Height
        .left = Grelha_Contactos.left
        .Width = Grelha_Contactos.Width
    End With
    
    With Grelha_Mensagens
        .top = Grelha_Contactos.top
        .Height = Grelha_Contactos.Height
        .left = Grelha_Contactos.left
        .Width = Grelha_Contactos.Width
    End With
     
    With Grelha_Ficheiros
        .top = Grelha_Contactos.top
        .Height = Grelha_Contactos.Height
        .left = Grelha_Contactos.left
        .Width = Grelha_Contactos.Width
    End With
    
    With Grelha_Recentes
        .top = Grelha_Contactos.top
        .Height = Grelha_Contactos.Height
        .left = Grelha_Contactos.left
        .Width = Grelha_Contactos.Width
    End With
    
    With Grelha_Favoritos
        .top = Grelha_Contactos.top
        .Height = Grelha_Contactos.Height
        .left = Grelha_Contactos.left
        .Width = Grelha_Contactos.Width
    End With
    
    With Grelha_Comunidade
        .top = Grelha_Contactos.top
        .Height = Grelha_Contactos.Height
        .left = Grelha_Contactos.left
        .Width = Grelha_Contactos.Width
    End With
    
    With Grelha_Amigos
        .top = Grelha_Contactos.top
        .Height = Grelha_Contactos.Height
        .left = Grelha_Contactos.left
        .Width = Grelha_Contactos.Width
    End With
    
    With Frame_Perfil
        .top = Grelha_Contactos.top
        .Height = Grelha_Contactos.Height
        .left = Grelha_Contactos.left
        .Width = Grelha_Contactos.Width
    End With
    
    With Grelha_Loja
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Grelha_Minha_Musica
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    'Frame bem vindo
    With Frame_Music_Link
        If Barra_Actualizar.Visible = True Then
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Actualizar.ScaleHeight - Barra_Drive.Height
        Else
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Drive.Height
        End If
        .top = Barra_Drive.top + Barra_Drive.Height
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Fundo_Frame_Music_Link
        .Stretch = True
        .top = 0
        .left = 0
        .Height = Frame_Music_Link.ScaleHeight
        .Width = Frame_Music_Link.ScaleWidth
    End With
    
    With Frame_Caixa_Pesquisa
        .Height = Form_Skin.Fundo_Frame_Caixa_Pesquisa.Height
        .top = (Frame_Music_Link.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Fundo_Frame_Caixa_Pesquisa.Width
        .left = (Frame_Music_Link.ScaleWidth - .ScaleWidth) / 2
    End With
    
    With Label_Texto(5)
        .top = Frame_Caixa_Pesquisa.top - .Height - 20
        .left = (Frame_Music_Link.ScaleWidth - .Width) / 2
    End With
    
    With Label_Texto(4)
        .top = Label_Texto(5).top - .Height - 5
        .left = (Frame_Music_Link.ScaleWidth - .Width) / 2
    End With
    
    With Botao_Pesquisar
        .top = (Frame_Caixa_Pesquisa.ScaleHeight - .Height) / 2
        .left = Frame_Caixa_Pesquisa.ScaleWidth - .Width - .top
    End With
    
    With Text_Pesquisar
        .top = (Frame_Caixa_Pesquisa.ScaleHeight - .Height) / 2
        .left = 15
        .Width = Frame_Caixa_Pesquisa.ScaleWidth - Botao_Pesquisar.Width - Botao_Pesquisar.top - 30
    End With
    
    With Separador_Barra_Lateral(3)
        .Height = Form_Skin.Bar_View_Cover.Height
        .top = 0
        .Width = Barra_Lateral.ScaleWidth
        .left = 0
    End With
    
    With Label_Topico_Barra_Lateral(3)
        .top = (Separador_Barra_Lateral(3).ScaleHeight - .Height) / 2
        .Width = Separador_Barra_Lateral(3).ScaleWidth
        .left = 0
    End With
    
    'Grelhas de pesquisa personalizada na biblioteca
    With Frame_Grelhas_Pesquisa
        .Height = Frame_Album.ScaleHeight - Linha_Frame_Album.Height
        .top = 0
        .Width = Frame_Album.ScaleWidth
        .left = 0
    End With
    
    Dim Largura_Frame_Pesquisa_Personalizada As Integer
    Largura_Frame_Pesquisa_Personalizada = (Frame_Grelhas_Pesquisa.ScaleWidth / 3)
    With Grelha_Artista
        .Height = Frame_Grelhas_Pesquisa.ScaleHeight
        .Width = Largura_Frame_Pesquisa_Personalizada
        .left = 0
    End With
    With Grelha_Album
        .Height = Grelha_Artista.Height
        .Width = Largura_Frame_Pesquisa_Personalizada
        .left = Grelha_Artista.left + Grelha_Artista.Width + 1
    End With
    
    With Grelha_Genero
        .Height = Grelha_Artista.Height
        .Width = Frame_Grelhas_Pesquisa.ScaleWidth - Grelha_Artista.Width - 1 - Grelha_Album.Width - 1
        .left = Grelha_Album.left + Grelha_Album.Width + 1
    End With
    
    With Image_Lupa
        .top = (Barra_Caixa_Pesquisar_Musica.ScaleHeight - .Height) / 2
        .left = Barra_Caixa_Pesquisar_Musica.ScaleWidth - .Width - .top
    End With
    
    With Grelha_Listas
        .Width = Grelha_Radio.Width
        .left = Grelha_Radio.left
    End With
    
    With Botao_Redimensionar
        .top = Barra_Informacoes.ScaleHeight - .Height
        .left = Barra_Informacoes.ScaleWidth - .Width
    End With
    
    With Frame_Evento
        .top = Barra_Informacoes.top - .ScaleHeight
        .left = Barra_Informacoes.left + Barra_Informacoes.ScaleWidth - .ScaleWidth
    End With
    
    With Shape_Evento
        .top = 0
        .left = 0
        .Width = Frame_Evento.Width
    End With
    
    'Frame para procurar novos programas
    With Frame_Programas
        If Barra_Actualizar.Visible = True Then
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Actualizar.ScaleHeight
        Else
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height
        End If
        .top = Barra_Lateral.top
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Barra_Top_Frame_Programas
        .top = 0
        .Stretch = True
        .left = 0
        .Width = Frame_Programas.ScaleWidth
    End With
    
    Separador_Frame_Programas(0).top = 0
    Separador_Frame_Programas(0).left = 0
    Dim tab_programas As Integer: For tab_programas = 1 To Separador_Frame_Programas.count - 1
        Separador_Frame_Programas(tab_programas).top = 0
        Separador_Frame_Programas(tab_programas).Stretch = True
        Separador_Frame_Programas(tab_programas).Width = Label_Frame_Programas(tab_programas).Width + 40
        Separador_Frame_Programas(tab_programas).left = Separador_Frame_Programas(tab_programas - 1).left + Separador_Frame_Programas(tab_programas - 1).Width
        Label_Frame_Programas(tab_programas).top = Barra_Top_Frame_Programas.top + ((Barra_Top_Frame_Programas.Height - Label_Frame_Programas(tab_programas).Height) / 2)
        Label_Frame_Programas(tab_programas).left = Separador_Frame_Programas(tab_programas).left + 20
    Next
    
    With Frame_Programas_Home
        If Frame_Programas.Height >= .Height Then
            .top = (Frame_Programas.ScaleHeight - .Height) / 2
        Else
            .top = Barra_Top_Frame_Programas.top + Barra_Top_Frame_Programas.Height
        End If
        .Width = ((Icon_Pasta_Categoria.count) * Icon_Pasta_Categoria(0).Width) + (40 * (Icon_Pasta_Categoria.count + 1))
        If Frame_Programas.Width >= .Width Then
            .left = (Frame_Programas.ScaleWidth - .ScaleWidth) / 2
        Else
            .left = 0
        End If
    End With
    
    With Label_Titulo_Frame_Programas(0)
        .top = 0
        .left = (Frame_Programas_Home.ScaleWidth - .Width) / 2
    End With
    
    With Label_Titulo_Frame_Programas(1)
        .top = Label_Titulo_Frame_Programas(0).top + Label_Titulo_Frame_Programas(0).Height + 5
        .left = (Frame_Programas_Home.ScaleWidth - .Width) / 2
    End With
    
    Icon_Pasta_Categoria(0).top = Label_Titulo_Frame_Programas(1).top + Label_Titulo_Frame_Programas(1).Height + 30
    Icon_Pasta_Categoria(0).left = 40
    Dim pastas As Integer: For pastas = 1 To Icon_Pasta_Categoria.count - 1
        Icon_Pasta_Categoria(pastas).top = Icon_Pasta_Categoria(0).top
        Icon_Pasta_Categoria(pastas).left = Icon_Pasta_Categoria(pastas - 1).left + Icon_Pasta_Categoria(pastas - 1).Width + 40
    Next
    
    With Label_Titulo_Frame_Programas(2)
        .top = Icon_Pasta_Categoria(0).top + Icon_Pasta_Categoria(0).Height + 10
        .left = 0
        .Width = Frame_Programas_Home.ScaleWidth
    End With
    
    'Resultado da pesquisa por categoria
    With Frame_Lista
        .Height = Frame_Programas.ScaleHeight - Barra_Top_Frame_Programas.Height
        .top = Barra_Top_Frame_Programas.top + Barra_Top_Frame_Programas.Height
        .left = 0
        .Width = Frame_Programas.ScaleWidth
    End With
    
    'Linhas da lista de programas
    Dim i As Integer
    For i = 0 To Pic_Linha.count - 1
        Pic_Linha(i).Width = Frame_Lista.ScaleWidth
        Pic_Linha(i).Height = Form_Skin.Linha_Normal.Height
    Next i
    With Label_Nenum_Resultado
        .top = 16
        .left = 16
    End With
    
    Dim j As Integer
    For j = 0 To Pic_Linha.count - 1
        With Botao_Mais_Informacoes(j)
            .Height = Form_Skin.Botao_Linha_Normal.Height
            .Width = Form_Skin.Botao_Linha_Normal.Width
            .Visible = False
        End With
        
        With Label_Mais_Informacoes(j)
            .top = (Botao_Mais_Informacoes(Index).ScaleHeight - .Height) / 2
            .Width = Botao_Mais_Informacoes(Index).ScaleWidth
            .left = 0
        End With
        
        With Botao_Remover_Transferencia(j)
            .top = Botao_Mais_Informacoes(j).top
            .Height = Form_Skin.Botao_Linha_2_Normal.Height
            .Width = Form_Skin.Botao_Linha_2_Normal.Width
            .left = Pic_Linha(j).ScaleWidth - Botao_Remover_Transferencia(j).ScaleWidth - 8
            .Visible = False
        End With
        
        With Label_Remover_Transferencia(j)
            .top = (Botao_Remover_Transferencia(Index).ScaleHeight - .Height) / 2
            .Width = Botao_Remover_Transferencia(Index).ScaleWidth
            .left = 0
        End With
        
        With Botao_Executar_Programa(j)
            .top = Botao_Mais_Informacoes(j).top
            .Height = Form_Skin.Botao_Linha_2_Normal.Height
            .Width = Form_Skin.Botao_Linha_2_Normal.Width
            .left = Botao_Remover_Transferencia(j).left - .ScaleWidth - 8
            .Visible = False
        End With
        
        With Label_Executar_Programa(j)
            .top = (Botao_Executar_Programa(Index).ScaleHeight - .Height) / 2
            .Width = Botao_Executar_Programa(Index).ScaleWidth
            .left = 0
        End With
    Next j
    
    'Ajustar as progressbars
    Dim xpto As Integer: For xpto = 0 To Pic_Linha.count - 1
        With Progresso(xpto)
            .Height = Form_Skin.Botao_Linha_2_Normal.Height
            .top = Botao_Remover_Transferencia(xpto).top
            .Width = Form_Skin.Botao_Linha_2_Normal.Width
            .left = Botao_Remover_Transferencia(xpto).left
        End With
    Next
    
    With Frame_Informacoes
        .Height = Frame_Programas.ScaleHeight - Barra_Top_Frame_Programas.Height
        .top = Barra_Top_Frame_Programas.top + Barra_Top_Frame_Programas.Height
        .left = 0
        .Width = Frame_Programas.ScaleWidth
    End With
    
    With Image_Logo
        .top = 40
        .left = 40
    End With
    
    With Label_Frame_Informacoes(0)
        .top = Image_Logo.top + 10
        .left = Image_Logo.left + Image_Logo.Width + 15
    End With
    
    With Label_Frame_Informacoes(1)
        .top = Label_Frame_Informacoes(0).top + Label_Frame_Informacoes(0).Height + 3
        .left = Label_Frame_Informacoes(0).left
    End With
    
    With Label_Frame_Informacoes(5)
        .top = Label_Frame_Informacoes(0).top + 5
        .left = Label_Frame_Informacoes(0).left + Label_Frame_Informacoes(0).Width + 5
    End With
    
    With Frame_Avaliacao
        .top = Image_Logo.top
        .Width = Form_Skin.Image_Estrelas_0.Width
        .left = Frame_Informacoes.ScaleWidth - .Width - 20
    End With
    Label_Frame_Informacoes(2).Width = Frame_Avaliacao.ScaleWidth
        
    With Shape_Transferir
        .Height = Form_Skin.Botao_Programas.Height + 10
        .left = 20
        .Width = (Frame_Informacoes.ScaleWidth - 40)
    End With
    
    With Label_Frame_Informacoes(3)
        .top = Shape_Transferir.top + ((Shape_Transferir.Height - .Height) / 2)
        .left = Shape_Transferir.left + (.top - Shape_Transferir.top)
    End With
    
    With Botao_Frame_Informacoes(0)
        .Height = Form_Skin.Botao_Programas.Height
        .top = Shape_Transferir.top + ((Shape_Transferir.Height - .Height) / 2)
        .Width = Form_Skin.Botao_Programas.Width
        .left = Shape_Transferir.left + Shape_Transferir.Width - .Width - (.top - Shape_Transferir.top)
    End With
    
    With Label_Botao_Frame_Informacoes(0)
        .top = (Botao_Frame_Informacoes(0).ScaleHeight - .Height) / 2
        .Width = Botao_Frame_Informacoes(0).Width
        .left = 0
    End With
    
    With ProgressBar1
        .Height = Botao_Frame_Informacoes(0).ScaleHeight
        .top = Botao_Frame_Informacoes(0).top
        .Width = Botao_Frame_Informacoes(0).ScaleWidth
        .left = Botao_Frame_Informacoes(0).left
    End With
    
    With Text_Informacao
        .top = Shape_Transferir.top + Shape_Transferir.Height + 40
        .left = Shape_Transferir.left
        .Width = Shape_Transferir.Width - Form_Skin.Foto_Programa.Width - 60
    End With
        
    With Shape_Foto
        .Height = Form_Skin.Foto_Programa.Height + 4
        .top = Text_Informacao.top
        .Width = Form_Skin.Foto_Programa.Width + 4
        .left = Shape_Transferir.left + Shape_Transferir.Width - .Width
    End With
    
    With Image_Tela
        .Height = Form_Skin.Foto_Programa.Height
        .top = Shape_Foto.top + 2
        .Width = Form_Skin.Foto_Programa.Width
        .left = Shape_Foto.left + 2
    End With
    
    With Label_Frame_Informacoes(4)
        .top = Text_Informacao.top + Text_Informacao.Height + 10
        .left = Shape_Transferir.left
    End With
    
    With Shape_Estado
        .Height = Shape_Transferir.Height
        .top = Frame_Informacoes.ScaleHeight - .Height - 20
        .Width = Shape_Transferir.Width
        .left = Shape_Transferir.left
    End With
    
    With Botao_Frame_Informacoes(2)
        .Height = Form_Skin.Botao_Programas.Height
        .top = Shape_Estado.top + ((Shape_Estado.Height - .Height) / 2)
        .Width = Form_Skin.Botao_Programas.Width
        .left = Botao_Frame_Informacoes(0).left
    End With
    
    With Label_Botao_Frame_Informacoes(2)
        .top = (Botao_Frame_Informacoes(2).ScaleHeight - .Height) / 2
        .Width = Botao_Frame_Informacoes(2).ScaleWidth
        .left = 0
    End With
    
    With Botao_Frame_Informacoes(1)
        .Height = Botao_Frame_Informacoes(2).ScaleHeight
        .top = Botao_Frame_Informacoes(2).top
        .Width = Botao_Frame_Informacoes(2).ScaleWidth
        .left = Botao_Frame_Informacoes(2).left - .ScaleWidth - 8
    End With
    
    With Label_Botao_Frame_Informacoes(1)
        .top = (Botao_Frame_Informacoes(1).ScaleHeight - .Height) / 2
        .Width = Botao_Frame_Informacoes(1).ScaleWidth
        .left = 0
    End With
    
    With Image_Download
        .top = Shape_Estado.top + ((Shape_Estado.Height - .Height) / 2)
        .left = Shape_Estado.left + 10
    End With
    
    With Label_Frame_Informacoes(6)
        .top = Shape_Estado.top + ((Shape_Estado.Height - .Height) / 2)
        .left = Image_Download.left + Image_Download.Width + 5
    End With
    
    
Exit Sub
Corrige_Erro:
End Sub

Private Sub Botao_Restaurar_Click()
    'Restaurar janela
    With Me
        .Height = 9000
        .Width = 13200
        .top = (Screen.Height - Me.Height) / 2
        .left = (Screen.Width - Me.Width) / 2
    End With
    Tela_Cheia = False
    Form_Preferencias.Text_Tela_Cheia.Text = "False"
    Botao_Maximizar.Visible = True
    Botao_Restaurar.Visible = False
    
    'Chamar o procedimento
'    Ajustar_Objectos_Na_Horizontal
'    Ajustar_Objectos_Na_Vertical
End Sub

Private Sub Botao_Maximizar_Click()
    'Maximixar formulário
    On Error Resume Next
    PosFormRelativeTaskBar Me
    Tela_Cheia = True
    Form_Preferencias.Text_Tela_Cheia.Text = "True"
    Botao_Maximizar.Visible = False
    Botao_Restaurar.Visible = True
    
    'Chamar o procedimento
'    Ajustar_Objectos_Na_Horizontal
'    Ajustar_Objectos_Na_Vertical
End Sub

Private Sub Botao_Minimizar_Click()
    'Minimizar o formulário
    If Form_Preferencias.Check_Tray.Value = 1 Then
        Botao_Tray_Click
    Else
        Me.WindowState = 1
    End If
End Sub

Private Sub Form_Resize()
    'Ajustar os objectos ao formulário
    If Me.WindowState = 1 Then Exit Sub
    Desenhar_Formulario
End Sub

Private Sub Frame_Grelha_Filmes_Click()
    Ocultar_menus
End Sub

Private Sub Frame_Grelha_Loja_Click()
    Ocultar_menus
End Sub

Private Sub Frame_Album_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Capa_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Capa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Frame_Grelhas_Pesquisa_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Home_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Home_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Frame_Music_Link_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Music_Link_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Frame_My_Drive_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Programas_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_My_Drive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Frame_Programas_Home_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Repor_Objectos
End Sub

Private Sub Frame_Programas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Repor_Objectos
End Sub

Private Sub Frame_Separador_Barra_Lateral_Click(Index As Integer)
    'Chamar procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Separador_Barra_Lateral_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Frame_Separadores_Click()
    'Chamar procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Separadores_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Frame_Slide_Album_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Slide_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Temas_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Frame_Wmp_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Fundo_Barra_Slider_Album_Dir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Ocultar_menus
    
    'Mover a frame consequntemente
    movimento = "right"
    Timer_Mover.Enabled = True
End Sub

Private Sub Fundo_Barra_Slider_Album_Dir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Para de mover a frame album
    Timer_Mover.Enabled = False
End Sub

Private Sub Fundo_Barra_Slider_Album_Esq_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Ocultar_menus
    
    'Mover a frame consequntemente
    If Frame_Slide.left >= 0 Then Exit Sub: Frame_Slide.left = 0: Timer_Mover.Enabled = False: Timer_Album.Enabled = False
    movimento = "true"
    Timer_Mover.Enabled = True
End Sub

Private Sub Fundo_Barra_Slider_Album_Esq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover a frame consequntemente
    Timer_Mover.Enabled = False
End Sub

Private Sub Grelha_Eventos_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Eventos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_Eventos
    Repor_Objectos
End Sub

Private Sub Grelha_Album_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Album_DblClick()
    'Carregar as músicas consoante a categoria selecionado
    Criterio = Grelha_Album.TextMatrix(Grelha_Album.Row, 1)
    Verifica_Rs_Musica
    Rs_Musica.Open "select * from Tabela_Musica where Album = '" & Criterio & "' order by Album Asc", Cnn_Biblioteca
    Carregar_Grelha_Musica
End Sub

Private Sub Grelha_Artista_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Artista_DblClick()
    'Carregar as músicas consoante a categoria selecionado
    Criterio = Grelha_Artista.TextMatrix(Grelha_Artista.Row, 1)
    Verifica_Rs_Musica
    Rs_Musica.Open "select * from Tabela_Musica where Artista = '" & Criterio & "' order by Artista Asc", Cnn_Biblioteca
    Carregar_Grelha_Musica
End Sub

Private Sub Grelha_Comunidade_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Comunidade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_comunidade
    Repor_Objectos
End Sub

Private Sub Grelha_Amigos_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Amigos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_Amigos
    Repor_Objectos
End Sub

Private Sub Grelha_Contactos_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Contactos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_contactos
    Repor_Objectos
End Sub

Private Sub Grelha_Favoritos_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Favoritos_DblClick()
    'Reproduzir o ficheiro selecionado
    If Grelha_Favoritos.Rows <= 1 Then Exit Sub
    Set Grelha_Reproduzida = Grelha_Favoritos
    Musica_Linha_Pressionada = Grelha_Favoritos.Row
    Reproduzir_Musica_da_Grelha
    
    'Chamar o procedimento
    'Activar_Linha_em_Reproducao Grelha_Favoritos
End Sub

Private Sub Grelha_Favoritos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_favoritos
    Repor_Objectos
End Sub

Private Sub Grelha_Ficheiros_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Ficheiros_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_ficheiros
    Repor_Objectos
End Sub

Private Sub Grelha_Filmes_Click()
    'Chamar o procedimento
    Ocultar_menus
    
    'Selecionar musica para reproduzir
    If Grelha_Filmes.Rows <= 1 Then Exit Sub
    With Grelha_Filmes
        Text_Classificacao.Text = .TextMatrix(.Row, 6)
    End With
    Verificar_Classificacao
End Sub

Private Sub Grelha_Filmes_DblClick()
    'Reproduzir o ficheiro selecionado
    If Grelha_Filmes.Rows <= 1 Then Exit Sub
    Set Grelha_Reproduzida = Grelha_Filmes
    Musica_Linha_Pressionada = Grelha_Reproduzida.Row
    Reproduzir_Musica_da_Grelha
    
    'Chamar o procedimento
    'Activar_Linha_em_Reproducao Grelha_Filmes
End Sub

Private Sub Grelha_Filmes_EnterCell()
    'Selecionar musica para reproduzir
    If Grelha_Filmes.Rows > 1 Then
        With Grelha_Filmes
            Text_Classificacao.Text = .TextMatrix(.Row, 6)
        End With
        Verificar_Classificacao
    End If
End Sub

Private Sub Grelha_Filmes_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    'Tocar o player
    If KeyCode = vbKeyReturn Then Grelha_Filmes_DblClick

    'Remover linha
    If KeyCode = vbKeyDelete Then
        Grelha_Filmes.RemoveItem Grelha_Filmes.Row
        Text_Classificacao.Text = ""
        Grelha_Filmes_EnterCell
        If Grelha_Reproduzida = Grelha_Filmes Then Musica_Linha_Pressionada = Grelha_Reproduzida.Row - 1
    End If
    
    'Tocar ou pausar a musica atraves do backspace
    If KeyCode = vbKeySpace Then
        If Timer_Slider_Video.Enabled = False Then
            Botao_Play_Click
        Else
            Botao_Pausa_Click
        End If
    End If
End Sub

Private Sub Grelha_Filmes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_Filmes
    Repor_Objectos
End Sub

Private Sub Grelha_Filmes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ordenar pela coluna clicada. Se a grelha estiver vazia não faz nada
    With Grelha_Filmes
        If .MouseRow <> 0 Then Exit Sub
        If .Rows <= 1 Then Exit Sub
        Dim coluna_ordenada As Integer: coluna_ordenada = .MouseCol
        'Ocultar a grid
        .Visible = False
        .Refresh
    
        'Ordenar usando a coluna clicada
        .Col = coluna_ordenada
        .ColSel = coluna_ordenada
        .Row = 0
        .RowSel = 0
    
        'Se esta é uma nova coluna de classificação, classificar em ordem crescente. Caso contrário, mudar a ordem
        If m_SortColumn <> coluna_ordenada Then
            m_SortOrder = flexSortGenericAscending
        ElseIf m_SortOrder = flexSortGenericAscending Then
            m_SortOrder = flexSortGenericDescending
        Else
            m_SortOrder = flexSortGenericAscending
        End If
        .Sort = m_SortOrder
    
        'Restaurar o nome da coluna sem o caracter de identidicação de ordenação
        If m_SortColumn >= 0 Then
            If m_SortColumn <= .Cols - 1 Then
                .TextMatrix(0, m_SortColumn) = Mid$(.TextMatrix(0, m_SortColumn), 3)
            End If
        End If
    
        'Identificar qual é a coluna ordenada
        m_SortColumn = coluna_ordenada
        If m_SortOrder = flexSortGenericAscending Then
            .TextMatrix(0, m_SortColumn) = "> " & .TextMatrix(0, m_SortColumn)
        Else
            .TextMatrix(0, m_SortColumn) = "< " & .TextMatrix(0, m_SortColumn)
        End If
        
        'Visualizar a grid
        .Visible = True
        
        'Selecionar a 1ª linha por inteiro
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Grelha_Genero_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Genero_DblClick()
    'Carregar as músicas consoante a categoria selecionado
    Criterio = Grelha_Genero.TextMatrix(Grelha_Genero.Row, 1)
    Verifica_Rs_Musica
    Rs_Musica.Open "select * from Tabela_Musica where Genero = '" & Criterio & "' order by Genero Asc", Cnn_Biblioteca
    Carregar_Grelha_Musica
End Sub

Private Sub Grelha_Lista_Em_Reproducao_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Lista_Em_Reproducao_DblClick()
    'Reproduzir o ficheiro selecionado
    If Grelha_Lista_Em_Reproducao.Rows <= 1 Then Exit Sub
    Set Grelha_Reproduzida = Grelha_Lista_Em_Reproducao
    Reproduzir_Musica_da_Grelha
End Sub

Private Sub Grelha_Lista_Em_Reproducao_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    'Tocar o player
    If KeyCode = vbKeyReturn Then Grelha_Lista_Em_Reproducao_DblClick

    'Remover linha
    If KeyCode = vbKeyDelete Then Grelha_Lista_Em_Reproducao.RemoveItem Grelha_Lista_Em_Reproducao.Row
    
    'Tocar ou pausar a musica atraves do backspace
    If KeyCode = vbKeySpace Then
        If Timer_Slider_Video.Enabled = False Then
            Botao_Play_Click
        Else
            Botao_Pausa_Click
        End If
    End If
End Sub

Private Sub Grelha_Lista_Em_Reproducao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_Lista_Em_Reproducao
End Sub

Private Sub Grelha_Lista_Em_Reproducao_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ordenar pela coluna clicada. Se a grelha estiver vazia não faz nada
    With Grelha_Lista_Em_Reproducao
        If .MouseRow <> 0 Then Exit Sub
        If .Rows <= 1 Then Exit Sub
        Dim coluna_ordenada As Integer: coluna_ordenada = .MouseCol
        'Ocultar a grid
        .Visible = False
        .Refresh
    
        'Ordenar usando a coluna clicada
        .Col = coluna_ordenada
        .ColSel = coluna_ordenada
        .Row = 0
        .RowSel = 0
    
        'Se esta é uma nova coluna de classificação, classificar em ordem crescente. Caso contrário, mudar a ordem
        If m_SortColumn <> coluna_ordenada Then
            m_SortOrder = flexSortGenericAscending
        ElseIf m_SortOrder = flexSortGenericAscending Then
            m_SortOrder = flexSortGenericDescending
        Else
            m_SortOrder = flexSortGenericAscending
        End If
        .Sort = m_SortOrder
    
        'Restaurar o nome da coluna sem o caracter de identidicação de ordenação
        If m_SortColumn >= 0 Then
            If m_SortColumn <= .Cols - 1 Then
                .TextMatrix(0, m_SortColumn) = Mid$(.TextMatrix(0, m_SortColumn), 3)
            End If
        End If
    
        'Identificar qual é a coluna ordenada
        m_SortColumn = coluna_ordenada
        If m_SortOrder = flexSortGenericAscending Then
            .TextMatrix(0, m_SortColumn) = "> " & .TextMatrix(0, m_SortColumn)
        Else
            .TextMatrix(0, m_SortColumn) = "< " & .TextMatrix(0, m_SortColumn)
        End If
        
        'Visualizar a grid
        .Visible = True
        
        'Selecionar a 1ª linha por inteiro
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Grelha_Lista_Em_Reproducao_OLEDragDrop(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Drag drop dos ficheiro
    Dim sFile As String, sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
    Dim Files As Variant
    For Each Files In Data.Files
       'Grelha_Lista_Em_Reproducao.AddItem Files
        Dim Ficheiro_Arrastado As String
        Ficheiro_Arrastado = Files
        cFile.FileName = Ficheiro_Arrastado

        Dim nova_linha As Integer
        cFile.FileName = True
        cFile.FileName = Ficheiro_Arrastado
        sTitle = Replace(cFile.Title, "'", " ", , , vbTextCompare)
        sArtist = Replace(cFile.Artist, "'", " ", , , vbTextCompare)
        sAlbum = Replace(cFile.album, "'", " ", , , vbTextCompare)
        sYear = Replace(cFile.Year, "'", " ", , , vbTextCompare)
        sGenre = Replace(cFile.Genre, "'", " ", , , vbTextCompare)
        sComment = Replace(cFile.Comments, "'", " ", , , vbTextCompare)
                   
        'Caso o ficheiro não tenha tags então o titulo será o nome do ficheiro, o qual é obtido atrvés do directório
        Dim nome_ficheiro As String: nome_ficheiro = Dir(Ficheiro_Arrastado, vbArchive)
        If sTitle = "" Then sTitle = Mid(nome_ficheiro, 1, InStrRev(nome_ficheiro, ".") - 1)
        If sArtist = "" Then sArtist = ""
        If sAlbum = "" Then sAlbum = ""
        If sYear = "" Then sYear = ""
        If sGenre = "" Then sGenre = ""
        If sComment = "" Then sComment = ""
                   
        'Adicionar as músicas na playlist
        With Grelha_Lista_Em_Reproducao
            .Rows = .Rows + 1
            nova_linha = .Rows - 1
            .TextMatrix(nova_linha, 0) = Ficheiro_Arrastado
            .TextMatrix(nova_linha, 1) = sTitle
            .TextMatrix(nova_linha, 2) = sArtist
            .TextMatrix(nova_linha, 3) = sAlbum
            .TextMatrix(nova_linha, 4) = sYear
            .TextMatrix(nova_linha, 5) = sGenre
            .TextMatrix(nova_linha, 6) = sComment
            .TextMatrix(nova_linha, 7) = Dir(Ficheiro_Arrastado, vbDirectory)
            .TextMatrix(nova_linha, 8) = "0"
            .Row = nova_linha
        End With
    Next
    Effect = vbDropEffectNone
End Sub

Private Sub Grelha_Listas_Click()
    'Chamar o procedimento
    Ocultar_menus
    
    'Selecionar musica para reproduzir
    If Grelha_Listas.Rows <= 1 Then Exit Sub
    With Grelha_Listas
        Text_Classificacao.Text = .TextMatrix(.Row, 8)
    End With
    Verificar_Classificacao
End Sub

Private Sub Grelha_Listas_DblClick()
    'Reproduzir o ficheiro selecionado
    If Grelha_Listas.Rows <= 1 Then Exit Sub
    Set Grelha_Reproduzida = Grelha_Listas
    Musica_Linha_Pressionada = Grelha_Listas.Row
    Reproduzir_Musica_da_Grelha
    
    'Chamar o procedimento
    'Activar_Linha_em_Reproducao Grelha_Listas
End Sub

Private Sub Grelha_Listas_EnterCell()
    'Selecionar musica para reproduzir
    If Grelha_Listas.Rows > 1 Then
        With Grelha_Listas
            Text_Classificacao.Text = .TextMatrix(.Row, 8)
        End With
        Verificar_Classificacao
    End If
End Sub

Private Sub Grelha_Listas_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    'Tocar o player
    If KeyCode = vbKeyReturn Then Grelha_Listas_DblClick

    'Remover linha
    If KeyCode = vbKeyDelete Then
        Grelha_Listas.RemoveItem Grelha_Listas.Row
        Text_Classificacao.Text = ""
        Grelha_Listas_EnterCell
        If Grelha_Reproduzida = Grelha_Listas Then Musica_Linha_Pressionada = Grelha_Reproduzida.Row - 1
    End If
    
    'Tocar ou pausar a musica atraves do backspace
    If KeyCode = vbKeySpace Then
        If Timer_Slider_Video.Enabled = False Then
            Botao_Play_Click
        Else
            Botao_Pausa_Click
        End If
    End If
End Sub

Private Sub Grelha_Listas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_Listas
    Repor_Objectos
End Sub

Private Sub Grelha_Listas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ordenar pela coluna clicada. Se a grelha estiver vazia não faz nada
    With Grelha_Listas
        If .MouseRow <> 0 Then Exit Sub
        If .Rows <= 1 Then Exit Sub
        Dim coluna_ordenada As Integer: coluna_ordenada = .MouseCol
        'Ocultar a grid
        .Visible = False
        .Refresh
    
        'Ordenar usando a coluna clicada
        .Col = coluna_ordenada
        .ColSel = coluna_ordenada
        .Row = 0
        .RowSel = 0
    
        'Se esta é uma nova coluna de classificação, classificar em ordem crescente. Caso contrário, mudar a ordem
        If m_SortColumn <> coluna_ordenada Then
            m_SortOrder = flexSortGenericAscending
        ElseIf m_SortOrder = flexSortGenericAscending Then
            m_SortOrder = flexSortGenericDescending
        Else
            m_SortOrder = flexSortGenericAscending
        End If
        .Sort = m_SortOrder
    
        'Restaurar o nome da coluna sem o caracter de identidicação de ordenação
        If m_SortColumn >= 0 Then
            If m_SortColumn <= .Cols - 1 Then
                .TextMatrix(0, m_SortColumn) = Mid$(.TextMatrix(0, m_SortColumn), 3)
            End If
        End If
    
        'Identificar qual é a coluna ordenada
        m_SortColumn = coluna_ordenada
        If m_SortOrder = flexSortGenericAscending Then
            .TextMatrix(0, m_SortColumn) = "> " & .TextMatrix(0, m_SortColumn)
        Else
            .TextMatrix(0, m_SortColumn) = "< " & .TextMatrix(0, m_SortColumn)
        End If
        
        'Visualizar a grid
        .Visible = True
        
        'Selecionar a 1ª linha por inteiro
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Grelha_Listas_OLEDragDrop(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Drag drop dos ficheiro
    Dim sFile As String, sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
    Dim Files As Variant
    For Each Files In Data.Files
       'Grelha_Listas.AddItem Files
        Dim Ficheiro_Arrastado As String
        Ficheiro_Arrastado = Files
        cFile.FileName = Ficheiro_Arrastado

        Dim nova_linha As Integer
        cFile.FileName = True
        cFile.FileName = Ficheiro_Arrastado
        sTitle = Replace(cFile.Title, "'", " ", , , vbTextCompare)
        sArtist = Replace(cFile.Artist, "'", " ", , , vbTextCompare)
        sAlbum = Replace(cFile.album, "'", " ", , , vbTextCompare)
        sYear = Replace(cFile.Year, "'", " ", , , vbTextCompare)
        sGenre = Replace(cFile.Genre, "'", " ", , , vbTextCompare)
        sComment = Replace(cFile.Comments, "'", " ", , , vbTextCompare)
                   
        'Caso o ficheiro não tenha tags então o titulo será o nome do ficheiro, o qual é obtido atrvés do directório
        Dim nome_ficheiro As String: nome_ficheiro = Dir(Ficheiro_Arrastado, vbArchive)
        If sTitle = "" Then sTitle = Mid(nome_ficheiro, 1, InStrRev(nome_ficheiro, ".") - 1)
        If sArtist = "" Then sArtist = ""
        If sAlbum = "" Then sAlbum = ""
        If sYear = "" Then sYear = ""
        If sGenre = "" Then sGenre = ""
        If sComment = "" Then sComment = ""
                   
        'Adicionar as músicas na playlist
        With Grelha_Listas
            .Rows = .Rows + 1
            nova_linha = .Rows - 1
            .TextMatrix(nova_linha, 0) = Ficheiro_Arrastado
            .TextMatrix(nova_linha, 1) = sTitle
            .TextMatrix(nova_linha, 2) = sArtist
            .TextMatrix(nova_linha, 3) = sAlbum
            .TextMatrix(nova_linha, 4) = sYear
            .TextMatrix(nova_linha, 5) = sGenre
            .TextMatrix(nova_linha, 6) = sComment
            .TextMatrix(nova_linha, 7) = Dir(Ficheiro_Arrastado, vbDirectory)
            .TextMatrix(nova_linha, 8) = "0"
            .Row = nova_linha
        End With
    Next
    Effect = vbDropEffectNone
    
'    'Guardar automaticamente a lista
'    If Grelha_Lista_Em_Reproducao.Rows > 1 Then
'        Personalizar_Grid Grelha_Lista_Em_Reproducao
'        Dim cFlexSettings As clsFlexSettings
'        Set cFlexSettings = New clsFlexSettings
'        Set cFlexSettings.FlexGrid = Grelha_Lista_Em_Reproducao
'        cFlexSettings.SaveSettings App.Path & "\Library\Standard.ini", True, True, True, True
'        Set clsFlexSettings = Nothing
'    End If
End Sub

Private Sub Grelha_Loja_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Loja_DblClick()
    'Reproduzir o ficheiro selecionado
    If Grelha_Loja.Rows <= 1 Then Exit Sub
    Set Grelha_Reproduzida = Grelha_Loja
    Musica_Linha_Pressionada = Grelha_Reproduzida.Row
    Reproduzir_Musica_da_Grelha
    
    'Chamar o procedimento
    'Activar_Linha_em_Reproducao Grelha_Loja
End Sub

Private Sub Grelha_Loja_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    'Tocar o player
    If KeyCode = vbKeyReturn Then Grelha_Loja_DblClick

    'Remover linha
    If KeyCode = vbKeyDelete Then
        Grelha_Loja.RemoveItem Grelha_Loja.Row
        'Text_Classificacao.Text = ""
        'Grelha_Loja_EnterCell
        If Grelha_Reproduzida = Grelha_Loja Then Musica_Linha_Pressionada = Grelha_Reproduzida.Row - 1
    End If
    
    'Tocar ou pausar a musica atraves do backspace
    If KeyCode = vbKeySpace Then
        If Timer_Slider_Video.Enabled = False Then
            Botao_Play_Click
        Else
            Botao_Pausa_Click
        End If
    End If
End Sub

Private Sub Grelha_Loja_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_Loja
    Repor_Objectos
End Sub

Private Sub Grelha_Loja_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ordenar pela coluna clicada. Se a grelha estiver vazia não faz nada
    With Grelha_Loja
        If .MouseRow <> 0 Then Exit Sub
        If .Rows <= 1 Then Exit Sub
        Dim coluna_ordenada As Integer: coluna_ordenada = .MouseCol
        'Ocultar a grid
        .Visible = False
        .Refresh
    
        'Ordenar usando a coluna clicada
        .Col = coluna_ordenada
        .ColSel = coluna_ordenada
        .Row = 0
        .RowSel = 0
    
        'Se esta é uma nova coluna de classificação, classificar em ordem crescente. Caso contrário, mudar a ordem
        If m_SortColumn <> coluna_ordenada Then
            m_SortOrder = flexSortGenericAscending
        ElseIf m_SortOrder = flexSortGenericAscending Then
            m_SortOrder = flexSortGenericDescending
        Else
            m_SortOrder = flexSortGenericAscending
        End If
        .Sort = m_SortOrder
    
        'Restaurar o nome da coluna sem o caracter de identidicação de ordenação
        If m_SortColumn >= 0 Then
            If m_SortColumn <= .Cols - 1 Then
                .TextMatrix(0, m_SortColumn) = Mid$(.TextMatrix(0, m_SortColumn), 3)
            End If
        End If
    
        'Identificar qual é a coluna ordenada
        m_SortColumn = coluna_ordenada
        If m_SortOrder = flexSortGenericAscending Then
            .TextMatrix(0, m_SortColumn) = "> " & .TextMatrix(0, m_SortColumn)
        Else
            .TextMatrix(0, m_SortColumn) = "< " & .TextMatrix(0, m_SortColumn)
        End If
        
        'Visualizar a grid
        .Visible = True
        
        'Selecionar a 1ª linha por inteiro
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Grelha_Mensagens_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Mensagens_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_mensagens
    Repor_Objectos
End Sub

Private Sub Grelha_Minha_Musica_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Minha_Musica_DblClick()
    'Reproduzir o ficheiro selecionado
    If Grelha_Minha_Musica.Rows <= 1 Then Exit Sub
    Set Grelha_Reproduzida = Grelha_Minha_Musica
    Musica_Linha_Pressionada = Grelha_Reproduzida.Row
    Reproduzir_Musica_da_Grelha
End Sub

Private Sub Grelha_Minha_Musica_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    'Tocar o player
    If KeyCode = vbKeyReturn Then Grelha_Minha_Musica_DblClick

    'Remover linha
    If KeyCode = vbKeyDelete Then
        Grelha_Minha_Musica.RemoveItem Grelha_Minha_Musica.Row
        'Text_Classificacao.Text = ""
        'Grelha_Minha_Musica_EnterCell
        If Grelha_Reproduzida = Grelha_Minha_Musica Then Musica_Linha_Pressionada = Grelha_Reproduzida.Row - 1
    End If
    
    'Tocar ou pausar a musica atraves do backspace
    If KeyCode = vbKeySpace Then
        If Timer_Slider_Video.Enabled = False Then
            Botao_Play_Click
        Else
            Botao_Pausa_Click
        End If
    End If
End Sub

Private Sub Grelha_Minha_Musica_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_Minha_Musica
    Repor_Objectos
End Sub

Private Sub Grelha_Minha_Musica_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ordenar pela coluna clicada. Se a grelha estiver vazia não faz nada
    With Grelha_Minha_Musica
        If .MouseRow <> 0 Then Exit Sub
        If .Rows <= 1 Then Exit Sub
        Dim coluna_ordenada As Integer: coluna_ordenada = .MouseCol
        'Ocultar a grid
        .Visible = False
        .Refresh
    
        'Ordenar usando a coluna clicada
        .Col = coluna_ordenada
        .ColSel = coluna_ordenada
        .Row = 0
        .RowSel = 0
    
        'Se esta é uma nova coluna de classificação, classificar em ordem crescente. Caso contrário, mudar a ordem
        If m_SortColumn <> coluna_ordenada Then
            m_SortOrder = flexSortGenericAscending
        ElseIf m_SortOrder = flexSortGenericAscending Then
            m_SortOrder = flexSortGenericDescending
        Else
            m_SortOrder = flexSortGenericAscending
        End If
        .Sort = m_SortOrder
    
        'Restaurar o nome da coluna sem o caracter de identidicação de ordenação
        If m_SortColumn >= 0 Then
            If m_SortColumn <= .Cols - 1 Then
                .TextMatrix(0, m_SortColumn) = Mid$(.TextMatrix(0, m_SortColumn), 3)
            End If
        End If
    
        'Identificar qual é a coluna ordenada
        m_SortColumn = coluna_ordenada
        If m_SortOrder = flexSortGenericAscending Then
            .TextMatrix(0, m_SortColumn) = "> " & .TextMatrix(0, m_SortColumn)
        Else
            .TextMatrix(0, m_SortColumn) = "< " & .TextMatrix(0, m_SortColumn)
        End If
        
        'Visualizar a grid
        .Visible = True
        
        'Selecionar a 1ª linha por inteiro
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Grelha_Musica_Click()
    'Chamar o procedimento
    Ocultar_menus
    
    'Selecionar musica para reproduzir
    If Grelha_Musica.Rows <= 1 Then Exit Sub
    With Grelha_Musica
        Text_Classificacao.Text = .TextMatrix(.Row, 8)
    End With
    Verificar_Classificacao
End Sub

Private Sub Grelha_Musica_EnterCell()
    'Selecionar musica para reproduzir
    If Grelha_Musica.Rows > 1 Then
        With Grelha_Musica
            Text_Classificacao.Text = .TextMatrix(.Row, 8)
        End With
        Verificar_Classificacao
    End If
End Sub

Private Sub Grelha_Musica_DblClick()
    'Reproduzir o ficheiro selecionado
    If Grelha_Musica.Rows <= 1 Then Exit Sub
    Set Grelha_Reproduzida = Grelha_Musica
    Musica_Linha_Pressionada = Grelha_Reproduzida.Row
    Reproduzir_Musica_da_Grelha
    
    'Chamar o procedimento
    'Activar_Linha_em_Reproducao Grelha_Musica
    
    'Carregar a lista de reprodução caso esta esteja vazia
    If Grelha_Lista_Em_Reproducao.Rows = 1 Then
        Dim i, Linha As Integer
        With Grelha_Lista_Em_Reproducao
            .Rows = 1
            i = 1
            For Linha = 1 To Grelha_Musica.Rows - 1
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Grelha_Musica.TextMatrix(i, 0)
                .TextMatrix(i, 1) = Grelha_Musica.TextMatrix(i, 1)
                .TextMatrix(i, 2) = Grelha_Musica.TextMatrix(i, 2)
                .TextMatrix(i, 3) = Grelha_Musica.TextMatrix(i, 3)
                .TextMatrix(i, 4) = Grelha_Musica.TextMatrix(i, 4)
                .TextMatrix(i, 5) = Grelha_Musica.TextMatrix(i, 5)
                .TextMatrix(i, 6) = Grelha_Musica.TextMatrix(i, 6)
                .TextMatrix(i, 7) = Grelha_Musica.TextMatrix(i, 7)
                .TextMatrix(i, 8) = Grelha_Musica.TextMatrix(i, 8)
                .TextMatrix(i, 9) = Grelha_Musica.TextMatrix(i, 9)
                i = i + 1
            Next Linha
        End With
        
        Grelha_Lista_Em_Reproducao.Visible = True
        Label_Carregar_Favoritos.Visible = False
    End If
End Sub

Public Sub Reproduzir_Musica_da_Grelha()
    'Procedimento para indicar qual a música escolhida na respectiva grelha
    'Reproduzir o ficheiro selecionado
    If Grelha_Reproduzida.Rows <= 1 Then Exit Sub
    'Definir a linha a ser reproduzida
    Faixa_em_Reproducao = Grelha_Reproduzida.TextMatrix(Grelha_Reproduzida.Row, 0)
    Tocar_Media
End Sub

Private Sub Grelha_Musica_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    'Tocar o player
    If KeyCode = vbKeyReturn Then Grelha_Musica_DblClick

    'Remover linha
    If KeyCode = vbKeyDelete Then
        Grelha_Musica.RemoveItem Grelha_Musica.Row
        Text_Classificacao.Text = ""
        Grelha_Musica_EnterCell
        If Grelha_Reproduzida = Grelha_Musica Then Musica_Linha_Pressionada = Grelha_Reproduzida.Row - 1
    End If
    
    'Tocar ou pausar a musica atraves do backspace
    If KeyCode = vbKeySpace Then
        If Timer_Slider_Video.Enabled = False Then
            Botao_Play_Click
        Else
            Botao_Pausa_Click
        End If
    End If
End Sub

Private Sub Grelha_Musica_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_Musica
    Repor_Objectos
End Sub

Private Sub Grelha_Musica_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ordenar pela coluna clicada. Se a grelha estiver vazia não faz nada
    With Grelha_Musica
        If .MouseRow <> 0 Then Exit Sub
        If .Rows <= 1 Then Exit Sub
        Dim coluna_ordenada As Integer: coluna_ordenada = .MouseCol
        'Ocultar a grid
        .Visible = False
        .Refresh
    
        'Ordenar usando a coluna clicada
        .Col = coluna_ordenada
        .ColSel = coluna_ordenada
        .Row = 0
        .RowSel = 0
    
        'Se esta é uma nova coluna de classificação, classificar em ordem crescente. Caso contrário, mudar a ordem
        If m_SortColumn <> coluna_ordenada Then
            m_SortOrder = flexSortGenericAscending
        ElseIf m_SortOrder = flexSortGenericAscending Then
            m_SortOrder = flexSortGenericDescending
        Else
            m_SortOrder = flexSortGenericAscending
        End If
        .Sort = m_SortOrder
    
        'Restaurar o nome da coluna sem o caracter de identidicação de ordenação
        If m_SortColumn >= 0 Then
            If m_SortColumn <= .Cols - 1 Then
                .TextMatrix(0, m_SortColumn) = Mid$(.TextMatrix(0, m_SortColumn), 3)
            End If
        End If
    
        'Identificar qual é a coluna ordenada
        m_SortColumn = coluna_ordenada
        If m_SortOrder = flexSortGenericAscending Then
            .TextMatrix(0, m_SortColumn) = "> " & .TextMatrix(0, m_SortColumn)
        Else
            .TextMatrix(0, m_SortColumn) = "< " & .TextMatrix(0, m_SortColumn)
        End If
        
        'Visualizar a grid
        .Visible = True
        
        'Selecionar a 1ª linha por inteiro
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Grelha_Radio_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Radio_DblClick()
    'Reproduzir a estação de rádio selecionada
    If Grelha_Radio.Rows <= 1 Then Exit Sub
    Set Grelha_Reproduzida = Grelha_Radio
    Musica_Linha_Pressionada = Grelha_Reproduzida.Row
    Reproduzir_Musica_da_Grelha
End Sub

Private Sub Grelha_Radio_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    'Tocar o player
    If KeyCode = vbKeyReturn Then Grelha_Radio_DblClick
    
    'Tocar ou pausar a musica atraves do backspace
    If KeyCode = vbKeySpace Then
        If Timer_Slider_Video.Enabled = False Then
            Botao_Play_Click
        Else
            Botao_Pausa_Click
        End If
    End If
End Sub

Private Sub Grelha_Radio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_Radio
    Repor_Objectos
End Sub

Private Sub Grelha_Radio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ordenar pela coluna clicada. Se a grelha estiver vazia não faz nada
    With Grelha_Radio
        If .MouseRow <> 0 Then Exit Sub
        If .Rows <= 1 Then Exit Sub
        Dim coluna_ordenada As Integer: coluna_ordenada = .MouseCol
        'Ocultar a grid
        .Visible = False
        .Refresh
    
        'Ordenar usando a coluna clicada
        .Col = coluna_ordenada
        .ColSel = coluna_ordenada
        .Row = 0
        .RowSel = 0
    
        'Se esta é uma nova coluna de classificação, classificar em ordem crescente. Caso contrário, mudar a ordem
        If m_SortColumn <> coluna_ordenada Then
            m_SortOrder = flexSortGenericAscending
        ElseIf m_SortOrder = flexSortGenericAscending Then
            m_SortOrder = flexSortGenericDescending
        Else
            m_SortOrder = flexSortGenericAscending
        End If
        .Sort = m_SortOrder
    
        'Restaurar o nome da coluna sem o caracter de identidicação de ordenação
        If m_SortColumn >= 0 Then
            If m_SortColumn <= .Cols - 1 Then
                .TextMatrix(0, m_SortColumn) = Mid$(.TextMatrix(0, m_SortColumn), 3)
            End If
        End If
    
        'Identificar qual é a coluna ordenada
        m_SortColumn = coluna_ordenada
        If m_SortOrder = flexSortGenericAscending Then
            .TextMatrix(0, m_SortColumn) = "> " & .TextMatrix(0, m_SortColumn)
        Else
            .TextMatrix(0, m_SortColumn) = "< " & .TextMatrix(0, m_SortColumn)
        End If
        
        'Visualizar a grid
        .Visible = True
        
        'Selecionar a 1ª linha por inteiro
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Grelha_Recentes_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Grelha_Recentes_DblClick()
    'Reproduzir o ficheiro selecionado
    If Grelha_Recentes.Rows <= 1 Then Exit Sub
    Set Grelha_Reproduzida = Grelha_Recentes
    Musica_Linha_Pressionada = Grelha_Recentes.Row
    Reproduzir_Musica_da_Grelha
    
    'Chamar o procedimento
    'Activar_Linha_em_Reproducao Grelha_Recentes
End Sub

Private Sub Grelha_Recentes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Activar scroll da grelha
    'Set mygrid = Grelha_recentes
    Repor_Objectos
End Sub

Private Sub Icon_Barra_Informacoes_Click(Index As Integer)
    'Selecionar botão
    Ocultar_menus
    
    Select Case Icon_Barra_Informacoes(Index).Index
        Case 0 'Nova playlist
            Menu_Ficheiro_Click 7
            
        Case 1 'Tocar uma música aleatória
            If musica_aleatoria = False Then
                musica_aleatoria = True
                Icon_Barra_Informacoes(1).Picture = Form_Skin.Button_Music_Randomize_Over.Picture
            Else
                musica_aleatoria = False
                Icon_Barra_Informacoes(1).Picture = Form_Skin.Button_Music_Randomize_Normal.Picture
            End If
        
        Case 2 'Recomeçar a tocar a biblioteca do inicio após esta chegar ao fim
            If musica_recomecar = False Then
                musica_recomecar = True
                Icon_Barra_Informacoes(2).Picture = Form_Skin.Button_Music_Repete_Over.Picture
            Else
                musica_recomecar = False
                Icon_Barra_Informacoes(2).Picture = Form_Skin.Button_Music_Repete_Normal.Picture
            End If
        
        Case 3 'Ver/ ocultar capa do album
            If Frame_Capa.Visible = False Then
                Frame_Capa.Visible = True
                Icon_Barra_Informacoes(3).Picture = Form_Skin.Button_Cover_Hide_Normal.Picture
                Icon_Barra_Informacoes(3).ToolTipText = Idioma_Ocultar_Capa
                Form_Preferencias.Check_Ver_Capa.Value = 1
                Form_Preferencias.Salvar_Valores
                Menu_Check(0).Visible = True
                Menu_Check(0).Picture = Form_Skin.Menu_Check_Normal.Picture
            Else
                Frame_Capa.Visible = False
                Icon_Barra_Informacoes(3).Picture = Form_Skin.Button_Cover_View_Normal.Picture
                Icon_Barra_Informacoes(3).ToolTipText = Idioma_Ver_Capa
                Form_Preferencias.Check_Ver_Capa.Value = 0
                Form_Preferencias.Salvar_Valores
                Menu_Check(0).Visible = False
            End If
            
        Case 4 'Adicionar um novo ficheiro media----------------------------------------------------------------------------------------
            'Carregar ficheiros de midia para a lista em reprodução
            If Barra_Playlist.Visible = False Then Exit Sub 'Or Grelha_Visivel <> Grelha_Listas
            Dim nome_grelha As MSFlexGrid
            
            If Barra_Playlist.Visible = True Then
                Set nome_grelha = Grelha_Lista_Em_Reproducao
            Else
                Set nome_grelha = nome_grelha
            End If
            
            With Explorador
                .Filter = ("*.wav,*.snd,*.au,*.aif,*.aifc,*.aiff,*.mid,*.rmi,*.mp3,*.m3u,*.m1v,*.mp2,*.mpa,*.mpe,*.mpeg,*.asf,*.asx,*.mov,*.qt,*.ra*.rm,*.ram,*.rmm,*.avi,*.mpg")
                ' A extensão default será a última (Todos)
                '.FilterIndex = List1.SelCount + 1
                
                ' Decide a pasta inicial
                If Text_Pesquisar_Musica <> "" Then
                    If Dir(Text_Pesquisar_Musica, vbDirectory) <> "" Then
                        ' Se for uma pasta, é ela mesma
                        .Path = Text_Pesquisar_Musica
                    Else
                        ' Se for um arquivo, extraia só o caminho
                        .Path = left(Text_Pesquisar_Musica, InStrRev(Text_Pesquisar_Musica, "\"))
                    End If
                End If
                
                .FileFlags = PATHMUSTEXIST
                .FileFlags = .FileFlags + EXPLORER
                .FileFlags = .FileFlags + ALLOWMULTISELECT
                
                ' Mostra o diálogo
                .DialogFile OpenFile
                If .cancel = True Then Exit Sub
                
                If Grelha_Lista_Em_Reproducao.Rows = 1 Then
                    Grelha_Lista_Em_Reproducao.Visible = True
                    Label_Carregar_Favoritos.Visible = False
                End If
                'Caso tenha sido selecionado algum ficheiro então adiciona-o á lista
                If Len(.FileName) <> 0 Then
                    Dim Musicas() As String
                    Musicas = Split(.FileName, "|")
                    Dim nova_linha As Integer
                    Dim nome_ficheiro As String
                    
                    Dim contador As Integer: For contador = 0 To UBound(Musicas)
                        nome_grelha.Rows = nome_grelha.Rows + 1
                        nova_linha = nome_grelha.Rows - 1
                        nome_grelha.TextMatrix(nova_linha, 0) = .Path & "\" & Musicas(contador)
                        nome_ficheiro = Dir(Musicas(contador), vbArchive)
                        nome_grelha.TextMatrix(nova_linha, 1) = Mid(nome_ficheiro, 1, InStrRev(nome_ficheiro, ".") - 1)
                    Next contador
                End If
            End With
        
        Case 5 'Ver/ ocultar barra do Playlist
            'If Grelha_Visivel <> Grelha_Musica Or Grelha_Visivel <> Grelha_Filmes Then Exit Sub
            If Barra_Playlist.Visible = True Then
                Barra_Playlist.Visible = False
                Icon_Barra_Informacoes(5).Picture = Form_Skin.Button_Playlist_View_Normal.Picture
                Icon_Barra_Informacoes(5).ToolTipText = Idioma_Ver_Lista
                Form_Preferencias.Check_Ver_Playlist.Value = 0
                Form_Preferencias.Salvar_Valores
                Menu_Check(1).Visible = False
            Else
                Barra_Playlist.Visible = True
                Icon_Barra_Informacoes(5).Picture = Form_Skin.Button_Playlist_Hide_Normal.Picture
                Icon_Barra_Informacoes(5).ToolTipText = Idioma_Ocultar_Lista
                Form_Preferencias.Check_Ver_Playlist.Value = 1
                Form_Preferencias.Salvar_Valores
                Menu_Check(1).Visible = True
                Menu_Check(1).Picture = Form_Skin.Menu_Check_Normal.Picture
            End If
    End Select
            
    'Ajustar os restantes objectos
    Desenhar_Formulario
    Ajustar_Objectos_Na_Horizontal
    Ajustar_Objectos_Na_Vertical
End Sub

Private Sub Icon_Barra_Informacoes_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Slecionar botoões e mostar a imagem down
    Select Case Icon_Barra_Informacoes(Index).Index
        Case 0
            Icon_Barra_Informacoes(0).Picture = Form_Skin.Button_New_Playlist_Down.Picture
            
        Case 1
            If musica_aleatoria = False Then
                Icon_Barra_Informacoes(1).Picture = Form_Skin.Button_Music_Randomize_Down.Picture
            Else
                Icon_Barra_Informacoes(1).Picture = Form_Skin.Button_Music_Randomize_Over_Down.Picture
            End If
            
        Case 2
            If musica_recomecar = False Then
                Icon_Barra_Informacoes(2).Picture = Form_Skin.Button_Music_Repete_Down.Picture
            Else
                Icon_Barra_Informacoes(2).Picture = Form_Skin.Button_Music_Repete_Over_Down.Picture
            End If
            
        Case 3
            If Frame_Capa.Visible = False Then
                Icon_Barra_Informacoes(3).Picture = Form_Skin.Button_Cover_View_Down.Picture
            Else
                Icon_Barra_Informacoes(3).Picture = Form_Skin.Button_Cover_Hide_Down.Picture
            End If
            
        Case 4
            Icon_Barra_Informacoes(4).Picture = Form_Skin.Button_Folder_Down.Picture
            
        Case 5
            If Barra_Playlist.Visible = False Then
                Icon_Barra_Informacoes(5).Picture = Form_Skin.Button_Playlist_View_Down.Picture
            Else
                Icon_Barra_Informacoes(5).Picture = Form_Skin.Button_Playlist_Hide_Down.Picture
            End If
    End Select
End Sub

Private Sub Icon_Barra_Informacoes_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar botão
    Select Case Icon_Barra_Informacoes(Index).Index
        Case 0
            Icon_Barra_Informacoes(0).Picture = Form_Skin.Button_New_Playlist_Normal.Picture
            
        Case 1 'Tocar uma música aleatória
            If musica_aleatoria = True Then
                Icon_Barra_Informacoes(1).Picture = Form_Skin.Button_Music_Randomize_Over.Picture
            Else
                Icon_Barra_Informacoes(1).Picture = Form_Skin.Button_Music_Randomize_Normal.Picture
            End If
        
        Case 2 'Recomeçar a tocar a biblioteca do inicio após esta chegar ao fim
            If musica_recomecar = True Then
                Icon_Barra_Informacoes(2).Picture = Form_Skin.Button_Music_Repete_Over.Picture
            Else
                Icon_Barra_Informacoes(2).Picture = Form_Skin.Button_Music_Repete_Normal.Picture
            End If
        
        Case 3 'Ver/ ocultar capa do album
            If Frame_Capa.Visible = True Then
                Icon_Barra_Informacoes(3).Picture = Form_Skin.Button_Cover_Hide_Normal.Picture
            Else
                Icon_Barra_Informacoes(3).Picture = Form_Skin.Button_Cover_View_Normal.Picture
            End If
            
        Case 4 'Adicionar um novo ficheiro media----------------------------------------------------------------------------------------
            Icon_Barra_Informacoes(4).Picture = Form_Skin.Button_Folder_Normal.Picture
                    
        Case 5 'Ver/ ocultar barra do Playlist
            If Barra_Playlist.Visible = True Then
                Icon_Barra_Informacoes(5).Picture = Form_Skin.Button_Playlist_Hide_Normal.Picture
            Else
                Icon_Barra_Informacoes(5).Picture = Form_Skin.Button_Playlist_View_Normal.Picture
            End If
    End Select
End Sub

Private Sub Icon_Pasta_Categoria_Click(Index As Integer)
    'Ocultar o separador referente ao programa que foi solicitado ver mais informações
    Separador_Frame_Programas(3).Visible = False
    Label_Frame_Programas(3).Visible = False
    Label_Frame_Programas(3).Caption = ""
    
    'Selecionar a categoria do programa
    Select Case Icon_Pasta_Categoria(Index).Index
        Case 0
            Me.MousePointer = 11
            Selecionar_Categoria "Ferramentas", ReadINI("Main", "Folder_Tools", Localizacao_Ficheiro_Lingua)
            Me.MousePointer = 0
        Case 1
            Me.MousePointer = 11
            Selecionar_Categoria "Som e video", ReadINI("Main", "Folder_Media", Localizacao_Ficheiro_Lingua)
            Me.MousePointer = 0
        Case 2
            Me.MousePointer = 11
            Selecionar_Categoria "Ferramentas", ReadINI("Main", "Folder_Accessibility", Localizacao_Ficheiro_Lingua)
            Me.MousePointer = 0
        Case 3
            Me.MousePointer = 11
            Selecionar_Categoria "Internet", ReadINI("Main", "Folder_Internet", Localizacao_Ficheiro_Lingua)
            Me.MousePointer = 0
        Case 4
            Me.MousePointer = 11
            Selecionar_Categoria "Jogos", ReadINI("Main", "Folder_Games", Localizacao_Ficheiro_Lingua)
            Me.MousePointer = 0
    End Select
End Sub

Private Sub Icon_Pasta_Categoria_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Indique que pasta está a ser selecionada
    Select Case Icon_Pasta_Categoria(Index).Index
        Case 0
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Tools", Localizacao_Ficheiro_Lingua)
        Case 1
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Media", Localizacao_Ficheiro_Lingua)
        Case 2
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Accessibility", Localizacao_Ficheiro_Lingua)
        Case 3
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Internet", Localizacao_Ficheiro_Lingua)
        Case 4
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Games", Localizacao_Ficheiro_Lingua)
    End Select
End Sub

Private Sub Icon_Visao_Click(Index As Integer)
    'Ver album art
    Ocultar_menus
    
    If Grelha_Visivel <> Grelha_Musica Then Exit Sub
    Repor_Icons_Visao
    visao_actual_da_biblioteca = Index
    Text_Visualizacao.Text = Index
    
    Dim i As Integer: For i = 0 To Menu_Check.count - 1
        Menu_Check(i).Picture = Form_Skin.Menu_Check_Normal.Picture
    Next
    
    Select Case Icon_Visao(Index).Index
        Case 0 'Simples
            Frame_Album.Visible = False
            Icon_Visao(0).Picture = Form_Skin.Icon_Visao_Down(0).Picture
            Menu_Check(2).Visible = True: Menu_Check(3).Visible = False: Menu_Check(4).Visible = False
            
        Case 1 'Avançada
            Frame_Album.Visible = True
            Frame_Grelhas_Pesquisa.Visible = True
            Frame_Slide_Album.Visible = False
            Icon_Visao(1).Picture = Form_Skin.Icon_Visao_Down(1).Picture
            Menu_Check(2).Visible = False: Menu_Check(3).Visible = True: Menu_Check(4).Visible = False
            
        Case 2 'Album art
            Frame_Album.Visible = True
            If Lista_Pastas.ListCount = 0 Then
                Frame_Slide_Album.Visible = False
                Label_Nenhum_Album.Visible = True
            Else
                Frame_Slide_Album.Visible = True
            End If
            Frame_Grelhas_Pesquisa.Visible = False
            Icon_Visao(2).Picture = Form_Skin.Icon_Visao_Down(2).Picture
            Menu_Check(2).Visible = False: Menu_Check(3).Visible = False: Menu_Check(4).Visible = True
    End Select
    
    Form_Preferencias.Salvar_Valores
    Ajustar_Objectos_Na_Vertical
End Sub

Public Sub Repor_Icons_Visao()
    'Repor icons normais
    Icon_Visao(0).Picture = Form_Skin.Icon_Visao_Normal(0).Picture
    Icon_Visao(1).Picture = Form_Skin.Icon_Visao_Normal(1).Picture
    Icon_Visao(2).Picture = Form_Skin.Icon_Visao_Normal(2).Picture
End Sub

Private Sub Icon_Visao_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver imagem down
    If Grelha_Visivel <> Grelha_Musica Then Exit Sub
    Select Case Icon_Visao(Index).Index
        Case 0 'Simples
            Icon_Visao(0).Picture = Form_Skin.Icon_Visao_Down(0).Picture
            
        Case 1 'Avançada
            Icon_Visao(1).Picture = Form_Skin.Icon_Visao_Down(1).Picture
            
        Case 2 'Album art
            Icon_Visao(2).Picture = Form_Skin.Icon_Visao_Down(2).Picture
    End Select
End Sub

Private Sub Icon_Visao_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ver imagem up
    If Grelha_Visivel <> Grelha_Musica Then Exit Sub
    Select Case Icon_Visao(Index).Index
        Case 0 'Simples
            If visao_actual_da_biblioteca <> "0" Then
                Icon_Visao(0).Picture = Form_Skin.Icon_Visao_Normal(0).Picture
            Else
                Icon_Visao(0).Picture = Form_Skin.Icon_Visao_Down(0).Picture
            End If
            
        Case 1 'Avançada
            If visao_actual_da_biblioteca <> "1" Then
                Icon_Visao(1).Picture = Form_Skin.Icon_Visao_Normal(1).Picture
            Else
                Icon_Visao(1).Picture = Form_Skin.Icon_Visao_Down(1).Picture
            End If
            
        Case 2 'Album art
            If visao_actual_da_biblioteca <> "2" Then
                Icon_Visao(2).Picture = Form_Skin.Icon_Visao_Normal(2).Picture
            Else
                Icon_Visao(2).Picture = Form_Skin.Icon_Visao_Down(2).Picture
            End If
    End Select
End Sub

Private Sub Video_FullScreen()
    'Colocar o player em full screen
    'If Timer_Slider_Video.Enabled = False Then Exit Sub
    'Wmp.fullScreen = True
    If tela_video_fullscreen = False Then
        With Frame_Wmp
            .Height = Me.ScaleHeight
            .top = 0
            .Width = Me.ScaleWidth
            .left = 0
        End With
    
        With Wmp
            .Height = Frame_Wmp.ScaleHeight
            .top = 0
            .Width = Frame_Wmp.ScaleWidth
            .left = 0
        End With
    
        With Barra_Mini_Player
            .Visible = False
            .Height = Form_Skin.Fundo_Barra_Mini_Player.Height
            .top = Frame_Wmp.top + Frame_Wmp.ScaleHeight - .ScaleHeight - 50 - Barra_Informacoes.ScaleHeight
            .Width = Form_Skin.Fundo_Barra_Mini_Player.Width
            .left = (Frame_Wmp.ScaleWidth - .ScaleWidth) / 2
        End With
        
        tela_video_fullscreen = True
        Botao_Player_Mini(4).Picture = Form_Skin.Icon_FullScreen_Off.Picture
        Botao_Player_Mini(4).ToolTipText = Idioma_Button_Fullscreen_Off
        
        With Close_Wmp
            .top = 0 'Frame_Wmp.top + 10
            .left = Me.ScaleWidth - .ScaleWidth - 10
        End With
        
    Else
        With Frame_Wmp
            .Height = Barra_Lateral.ScaleHeight
            .top = Barra_Lateral.top
            .Width = Me.ScaleWidth - 2
            .left = 1
        End With
    
        With Wmp
            .Height = Frame_Wmp.ScaleHeight
            .top = 0
            .Width = Frame_Wmp.ScaleWidth
            .left = 0
        End With
    
        With Barra_Mini_Player
            .Visible = False
            .Height = Form_Skin.Fundo_Barra_Mini_Player.Height
            .top = Frame_Wmp.top + Frame_Wmp.ScaleHeight - .ScaleHeight - 50
            .Width = Form_Skin.Fundo_Barra_Mini_Player.Width
            .left = (Frame_Wmp.ScaleWidth - .ScaleWidth) / 2
        End With
        
        tela_video_fullscreen = False
        Botao_Player_Mini(4).Picture = Form_Skin.Icon_FullScreen_On.Picture
        Botao_Player_Mini(4).ToolTipText = Idioma_Button_Fullscreen_On
        
        With Close_Wmp
            .top = Frame_Wmp.top
            .left = Me.ScaleWidth - .ScaleWidth - 10
        End With
    End If
    
End Sub

Private Sub Image_Album_Click(Index As Integer)
    'Alterar a imagem do album activo
    Ocultar_menus
    
    Image_Album(album_activo).Picture = Form_Skin.Image_Album.Picture
    album_activo = Index
    Image_Album(Index).Picture = Form_Skin.Image_Album_Over.Picture
End Sub

Private Sub Image_Album_DblClick(Index As Integer)
    'Carregar as músicas do album seleciondo
    Index_Album = Index
    Lista_Pastas.ListIndex = Index_Album
    File_Ficheiros.Path = (Label_Directorio_Album(Index).Caption) + "\"
    Carregar_Grelha_Albuns
    
    'Ver qual é a capa selecionada
    Label_Album.Caption = Label_Nome_Album(album_activo).Text
End Sub

Private Sub Image_Album_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Mover os albuns através das teclas das setas
    If KeyCode = vbKeyRight Then
        Timer_Mover.Enabled = False
        
        'Alterar a imagem do album activo
        Image_Album(album_activo).Picture = Form_Skin.Image_Album.Picture
        If album_activo = Image_Album.count - 1 Then Image_Album(Image_Album.count - 1).Picture = Form_Skin.Image_Album_Over.Picture: Exit Sub
        album_activo = album_activo + 1
        Image_Album(album_activo).Picture = Form_Skin.Image_Album_Over.Picture
    End If
    
    'Mover o slide
    If KeyCode = vbKeyLeft Then
        Timer_Mover.Enabled = False
        
        'Alterar a imagem do album activo
        Image_Album(album_activo).Picture = Form_Skin.Image_Album.Picture
        If album_activo = 0 Then
            Image_Album(0).Picture = Form_Skin.Image_Album_Over.Picture
            Exit Sub
        End If
        album_activo = album_activo - 1
        Image_Album(album_activo).Picture = Form_Skin.Image_Album_Over.Picture
    End If
    
    'Atalho para carregar as músicas do album selecionado
    If KeyCode = vbKeyReturn Then
        'Image_Album_Click Index
        Image_Album_DblClick album_activo
    End If
End Sub

Private Sub Imagem_Votar_Click()
    'Votar no programa
    On Error GoTo Corrige_Erro
    If Label_Id_Programa.Caption = "" Or Label_Votos.Caption = "" Then Exit Sub
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    
    'Adicionar um voto á avaliação do programa
    Label_Votos.Caption = Val(Label_Votos.Caption) + 1
    servidor.Open "GET", "http://www.nikyts.com/nplayer/applibrary/" & "actualizaravaliacao.asp?id_programa=" & Label_Id_Programa.Caption & "&avaliacao=" & Label_Votos.Caption, False
    servidor.send

    'Actualizar a senha
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        With Form_Principal
            If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
                
                'Recebe confirmação de que o voto foi recebido com sucesso
                Mensagem_de_Aviso "Information", "O seu voto foi atribuido ao programa com sucesso!" & vbNewLine & "Obrigado pela sua contribuição."
                
                'Actualizar a avaliação do programa
                If Label_Votos.Caption = "1" Then
                    Label_Frame_Informacoes(2).Caption = Idioma_Label_Rate & ": " & Label_Votos.Caption
                Else
                    Label_Frame_Informacoes(2).Caption = Idioma_Label_Rate & ": " & Label_Votos.Caption
                End If
                Label_Frame_Informacoes(2).Width = Frame_Avaliacao.ScaleWidth
    
                'Avaliação do programa, Estrelas
                If Val(Label_Votos.Caption) < 20 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_0.Picture
                
                ElseIf Val(Label_Votos.Caption) >= 20 And Val(Label_Votos.Caption) < 40 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_1.Picture
                
                ElseIf Val(Label_Votos.Caption) >= 40 And Val(Label_Votos.Caption) < 60 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_2.Picture
                
                ElseIf Val(Label_Votos.Caption) >= 60 And Val(Label_Votos.Caption) < 80 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_3.Picture
                
                ElseIf Val(Label_Votos.Caption) >= 80 And Val(Label_Votos.Caption) < 100 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_4.Picture
                
                ElseIf Val(Label_Votos.Caption) > 100 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_5.Picture
                End If
            End If
        End With
    End If
    Set servidor = Nothing
    
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

Private Sub Label_Actualizar_Programa_Click()
    'Actualizar programa
    Shell App.Path & "\Options\Update.exe"
    Form_Principal.Botao_Fechar_Click
End Sub

Public Sub Repor_Cores_Labels_Separadores()
    'Procedimento para repor a cor original das labels dos separadores
    Dim i As Integer: For i = 3 To 14
        Label_Barra_Drive(i).ForeColor = vbWhite
    Next
End Sub

Public Sub Label_Barra_Drive_Click(Index As Integer)
    'Selecionar botões
    'On Error GoTo Corrige_Erro
    Repor_Objectos
    Ocultar_menus
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    
    Barra_Botoes_Musica.Visible = True
    Barra_Lateral.Visible = True
    
    Select Case Botao_Barra_Drive(Index).Index
        Case 3 'Contacto-------------------------------------------------------------------------------------------------------------
            If Utilizador_Logado = False Then
                Separador_Clicado = "my_contacts"
                Form_Login.Show vbModal
            Else
                Ocultar_Frame_Central
                Grelha_Contactos.Visible = True
                Set Grelha_Visivel = Grelha_Contactos
                Repor_Cores_Labels_Separadores
                Label_Barra_Drive(3).ForeColor = Azul
                Label_Botao(3).Visible = True
                Label_Botao(6).Visible = True
                Label_Botao(13).Visible = True
                Label_Botao(14).Visible = True
                Label_Botao(15).Visible = True
                Label_Botao(16).Visible = True
                Label_Contador.Caption = Grelha_Contactos.Rows - 1 & " " & Idioma_Total_Contactos
            End If
            
        Case 4 'Eventos-------------------------------------------------------------------------------------------------------------
            If Utilizador_Logado = False Then
                Separador_Clicado = "my_events"
                Form_Login.Show vbModal
            Else
                Ocultar_Frame_Central
                Grelha_Eventos.Visible = True
                Set Grelha_Visivel = Grelha_Eventos
                Repor_Cores_Labels_Separadores
                Label_Barra_Drive(4).ForeColor = Azul
                Label_Botao(3).Visible = True
                Label_Botao(6).Visible = True
                Label_Botao(17).Visible = True
                Label_Botao(18).Visible = True
                Label_Botao(19).Visible = True
                Label_Contador.Caption = Grelha_Eventos.Rows - 1 & " " & Idioma_Total_Eventos
            End If
            
        Case 5 'Ficheiros-------------------------------------------------------------------------------------------------------------
            Ocultar_Frame_Central
            Grelha_Ficheiros.Visible = True
            Set Grelha_Visivel = Grelha_Ficheiros
            Repor_Cores_Labels_Separadores
            Label_Barra_Drive(5).ForeColor = Azul
            Label_Botao(3).Visible = True
            Label_Botao(6).Visible = True
            Label_Contador.Caption = Grelha_Ficheiros.Rows - 1 & " " & Idioma_Total_Ficheiros_Online
            
        Case 6 'Compartilhados-------------------------------------------------------------------------------------------------------------

            
        Case 7 'Recomendo-------------------------------------------------------------------------------------------------------------
            Ocultar_Frame_Central
            Grelha_Recentes.Visible = True
            Set Grelha_Visivel = Grelha_Recentes
            Repor_Cores_Labels_Separadores
            Label_Barra_Drive(7).ForeColor = Azul
            Label_Botao(1).Visible = True
            Label_Botao(2).Visible = True
            Label_Botao(5).Visible = True
            Label_Botao(6).Visible = True
            Label_Botao(3).Visible = True
            
            'Efectuar pesquisa na base de dados consuante os dados introduzidos
            If Grelha_Recentes.Rows <= 1 Then
                Me.MousePointer = 11
                'Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
                servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "carregarrecentes.asp?total=" & "30", False '& Form_Opcoes.Text_Num_Max_Linhas.Text
                servidor.send
                'Verificar os dados acesso
                If servidor.responseText = "false" Then
                    Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
                ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
                    If servidor.readyState = 4 And servidor.Status = 200 Then
                        Grelha_Recentes.Clear
                        Formatar_Grelha Grelha_Recentes
                        Dados_Servidor_Musicas_Recentes servidor.responseText
                    End If
                End If
                Set servidor = Nothing
                Me.MousePointer = 0
            End If
            Label_Contador.Caption = Grelha_Recentes.Rows - 1 & " " & Idioma_Total_Musicas
            
        Case 8 'Favoritos-------------------------------------------------------------------------------------------------------------
            Ocultar_Frame_Central
            Grelha_Favoritos.Visible = True
            Set Grelha_Visivel = Grelha_Favoritos
            Repor_Cores_Labels_Separadores
            Label_Barra_Drive(8).ForeColor = Azul
            Label_Botao(1).Visible = True
            Label_Botao(2).Visible = True
            Label_Botao(5).Visible = True
            Label_Botao(6).Visible = True
            Label_Botao(3).Visible = True
            
            If Grelha_Favoritos.Rows <= 1 Then
                Me.MousePointer = 11
                'Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
                servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "musicasrecomendadas.asp?Recebe_Pesquisa=sim", False
                servidor.send
                'Verificar os dados acesso
                If servidor.responseText = "false" Then
                    Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
                ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
                    If servidor.readyState = 4 And servidor.Status = 200 Then
                        Grelha_Favoritos.Clear
                        Formatar_Grelha Grelha_Favoritos
                        Dados_Servidor_Musicas_Recomendadas servidor.responseText
                    End If
                End If
                Set servidor = Nothing
                Me.MousePointer = 0
            End If
            Label_Contador.Caption = Grelha_Favoritos.Rows - 1 & " " & Idioma_Total_Musicas
            
        Case 9 'A minha música-------------------------------------------------------------------------------------------------------------
            If Utilizador_Logado = False Then
                Separador_Clicado = "my_music"
                Form_Login.Show vbModal
            Else
                Ocultar_Frame_Central
                Grelha_Minha_Musica.Visible = True
                Set Grelha_Visivel = Grelha_Minha_Musica
                Repor_Cores_Labels_Separadores
                Label_Barra_Drive(9).ForeColor = Azul
                Label_Botao(1).Visible = True
                Label_Botao(4).Visible = True
                Label_Botao(5).Visible = True
                Label_Botao(6).Visible = True
                Label_Botao(3).Visible = True
                Label_Contador.Caption = Grelha_Minha_Musica.Rows - 1 & " " & Idioma_Total_Musicas
            End If
            
        Case 10 'Resultado da pesquisa-------------------------------------------------------------------------------------------------------------
            Ocultar_Frame_Central
            Grelha_Loja.Visible = True
            Set Grelha_Visivel = Grelha_Loja
            Repor_Cores_Labels_Separadores
            Label_Barra_Drive(10).ForeColor = Azul
            Label_Botao(1).Visible = True
            Label_Botao(2).Visible = True
            Label_Botao(5).Visible = True
            Label_Botao(6).Visible = True
            Label_Botao(3).Visible = True
            Label_Contador.Caption = Grelha_Loja.Rows - 1 & " " & Idioma_Total_Musicas
    
        Case 11 'Comunidade-------------------------------------------------------------------------------------------------------------
            Ocultar_Frame_Central
            Grelha_Comunidade.Visible = True
            Set Grelha_Visivel = Grelha_Comunidade
            Repor_Cores_Labels_Separadores
            Label_Barra_Drive(11).ForeColor = Azul
            Label_Botao(3).Visible = True
            Label_Botao(5).Visible = True
            Label_Botao(6).Visible = True
            Label_Botao(9).Visible = True
            Label_Botao(10).Visible = True
            Label_Botao(11).Visible = True
            
            If Grelha_Comunidade.Rows <= 1 Then
                Me.MousePointer = 11
                'Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
                servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "comunidade.asp", False
                servidor.send
                'Verificar os dados acesso
                If servidor.responseText = "false" Then
                    Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
                ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
                    If servidor.readyState = 4 And servidor.Status = 200 Then
                        Grelha_Comunidade.Clear
                        Formatar_Grelha_Comunidade
                        Dados_Servidor_Comunidade servidor.responseText
                    End If
                End If
                Set servidor = Nothing
                Me.MousePointer = 0
            End If
            Label_Contador.Caption = Grelha_Comunidade.Rows - 1 & " " & Idioma_Total_Utilizadores
            
        Case 12 'Amigos
            If Utilizador_Logado = False Then
                Separador_Clicado = "friends"
                Form_Login.Show vbModal
            Else
                Ocultar_Frame_Central
                Grelha_Amigos.Visible = True
                Set Grelha_Visivel = Grelha_Amigos
                Repor_Cores_Labels_Separadores
                Label_Barra_Drive(12).ForeColor = Azul
                Label_Botao(3).Visible = True
                Label_Botao(4).Visible = True
                Label_Botao(5).Visible = True
                Label_Botao(6).Visible = True
                Label_Botao(9).Visible = True
                Label_Botao(11).Visible = True
                Label_Botao(12).Visible = True
                Label_Contador.Caption = Grelha_Amigos.Rows - 1 & " " & Idioma_Total_Amigos
            End If
            
        Case 13 'Mensagens
            If Utilizador_Logado = False Then
                Separador_Clicado = "messages"
                Form_Login.Show vbModal
            Else
                Ocultar_Frame_Central
                Grelha_Mensagens.Visible = True
                Set Grelha_Visivel = Grelha_Mensagens
                Repor_Cores_Labels_Separadores
                Label_Barra_Drive(13).ForeColor = Azul
                Label_Botao(3).Visible = True
                Label_Botao(6).Visible = True
                Label_Contador.Caption = Grelha_Mensagens.Rows - 1 & " " & Idioma_Total_Messagens
            End If
            
        Case 14 'Ver perfil
            Ocultar_Frame_Central
            Frame_Perfil.Visible = True
            Repor_Cores_Labels_Separadores
            Label_Barra_Drive(14).ForeColor = Azul
            Label_Botao(11).Visible = True
            Label_Botao(10).Visible = True
            Label_Contador.Caption = ""
    End Select
    
    'Verificar_Contador
    Text_Pesquisar_Musica.Text = Empty
    Ajustar_Objectos_Na_Horizontal
    Ajustar_Objectos_Na_Vertical
    
    If Grelha_Amigos.Visible = True Then Label_Botao(4).left = Label_Botao(11).left + Label_Botao(11).Width + 20
            
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

Private Sub Dados_Servidor_Comunidade(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/pesquisa/resultado")
        Dim i As Integer: i = Grelha_Comunidade.Rows
        
        For Each node In nodeList
            DoEvents
            Grelha_Comunidade.Rows = Grelha_Comunidade.Rows + 1
            Dim Data_Dia, Data_Mes, Dia_Ano As String
            
            If Not IsEmpty(node.selectSingleNode("utilizador")) Then Grelha_Comunidade.TextMatrix(i, 1) = node.selectSingleNode("utilizador").Text
            If Not IsEmpty(node.selectSingleNode("nome")) Then Grelha_Comunidade.TextMatrix(i, 2) = node.selectSingleNode("nome").Text
            If Not IsEmpty(node.selectSingleNode("genero")) Then Grelha_Comunidade.TextMatrix(i, 3) = node.selectSingleNode("genero").Text
            If Not IsEmpty(node.selectSingleNode("dia")) Then Data_Dia = node.selectSingleNode("dia").Text
            If Not IsEmpty(node.selectSingleNode("mes")) Then Data_Mes = node.selectSingleNode("mes").Text
            If Not IsEmpty(node.selectSingleNode("ano")) Then Dia_Ano = node.selectSingleNode("ano").Text
            If Data_Dia <> Empty Then
                Grelha_Comunidade.TextMatrix(i, 4) = Data_Dia & "-" & Data_Mes & "-" & Dia_Ano
            End If
            If Not IsEmpty(node.selectSingleNode("pais")) Then Grelha_Comunidade.TextMatrix(i, 5) = node.selectSingleNode("pais").Text
            If Not IsEmpty(node.selectSingleNode("foto")) Then Grelha_Comunidade.TextMatrix(i, 6) = node.selectSingleNode("foto").Text
            If Not IsEmpty(node.selectSingleNode("email")) Then Grelha_Comunidade.TextMatrix(i, 7) = node.selectSingleNode("email").Text
            i = i + 1
        Next
        
    Else
        'Caso nenhum encontre num ficheiro referente á pesquisa efectuada
        Grelha_Comunidade.Rows = 1
        'Label_Contador_Frame_Top_Pesquisa.Caption = "nenhum utilizador registado"
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Private Sub Dados_Servidor_Musicas_Recomendadas(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/pesquisa/resultado")
        Dim i As Integer: i = Grelha_Favoritos.Rows
        
        For Each node In nodeList
            DoEvents
            Grelha_Favoritos.Rows = Grelha_Favoritos.Rows + 1

            If Not IsEmpty(node.selectSingleNode("servidor")) Then Grelha_Favoritos.TextMatrix(i, 0) = node.selectSingleNode("servidor").Text
            If Not IsEmpty(node.selectSingleNode("titulo")) Then Grelha_Favoritos.TextMatrix(i, 1) = node.selectSingleNode("titulo").Text
            If Not IsEmpty(node.selectSingleNode("artista")) Then Grelha_Favoritos.TextMatrix(i, 2) = node.selectSingleNode("artista").Text
            If Not IsEmpty(node.selectSingleNode("data")) Then Grelha_Favoritos.TextMatrix(i, 3) = node.selectSingleNode("data").Text
            If Not IsEmpty(node.selectSingleNode("adicionado")) Then Grelha_Favoritos.TextMatrix(i, 4) = node.selectSingleNode("adicionado").Text
            If Not IsEmpty(node.selectSingleNode("id")) Then Grelha_Favoritos.TextMatrix(i, 5) = node.selectSingleNode("id").Text
            i = i + 1
        Next
    Else
        'Caso nenhum encontre num ficheiro referente á pesquisa efectuada
        Grelha_Favoritos.Rows = 1
        'Label_Contador_Frame_Top_Pesquisa.Caption = "nenhuma pesquisa encontrada"
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Private Sub Dados_Servidor_Musicas_Recentes(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/pesquisa/resultado")
        Dim i As Integer: i = Grelha_Recentes.Rows
        
        For Each node In nodeList
            DoEvents
            Grelha_Recentes.Rows = Grelha_Recentes.Rows + 1

            If Not IsEmpty(node.selectSingleNode("servidor")) Then Grelha_Recentes.TextMatrix(i, 0) = node.selectSingleNode("servidor").Text
            If Not IsEmpty(node.selectSingleNode("titulo")) Then Grelha_Recentes.TextMatrix(i, 1) = node.selectSingleNode("titulo").Text
            If Not IsEmpty(node.selectSingleNode("artista")) Then Grelha_Recentes.TextMatrix(i, 2) = node.selectSingleNode("artista").Text
            If Not IsEmpty(node.selectSingleNode("data")) Then Grelha_Recentes.TextMatrix(i, 3) = node.selectSingleNode("data").Text
            If Not IsEmpty(node.selectSingleNode("adicionado")) Then Grelha_Recentes.TextMatrix(i, 4) = node.selectSingleNode("adicionado").Text
            If Not IsEmpty(node.selectSingleNode("id")) Then Grelha_Recentes.TextMatrix(i, 5) = node.selectSingleNode("id").Text
            'If i > 19 Then Exit For
            i = i + 1
        Next
    Else
        'Caso nenhum encontre num ficheiro referente á pesquisa efectuada
        Grelha_Recentes.Rows = 1
        'Label_Contador_Frame_Top_Pesquisa.Caption = "nenhuma pesquisa encontrada"
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Private Sub Label_Barra_Drive_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    Ocultar_menus
    
    Select Case Label_Barra_Drive(Index).Index
        Case 3 'Contacto
            Botao_Barra_Drive(3).Picture = Form_Skin.Botao_Barra_Down.Picture
        
        Case 4 'Agenda
            Botao_Barra_Drive(4).Picture = Form_Skin.Botao_Barra_Down.Picture
        
        Case 5 'Mensagens
            Botao_Barra_Drive(5).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 6 'Ficheiros
            Botao_Barra_Drive(6).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 7 'Recomendo
            Botao_Barra_Drive(7).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 8 'Favoritos
            Botao_Barra_Drive(8).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 9 'A minha música
            Botao_Barra_Drive(9).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 10 'Resultado da pesquisa
            Botao_Barra_Drive(10).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 11 'Comunidade
            Botao_Barra_Drive(11).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 12 'Os meus amigos
            Botao_Barra_Drive(12).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 13 'Mensagens
            Botao_Barra_Drive(13).Picture = Form_Skin.Botao_Barra_Down.Picture
            
        Case 14 'Ver perfil
            Botao_Barra_Drive(14).Picture = Form_Skin.Botao_Barra_Down.Picture
    End Select
End Sub

Private Sub Label_Barra_Drive_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    Ocultar_menus
    
    Select Case Label_Barra_Drive(Index).Index
        Case 3 'Contacto
            Botao_Barra_Drive(3).Picture = Form_Skin.Botao_Barra_Normal.Picture
        
        Case 4 'Agenda
            Botao_Barra_Drive(4).Picture = Form_Skin.Botao_Barra_Normal.Picture
        
        Case 5 'Mensagens
            Botao_Barra_Drive(5).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 6 'Ficheiros
            Botao_Barra_Drive(6).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 7 'Recomendo
            Botao_Barra_Drive(7).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 8 'Favoritos
            Botao_Barra_Drive(8).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 9 'A minha música
            Botao_Barra_Drive(9).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 10 'Resultado da pesquisa
            Botao_Barra_Drive(10).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 11 'Comunidade
            Botao_Barra_Drive(11).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 12 'Os meus amigos
            Botao_Barra_Drive(12).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 13 'Mensagens
            Botao_Barra_Drive(13).Picture = Form_Skin.Botao_Barra_Normal.Picture
            
        Case 14 'Ver perfil
            Botao_Barra_Drive(14).Picture = Form_Skin.Botao_Barra_Normal.Picture
    End Select
End Sub

Public Sub Label_Botao_Click(Index As Integer)
    'Eventos disponiveis
    'On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    Repor_Objectos
    
    Select Case Label_Botao(Index).Index
        Case 0 'Abrir formulário de actualização da biblioteca do tópico selecionado
            With Form_Actualizar_Biblioteca
                If Grelha_Reproduzida = Grelha_Visivel Then Botao_Pausa_Click
                
                If Grelha_Visivel = Grelha_Musica Then
                    .Biblioteca_Selecionada = "Musica"
                    .Pesquisar_pela_Extensao = "mp3"
                    .File1.Pattern = "*.mp3"
                Else
                    .Biblioteca_Selecionada = "Filmes"
                    .Pesquisar_pela_Extensao = "avi"
                    .File1.Pattern = "*.avi"
                End If
                
                .Text_Pesquisar_Musica.Text = ""
                .Label_Ficheiro.Caption = ""
                .Label_Contador.Caption = ""
                .Show vbModal
            End With
        
        '-----------------------------------------------------------------------------------------------------
        Case 1 'Transferir a música selecionada
            If Grelha_Visivel.Rows = 1 Then Exit Sub
            Dim novo_download As New Form_Download
            With novo_download
                .Text_Servidor.Text = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0)
                .Show
            End With
        
        '-----------------------------------------------------------------------------------------------------
        Case 2 'Efectuar uma nova pesquisa
            If Utilizador_Logado = False Then
                Separador_Clicado = "adicionar_my_music"
                Form_Login.Show vbModal
            Else
                'Adicionar as músicas selecionadas para onde se pretende
                Adicionar_Musica_Selecionada Grelha_Minha_Musica
                Adicionar_na_Tabela_Lista
            End If
        
        '-----------------------------------------------------------------------------------------------------
        Case 3 'Conta do utilizador
            Label_Botao(3).FontUnderline = False
            Label_Botao(6).FontUnderline = False
            If Utilizador_Logado = False Then
                Form_Criar.Show vbModal
            Else
                Form_Perfil.Show vbModal
            End If
    
        '-----------------------------------------------------------------------------------------------------
        Case 4 'Remover
            Remover_Ficheiros_Selecionados
        
        '-----------------------------------------------------------------------------------------------------
        Case 5 'Adicionar um novo link de música á base de dados
            If Utilizador_Logado = False Then
                Separador_Clicado = "adicionar_link"
                Form_Login.Show vbModal
            Else
                With Form_Adicionar
                    .Text_Data.Text = Date
                    .Show vbModal
                End With
            End If
    
        '-----------------------------------------------------------------------------------------------------
        Case 6 'Iniciar/ terminar sessão
            Label_Botao(6).FontUnderline = False
            If Utilizador_Logado = False Then
                Form_Login.Show vbModal
            Else
                'Encerrar sessão do utilizador
                Label_Botao(6).Caption = ReadINI("Main", "Button_Login", Localizacao_Ficheiro_Lingua)
                Label_Botao(3).Caption = ReadINI("Main", "Button_Create_Account", Localizacao_Ficheiro_Lingua)
                Ajustar_Objectos_Na_Horizontal
                Utilizador_Logado = False
                Form_Perfil.Label_Utilizador.Caption = ""
                Form_Perfil.Label_Minha_Senha.Caption = ""
                Grelha_Minha_Musica.Clear
                Formatar_Grelha Grelha_Minha_Musica
                Formatar_Grelha_Contactos
                Formatar_Grelha_Eventos
                Formatar_Grelha_Mensagens
                Formatar_Grelha_Ficheiros
                Formatar_Grelha_Amigos
                
                Unload Form_Perfil
                Ajustar_Objectos_Na_Horizontal
                'Caso seja músicas da minha pasta que estejam a ser reproduzidas então para o player
                If Grelha_Reproduzida = Grelha_Minha_Musica Then
                    Parar_o_Player
                End If
                Botao_Mensagens.Visible = False
                Label_Mensagens.Caption = 0
                Label_Contador.Caption = ""
                Verificar_Contador
            End If
            
        '-----------------------------------------------------------------------------------------------------
        Case 7 'Nova lista
            Grelha_Listas.Clear
            Formatar_Grelha_Musica Grelha_Listas
            Grelha_Listas.Rows = 1
            
            Dim retVal As String
            retVal = Dir(App.Path & "\Library\Playlist\" & nome_nova_lista)
            
            'Verificar se existe alguma lista com este nome
            If retVal = nome_nova_lista Then 'já existe
               numero_lista = numero_lista + 1
               nome_nova_lista = Idioma_Name_Of_New_Playlist & " " & numero_lista & ".ini"
               Call Label_Botao_Click(7)
               Exit Sub
            Else 'não existe
                Dim classe_nova_lista As clsFlexSettings
                Set classe_nova_lista = New clsFlexSettings
                Set classe_nova_lista.FlexGrid = Grelha_Listas
                classe_nova_lista.SaveSettings App.Path & "\Library\Playlist\" & nome_nova_lista, True, True, True, True
                Set classe_nova_lista = Nothing
            End If
                        
            File_Lista.Refresh
            
            If File_Lista.ListCount = 0 Then Exit Sub
            If Label_Topico_Lista(0).Visible = False Then
                Label_Topico_Lista(0).Caption = Espaco & Mid(nome_nova_lista, 1, InStrRev(nome_nova_lista, ".") - 1)
                Label_Topico_Lista(0).Visible = True
                Shape_Topico_Lista(0).Visible = True
                Icon_Topico_Lista(0).Visible = True
                Label_Topico_Lista(0).ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
                Icon_Topico_Lista(0).Picture = Form_Skin.Icon_Topico_Lista_Over.Picture
                
            Else
                Repor_a_Cor_Dos_Topicos
                
                Dim Objecto As Integer
                Objecto = Label_Topico_Lista.count '+ 1
                Load Shape_Topico_Lista(Objecto)
                Shape_Topico_Lista(Objecto).Move Shape_Topico_Lista(Objecto - 1).left, Shape_Topico_Lista(Objecto - 1).top + Shape_Topico_Lista(Objecto - 1).Height
                Shape_Topico_Lista(Objecto).Visible = True
                
                Load Label_Topico_Lista(Objecto)
                Label_Topico_Lista(Objecto).Move Label_Topico_Lista(Objecto - 1).left, Shape_Topico_Lista(Objecto).top + ((Shape_Topico_Lista(Objecto).Height - Label_Topico_Lista(Objecto).Height) / 2)
                Label_Topico_Lista(Objecto).Visible = True
                Label_Topico_Lista(Objecto).Caption = Espaco & Mid(nome_nova_lista, 1, InStrRev(nome_nova_lista, ".") - 1)
                
                Load Icon_Topico_Lista(Objecto)
                Icon_Topico_Lista(Objecto).Move Icon_Topico_Lista(Objecto - 1).left, Shape_Topico_Lista(Objecto).top + ((Shape_Topico_Lista(Objecto).Height - Icon_Topico_Lista(Objecto).Height) / 2)
                Icon_Topico_Lista(Objecto).Visible = True
                
                Shape_Topico_Lista(Objecto).ZOrder 1
                Label_Topico_Lista(Objecto).ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
                Icon_Topico_Lista(Objecto).Picture = Form_Skin.Icon_Topico_Lista_Over.Picture
            End If
            
            'Ajustar o tamanho da Frame_Separador_Barra_Lateral(2)
            Frame_Separador_Barra_Lateral(2).Height = (Shape_Topico_Lista.count * Separador_Barra_Lateral(0).ScaleHeight) + Separador_Barra_Lateral(0).ScaleHeight
            Label_Topico_Lista_DblClick (Objecto)
            
        '-----------------------------------------------------------------------------------------------------
        Case 8 'Guardar lista
            If Grelha_Listas.Rows <= 1 Then Exit Sub
            Dim nome_lista As String: nome_lista = Replace(Label_Topico_Lista(index_lista_selecionada).Caption, Espaco, "")
        
            If Grelha_Listas.Rows > 1 Then
                Personalizar_Grid Grelha_Listas
                Dim nova_classe As clsFlexSettings
                Set nova_classe = New clsFlexSettings
                Set nova_classe.FlexGrid = Grelha_Listas
                nova_classe.SaveSettings App.Path & "\Library\Playlist\" & nome_lista & ".ini", True, True, True, True
                Set nova_classe = Nothing
            End If
            
        '-----------------------------------------------------------------------------------------------------
        Case 9 'Ver o perfil
            If Grelha_Comunidade.Rows <= 1 Then Exit Sub
            If Grelha_Comunidade.TextMatrix(Grelha_Comunidade.Row, 8) = "Privado" Then
                Mensagem_de_Aviso "Information", ReadINI("Message", "Info_Profile_Private1", Localizacao_Ficheiro_Lingua) & vbNewLine & ReadINI("Message", "Info_Profile_Private2", Localizacao_Ficheiro_Lingua)
            Else
                With Grelha_Comunidade
                    Label_Nickname.Caption = .TextMatrix(.Row, 1)
                    Label_Perfil(0).Caption = .TextMatrix(.Row, 2)
                    Label_Perfil(1).Caption = .TextMatrix(.Row, 7)
                    Label_Perfil(2).Caption = .TextMatrix(.Row, 3)
                    Label_Perfil(3).Caption = .TextMatrix(.Row, 4)
                    Label_Perfil(4).Caption = .TextMatrix(.Row, 5)
                End With
                Label_Barra_Drive_Click (14)
                Botao_Barra_Drive(14).Visible = True
                Label_Barra_Drive(14).Visible = True
            End If
            
        '-----------------------------------------------------------------------------------------------------
        Case 10 'Enviar convite
            If Utilizador_Logado = False Then
                Separador_Clicado = "send_invitation"
                Form_Login.Show vbModal
            Else
                Me.MousePointer = 11
                'Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
                servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "enviarmensagem.asp?Utilizador=" & Grelha_Comunidade.TextMatrix(Grelha_Comunidade.Row, 1) & "&Assunto=" & "Olá!" & "&Mensagem=" & "O utilizador " & "Nikyts" & " enviou-lhe um convite de amizade." & "&Data=" & Date & "&Anexo=" & "" & "&Visualizada=" & "0", False
                servidor.send
                'Verificar os dados acesso
                If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
                    If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
                        Mensagem_de_Aviso "Invitation", ReadINI("Message", "Info_Invitation_Sent_Successfully", Localizacao_Ficheiro_Lingua)
                        Me.MousePointer = 0
                    End If
                End If
            End If
        
        '-----------------------------------------------------------------------------------------------------
        Case 11 'O que ando a ouvir
            
        
        '-----------------------------------------------------------------------------------------------------
        Case 12 'Enviar mensagem
            
    End Select
    
    
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

Private Sub Label_Botao_Frame_Informacoes_Click(Index As Integer)
    'Efectuar operações da frame informações do programa
    'On Error GoTo Corrige_Erro
    Ocultar_menus
    
    Select Case Label_Botao_Frame_Informacoes(Index).Index
        Case 0 'Tansferir-----------------------------------------------------------------------------------------------------------
            Select Case Label_Botao_Frame_Informacoes(0).Caption
                Case Idioma_Button_Transfer_Program
                    Botao_Frame_Informacoes(2).Enabled = False
                    Label_Botao_Frame_Informacoes(2).Enabled = False
                    Botao_Frame_Informacoes(1).Enabled = True
                    Label_Botao_Frame_Informacoes(1).Enabled = True
                
                    Label_Frame_Informacoes(6).Caption = "Transferindo o ficheiro..."
                    Barra_Estado_Visivel True
                    Me.MousePointer = 11
                    Verificar_Pastas
                    Botao_Frame_Informacoes(0).Visible = False
                    ProgressBar1.Visible = True
                    Text_Servidor.Text = "http://www.nikyts.com/nplayer/applibrary/programas/" & Label_Frame_Informacoes(3).Caption
                    dl.DownloadFile Text_Servidor.Text, App.Path & "\Programs\" & Label_Frame_Informacoes(3).Caption '& GetFileName(Label_Frame_Informacoes(3).Caption)
                    On Error GoTo 0 'Tratamento de erros
                    
                Case Idioma_Button_Remove_Program
                    Me.MousePointer = 11
                    'Remover a pasta, sub-pasta e respectivos ficheiros referentes ao programa
                    DeleteFolderTree App.Path & "\Programs\" & Label_Frame_Informacoes(0).Caption
                    Label_Frame_Informacoes(3).Caption = Label_Frame_Informacoes(0).Caption & ".zip"
                    Label_Botao_Frame_Informacoes(0).Caption = Idioma_Button_Transfer_Program
                    Barra_Estado_Visivel False
                    Me.MousePointer = 0
            End Select
            
        Case 1 'Cancelar-------------------------------------------------------------------------------------------------------------
            Barra_Estado_Visivel False
            dl.cancel
            ProgressBar1.Value = 0
            Label_Botao_Frame_Informacoes(0).Caption = Idioma_Button_Transfer_Program
            
            Botao_Frame_Informacoes(0).Visible = True
            ProgressBar1.Visible = False
            Barra_Estado_Visivel False
            Label_Frame_Informacoes(6).Caption = "Operação cancelada"
            Me.MousePointer = 0
            
        Case 2 'Executar-------------------------------------------------------------------------------------------------------------
            Shell App.Path & "\Programs\" & Label_Frame_Informacoes(0).Caption & "\" & Label_Frame_Informacoes(0).Caption & ".exe"
    End Select
    
Exit Sub
Corrige_Erro:

End Sub

Private Sub Label_Botao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar a label
    If Index = 20 Then Exit Sub
    Repor_Objectos
    Label_Botao(Index).FontUnderline = True
End Sub

Private Sub Label_Carregar_Favoritos_Click()
    'Carregar favoritos
    Criterio = "0"
    Verifica_Rs_Musica
    Rs_Musica.Open "select * from Tabela_Musica where Classificacao <> '" & Criterio & "' order by Artista Asc", Cnn_Biblioteca
    
    If Rs_Musica.RecordCount > 0 Then
        Me.MousePointer = 11
        Dim i As Integer
        With Grelha_Lista_Em_Reproducao
            .Rows = 1
            i = 1
            Do While Not Rs_Musica.EOF
                .Rows = Rs_Musica.RecordCount + 1
                If Rs_Musica(0).Value <> "" Then .TextMatrix(i, 0) = Rs_Musica(0).Value
                If Rs_Musica(1).Value <> "" Then .TextMatrix(i, 1) = Rs_Musica(1).Value
                If Rs_Musica(2).Value <> "" Then .TextMatrix(i, 2) = Rs_Musica(2).Value
                If Rs_Musica(3).Value <> "" Then .TextMatrix(i, 3) = Rs_Musica(3).Value
                If Rs_Musica(4).Value <> "" Then .TextMatrix(i, 4) = Rs_Musica(4).Value
                If Rs_Musica(5).Value <> "" Then .TextMatrix(i, 5) = Rs_Musica(5).Value
                If Rs_Musica(6).Value <> "" Then .TextMatrix(i, 6) = Rs_Musica(6).Value
                If Rs_Musica(7).Value <> "" Then .TextMatrix(i, 7) = Rs_Musica(7).Value
                If Rs_Musica(8).Value <> "" Then .TextMatrix(i, 8) = Rs_Musica(8).Value
                If Rs_Musica(9).Value <> "" Then .TextMatrix(i, 9) = Rs_Musica(9).Value
                i = i + 1
                Rs_Musica.MoveNext
            Loop
        End With
    
        Grelha_Lista_Em_Reproducao.Visible = True
        Label_Carregar_Favoritos.Visible = False
        Me.MousePointer = 0
    End If
End Sub

Private Sub Label_Descricao_Click(Index As Integer)
    'Atalho
    Pic_Linha_Click (Index)
End Sub

Private Sub Label_Evento_Click(Index As Integer)
    'Ocultar a frame evento
    Select Case Label_Evento(Index).Index
        Case 3
            Frame_Evento.Visible = False
    End Select
End Sub

Private Sub Label_Executar_Programa_Click(Index As Integer)
    'Executar o programa automaticamente
    On Error GoTo Corrige_Erro
    Shell App.Path & "\Programs\" & Label_Nome(Index).Caption & "\" & Label_Nome(Index).Caption & ".exe"
      
Exit Sub
Corrige_Erro:
Select Case err.Number
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Label_Faixa_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Label_Faixa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Caso o nome da faixa em reprodução seja demasiado grande que não caiba no visor de reprodução á a possibilidade de o
    'utilizador visualiza-la atraves do mouse
    If Label_Faixa.Caption <> "" Then Label_Faixa.ToolTipText = Label_Faixa.Caption
End Sub

Private Sub Label_Frame_Informacoes_Click(Index As Integer)
    'Abrir site do programa selecionado
    Select Case Label_Frame_Informacoes(Index).Index
        Case 4 'Site do programa
            If Label_Site_Programa.Caption = Empty Then Exit Sub
            Call ShellExecute(0, "open", Label_Site_Programa.Caption, vbNullString, vbNullString, SW_NORMAL)
    End Select
End Sub

Private Sub Label_Frame_Programas_Click(Index As Integer)
    'Selecionar separadores
    Select Case Label_Frame_Programas(Index).Index
        Case 1 'Instalados
            Frame_Programas_Home.Visible = False
            Frame_Lista.Visible = False
            Frame_Informacoes.Visible = False
            Label_Botao(20).Visible = False: Imagem_Votar.Visible = False
            
        
        Case 2 'Categoria selecionada
            If Frame_Lista.Visible = True Then Exit Sub
            Frame_Programas_Home.Visible = False
            Frame_Lista.Visible = True
            Frame_Informacoes.Visible = False
            Label_Botao(20).Visible = False: Imagem_Votar.Visible = False
            
        Case 3 'Programa selecionado
            If Frame_Informacoes.Visible = True Then Exit Sub
            Frame_Programas_Home.Visible = False
            Frame_Lista.Visible = False
            Frame_Informacoes.Visible = True
            Label_Botao(20).Visible = True: Imagem_Votar.Visible = True
    End Select
End Sub

Private Sub Label_Legendas_Click()
    'Legendas on-line
    Procurar_Legendas
End Sub

Private Sub Label_Legendas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mostar a imagem down
    Botao_Legendas.Picture = Form_Skin.Button_Menu_Down.Picture
    Icon_Legendas.Picture = Form_Skin.Icon_Subtitles_Down.Picture
End Sub

Private Sub Label_Legendas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Legendas on-line
    Botao_Legendas.Picture = Form_Skin.Button_Menu_Normal.Picture
    Icon_Legendas.Picture = Form_Skin.Icon_Subtitles_Normal.Picture
End Sub

Private Sub Label_Mais_Informacoes_Click(Index As Integer)
    'Ver informações do programa
    Selecionar_Programa Label_Nome(Index), Label_Descricao(Index), Label_Programa(Index), Label_Downloads(Index), Label_Observacoes(Index), _
        Label_Icon(Index), Label_Logotipo(Index), Label_Tela(Index), Label_Avaliacao(Index), Label_Id(Index), Label_Site(Index)
        
    'Carregar o logotipo e respectiva tela do programa
    Image_Logo.Picture = Logotipo_Programa(Index).Picture
    Image_Tela.Picture = Tela_Programa(Index).Picture
    
    With Label_Frame_Informacoes(5)
        .left = Label_Frame_Informacoes(0).left + Label_Frame_Informacoes(0).Width + 10
    End With
End Sub

Public Sub Selecionar_Programa(Label_Nome As Label, Label_Descricao As Label, Label_Programa As Label, Label_Downloads As Label, _
                                Label_Observacoes As Label, Label_Icon As Label, Label_Logotipo As Label, Label_Tela As Label, _
                                Label_Avaliacao As Label, Label_Id As Label, Label_Site As Label)
    'Procedimento para escolher a categoria do programa
    Ocultar_Frame_Central
    Frame_Programas.Visible = True
    Frame_Informacoes.Visible = True
    Frame_Lista.Visible = False
    Label_Botao(20).Visible = True
    Imagem_Votar.Visible = True
    Me.MousePointer = 11
    
    Label_Frame_Programas(3).Caption = Label_Nome.Caption
    Label_Frame_Programas(3).Visible = True
    With Separador_Frame_Programas(3)
        Separador_Frame_Programas(3).Stretch = True
        Separador_Frame_Programas(3).Width = Label_Frame_Programas(3).Width + 40
        Separador_Frame_Programas(3).left = Separador_Frame_Programas(2).left + Separador_Frame_Programas(2).Width
        Label_Frame_Programas(3).left = Separador_Frame_Programas(3).left + 20
        .Visible = True
    End With

    Label_Frame_Informacoes(0).Caption = Label_Nome.Caption
    Label_Frame_Informacoes(1).Caption = Label_Descricao.Caption
    
    Label_Transferencias.Caption = Label_Downloads.Caption
    If Val(Label_Downloads.Caption) = 1 Then
         Label_Frame_Informacoes(5).Caption = "(" & Label_Downloads.Caption & " download)"
    Else
        Label_Frame_Informacoes(5).Caption = "(" & Label_Downloads.Caption & " downloads)"
    End If
    Label_Frame_Informacoes(5).left = Label_Frame_Informacoes(0).left + Label_Frame_Informacoes(0).Width + 5
    
    Label_Frame_Informacoes(3).Caption = Label_Programa.Caption
    Text_Informacao.Text = Label_Observacoes.Caption
    
    'Avaliação do programa, Estrelas
    If Val(Label_Avaliacao.Caption) < 20 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_0.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) >= 20 And Val(Label_Avaliacao.Caption) < 40 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_1.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) >= 40 And Val(Label_Avaliacao.Caption) < 60 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_2.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) >= 60 And Val(Label_Avaliacao.Caption) < 80 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_3.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) >= 80 And Val(Label_Avaliacao.Caption) < 100 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_4.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) > 100 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_5.Picture
    End If
    
    'Total de avaliações
    Label_Votos.Caption = Label_Avaliacao.Caption
    If Label_Avaliacao.Caption = "1" Then
        Label_Frame_Informacoes(2).Caption = Idioma_Label_Rate & ": " & Label_Avaliacao.Caption
    Else
        Label_Frame_Informacoes(2).Caption = Idioma_Label_Rate & ": " & Label_Avaliacao.Caption
    End If
    
    'Receber o id do programa para depois obter os comentários sobre o mesmo
    Label_Id_Programa.Caption = Label_Id
    
    Label_Site_Programa.Caption = Label_Site
    If Label_Site_Programa.Caption = Empty Then
        Label_Frame_Informacoes(4).Visible = False
    Else
        Label_Frame_Informacoes(4).Visible = True
    End If
    
    Verificar_Se_Programa_Existe
    Me.MousePointer = 0
End Sub

Private Sub Label_Mensagens_Click()
    'Atalho para ver as mensagens
    Label_Topico_MusicLink_Click
    Label_Barra_Drive_Click (13)
End Sub

Private Sub Label_Mensagens_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mostar a imagem down
    Botao_Mensagens.Picture = Form_Skin.Button_Menu_Standard_Down.Picture
    Icon_Mensagens.Picture = Form_Skin.Icon_Mensagem_Down.Picture
End Sub

Private Sub Label_Mensagens_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Legendas on-line
    Botao_Mensagens.Picture = Form_Skin.Button_Menu_Standard_Normal.Picture
    Icon_Mensagens.Picture = Form_Skin.Icon_Mensagem_Normal.Picture
End Sub

Private Sub Label_Menu_Click(Index As Integer)
    'Ver menu consoante o menu selecionado
    Dim i As Integer: For i = 0 To Menu_Check.count - 1
        Menu_Check(i).Picture = Form_Skin.Menu_Check_Normal.Picture
    Next
    
    Select Case Label_Menu(Index).Index
        Case 0 'Ficheiro-------------------------------------------------------------------------------------------
            If Frame_Menu(0).Visible = False Then
                Frame_Menu(0).Visible = True
                Menu_Activo = True
                Shape_Menu(0).BackStyle = 1
                Label_Menu(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
            Else
                Ocultar_menus
                Menu_Activo = False
                Shape_Menu(0).BackStyle = 0
                Label_Menu(0).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            End If
        
        Case 1 'Editar---------------------------------------------------------------------------------------------
            If Frame_Menu(1).Visible = False Then
                Frame_Menu(1).Visible = True
                Menu_Activo = True
                Shape_Menu(1).BackStyle = 1
                Label_Menu(1).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
            Else
                Ocultar_menus
                Menu_Activo = False
                Shape_Menu(1).BackStyle = 0
                Label_Menu(1).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            End If
        
        Case 2 'Ver-------------------------------------------------------------------------------------------------
            If Frame_Menu(2).Visible = False Then
                Frame_Menu(2).Visible = True
                Menu_Activo = True
                Shape_Menu(2).BackStyle = 1
                Label_Menu(2).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
            Else
                Ocultar_menus
                Menu_Activo = False
                Shape_Menu(2).BackStyle = 0
                Label_Menu(2).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            End If
        
        Case 3 'Controlos--------------------------------------------------------------------------------------------
            If Frame_Menu(3).Visible = False Then
                Frame_Menu(3).Visible = True
                Menu_Activo = True
                Shape_Menu(3).BackStyle = 1
                Label_Menu(3).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
            Else
                Ocultar_menus
                Menu_Activo = False
                Shape_Menu(3).BackStyle = 0
                Label_Menu(3).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            End If
                
        Case 4 'Ferramentas-------------------------------------------------------------------------------------------
            If Frame_Menu(4).Visible = False Then
                Frame_Menu(4).Visible = True
                Menu_Activo = True
                Shape_Menu(4).BackStyle = 1
                Label_Menu(4).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
            Else
                Ocultar_menus
                Menu_Activo = False
                Shape_Menu(4).BackStyle = 0
                Label_Menu(4).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            End If
        
        Case 5 'Ajuda-------------------------------------------------------------------------------------------------
            If Frame_Menu(5).Visible = False Then
                Frame_Menu(5).Visible = True
                Menu_Activo = True
                Shape_Menu(5).BackStyle = 1
                Label_Menu(5).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
            Else
                Ocultar_menus
                Menu_Activo = False
                Shape_Menu(5).BackStyle = 0
                Label_Menu(5).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            End If
    End Select
End Sub

Private Sub Label_Menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Abrir automaticamente os menus caso a variavel Menu_Activo esteja activa
    If Menu_Activo = True Then
        Select Case Label_Menu(Index).Index
            Case 0 'Ficheiro
                Menu_Visivel True, False, False, False, False, False
                Shape_Menu_Activo 1, 0, 0, 0, 0, 0
                Label_Menu(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
                Label_Menu(1).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(2).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(3).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(4).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(5).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            Case 1 'Editar
                Menu_Visivel False, True, False, False, False, False
                Shape_Menu_Activo 0, 1, 0, 0, 0, 0
                Label_Menu(0).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(1).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
                Label_Menu(2).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(3).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(4).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(5).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            Case 2 'Ver
                Menu_Visivel False, False, True, False, False, False
                Shape_Menu_Activo 0, 0, 1, 0, 0, 0
                Label_Menu(0).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(1).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(2).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
                Label_Menu(3).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(4).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(5).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            Case 3 'Controlos
                Menu_Visivel False, False, False, True, False, False
                Shape_Menu_Activo 0, 0, 0, 1, 0, 0
                Label_Menu(0).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(1).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(2).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(3).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
                Label_Menu(4).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(5).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            Case 4 'Ferramentas
                Menu_Visivel False, False, False, False, True, False
                Shape_Menu_Activo 0, 0, 0, 0, 1, 0
                Label_Menu(0).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(1).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(2).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(3).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(4).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
                Label_Menu(5).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
            Case 5 'Ajuda
                Menu_Visivel False, False, False, False, False, True
                Shape_Menu_Activo 0, 0, 0, 0, 0, 1
                Label_Menu(0).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(1).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(2).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(3).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(4).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
                Label_Menu(5).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
        End Select
    End If
End Sub

Private Sub Menu_Visivel(Menu_Ficheiro As Boolean, Menu_Editar As Boolean, Menu_Ver As Boolean, Menu_Controlos As Boolean, _
                            Menu_Ferramentas As Boolean, Menu_Ajuda As Boolean)
    'Procedimento para tornar visivel o menu que se pretende
    Frame_Menu(0).Visible = Menu_Ficheiro
    Frame_Menu(1).Visible = Menu_Editar
    Frame_Menu(2).Visible = Menu_Ver
    Frame_Menu(3).Visible = Menu_Controlos
    Frame_Menu(4).Visible = Menu_Ferramentas
    Frame_Menu(5).Visible = Menu_Ajuda
End Sub

Private Sub Shape_Menu_Activo(Menu_Ficheiro As Integer, Menu_Editar As Integer, Menu_Ver As Integer, Menu_Controlos As Integer, _
                            Menu_Ferramentas As Integer, Menu_Ajuda As Integer)
    'Procedimento para tornar visivel o menu que se pretende
    Shape_Menu(0).BackStyle = Menu_Ficheiro
    Shape_Menu(1).BackStyle = Menu_Editar
    Shape_Menu(2).BackStyle = Menu_Ver
    Shape_Menu(3).BackStyle = Menu_Controlos
    Shape_Menu(4).BackStyle = Menu_Ferramentas
    Shape_Menu(5).BackStyle = Menu_Ajuda
End Sub

Public Sub Adicionar_Musica_Selecionada(Grelha_Selecionada As MSFlexGrid)
    'Procedimento para adicionar a musica selecionada á grelha destinada
    i = Grelha_Minha_Musica.Rows
    Grelha_Minha_Musica.Rows = Grelha_Minha_Musica.Rows + 1
    Grelha_Minha_Musica.TextMatrix(i, 0) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0)
    Grelha_Minha_Musica.TextMatrix(i, 1) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 1)
    Grelha_Minha_Musica.TextMatrix(i, 2) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 2)
    Grelha_Minha_Musica.TextMatrix(i, 3) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 3)
    Grelha_Minha_Musica.TextMatrix(i, 4) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 4)
End Sub

Public Sub Adicionar_na_Tabela_Lista()
    'Procedimento para adicionar a musica selecionada à biblioteca my music
    'On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "adicionarnalista.asp?Utilizador=" & Form_Perfil.Label_Utilizador.Caption & "&ID_Loja=" & Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 5), False
    servidor.send 'envia o pedido para o servidor

    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
            '"A música selecionada foi adicionada à" & vbNewLine & "sua biblioteca com sucesso."
            Form_Notificacao.Show
        End If
    End If
    Set servidor = Nothing
    
    
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

Private Sub Remover_Ficheiros_Selecionados()
    'Remover o ficheiro selecionado
    If Grelha_Visivel.Rows <= 1 Then Exit Sub
    If Grelha_Visivel = Grelha_Musica Or Grelha_Visivel = Grelha_Filmes Or Grelha_Visivel = Grelha_Minha_Musica Or Grelha_Visivel = Grelha_Listas Then
        
        Select Case Grelha_Visivel
            Case Grelha_Musica
                Form_Mensagem.Check_Remover.Visible = True: Form_Mensagem.Pic_Remover.Visible = True
                Mensagem_de_Aviso "Question", ReadINI("Message", "Quest_Remove_File", Localizacao_Ficheiro_Lingua) & vbNewLine & Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 1) & " ?"
                If Resposta = "Sim" Then
                    If Remover_da_Biblioteca = True Then
                        Cnn_Biblioteca.Execute "Delete From Tabela_Musica Where Id = '" & Grelha_Musica.TextMatrix(Grelha_Musica.Row, 9) & "'"
                        Rs_Musica.Requery
                        Verifica_Rs_Musica
                        Rs_Musica.Open "select * from Tabela_Musica order by Titulo", Cnn_Biblioteca
                    End If
                    Grelha_Musica.RemoveItem (Grelha_Musica.Row)
                    Text_Classificacao.Text = ""
                    Grelha_Musica_EnterCell
                    If Grelha_Reproduzida = Grelha_Musica Then
                        Musica_Linha_Pressionada = Grelha_Reproduzida.Row - 1
                    End If
                End If
            
            
            Case Grelha_Filmes
                Form_Mensagem.Check_Remover.Visible = True: Form_Mensagem.Pic_Remover.Visible = True
                Mensagem_de_Aviso "Question", "Pretende realmente remover o ficheiro da lista?" & vbNewLine & Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 1) & " ?"
                If Resposta = "Sim" Then
                    If Remover_da_Biblioteca = True Then
                        Cnn_Biblioteca.Execute "Delete From Tabela_Filmes Where Id = '" & Grelha_Filmes.TextMatrix(Grelha_Filmes.Row, 7) & "'"
                        Rs_Filmes.Requery
                        Verifica_Rs_Filmes
                        Rs_Filmes.Open "select * from Tabela_Filmes order by Titulo", Cnn_Biblioteca
                    End If
                    Grelha_Filmes.RemoveItem (Grelha_Filmes.Row)
                    Text_Classificacao.Text = ""
                    Grelha_Filmes_EnterCell
                    If Grelha_Reproduzida = Grelha_Filmes Then
                        Musica_Linha_Pressionada = Grelha_Reproduzida.Row - 1
                    End If
                End If
            
            
            Case Grelha_Minha_Musica
                Form_Mensagem.Check_Remover.Visible = False: Form_Mensagem.Pic_Remover.Visible = False
                Mensagem_de_Aviso "Question", ReadINI("Message", "Quest_Remove_MyMusic", Localizacao_Ficheiro_Lingua) & vbNewLine & Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 1)
                If Resposta = "Sim" Then
                    'Envia o pedido para o servidor
                    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
                    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "removerminhamusica.asp?ID=" & Grelha_Minha_Musica.TextMatrix(Grelha_Minha_Musica.Row, 5) & "&Utilizador=" & Form_Perfil.Label_Utilizador.Caption, False
                    servidor.send
                    
                    'Remover a música da grelha
                    Grelha_Minha_Musica.RemoveItem (Grelha_Minha_Musica.Row)
            
                    'Verificar os dados acesso
                    If servidor.responseText = "NaoExiste" Then
                        'Não encontrou a base de dados no servidor
                    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
                        With Form_Principal
                            If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
                                'Mensagem_de_Aviso "Information", "A música selecionada foi removida com sucesso da sua biblioteca."
                            End If
                        End With
                    End If
                End If
                
            Case Grelha_Listas
                Form_Mensagem.Check_Remover.Visible = True: Form_Mensagem.Pic_Remover.Visible = True
                Mensagem_de_Aviso "Question", ReadINI("Message", "Quest_Remove_File", Localizacao_Ficheiro_Lingua) & vbNewLine & Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 1) & " ?"
                If Resposta = "Sim" Then
                    Grelha_Listas.RemoveItem (Grelha_Listas.Row)
                    Text_Classificacao.Text = ""
                    Grelha_Listas_EnterCell
                    If Grelha_Reproduzida = Grelha_Listas Then Musica_Linha_Pressionada = Grelha_Reproduzida.Row - 1
                    
                    If Remover_da_Biblioteca = True Then
                        Dim nome_lista As String: nome_lista = Replace(Label_Topico_Lista(index_lista_selecionada).Caption, Espaco, "")
                        Personalizar_Grid Grelha_Listas
                        Dim nova_classe As clsFlexSettings
                        Set nova_classe = New clsFlexSettings
                        Set nova_classe.FlexGrid = Grelha_Listas
                        nova_classe.SaveSettings App.Path & "\Library\Playlist\" & nome_lista & ".ini", True, True, True, True
                        Set nova_classe = Nothing
                    End If
                End If
        End Select
    End If
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada

    Case Else
        Mensagem_de_Aviso "Error", ReadINI("Message", "Error_Remove_Selected_File", Localizacao_Ficheiro_Lingua)
End Select
End Sub

Private Sub Label_Solicitar_Click()
    'Solicitar a activação do serviço
    'On Error GoTo Corrige_Erro
    Me.MousePointer = 11
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/suporte/" & "enviarmensagem.asp?Email=" & Form_Perfil.Text_Email.Text & "&Assunto=" & App.ProductName & " - " & "Solicitar" & "&Mensagem=" & "Solicitar a activação do serviço", False
    servidor.send 'envia o pedido para o servidor

    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
            Mensagem_de_Aviso "Information", Idioma_Mensagem_Enviada & vbNewLine & vbNewLine & "Assim que for possivel verificar o seu pedido," & vbNewLine & "receberá um email com a confirmação da activação do serviço."
            Me.MousePointer = 0
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
End Sub

Public Sub Ver_Tela_de_Video()
    'Procedimento para ver a tela de video
    Frame_Wmp.Visible = True
    Botao_Legendas.Visible = True
    Dim star As Integer: For star = 0 To Estrela.count - 1
        Estrela(star).Visible = True
    Next star
    Barra_Mini_Player.Visible = True
    Close_Wmp.Visible = True
    Barra_Actualizar.Visible = False
End Sub

Private Sub Label_Nome_Click(Index As Integer)
    'Atalho
    Pic_Linha_Click (Index)
End Sub

Private Sub Label_Remover_Transferencia_Click(Index As Integer)
    'Remover a pasta, sub-pasta e respectivos ficheiros referentes ao programa
    Me.MousePointer = 11
    Select Case Label_Remover_Transferencia(Index).Caption
        Case Idioma_Button_Transfer_Program
            Me.MousePointer = 11
            Verificar_Pastas
            Botao_Remover_Transferencia(Index).Visible = False
            
            'Proceder á transferência dos respectivos programas
            Linha_Programa_Selecionado = Label_Remover_Transferencia(Index).Index
            Text_Servidor.Text = "http://www.nikyts.com/nplayer/applibrary/programas/" & Label_Programa(Index).Caption
            progress_activo = Index
            Progresso(progress_activo).Visible = True
            Download_Programa.DownloadFile Text_Servidor.Text, App.Path & "\Programs\" & Label_Programa(Index).Caption
            On Error GoTo 0 'Tratamento de erros
        
        
        Case Idioma_Button_Remove_Program
            DeleteFolderTree App.Path & "\Programs\" & Label_Nome(Index).Caption
            Label_Remover_Transferencia(Index).Caption = Idioma_Button_Transfer_Program
            Botao_Executar_Programa(Index).Enabled = False
            Label_Executar_Programa(Index).Enabled = False
    End Select
    Me.MousePointer = 0
    
Exit Sub
errHand:
End Sub

Private Sub Label_Titulo_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Label_Titulo_DblClick()
    'Maximixar/ Restaurar Formulários
    If Tela_Cheia = True Then
        Botao_Restaurar_Click
    Else
        Botao_Maximizar_Click
    End If
    
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Principal
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    If Tela_Cheia = False Then Mover_Formulario Form_Principal
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Principal
End Sub

Private Sub Barra_ControlBox_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Barra_ControlBox_DblClick()
    'Atalho para
    Label_Titulo_DblClick
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Principal
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    If Tela_Cheia = False Then Mover_Formulario Form_Principal
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Principal
End Sub

Public Sub Ajustar_Separadores_Barra_Lateral()
    'Procedimento para ajustar os separadores da barra lateral
    'Biblioteca-----------------------------------------------------------------------------------------------------------------------
    With Separador_Barra_Lateral(0)
        .Height = Form_Skin.Bar_View_Cover.Height
        .top = 0
        .Width = Barra_Lateral.ScaleWidth
        .left = 0
    End With
    
    With Frame_Separador_Barra_Lateral(0)
        If .Visible = True Then .Height = (2 * Separador_Barra_Lateral(0).ScaleHeight) + Separador_Barra_Lateral(0).ScaleHeight
        .top = Separador_Barra_Lateral(0).top + Separador_Barra_Lateral(0).Height
        .Width = Separador_Barra_Lateral(0).ScaleWidth
        .left = Separador_Barra_Lateral(0).left
    End With
    
    With Shape_Topico(0)
        .Height = Separador_Barra_Lateral(0).ScaleHeight
        .top = 0
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = 0
    End With
    
    With Shape_Topico(1)
        .Height = Shape_Topico(0).Height
        .top = Shape_Topico(0).top + Shape_Topico(0).Height
        .Width = Shape_Topico(0).Width
        .left = Shape_Topico(0).left
    End With
    
    'Serviços-----------------------------------------------------------------------------------------------------------------------
    With Separador_Barra_Lateral(1)
        .Height = Separador_Barra_Lateral(0).ScaleHeight
        .top = Frame_Separador_Barra_Lateral(0).top + Frame_Separador_Barra_Lateral(0).Height
        .Width = Separador_Barra_Lateral(0).ScaleWidth
        .left = Separador_Barra_Lateral(0).left
    End With

    With Frame_Separador_Barra_Lateral(1)
        If .Visible = True Then .Height = (4 * Separador_Barra_Lateral(0).ScaleHeight) + Separador_Barra_Lateral(0).ScaleHeight
        .top = Separador_Barra_Lateral(1).top + Separador_Barra_Lateral(1).ScaleHeight
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = Frame_Separador_Barra_Lateral(0).left
    End With
    
    With Shape_Topico(2)
        .Height = Shape_Topico(0).Height
        .top = 0
        .Width = Shape_Topico(0).Width
        .left = Shape_Topico(0).left
    End With
    
    With Shape_Topico(3)
        .Height = Shape_Topico(0).Height
        .top = Shape_Topico(2).top + Shape_Topico(2).Height
        .Width = Shape_Topico(0).Width
        .left = Shape_Topico(0).left
    End With
    
    With Shape_Topico(4)
        .Height = Shape_Topico(0).Height
        .top = Shape_Topico(3).top + Shape_Topico(3).Height
        .Width = Shape_Topico(0).Width
        .left = Shape_Topico(0).left
    End With
    
    With Shape_Topico(5)
        .Height = Shape_Topico(0).Height
        .top = Shape_Topico(4).top + Shape_Topico(4).Height
        .Width = Shape_Topico(0).Width
        .left = Shape_Topico(0).left
    End With
        
    'Listas para reprodução----------------------------------------------------------------------------------------------------------
    With Separador_Barra_Lateral(2)
        .Height = Separador_Barra_Lateral(0).ScaleHeight
        .top = Frame_Separador_Barra_Lateral(1).top + Frame_Separador_Barra_Lateral(1).ScaleHeight
        .Width = Separador_Barra_Lateral(0).ScaleWidth
        .left = Separador_Barra_Lateral(0).left
    End With

    With Frame_Separador_Barra_Lateral(2)
        If .Visible = True Then .Height = (Shape_Topico_Lista.count * Separador_Barra_Lateral(0).ScaleHeight) + Separador_Barra_Lateral(0).ScaleHeight
        .top = Separador_Barra_Lateral(2).top + Separador_Barra_Lateral(2).Height
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = Frame_Separador_Barra_Lateral(0).left
    End With
    
    With Shape_Topico_Lista(0)
        .Height = Shape_Topico(0).Height
        .top = 0
        .Width = Separador_Barra_Lateral(0).Width
        .left = Shape_Topico(0).left
    End With
    
    '------------
    With Label_Topico_Barra_Lateral(0)
        .top = (Separador_Barra_Lateral(0).ScaleHeight - .Height) / 2
    End With
    
    With Label_Topico_Barra_Lateral(1)
        .top = Label_Topico_Barra_Lateral(0).top
    End With
    
    With Label_Topico_Barra_Lateral(2)
        .top = Label_Topico_Barra_Lateral(0).top
    End With
    
    '------------------------------------------------------------
    With Icon_Topico(0)
        .top = Shape_Topico(0).top + ((Shape_Topico(0).Height - .Height) / 2)
        .left = 16
    End With
    
    With Icon_Topico(1)
        .top = Shape_Topico(1).top + ((Shape_Topico(1).Height - .Height) / 2)
        .left = Icon_Topico(0).left
    End With
    
    With Icon_Topico(2)
        .top = Shape_Topico(2).top + ((Shape_Topico(2).Height - .Height) / 2)
        .left = Icon_Topico(0).left
    End With
    
    With Icon_Topico(4)
        .top = Shape_Topico(4).top + ((Shape_Topico(4).Height - .Height) / 2)
        .left = Icon_Topico(0).left
    End With
    
    With Icon_Topico(3)
        .top = Shape_Topico(3).top + ((Shape_Topico(3).Height - .Height) / 2)
        .left = Icon_Topico(0).left
    End With
    
    With Icon_Topico(5)
        .top = Shape_Topico(5).top + ((Shape_Topico(5).Height - .Height) / 2)
        .left = Icon_Topico(0).left
    End With
    
    With Icon_Topico_Lista(0)
        .top = Shape_Topico_Lista(0).top + ((Shape_Topico_Lista(0).Height - .Height) / 2)
        .left = Icon_Topico(0).left
    End With
    
    '------------------------------------------------------------
    With Label_Topico_Musica
        .top = Shape_Topico(0).top + ((Shape_Topico(0).Height - .Height) / 2)
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = 0 'Icon_Topico(0).Left + Icon_Topico(0).Width + 6
    End With
    
    With Label_Topico_Filmes
        .top = Shape_Topico(1).top + ((Shape_Topico(1).Height - .Height) / 2)
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = Label_Topico_Musica.left
    End With
    
    With Label_Topico_Radio
        .top = Shape_Topico(2).top + ((Shape_Topico(2).Height - .Height) / 2)
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = Label_Topico_Musica.left
    End With
    
    With Label_Topico_Drive
        .top = Shape_Topico(4).top + ((Shape_Topico(4).Height - .Height) / 2)
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = Label_Topico_Musica.left
    End With
    
    With Label_Topico_MusicLink
        .top = Shape_Topico(3).top + ((Shape_Topico(3).Height - .Height) / 2)
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = Label_Topico_Musica.left
    End With
    
    With Label_Topico_Programas
        .top = Shape_Topico(5).top + ((Shape_Topico(5).Height - .Height) / 2)
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = Label_Topico_Musica.left
    End With
    
    With Label_Topico_Lista(0)
        .top = Shape_Topico_Lista(0).top + ((Shape_Topico_Lista(0).Height - .Height) / 2)
        .Width = Frame_Separador_Barra_Lateral(0).ScaleWidth
        .left = Label_Topico_Musica.left
    End With
End Sub

Private Sub Label_Topico_Barra_Lateral_Click(Index As Integer)
    'Selecionar separadores
    Ocultar_menus
    
    Select Case Label_Topico_Barra_Lateral(Index).Index
        Case 0 'Biblioteca
            Menu_Ficheiro_Click 1
            
            If Lista_Pastas.ListCount > 1 Then Label_Album.Caption = Label_Topico_Barra_Lateral(0).Caption
            Botao_Pausa_Click
            If Grelha_Musica.Rows > 1 Then
                Musica_Linha_Pressionada = 1
                Musica_Linha_Selecionada = 1
            Else
                Musica_Linha_Pressionada = 0
                Musica_Linha_Selecionada = 0
            End If
    End Select
End Sub

Private Sub Label_Topico_Barra_Lateral_DblClick(Index As Integer)
    'Ver/ ocultar conteudo dos separadores
    
    Select Case Label_Topico_Barra_Lateral(Index).Index
        Case 0 'Biblioteca
            If Frame_Separador_Barra_Lateral(0).Visible = True Then
                Frame_Separador_Barra_Lateral(0).Visible = False
                Frame_Separador_Barra_Lateral(0).Height = 0
            Else
                Frame_Separador_Barra_Lateral(0).Visible = True
                Frame_Separador_Barra_Lateral(0).Height = (2 * Separador_Barra_Lateral(0).ScaleHeight) + Separador_Barra_Lateral(0).ScaleHeight
            End If
            
        Case 1 'Serviços
            If Frame_Separador_Barra_Lateral(1).Visible = True Then
                Frame_Separador_Barra_Lateral(1).Visible = False
                Frame_Separador_Barra_Lateral(1).Height = 0
            Else
                Frame_Separador_Barra_Lateral(1).Visible = True
                Frame_Separador_Barra_Lateral(1).Height = (4 * Separador_Barra_Lateral(0).ScaleHeight) + Separador_Barra_Lateral(0).ScaleHeight
            End If
    
        Case 2 'Listas
            If Frame_Separador_Barra_Lateral(2).Visible = True Then
                Frame_Separador_Barra_Lateral(2).Visible = False
                Frame_Separador_Barra_Lateral(2).Height = 0
            Else
                Frame_Separador_Barra_Lateral(2).Visible = True
                Frame_Separador_Barra_Lateral(2).Height = (Shape_Topico_Lista.count * Separador_Barra_Lateral(0).ScaleHeight) + Separador_Barra_Lateral(0).ScaleHeight
            End If
    End Select
    
    Ajustar_Separadores_Barra_Lateral
End Sub

Public Sub Label_Topico_Drive_Click()
    'Chamar o procedimento
    Ocultar_menus
    If Frame_My_Drive.Visible = True Then Exit Sub
    
    Ocultar_Objectos
    Repor_Cores_Labels_Separadores
    Servico_Activo = "My other drive"
    Repor_a_Cor_Dos_Topicos
    Shape_Topico(4).Visible = True
    Label_Topico_Drive.ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
    Icon_Topico(4).Picture = Form_Skin.Icon_Topico_Drive_Over.Picture
    
    Frame_My_Drive.Visible = True
    Barra_Drive.Visible = True
    'Dim h As Integer: For h = 0 To 2
        Botao_Barra_Drive(2).Visible = True
    'Next
    Dim j As Integer: For j = 3 To 5 '6
        Label_Barra_Drive(j).Visible = True
        Botao_Barra_Drive(j).Visible = True
    Next
    Label_Contador.Caption = ""
    Text_Pesquisar_Musica.Text = Empty
    Barra_Botoes_Musica.Visible = True
    Barra_Lateral.Visible = True
    
    Label_Botao(3).Visible = True
    Label_Botao(6).Visible = True
    
    Ajustar_Objectos_Na_Horizontal
    Ajustar_Objectos_Na_Vertical
End Sub

Private Sub Label_Topico_Filmes_Click()
    'Chamar o procedimento
    Ocultar_menus
    Ocultar_Objectos
    Repor_Cores_Labels_Separadores
    
    'Caso a grelha da música esteja vazia e caso a biblioteca seja <> empty
    If Grelha_Filmes.Rows = 1 Then Menu_Ficheiro_Click 1
    
    'Selecionar tópico
    Repor_a_Cor_Dos_Topicos
    Shape_Topico(1).Visible = True
    Label_Topico_Filmes.ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
    Icon_Topico(1).Picture = Form_Skin.Icon_Topico_Filmes_Over.Picture
    
    Set Grelha_Visivel = Grelha_Filmes
    Grelha_Filmes.Visible = True
    If Form_Preferencias.Check_Ver_Playlist.Value = 1 Then Barra_Playlist.Visible = True
    
    Dim star As Integer: For star = 0 To Estrela.count - 1
        Estrela(star).Visible = True
    Next star
    
    If Grelha_Filmes.Rows > 1 Then Text_Classificacao.Text = Grelha_Filmes.TextMatrix(Grelha_Filmes.Row, 6)
    Verificar_Classificacao
    
    Verificar_Contador
    
    Label_Botao(0).Visible = True
    Text_Pesquisar_Musica.Text = Empty
    Label_Botao(4).Visible = True
    Barra_Botoes_Musica.Visible = True
    Barra_Lateral.Visible = True
    Botao_Legendas.Visible = True
    Ajustar_Objectos_Na_Horizontal
    Ajustar_Objectos_Na_Vertical
End Sub

Private Sub Label_Topico_Lista_Click(Index As Integer)
    'Chamar o procedimento
    Ocultar_menus
    
    index_lista_selecionada = Index
    
    'Selecionar tópico
    Repor_a_Cor_Dos_Topicos
    Shape_Topico_Lista(Label_Topico_Lista(Index).Index).Visible = True
    Label_Topico_Lista(Label_Topico_Lista(Index).Index).ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
    Icon_Topico_Lista(Label_Topico_Lista(Index).Index).Picture = Form_Skin.Icon_Topico_Lista_Over.Picture

    Ocultar_Objectos
    Repor_Cores_Labels_Separadores
    Set Grelha_Visivel = Grelha_Listas
    Grelha_Listas.Visible = True

    Dim star As Integer: For star = 0 To Estrela.count - 1
        Estrela(star).Visible = True
    Next star
    
    Text_Classificacao.Text = ""
    Verificar_Classificacao
    Verificar_Contador
    Label_Botao(4).Visible = True
    Label_Botao(7).Visible = True
    Label_Botao(8).Visible = True
    
    'Carregar a lista selecionada
    Dim Lista_Selecionada As String: Lista_Selecionada = Replace(Label_Topico_Lista(Index).Caption, Espaco, "")
    
    Grelha_Listas.Clear
    Grelha_Listas.Rows = 1
    Formatar_Grelha_Musica Grelha_Listas
    
    Dim cFlexSettings As clsFlexSettings
    Set cFlexSettings = New clsFlexSettings
    Set cFlexSettings.FlexGrid = Grelha_Listas
    cFlexSettings.LoadSettings App.Path & "\Library\Playlist\" & Lista_Selecionada & ".ini", True, True, True, True
    Set cFlexSettings = Nothing
    
    Personalizar_Grid Grelha_Listas
    Formatar_Grelha_Musica Grelha_Listas
    With Grelha_Listas
        If .Rows > 1 Then
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    Barra_Botoes_Musica.Visible = True
    Barra_Lateral.Visible = True
    
    Ajustar_Objectos_Na_Horizontal
    Ajustar_Objectos_Na_Vertical
    
    Dim vista As Integer: For vista = 0 To Icon_Visao.count - 1
        Icon_Visao(vista).Visible = True
    Next vista
    
    If Grelha_Listas.Rows > 1 Then Text_Classificacao.Text = Grelha_Listas.TextMatrix(Grelha_Listas.Row, 8)
    Verificar_Classificacao
End Sub

Private Sub Label_Topico_Lista_DblClick(Index As Integer)
    'Alterar o nome da lista
    index_lista_editada = Index
    
    With Text_Nome_Lista
        .Visible = False
        .Height = Label_Topico_Lista(Index).Height
        .top = Label_Topico_Lista(Index).top
        .Width = 100
        .left = Label_Topico_Lista(Index).left + 37
        .Text = Replace(Label_Topico_Lista(Index).Caption, Espaco, "")
        .Visible = True
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Label_Topico_Nova_Lista_Click()
    'Selecionar tópico
    Repor_a_Cor_Dos_Topicos
    Shape_Topico_Nova_Lista.backcolor = Form_Skin.Cor_Fundo_Topico_Over.backcolor
    Label_Topico_Nova_Lista.ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
End Sub

Private Sub Label_Topico_Musica_Click()
    'Chamar o procedimento
    Ocultar_menus
    
    'Caso a grelha da música esteja vazia e caso a biblioteca seja <> empty
    If Grelha_Musica.Rows = 1 Then Menu_Ficheiro_Click 1
    
    'Selecionar tópico
    Repor_a_Cor_Dos_Topicos
    Shape_Topico(0).Visible = True
    Label_Topico_Musica.ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
    Icon_Topico(0).Picture = Form_Skin.Icon_Topico_Musica_Over.Picture
    
    If Grelha_Musica.Visible = True Then Exit Sub
    Ocultar_Objectos
    Repor_Cores_Labels_Separadores
    Set Grelha_Visivel = Grelha_Musica
    Grelha_Musica.Visible = True
    If Form_Preferencias.Check_Ver_Playlist.Value = 1 Then Barra_Playlist.Visible = True
    
    Dim star As Integer: For star = 0 To Estrela.count - 1
        Estrela(star).Visible = True
    Next star
    
    If Grelha_Musica.Rows > 1 Then Text_Classificacao.Text = Grelha_Musica.TextMatrix(Grelha_Musica.Row, 8)
    Verificar_Classificacao
    Verificar_Contador

    Label_Botao(0).Visible = True
    Text_Pesquisar_Musica.Text = Empty
    Label_Botao(4).Visible = True
    Barra_Botoes_Musica.Visible = True
    Barra_Lateral.Visible = True
    Frame_Album.Visible = True
    Ajustar_Objectos_Na_Horizontal
    
    Dim vista As Integer: For vista = 0 To Icon_Visao.count - 1
        Icon_Visao(vista).Visible = True
    Next vista
End Sub

Private Sub Label_Topico_MusicLink_Click()
    'Chamar o procedimento
    Ocultar_menus
    
    If Frame_Music_Link.Visible = True Then Exit Sub
    Ocultar_Objectos
    Repor_Cores_Labels_Separadores
    Servico_Activo = "Music link"
    
    'Selecionar tópico
    Repor_a_Cor_Dos_Topicos
    Shape_Topico(3).Visible = True
    Label_Topico_MusicLink.ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
    Icon_Topico(3).Picture = Form_Skin.Icon_Topico_MusicLink_Over.Picture
    
    Frame_Music_Link.Visible = True
    Barra_Drive.Visible = True
    'Dim h As Integer: For h = 0 To 2
        Botao_Barra_Drive(2).Visible = True
    'Next
    Dim j As Integer: For j = 7 To 13
        Botao_Barra_Drive(j).Visible = True
        Label_Barra_Drive(j).Visible = True
    Next

    Text_Pesquisar.Text = Idioma_Pesquisa_Musica
    Text_Pesquisar_Musica.Text = Empty
    Label_Botao(5).Visible = True
    Label_Botao(6).Visible = True
    Label_Botao(3).Visible = True
    Label_Contador.Caption = ""
    Barra_Botoes_Musica.Visible = True
    Barra_Lateral.Visible = True
    
    Ajustar_Objectos_Na_Horizontal
    Ajustar_Objectos_Na_Vertical
End Sub

Private Sub Label_Topico_Programas_Click()
    'Chamar o procedimento
    Ocultar_menus
    If Frame_Programas.Visible = True Then Exit Sub
    
    Ocultar_Objectos
    Repor_Cores_Labels_Separadores
    Servico_Activo = "App library"
    Repor_a_Cor_Dos_Topicos
    Shape_Topico(5).Visible = True
    Label_Topico_Programas.ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
    Icon_Topico(5).Picture = Form_Skin.Icon_Topico_Programas_Over.Picture
    
    Frame_Programas.Visible = True
    Label_Contador.Caption = ""
    Text_Pesquisar_Musica.Text = Empty
    Barra_Botoes_Musica.Visible = True
    Barra_Lateral.Visible = True
    
    'Label_Botao(3).Visible = True
    'Label_Botao(6).Visible = True
    
    Ajustar_Objectos_Na_Horizontal
    Ajustar_Objectos_Na_Vertical
End Sub

Private Sub Label_Topico_Radio_Click()
    'Chamar o procedimento
    Ocultar_menus
    Ocultar_Objectos
    Repor_Cores_Labels_Separadores
    
    'Selecionar tópico
    Repor_a_Cor_Dos_Topicos
    Shape_Topico(2).Visible = True
    Label_Topico_Radio.ForeColor = Form_Skin.Cor_Letra_Topico_Over.backcolor
    Icon_Topico(2).Picture = Form_Skin.Icon_Topico_radio_Over.Picture
    Set Grelha_Visivel = Grelha_Radio
    
    Grelha_Radio.Visible = True
    
    Verificar_Contador
    Text_Pesquisar_Musica.Text = Empty
    Barra_Botoes_Musica.Visible = True
    Barra_Lateral.Visible = True
    Ajustar_Objectos_Na_Horizontal
    Ajustar_Objectos_Na_Vertical
End Sub

Private Sub Menu_Ajuda_Click(Index As Integer)
    'Seleciona menu ajuda
    Ocultar_menus
    
    Select Case Menu_Ajuda(Index).Index
        Case 0 'Site official
            Call ShellExecute(0, "open", "http://www.nplayer.comuv.com", vbNullString, vbNullString, SW_NORMAL)
            
        Case 1 'Reportar erro
            Form_Reportar_Erro.Show vbModal
            
        Case 3 'Verificar se existem actualizações
            Ocultar_menus
            Me.MousePointer = 11
            Verificar_Actualizacoes
            Barra_Actualizar.Visible = True
            Ajustar_Objectos_Na_Vertical
            
        Case 5 'Sobre
            Form_Sobre.Show vbModal
    End Select
End Sub

Private Sub Menu_Ajuda_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Ajuda = Index Then Exit Sub
    Sombra_Ajuda(Linha_Selecionada_Ajuda).Visible = False
    Menu_Ajuda(Linha_Selecionada_Ajuda).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Sombra_Ajuda(Index).Visible = True
    Menu_Ajuda(Index).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    Linha_Selecionada_Ajuda = Index
End Sub

Private Sub Menu_Controlos_Click(Index As Integer)
    'Seleciona menu controlos
    Ocultar_menus
    
    Select Case Menu_Controlos(Index).Index
        Case 0
            Botao_Play_Click
        Case 1
            Botao_Antes_Click
        Case 2
            Botao_Seguinte_Click
        Case 4
            Botao_Mudo_Click
    End Select
End Sub

Private Sub Menu_Controlos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Controlos = Index Then Exit Sub
    Sombra_Controlos(Linha_Selecionada_Controlos).Visible = False
    Menu_Controlos(Linha_Selecionada_Controlos).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Sombra_Controlos(Index).Visible = True
    Menu_Controlos(Index).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    Linha_Selecionada_Controlos = Index
End Sub

Private Sub Menu_Editar_Click(Index As Integer)
    'Seleciona menu ajuda
    Ocultar_menus
    
    Select Case Menu_Editar(Index).Index
        Case 0 'Adicionar o ficheiro à playlist
            Dim i As Integer
            With Grelha_Lista_Em_Reproducao
                .Rows = .Rows + 1
                i = .Rows - 1
                .TextMatrix(i, 0) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0)
                .TextMatrix(i, 1) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 1)
                .TextMatrix(i, 2) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 2)
                .TextMatrix(i, 3) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 3)
                .TextMatrix(i, 4) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 4)
                .TextMatrix(i, 5) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 5)
                .TextMatrix(i, 6) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 6)
                .TextMatrix(i, 7) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 7)
                .TextMatrix(i, 8) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 8)
                .TextMatrix(i, 9) = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 9)
            End With
            
        Case 1 'Remover da biblioteca
            Label_Botao_Click (4)
            
        Case 3 'Copiar url da musica
            If Grelha_Visivel = Grelha_Radio Or Grelha_Visivel = Grelha_Loja Or Grelha_Visivel = Grelha_Minha_Musica Or Grelha_Visivel = Grelha_Recentes Or Grelha_Visivel = Grelha_Favoritos Then
                If Grelha_Visivel = Grelha_Radio Then
                    Form_Mensagem.Text_Servidor.Text = "http://www.gotradio.com/player/launch.asp?refer=web&id=" & Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0)
                Else
                    Form_Mensagem.Text_Servidor.Text = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0)
                End If
                'Form_Mensagem.Text_Servidor.SelStart = 0
                'Form_Mensagem.Text_Servidor.SelLength = Len(Form_Mensagem.Text_Servidor.Text)
                Mensagem_de_Aviso "Hyperlink", Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 1)
            End If
                    
        Case 5 'Limpar playlit
            Grelha_Lista_Em_Reproducao.Clear
            Formatar_Grelha_Musica Grelha_Lista_Em_Reproducao
            Grelha_Lista_Em_Reproducao.Rows = 1
    End Select
End Sub

Private Sub Menu_Editar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Editar = Index Then Exit Sub
    Sombra_Editar(Linha_Selecionada_Editar).Visible = False
    Menu_Editar(Linha_Selecionada_Editar).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Sombra_Editar(Index).Visible = True
    Menu_Editar(Index).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    Linha_Selecionada_Editar = Index
End Sub

Private Sub Menu_Ferramentas_Click(Index As Integer)
    'Seleciona menu ajuda
    Ocultar_menus
        
    'Selecionar opção
    Select Case Menu_Ferramentas(Index).Index
        Case 0 'Propriedades do ficheiro
            If Grelha_Visivel = Grelha_Musica Or Grelha_Visivel = Grelha_Filmes Or Grelha_Visivel = Grelha_Listas Then
                If Grelha_Visivel.Rows <= 1 Then Exit Sub
                'Ver propriedades do ficheiro
                With Form_Atributos
                    .Text_Ficheiro.Text = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0)
                    .Ver_Propriedades
                    .Label2(1).Caption = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 1)
                    .Show vbModal
                End With
            End If
        
        Case 1 'Tag editor
            If Grelha_Visivel = Grelha_Musica Or Grelha_Visivel = Grelha_Listas Then
                If Grelha_Visivel.Rows = 1 Then Exit Sub
                With Form_Tag
                    .Text_Ficheiro.Text = Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0)
                    .Ler_Tags
                    .Show 'vbModal
                End With
            End If
    
        Case 2 'Media manager
            
        Case 4 'Opções
            Form_Preferencias.Show vbModal
    End Select
End Sub

Private Sub Menu_Ferramentas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Ferramentas = Index Then Exit Sub
    Sombra_Ferramentas(Linha_Selecionada_Ferramentas).Visible = False
    Menu_Ferramentas(Linha_Selecionada_Ferramentas).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Sombra_Ferramentas(Index).Visible = True
    Menu_Ferramentas(Index).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    Linha_Selecionada_Ferramentas = Index
End Sub

Public Sub Menu_Ficheiro_Click(Index As Integer)
    'Seleciona menu controlos
    On Error Resume Next
    Ocultar_menus
    
    Select Case Menu_Ficheiro(Index).Index
        Case 0 'Nova biblioteca
            Label_Botao_Click 0
        
        Case 1 'Actualizar os dados da biblioteca
            If Grelha_Visivel = Grelha_Musica Then
                Verifica_Rs_Musica
                Rs_Musica.Open "select * from Tabela_Musica order by Titulo", Cnn_Biblioteca
                Carregar_Grelha_Musica
                Carregar_Grelha_Artista
                Carregar_Grelha_Genero
                Carregar_Grelha_Album
            End If
            
            If Grelha_Visivel = Grelha_Filmes Then
                Verifica_Rs_Filmes
                Rs_Filmes.Open "select * from Tabela_Filmes order by Titulo", Cnn_Biblioteca
                Carregar_Grelha_Filmes
            End If
            
            Text_Pesquisar_Musica.Text = Empty
            Verificar_Contador
        
        Case 3 'Adicionar media
            Icon_Barra_Informacoes_Click 4
        
        Case 4 'Nova lista
            Label_Botao_Click 7
        
        Case 5 'Guardar lista de reprodução
            Label_Botao_Click 8
                        
        Case 7 'Abrir o explorador na localização do ficheiro
            If Grelha_Visivel.Rows = 1 Then Exit Sub
            If Grelha_Visivel = Grelha_Musica Or Grelha_Visivel = Grelha_Filmes Or Grelha_Visivel = Grelha_Listas Then
                Dim X As Long
                Dim Pasta_do_Ficheiro As String
                'Verificar primeiro se exite o ficheiro selecionado
                If Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0) <> "" Then
                    Pasta_do_Ficheiro = Replace(Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0), Dir(Grelha_Visivel.TextMatrix(Grelha_Visivel.Row, 0), vbDirectory), "")
                    X = Shell("explorer.exe " & Pasta_do_Ficheiro, vbNormalFocus)
                Else
                    Mensagem_de_Aviso "Error", ReadINI("Messgae", "Error_File_Not_Found", Localizacao_Ficheiro_Lingua)
                End If
            End If
            
        Case 9 'Fechar o programa
            Botao_Fechar_Click
    End Select
End Sub

Private Sub Menu_Ficheiro_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Ficheiro = Index Then Exit Sub
    Sombra_Ficheiro(Linha_Selecionada_Ficheiro).Visible = False
    Menu_Ficheiro(Linha_Selecionada_Ficheiro).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Sombra_Ficheiro(Index).Visible = True
    Menu_Ficheiro(Index).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    Linha_Selecionada_Ficheiro = Index
End Sub

Private Sub Menu_Ver_Click(Index As Integer)
    'Seleciona menu ajuda
    Ocultar_menus
    
    Select Case Menu_Ver(Index).Index
        Case 0 'Ver o formulário em modo mascara
            Form_Mini_Player.Show
            Me.Hide
            
        Case 2 'Ver capa
            Icon_Barra_Informacoes_Click 3
        
        Case 3 'Ver lista de reprodução
            Icon_Barra_Informacoes_Click 5
        
        Case 5 'Ver simples(apenas a grelha musica)
            Icon_Visao_Click 0
        
        Case 6 'Ver pesquisa avançada
            Icon_Visao_Click 1
            
        Case 7 'Ver album arte
            Icon_Visao_Click 2
            
        Case 9 'Visualizar a tela de video
            Ver_Tela_de_Video
    End Select
End Sub

Private Sub Menu_Ver_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Ver = Index Then Exit Sub
    Sombra_Ver(Linha_Selecionada_Ver).Visible = False
    Menu_Ver(Linha_Selecionada_Ver).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Sombra_Ver(Index).Visible = True
    Menu_Ver(Index).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    Linha_Selecionada_Ver = Index
    
    Menu_Check(0).Picture = Form_Skin.Menu_Check_Normal.Picture
    Menu_Check(1).Picture = Form_Skin.Menu_Check_Normal.Picture
    Menu_Check(2).Picture = Form_Skin.Menu_Check_Normal.Picture
    Menu_Check(3).Picture = Form_Skin.Menu_Check_Normal.Picture
    Menu_Check(4).Picture = Form_Skin.Menu_Check_Normal.Picture
    Select Case Menu_Ver(Index).Index
        Case 2
            Menu_Check(0).Picture = Form_Skin.Menu_Check_Over.Picture
        Case 3
            Menu_Check(1).Picture = Form_Skin.Menu_Check_Over.Picture
        Case 5
            Menu_Check(2).Picture = Form_Skin.Menu_Check_Over.Picture
        Case 6
            Menu_Check(3).Picture = Form_Skin.Menu_Check_Over.Picture
        Case 7
            Menu_Check(4).Picture = Form_Skin.Menu_Check_Over.Picture
    End Select
End Sub

Private Sub Pic_Capa_Album_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Picture_Slide_Som_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Picture_Slide_Som_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Colocar o slide som na posição pretendida
    Picture_Slide_Som.CurrentX = X
    Slide_Som.left = Picture_Slide_Som.CurrentX
    Form_Principal.Slide_Som_Mini.left = Picture_Slide_Som.CurrentX
    Form_Mini_Player.Slide_Som.left = Picture_Slide_Som.CurrentX
    Form_PopUp.Slide_Som.left = Picture_Slide_Som.CurrentX
    
    Verificar_Volume
    Form_Principal.Text_Slide_Som.Text = Slide_Som.left
    Form_Mini_Player.Text_Slide_Som.Text = Slide_Som.left
    
    'Caso o player esteja sem som
    If Mudo = True Then
        Form_Principal.Wmp.settings.mute = False
        Form_Wmp.Wmp.settings.mute = True
        Mudo = False
        
        Form_Principal.Menu_Controlos(4).Caption = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_Principal.Botao_Mudo_Mini.ToolTipText = Idioma_Mudo_On
        Form_Mini_Player.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_PopUp.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_Principal.Botao_Mudo_Mini.Picture = Form_Skin.Som_On_Normal_Mini.Picture
        Form_Mini_Player.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_PopUp.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
    End If
    Slide_Som.Picture = Form_Skin.Slide_Som_Normal.Picture
    Form_Mini_Player.Slide_Som.Picture = Form_Skin.Slide_Som_Normal.Picture
    Form_PopUp.Slide_Som.Picture = Form_Skin.Slide_Som_Normal.Picture
End Sub

Private Sub Separador_Barra_Lateral_Click(Index As Integer)
    'Selecionar separadores
    Ocultar_menus
    
    Select Case Separador_Barra_Lateral(Index).Index
        Case 0 'Biblioteca
            Label_Topico_Barra_Lateral_Click (0)
    End Select
End Sub

Private Sub Separador_Barra_Lateral_DblClick(Index As Integer)
    'Ver/ ocultar conteudo dos separadores
    Select Case Separador_Barra_Lateral(Index).Index
        Case 0 'Biblioteca
            Label_Topico_Barra_Lateral_DblClick (0)
            
        Case 1 'Serviços
            Label_Topico_Barra_Lateral_DblClick (1)
        
        Case 2 'Listas
            Label_Topico_Barra_Lateral_DblClick (2)
    End Select
End Sub

Private Sub Separador_Barra_Lateral_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Repor_Objectos
End Sub

Private Sub Separador_Frame_Programas_Click(Index As Integer)
    'Selecionar opção do separador
    Select Case Separador_Frame_Programas(Index).Index
        Case 0 'Home
            If Frame_Programas_Home.Visible = True Then Exit Sub
            Frame_Programas_Home.Visible = True
            Label_Titulo_Frame_Programas(2).Caption = ""
            Frame_Lista.Visible = False
            Frame_Informacoes.Visible = False
            Label_Botao(20).Visible = False: Imagem_Votar.Visible = False
            
        Case 1 'Programas instalados
            Label_Frame_Programas_Click (1)
        
        Case 2 'Categoria selecionda
            Label_Frame_Programas_Click (2)
        
        Case 3 'Programa selecionado
            Label_Frame_Programas_Click (3)
    End Select
End Sub

Public Sub Repor_a_Cor_Dos_Topicos()
    'Procedimento para repor as cores originais dos tópicos da barra lateral
    With Form_Skin
        Shape_Topico(0).Visible = False
        Label_Topico_Musica.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Shape_Topico(1).Visible = False
        Label_Topico_Filmes.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Shape_Topico(2).Visible = False
        Label_Topico_Radio.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Shape_Topico(4).Visible = False
        Label_Topico_Drive.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Shape_Topico(3).Visible = False
        Label_Topico_MusicLink.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Shape_Topico(5).Visible = False
        Label_Topico_Programas.ForeColor = .Cor_Letra_Topico_Normal.backcolor
        
        Dim X As Integer: For X = 0 To Shape_Topico_Lista.count - 1
            Shape_Topico_Lista(X).Visible = False
            Label_Topico_Lista(X).ForeColor = .Cor_Letra_Topico_Normal.backcolor
        Next X
        
        Icon_Topico(0).Picture = .Icon_Topico_Musica_Normal.Picture
        Icon_Topico(1).Picture = .Icon_Topico_Filmes_Normal.Picture
        Icon_Topico(2).Picture = .Icon_Topico_radio_Normal.Picture
        Icon_Topico(3).Picture = .Icon_Topico_MusicLink_Normal.Picture
        Icon_Topico(4).Picture = .Icon_Topico_Drive_Normal.Picture
        Icon_Topico(5).Picture = .Icon_Topico_Programas_Normal.Picture
        Dim imagem_da_lista As Integer: For imagem_da_lista = 0 To Icon_Topico_Lista.count - 1
            Icon_Topico_Lista(imagem_da_lista).Picture = .Icon_Topico_Lista_Normal.Picture
        Next
    End With
End Sub

Private Sub Botao_Barra_Drive_Click(Index As Integer)
    'Selecionar botões
    Repor_Objectos
    Ocultar_menus
    Barra_Botoes_Musica.Visible = True
    Barra_Lateral.Visible = True
    
    Select Case Botao_Barra_Drive(Index).Index
        Case 0 'Ver antes
            
        Case 1 'Ver seguinte
            
        Case 2 'Home
            If Servico_Activo = "My other drive" Then
                If Frame_My_Drive.Visible = True Then Exit Sub
                Ocultar_Frame_Central
                Frame_My_Drive.Visible = True
                Frame_Music_Link.Visible = False
                Frame_Programas.Visible = False
                Label_Botao(3).Visible = True
                Label_Botao(5).Visible = True
                Label_Botao(6).Visible = True
            
            ElseIf Servico_Activo = "Music link" Then
                If Frame_Music_Link.Visible = True Then Exit Sub
                Ocultar_Frame_Central
                Frame_My_Drive.Visible = False
                Frame_Music_Link.Visible = True
                Frame_Programas.Visible = False
                Label_Botao(3).Visible = True
                Label_Botao(5).Visible = True
                Label_Botao(6).Visible = True
            
            ElseIf Servico_Activo = "App library" Then
                If Frame_Music_Link.Visible = True Then Exit Sub
                Ocultar_Frame_Central
                Frame_Music_Link.Visible = False
                Frame_My_Drive.Visible = False
                Frame_Programas.Visible = True
            End If
            Repor_Cores_Labels_Separadores
                
        Case 3 'Contacto
            Label_Barra_Drive_Click (3)
        
        Case 4 'Agenda
            Label_Barra_Drive_Click (4)
        
        Case 5 'Ficheiros
            Label_Barra_Drive_Click (5)
            
        Case 6 'Compartilhados
            Label_Barra_Drive_Click (6)
            
        Case 7 'Recomendo
            Label_Barra_Drive_Click (7)
            
        Case 8 'Favoritos
            Label_Barra_Drive_Click (8)
            
        Case 9 'A minha música
            Label_Barra_Drive_Click (9)
            
        Case 10 'Resultado da pesquisa
            Label_Barra_Drive_Click (10)
            
        Case 11 'Comunidade
            Label_Barra_Drive_Click (11)
            
        Case 12 'Amigos
            Label_Barra_Drive_Click (12)
            
        Case 13 'Mensagens
            Label_Barra_Drive_Click (13)
            
        Case 14 'Ver perfil
            Label_Barra_Drive_Click (14)
    End Select
End Sub

Private Sub Slide_Album_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Pegar a posição do slider
    DNa_album = True
    Txa_album = X
End Sub

Private Sub Slide_Album_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o slider do som
    If DNa_album Then
        NewLeft_album = Slide_Album.left + X - Txa_album
        If NewLeft_album < 0 Then
            NewLeft_album = 0
        End If
        If NewLeft_album > Barra_Slider_Album_Center.ScaleWidth - Slide_Album.ScaleWidth Then
            NewLeft_album = Barra_Slider_Album_Center.ScaleWidth - Slide_Album.ScaleWidth
        End If
        Slide_Album.left = NewLeft_album
    End If
    
    'Posicionar a frame dos albuns
    Frame_Slide.left = -Slide_Album.left
End Sub

Private Sub Slide_Album_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o slide na posição pretendida
    DNa_album = False
End Sub

Private Sub Slide_Som_Mini_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Captar a posição do slider do som
    DNa_Som = True
    Txa_Som = X
End Sub

Private Sub Slide_Som_Mini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o slider do som
    If DNa_Som Then
        NewLeft_Som = Slide_Som.left + X - Txa_Som '- 6
        If NewLeft_Som < 0 Then
            NewLeft_Som = 0
        End If
        If NewLeft_Som > Picture_Slide_Som.Width - Slide_Som.Width Then
            NewLeft_Som = Picture_Slide_Som.Width - Slide_Som.Width
        End If
        Form_Principal.Slide_Som.left = NewLeft_Som
        Form_Principal.Slide_Som_Mini.left = NewLeft_Som
        Form_Mini_Player.Slide_Som.left = NewLeft_Som
        Form_PopUp.Slide_Som.left = NewLeft_Som
    End If
    Verificar_Volume
End Sub

Private Sub Slide_Som_Mini_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Posicionar o slider do som na posição largada
    On Error Resume Next
'''    Dim offseti As Single
    DNa_Som = False
    Verificar_Volume
    Form_Principal.Text_Slide_Som.Text = Slide_Som.left
    Form_Mini_Player.Text_Slide_Som.Text = Slide_Som.left
        
    'Caso o player esteja sem som
    If Mudo = True Then
        Form_Principal.Wmp.settings.mute = False
        Form_Wmp.Wmp.settings.mute = True
        Mudo = False
        
        Form_Principal.Menu_Controlos(4).Caption = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_Principal.Botao_Mudo_Mini.ToolTipText = Idioma_Mudo_On
        Form_Mini_Player.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_PopUp.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_Principal.Botao_Mudo_Mini.Picture = Form_Skin.Som_On_Normal_Mini.Picture
        Form_Mini_Player.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_PopUp.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
    End If
End Sub

Private Sub SliderBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Colocar o slide som na posição pretendida
    SliderBar.CurrentX = X
    Slide.left = SliderBar.CurrentX
    Slide_Mini.left = SliderBar.CurrentX
    Form_Mini_Player.Slide.left = SliderBar.CurrentX
    Image_Progresso.Width = Slide.left
    Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
    
    'Colocar o slide na posição largada
    On Error Resume Next
    Dim offseti As Single
    DNa = False
    offseti = (Slide.left - Form_Principal.Image_Barra_Slide.left - 3) / (Form_Principal.Image_Barra_Slide.Width - 10 - Slide.Width)
    Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    Form_Wmp.Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    Image_Progresso.Width = Slide.left
    Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
End Sub

Private Sub Sombra_Ajuda_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Ajuda = Index Then Exit Sub
    Sombra_Ajuda(Linha_Selecionada_Ajuda).Visible = False
    Sombra_Ajuda(Index).Visible = True
    Linha_Selecionada_Ajuda = Index
End Sub

Private Sub Sombra_Controlos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Controlos = Index Then Exit Sub
    Sombra_Controlos(Linha_Selecionada_Controlos).Visible = False
    Sombra_Controlos(Index).Visible = True
    Linha_Selecionada_Controlos = Index
End Sub

Private Sub Sombra_Editar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Editar = Index Then Exit Sub
    Sombra_Editar(Linha_Selecionada_Editar).Visible = False
    Sombra_Editar(Index).Visible = True
    Linha_Selecionada_Editar = Index
End Sub

Private Sub Sombra_Ferramentas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Ferramentas = Index Then Exit Sub
    Sombra_Ferramentas(Linha_Selecionada_Ferramentas).Visible = False
    Sombra_Ferramentas(Index).Visible = True
    Linha_Selecionada_Ferramentas = Index
End Sub

Private Sub Sombra_Ficheiro_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Ficheiro = Index Then Exit Sub
    Sombra_Ficheiro(Linha_Selecionada_Ficheiro).Visible = False
    Sombra_Ficheiro(Index).Visible = True
    Linha_Selecionada_Ficheiro = Index
End Sub

Private Sub Sombra_Ver_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Ver = Index Then Exit Sub
    Sombra_Ver(Linha_Selecionada_Ver).Visible = False
    Sombra_Ver(Index).Visible = True
    Linha_Selecionada_Ver = Index
    
    Menu_Check(0).Picture = Form_Skin.Menu_Check_Normal.Picture
    Menu_Check(1).Picture = Form_Skin.Menu_Check_Normal.Picture
    Menu_Check(2).Picture = Form_Skin.Menu_Check_Normal.Picture
    Menu_Check(3).Picture = Form_Skin.Menu_Check_Normal.Picture
    Menu_Check(4).Picture = Form_Skin.Menu_Check_Normal.Picture
    Select Case Menu_Ver(Index).Index
        Case 2
            Menu_Check(0).Picture = Form_Skin.Menu_Check_Over.Picture
        Case 3
            Menu_Check(1).Picture = Form_Skin.Menu_Check_Over.Picture
        Case 5
            Menu_Check(2).Picture = Form_Skin.Menu_Check_Over.Picture
        Case 6
            Menu_Check(3).Picture = Form_Skin.Menu_Check_Over.Picture
        Case 7
            Menu_Check(4).Picture = Form_Skin.Menu_Check_Over.Picture
    End Select
End Sub

Private Sub Text_Nome_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    'Alterar o nome da lista
    On Error GoTo Corrige_Erro
    If Len(Trim(Text_Nome_Lista.Text)) = 0 Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        Dim nome_actual, novo_nome, localizacao_da_lista As String
        nome_actual = Replace(Label_Topico_Lista(index_lista_editada).Caption, Espaco, "")
        localizacao_da_lista = App.Path & "\Library\Playlist\"
        
        'Renomear o ficheiro e alterar o nome da lista
        Name localizacao_da_lista & nome_actual & ".ini" As localizacao_da_lista & Text_Nome_Lista.Text & ".ini"
        Label_Topico_Lista(index_lista_editada).Caption = Espaco & Text_Nome_Lista.Text
        Text_Nome_Lista.Visible = False
    End If
    
Exit Sub
Corrige_Erro:
End Sub

Private Sub Text_Pesquisar_GotFocus()
    'Ao receber o focus na caixa de texto limpa a mesma
    If Text_Pesquisar.Text = Idioma_Pesquisa_Musica Then
        Text_Pesquisar.Text = Empty
    End If
End Sub

Private Sub Text_Pesquisar_LostFocus()
    'Ao receber o focus na caixa de texto limpa a mesma
    'If Len(Trim(Text_Pesquisar.text)) = 0 Then
        Text_Pesquisar.Text = Idioma_Pesquisa_Musica
    'End If
End Sub

Private Sub Text_Pesquisar_KeyPress(KeyAscii As Integer)
    'Atalho das teclas
    If KeyAscii = vbKeyReturn Then Botao_Pesquisar_Click
End Sub

Private Sub Text_Pesquisar_Musica_Change()
    'Efectuar pesquisa personalizada das músicas
    If Grelha_Visivel = Grelha_Musica Or Grelha_Visivel = Grelha_Filmes Then
        If Grelha_Visivel.Row >= 1 Then
            Criterio = Text_Pesquisar_Musica.Text
            Text_Filtro
        End If
    End If
    
    If Text_Pesquisar_Musica.Text = Empty Then
        Menu_Ficheiro_Click 1
    End If
End Sub

Public Sub Text_Filtro()
    'Efectuar o filtro á pesquisa
    If Grelha_Visivel = Grelha_Musica Then
        Verifica_Rs_Musica
        Rs_Musica.Open "select * from Tabela_Musica where Titulo like '" & Replace(Criterio, "'", "''") & "%' Or Artista like '" & Replace(Criterio, "'", "''") & "%' order by Titulo Asc", Cnn_Biblioteca
        Carregar_Grelha_Musica
    End If
    
    If Grelha_Visivel = Grelha_Filmes Then
        Verifica_Rs_Filmes
        Rs_Filmes.Open "select * from Tabela_Filmes where Titulo like '" & Criterio & "%' order by Titulo Asc", Cnn_Biblioteca
        Carregar_Grelha_Filmes
    End If
    
    If Grelha_Visivel.Row <= 1 Then
        Musica_Linha_Pressionada = 0
    End If
End Sub

Private Sub Text_Pesquisar_Musica_Click()
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Text_Pesquisar_Musica_KeyDown(KeyCode As Integer, Shift As Integer)
    'Efectuar pesquisa
    On Error GoTo Corrige_Erro
    If KeyCode = vbKeyReturn Then
        If Frame_Music_Link.Visible = True Or Grelha_Minha_Musica.Visible = True Or Grelha_Loja.Visible = True Or Grelha_Recentes.Visible = True Or Grelha_Favoritos.Visible = True Then
            'Efectuar pesquisa na base de dados consuante os dados introduzidos
            If Len(Trim(Text_Pesquisar_Musica.Text)) = 0 Then Exit Sub
            Me.MousePointer = 11
            
            Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
            servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "pesquisarmusica.asp?Recebe_Pesquisa=" & Text_Pesquisar_Musica.Text, False
            servidor.send 'envia o pedido para o servidor
            
            'Verificar os dados acesso
            If servidor.responseText = "false" Then
                Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
            ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
                If servidor.readyState = 4 And servidor.Status = 200 Then
                    Grelha_Loja.Clear
                    Formatar_Grelha Grelha_Loja
                    Carregar_Loja_Online servidor.responseText
                    Me.MousePointer = 0
        
                    Text_Pesquisar_Musica.Text = Empty
                    Label_Barra_Drive_Click (10)
                End If
            End If
        End If
    End If
    Set servidor = Nothing
    
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

Private Sub Text_Pesquisar_Musica_KeyPress(KeyAscii As Integer)
    'Permitir que seja digitado apenas números, letras,Backspace,enter,espace
    If KeyAscii = 8 Then 'Backspace
        If KeyAscii = 13 Then 'Enter
            If KeyAscii = 32 Then 'Espace
                If (KeyAscii < 48 Or KeyAscii > 57) Then '( SÓ NUMEROS)
                    If (KeyAscii < 65 Or KeyAscii > 90) Then '(SÓ LETRAS MAIUSCULAS)
                        If (KeyAscii < 97 Or KeyAscii > 122) Then '(SÓ LETRAS MINUSCULAS)
                            If (KeyAscii <> 8) Then '(APAGA CARACTER POR CARACTER)
                                KeyAscii = 0
                            End If
                         End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub Verificar_Contador()
    'Contar total de ficheiros das grelhas
    Label_Contador.Caption = ""
    
    Select Case Grelha_Visivel
        Case Grelha_Musica
           Label_Contador.Caption = Grelha_Musica.Rows - 1 & " " & Idioma_Total_Musicas

        Case Grelha_Filmes
           Label_Contador.Caption = Grelha_Filmes.Rows - 1 & " " & Idioma_Total_Filmes

        Case Grelha_Radio
           Label_Contador.Caption = Grelha_Radio.Rows - 1 & " " & Idioma_Total_Estacoes_Radio

        Case Grelha_Contactos
           Label_Contador.Caption = Grelha_Contactos.Rows - 1 & " " & Idioma_Total_Contactos

        Case Grelha_Minha_Musica
           Label_Contador.Caption = Grelha_Minha_Musica.Rows - 1 & " " & Idioma_Total_Musicas

        Case Grelha_Loja
           Label_Contador.Caption = Grelha_Loja.Rows - 1 & " " & Idioma_Total_Musicas

        Case Grelha_Listas
            Label_Contador.Caption = Grelha_Listas.Rows - 1 & " " & Idioma_Total_Musicas

        Case Grelha_Eventos
            Label_Contador.Caption = Grelha_Eventos.Rows - 1 & " " & Idioma_Total_Eventos

        Case Grelha_Mensagens
            Label_Contador.Caption = Grelha_Mensagens.Rows - 1 & " " & Idioma_Total_Messagens

        Case Grelha_Ficheiros
            Label_Contador.Caption = Grelha_Ficheiros.Rows - 1 & " " & Idioma_Total_Ficheiros_Online

        Case Grelha_Recentes
            Label_Contador.Caption = Grelha_Recentes.Rows - 1 & " " & Idioma_Total_Musicas

        Case Grelha_Favoritos
            Label_Contador.Caption = Grelha_Favoritos.Rows - 1 & " " & Idioma_Total_Musicas

        Case Grelha_Comunidade
            Label_Contador.Caption = Grelha_Comunidade.Rows - 1 & " " & Idioma_Total_Utilizadores

        Case Grelha_Amigos
            Label_Contador.Caption = Grelha_Comunidade.Rows - 1 & " " & Idioma_Total_Amigos
    End Select
End Sub

Private Sub Topico_Radio_Click()
    'Atalho
    Label_Topico_Radio_Click
End Sub

Public Sub Formatar_Grelha_Musica(Grelha As MSFlexGrid)
    'Procedimento para carregar o cabeçalho da grelha da lista de reprodução
    With Grelha
        .RowHeight(0) = 270
        .AllowUserResizing = flexResizeColumns
        .Cols = 10
        .TextMatrix(0, 0) = "Dir"
        .ColWidth(0) = 0
        .ColAlignment(0) = vbleft
        .TextMatrix(0, 1) = Idioma_Grid_Music_Col_1
        .ColWidth(1) = 8000
        .ColAlignment(1) = vbleft
        .TextMatrix(0, 2) = Idioma_Grid_Music_Col_2
        .ColWidth(2) = 1200
        .ColAlignment(2) = vbleft
        .TextMatrix(0, 3) = Idioma_Grid_Music_Col_3
        .ColWidth(3) = 2000
        .ColAlignment(3) = vbleft
        .TextMatrix(0, 4) = Idioma_Grid_Music_Col_4
        .ColWidth(4) = 1200
        .ColAlignment(4) = vbleft
        .TextMatrix(0, 5) = Idioma_Grid_Music_Col_5
        .ColWidth(5) = 2000
        .ColAlignment(5) = vbleft
        .TextMatrix(0, 6) = Idioma_Grid_Music_Col_6
        .ColWidth(6) = 3000
        .ColAlignment(6) = vbleft
        .TextMatrix(0, 7) = Idioma_Grid_Music_Col_7
        .ColWidth(7) = 2000
        .ColAlignment(7) = vbleft
        .TextMatrix(0, 8) = Idioma_Grid_Music_Col_8
        .ColWidth(8) = 1000
        .ColAlignment(8) = vbleft
        .TextMatrix(0, 9) = "Id"
        .ColWidth(9) = 1000
        .ColAlignment(9) = vbleft
    End With
End Sub

Public Sub Carregar_Grelha_Musica()
    'Procedimento para carregar a grelha de transferenicas com os dados da base de dados
    On Error GoTo Corrige_Erro
    Dim i As Integer
    With Grelha_Musica
        If Rs_Musica.RecordCount = 0 Then
            .Clear
            .Rows = 1
            Formatar_Grelha_Musica Grelha_Musica
            Exit Sub

        Else
            .Clear
            .Rows = 1
            i = 1
            Do While Not Rs_Musica.EOF
                .Rows = Rs_Musica.RecordCount + 1
                If Rs_Musica(0).Value <> "" Then .TextMatrix(i, 0) = Rs_Musica(0).Value
                If Rs_Musica(1).Value <> "" Then .TextMatrix(i, 1) = Rs_Musica(1).Value
                If Rs_Musica(2).Value <> "" Then .TextMatrix(i, 2) = Rs_Musica(2).Value
                If Rs_Musica(3).Value <> "" Then .TextMatrix(i, 3) = Rs_Musica(3).Value
                If Rs_Musica(4).Value <> "" Then .TextMatrix(i, 4) = Rs_Musica(4).Value
                If Rs_Musica(5).Value <> "" Then .TextMatrix(i, 5) = Rs_Musica(5).Value
                If Rs_Musica(6).Value <> "" Then .TextMatrix(i, 6) = Rs_Musica(6).Value
                If Rs_Musica(7).Value <> "" Then .TextMatrix(i, 7) = Rs_Musica(7).Value
                If Rs_Musica(8).Value <> "" Then .TextMatrix(i, 8) = Rs_Musica(8).Value
                If Rs_Musica(9).Value <> "" Then .TextMatrix(i, 9) = Rs_Musica(9).Value
                i = i + 1
                Rs_Musica.MoveNext
            Loop
            Formatar_Grelha_Musica Grelha_Musica
        End If
        Verificar_Contador
    End With
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Public Sub Formatar_Grelha_Filmes()
    'Procedimento para carregar o cabeçalho da grelha da lista de reprodução
    With Grelha_Filmes
        .RowHeight(0) = 270
        .AllowUserResizing = flexResizeColumns
        .Cols = 8
        .TextMatrix(0, 0) = "Dir"
        .ColWidth(0) = 0
        .ColAlignment(0) = vbleft
        .TextMatrix(0, 1) = Idioma_Grid_Movies_Col_1
        .ColWidth(1) = 8000
        .ColAlignment(1) = vbleft
        .TextMatrix(0, 2) = Idioma_Grid_Movies_Col_2
        .ColWidth(2) = 2000
        .ColAlignment(2) = vbleft
        .TextMatrix(0, 3) = Idioma_Grid_Movies_Col_3
        .ColWidth(3) = 2000
        .ColAlignment(3) = vbleft
        .TextMatrix(0, 4) = Idioma_Grid_Movies_Col_4
        .ColWidth(4) = 2000
        .ColAlignment(4) = vbleft
        .TextMatrix(0, 5) = Idioma_Grid_Movies_Col_5
        .ColWidth(5) = 1600
        .ColAlignment(5) = vbleft
        .TextMatrix(0, 6) = Idioma_Grid_Movies_Col_6
        .ColWidth(6) = 1600
        .ColAlignment(6) = vbleft
        .TextMatrix(0, 7) = "Id"
        .ColWidth(7) = 6000
        .ColAlignment(7) = vbleft
    End With
End Sub

Public Sub Carregar_Grelha_Filmes()
    'Procedimento para carregar a grelha de transferenicas com os dados da base de dados
    'On Error GoTo Corrige_Erro
    With Grelha_Filmes
        If Rs_Filmes.RecordCount = 0 Then
            .Clear
            .Rows = 1
            Formatar_Grelha_Filmes
            Exit Sub

        Else
            .Clear
            .Rows = 1
            i = 1
            Do While Not Rs_Filmes.EOF
                .Rows = Rs_Filmes.RecordCount + 1
                If Rs_Filmes(0).Value <> "" Then .TextMatrix(i, 0) = Rs_Filmes(0).Value
                If Rs_Filmes(1).Value <> "" Then .TextMatrix(i, 1) = Rs_Filmes(1).Value
                If Rs_Filmes(2).Value <> "" Then .TextMatrix(i, 2) = Rs_Filmes(2).Value
                If Rs_Filmes(3).Value <> "" Then .TextMatrix(i, 3) = Rs_Filmes(3).Value
                If Rs_Filmes(4).Value <> "" Then .TextMatrix(i, 4) = Rs_Filmes(4).Value
                If Rs_Filmes(5).Value <> "" Then .TextMatrix(i, 5) = Rs_Filmes(5).Value
                If Rs_Filmes(6).Value <> "" Then .TextMatrix(i, 6) = Rs_Filmes(6).Value
                If Rs_Filmes(7).Value <> "" Then .TextMatrix(i, 7) = Rs_Filmes(7).Value
                i = i + 1
                Rs_Filmes.MoveNext
            Loop
            Formatar_Grelha_Filmes
        End If
        Verificar_Contador
    End With
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Public Sub Botao_Mudo_Click()
    'Colocar o media player como mudo ou ouvir
    Form_Wmp.Wmp.settings.mute = True
    
    If Mudo = False Then
        Wmp.settings.mute = True
        Mudo = True
        
        Form_Principal.Menu_Controlos(4).Caption = Idioma_Mudo_Off
        Form_Principal.Botao_Mudo.ToolTipText = Idioma_Mudo_Off
        Form_Principal.Botao_Mudo_Mini.ToolTipText = Idioma_Mudo_Off
        Form_Mini_Player.Botao_Mudo.ToolTipText = Idioma_Mudo_Off
        Form_PopUp.Botao_Mudo.ToolTipText = Idioma_Mudo_Off
        
        Form_Principal.Botao_Mudo.Picture = Form_Skin.Som_Off_Normal.Picture
        Form_Principal.Botao_Mudo_Mini.Picture = Form_Skin.Som_Off_Normal_Mini.Picture
        Form_Mini_Player.Botao_Mudo.Picture = Form_Skin.Som_Off_Normal.Picture
        Form_PopUp.Botao_Mudo.Picture = Form_Skin.Som_Off_Normal.Picture
        
        Form_Principal.Slide_Som.left = 0
        Form_Principal.Slide_Som_Mini.left = 0
        Form_Mini_Player.Slide_Som.left = 0
        Form_PopUp.Slide_Som.left = 0
        
    Else
        Wmp.settings.mute = False
        Mudo = False
        
        Form_Principal.Menu_Controlos(4).Caption = Idioma_Mudo_On
        Form_Principal.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_Principal.Botao_Mudo_Mini.ToolTipText = Idioma_Mudo_On
        Form_Mini_Player.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_PopUp.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_Principal.Botao_Mudo_Mini.Picture = Form_Skin.Som_On_Normal_Mini.Picture
        Form_Mini_Player.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_PopUp.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        
        Form_Principal.Slide_Som.left = Val(Text_Slide_Som.Text)
        Form_Principal.Slide_Som_Mini.left = Val(Text_Slide_Som.Text)
        Form_Mini_Player.Slide_Som.left = Val(Text_Slide_Som.Text)
        Form_PopUp.Slide_Som.left = Val(Text_Slide_Som.Text)
    End If
End Sub

Public Sub Tocar_Media()
    'Procedimento para reproduzir os ficheiros
    'On Error GoTo Corrige_Erro
    Pic_Capa_Album.Picture = Nothing
    Pic_Capa_Album.Picture = Form_Skin.Image_Sem_Capa.Picture
    Musica_Linha_Pressionada = Grelha_Reproduzida.Row
    
'    If Form_PopUp.Modo_de_Trabalho = False Then
'        Form_PopUp.Hide
'    End If
    
    Slide.left = 0: Form_Mini_Player.Slide.left = 0
    Image_Progresso.Width = 1: Form_Mini_Player.Image_Progresso.Width = 1
    Image_Progresso.left = 0: Form_Mini_Player.Image_Progresso.left = 0
    VideoDuration = 0
    Posicao_do_Player = 0
    Wmp.Controls.stop: Form_Wmp.Wmp.Controls.stop

    'Reproduzir o som
    Label_Duracao.Caption = "00:00"
    Form_Mini_Player.Label_Duracao.Caption = Label_Duracao.Caption
    Tempo_Estimado.Caption = "00:00"
    Form_Mini_Player.Tempo_Estimado.Caption = Tempo_Estimado.Caption
    
    'Verificar se é a grelha rádio que está activa
    If Grelha_Reproduzida = Grelha_Radio Then
        Wmp.URL = "http://www.gotradio.com/player/launch.asp?refer=web&id=" & Faixa_em_Reproducao
        Form_Wmp.Wmp.URL = "http://www.gotradio.com/player/launch.asp?refer=web&id=" & Faixa_em_Reproducao
    Else
        Wmp.URL = Faixa_em_Reproducao
        Form_Wmp.Wmp.URL = Faixa_em_Reproducao
    End If
    
    Timer_Slider_Video.Enabled = True
    Wmp.Controls.play
    Form_Wmp.Wmp.Controls.play: Form_Wmp.Wmp.settings.mute = True
    
    Slide.Visible = True
    Slide_Mini.Visible = True
    Image_Progresso.Visible = True
    
    Form_Mini_Player.Slide.Visible = True
    Form_Mini_Player.Image_Progresso.Visible = True
    
    'Indicar a faixa que está a ser reproduzida
    If Grelha_Reproduzida = Grelha_Radio Or Grelha_Reproduzida = Grelha_Loja Or Grelha_Reproduzida = Grelha_Minha_Musica Then
        Label_Faixa.Caption = Grelha_Reproduzida.TextMatrix(Grelha_Reproduzida.Row, 1) & " (" & Idioma_Conectando & "...)"
    Else
        Label_Faixa.Caption = Grelha_Reproduzida.TextMatrix(Grelha_Reproduzida.Row, 1)
    End If
    Form_Mini_Player.Label_Faixa.Caption = Label_Faixa.Caption
    
    Botao_Play.Visible = False: Form_Mini_Player.Botao_Play.Visible = False: Form_PopUp.Botao_Play.Visible = False: Botao_Player_Mini(1).Visible = False
    Botao_Pausa.Visible = True: Form_Mini_Player.Botao_Pausa.Visible = True: Form_PopUp.Botao_Pausa.Visible = True: Botao_Player_Mini(2).Visible = True
    
    Musica_Play = True
    Form_Wmp.Wmp.settings.mute = True
    
    If Mudo = True Then
        Wmp.settings.mute = True
    Else
        Wmp.settings.mute = False
    End If
    
    'ver form popup
    If Modo_Tray = True Then
        With Form_PopUp
            .Label_Contador.Caption = ""
            .Label_Faixa.Caption = ""
            .Label_Artista.Caption = ""
            .Label_Contador.Caption = Grelha_Reproduzida.Row & " / " & Grelha_Reproduzida.Rows - 1
            .Label_Faixa.Caption = Label_Faixa.Caption
            .Label_Artista.Caption = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 2)
            .Tempo = 0
            .Timer1.Enabled = True
            .Show
        End With
    End If
    
    'Chamar o procedimento
    Ver_Capa_Album
    
    'Caso seja um video que esteja a ser reproduzido mostra de imediato a tela de video
    If Grelha_Reproduzida = Grelha_Filmes Then Ver_Tela_de_Video
    If Grelha_Reproduzida = Grelha_Lista_Em_Reproducao Or Grelha_Reproduzida = Grelha_Listas Or Grelha_Reproduzida = Grelha_Musica Then
        Dim Extensao_Ficheiro As String
        Extensao_Ficheiro = Right(Wmp.URL, 4) 'Pegar a extensao do ficheiro
        If Extensao_Ficheiro = ".avi" Or Extensao_Ficheirot = ".mpg" Or Extensao_Ficheiro = ".wmv" Or Extensao_Ficheiro = ".asf" Or Extensao_Ficheiro = ".mov" Or Extensao_Ficheiro = "mpeg" Then
            Ver_Tela_de_Video
        End If
    End If
    
Exit Sub
'Mostrar possiveis erros que possam vir a surgir
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case 91 'Erro a carregar a capa do album
        Pic_Capa_Album.Cls
        Pic_Capa_Album.Picture = Form_Skin.Image_Sem_Capa.Picture
        
        Case Else 'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Timer_Album_Timer()
    'Mover os albuns para direita/ esquerda
    With Frame_Slide
        If direcao_do_movimento = "left" Then
            If .left >= 0 Then .left = 0: Timer_Album.Enabled = False: Timer_Mover.Enabled = False: Exit Sub
            If .left <= nova_posicao Then
                .left = .left + 10
            Else
                Frame_Slide.left = nova_posicao
                Timer_Album.Enabled = False
            End If
            
        Else
            If .left <= (Frame_Slide_Album.ScaleWidth - Frame_Slide.ScaleWidth) Then .left = (Frame_Slide_Album.ScaleWidth - Frame_Slide.ScaleWidth): Timer_Album.Enabled = False: Timer_Mover.Enabled = False: Exit Sub
            If .left >= nova_posicao Then
                .left = .left - 10
            Else
                Frame_Slide.left = nova_posicao
                Timer_Album.Enabled = False
            End If
        End If
    End With
End Sub

Private Sub Timer_Mover_Timer()
    'Activar o movimento dos albuns a deslizar
    If movimento = "right" Then
        direcao_do_movimento = "right"
        nova_posicao = Frame_Slide.left - Form_Skin.Image_Album.Width - 10
        Timer_Album.Enabled = True
    Else
        'Mover a frame album para a direita
        direcao_do_movimento = "left"
        nova_posicao = Frame_Slide.left + Form_Skin.Image_Album.Width + 10
        Timer_Album.Enabled = True
    End If
End Sub

Private Sub Timer_Slider_Video_Timer()
    'Mostar a duração da música
    On Error Resume Next
    Tempo_Estimado.Caption = Duration(Wmp.Controls.CurrentPosition)
    Form_Mini_Player.Tempo_Estimado.Caption = Tempo_Estimado.Caption
    
    If VideoDuration >= 1 Then
        Wmp.Controls.play
        Form_Wmp.Wmp.Controls.play: Form_Wmp.Wmp.settings.mute = True
    End If
    Botao_Play.Visible = False: Form_Mini_Player.Botao_Play.Visible = False: Form_PopUp.Botao_Play.Visible = False: Botao_Player_Mini(1).Visible = False
    Botao_Pausa.Visible = True: Form_Mini_Player.Botao_Pausa.Visible = True: Form_PopUp.Botao_Pausa.Visible = True: Botao_Player_Mini(2).Visible = True
    
    Label_Duracao.Caption = Wmp.Controls.currentItem.durationString
    Form_Mini_Player.Label_Duracao.Caption = Label_Duracao.Caption
    
    'Mostrar a posição em que está a música
    'On Error Resume Next
    Dim tm As Integer, tt As Integer, tp As Single, Offset As Integer
    Dim tm_2 As Integer, tt_2 As Integer, tp_2 As Single, offset_2 As Integer
    
    tm = Int(Wmp.Controls.CurrentPosition)
    tt = Int(Wmp.currentMedia.Duration)
    tm_2 = Int(Wmp.Controls.CurrentPosition)
    tt_2 = Int(Wmp.currentMedia.Duration)
    
    Posicao_do_Player = Int(Wmp.currentMedia.Duration)
    
    If tm <> -1 Then
        tp = tm / tt
        tp_2 = tm_2 / tt_2
        
        Offset = Int((Form_Principal.Image_Barra_Slide.Width - 5 - Slide.Width) * tp)
        offset_2 = Int((Form_Mini_Player.Image_Barra_Slide.Width - 5 - Form_Mini_Player.Slide.Width) * tp_2)
        
        If Not DNa Then
            Form_Principal.Slide.left = Offset + Form_Principal.Image_Barra_Slide.left + 3
            Form_Principal.Slide_Mini.left = Offset + Form_Principal.Image_Barra_Slide_Mini.left + 3
            Form_Mini_Player.Slide.left = offset_2 + Form_Mini_Player.Image_Barra_Slide.left + 3
            Form_Principal.Image_Progresso.Width = Form_Principal.Slide.left
            Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
        End If
        
        'Verificar a finalização da música
        If Grelha_Reproduzida <> Grelha_Radio Then
            If Posicao_do_Player > 1 Then
                If tm >= (Posicao_do_Player - 1) Then
                    If musica_aleatoria = True Then 'Reproduzir música aleatóriamente
                        'MsgBox Aleatorio(CLng("1"), CLng("10")), vbInformation
                        'Dim nova_linha As Integer: nova_linha
                        Musica_Linha_Pressionada = Aleatorio(CLng("1"), CLng(Grelha_Reproduzida.Rows - 1))
                        With Grelha_Reproduzida
                            .Row = Musica_Linha_Pressionada
                            '.Row = 1
                            .Col = 0
                            .ColSel = .Cols - 1 'Selecionar a linha por inteiro
                            Musica_Linha_Pressionada = .Row
                        End With
                        Faixa_em_Reproducao = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 0)
                        Tocar_Media
                    Else
                        If Grelha_Reproduzida.Row = Grelha_Reproduzida.Rows - 1 Then
                            If musica_recomecar = True Then 'lopping á lista em reprodução
                                With Grelha_Reproduzida
                                    .Row = Musica_Linha_Pressionada
                                    .Row = 1
                                    .Col = 0
                                    .ColSel = .Cols - 1 'Selecionar a linha por inteiro
                                    Musica_Linha_Pressionada = .Row
                                End With
                                Faixa_em_Reproducao = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 0)
                                Tocar_Media
                            Else
                                Parar_o_Player
                            End If
                        Else
                            Botao_Play.Visible = True: Form_Mini_Player.Botao_Play.Visible = True: Form_PopUp.Botao_Play.Visible = True: Botao_Player_Mini(1).Visible = True
                            Botao_Pausa.Visible = False: Form_Mini_Player.Botao_Pausa.Visible = False: Form_PopUp.Botao_Pausa.Visible = False: Botao_Player_Mini(2).Visible = False
                            Botao_Seguinte_Click
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Slide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Pegar a posição do slider
    DNa = True
    Txa = X
End Sub

Private Sub Slide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Possicionar a música na posição pretendida
    If DNa Then
        NewLeft = Slide.left + X - Txa
        If NewLeft < Form_Principal.Image_Barra_Slide.left + 3 Then
            NewLeft = Form_Principal.Image_Barra_Slide.left + 3
        End If
        If NewLeft > Form_Principal.Image_Barra_Slide.Width + Form_Principal.Image_Barra_Slide.left - 7 - Slide.Width Then
            NewLeft = Form_Principal.Image_Barra_Slide.Width + Form_Principal.Image_Barra_Slide.left - 7 - Slide.Width
        End If
        Slide.left = NewLeft
        Slide_Mini.left = NewLeft
        Form_Mini_Player.Slide.left = NewLeft
        Image_Progresso.Width = Slide.left
        Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
    End If
End Sub

Private Sub Slide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Colocar o slide na posição largada
    On Error Resume Next
    Dim offseti As Single
    DNa = False
    offseti = (Slide.left - Form_Principal.Image_Barra_Slide.left - 3) / (Form_Principal.Image_Barra_Slide.Width - 10 - Slide.Width)
    Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    Form_Wmp.Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    Image_Progresso.Width = Slide.left
    Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
End Sub

Public Sub Botao_Pausa_Click()
    On Error Resume Next
    'Pausa do media player
    Wmp.Controls.pause
    Form_Wmp.Wmp.Controls.pause
    
    Label_Faixa.Caption = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 1) & " (" & Botao_Pausa.ToolTipText & ")"
    
    Form_Mini_Player.Label_Faixa.Caption = Label_Faixa.Caption
    Timer_Slider_Video.Enabled = False
    Botao_Play.Visible = True: Form_Mini_Player.Botao_Play.Visible = True: Form_PopUp.Botao_Play.Visible = True: Botao_Player_Mini(1).Visible = True
    Botao_Pausa.Visible = False: Form_Mini_Player.Botao_Pausa.Visible = False: Form_PopUp.Botao_Pausa.Visible = False: Botao_Player_Mini(2).Visible = False
    Musica_Play = False
End Sub

Public Sub Botao_Play_Click()
    On Error GoTo Corrige_Erro
    'Verificar a faixa de reprodução
    If Grelha_Reproduzida.Rows = 1 Then Exit Sub
    If Grelha_Reproduzida.Row = 0 Then Botao_Seguinte_Click
    
    'Associar ficheiro á faixa de reproduçºao
    If Faixa_em_Reproducao = "" Then
        If Grelha_Reproduzida = Grelha_Radio Then
            Faixa_em_Reproducao = "http://www.gotradio.com/player/launch.asp?refer=web&id=" & Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 0)
        Else
            Faixa_em_Reproducao = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 0)
        End If
        Tocar_Media
    End If
        
    'Reproduzir o ficheiro existente no player
    Wmp.Controls.play
    Form_Wmp.Wmp.Controls.play: Form_Wmp.Wmp.settings.mute = True
    Label_Faixa.Caption = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 1) & " (" & Idioma_Reproduzindo & ")"
    
    Form_Mini_Player.Label_Faixa.Caption = Label_Faixa.Caption
    Timer_Slider_Video.Enabled = True
    Botao_Play.Visible = False: Form_Mini_Player.Botao_Play.Visible = False: Form_PopUp.Botao_Play.Visible = False: Botao_Player_Mini(1).Visible = False
    Botao_Pausa.Visible = True: Form_Mini_Player.Botao_Pausa.Visible = True: Form_PopUp.Botao_Pausa.Visible = True: Botao_Player_Mini(2).Visible = True

    Slide.Visible = True
    Slide_Mini.Visible = True
    Image_Progresso.Visible = True
    
    Form_Mini_Player.Slide.Visible = True
    Form_Mini_Player.Image_Progresso.Visible = True
    
    'Chamar o procedimento
    Ver_Capa_Album
    
    'Verificar se é um filme que está a ser reproduzido
    If Grelha_Reproduzida = Grelha_Filmes Then
        If Frame_Wmp.Visible = False Then Ver_Tela_de_Video
    End If
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Public Sub Botao_Seguinte_Click()
    'Passar para a faixa seguinte
    If Grelha_Reproduzida.Rows = 1 Then Exit Sub
    
    'Caso esteja na última linha não avança mais
    If Grelha_Reproduzida.Row = Grelha_Reproduzida.Rows - 1 Then
        Wmp.Controls.stop: Form_Wmp.Wmp.Controls.stop
        Botao_Play.Visible = True: Form_Mini_Player.Botao_Play.Visible = True: Form_PopUp.Botao_Play.Visible = True: Botao_Player_Mini(1).Visible = True
        Botao_Pausa.Visible = False: Form_Mini_Player.Botao_Pausa.Visible = False: Form_PopUp.Botao_Pausa.Visible = False: Botao_Player_Mini(2).Visible = False
        Slide.left = 0: Form_Mini_Player.Slide.left = 0
        Image_Progresso.Width = 1: Form_Mini_Player.Image_Progresso.Width = 1
        Image_Progresso.left = 0: Form_Mini_Player.Image_Progresso.left = 0
        VideoDuration = 0
        Posicao_do_Player = 0
        Label_Duracao.Caption = "00:00" & "  |  "
        Tempo_Estimado.Caption = "00:00"
        Exit Sub
        
    Else
        'Selecionar a linha seguinte
        With Grelha_Reproduzida
            .Row = Musica_Linha_Pressionada
            .Row = .Row + 1
            .Col = 0
            .ColSel = .Cols - 1 'Selecionar a linha por inteiro
            Musica_Linha_Pressionada = .Row
        End With
        Faixa_em_Reproducao = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 0)
        Tocar_Media
    End If
End Sub

Public Sub Botao_Antes_Click()
    'Reproduzir a faixa anterior
    If Grelha_Reproduzida.Rows = 1 Then Exit Sub
    
    'Caso esteja na 1ºlinha não recua mais
    If Grelha_Reproduzida.Row = 1 Then
        Wmp.Controls.stop: Form_Wmp.Wmp.Controls.stop
        Botao_Play.Visible = True: Form_Mini_Player.Botao_Play.Visible = True: Form_PopUp.Botao_Play.Visible = True: Botao_Player_Mini(1).Visible = True
        Botao_Pausa.Visible = False: Form_Mini_Player.Botao_Pausa.Visible = False: Form_PopUp.Botao_Pausa.Visible = False: Botao_Player_Mini(2).Visible = False
        Slide.left = 0: Form_Mini_Player.Slide.left = 0
        Image_Progresso.Width = 1: Form_Mini_Player.Image_Progresso.Width = 1
        Image_Progresso.left = 0: Form_Mini_Player.Image_Progresso.left = 0
        VideoDuration = 0
        Posicao_do_Player = 0
        Label_Duracao.Caption = "00:00" & "  |  "
        Tempo_Estimado.Caption = "00:00"
        Exit Sub
        
    Else
        'Selecionar a linha anterior
        With Grelha_Reproduzida
            .Row = Musica_Linha_Pressionada
            .Row = .Row - 1
            .Col = 0
            .ColSel = .Cols - 1 'Selecionar a linha por inteiro
            Musica_Linha_Pressionada = .Row
        End With
        Faixa_em_Reproducao = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 0)
        Tocar_Media
    End If
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pichook é uma picture box, utilizada pelo Windows para reconhecer o ícone na barra de tarefas.
    Static Rec As Boolean, Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    If Rec = False Then
        Rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                'Remover do sistema o icon do programa
                Remover_Tray_Icon
    
            Case WM_LBUTTONDOWN:
                'Chamar o procedimento
                Mostrar_Faixa_Musica_Formulario_Popup
                With Form_PopUp
                    .Show
                    .Tempo = 0
                    .Timer1.Enabled = True
                End With
                
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                'Ver o menu icon se for pressionado o botão direito
                Form_Skin.PopupMenu Form_Skin.Menu_Icon
        End Select
        Rec = False
    End If
End Sub

Public Sub Remover_Tray_Icon()
    'Remover do sistema o icon do programa
    Form_PopUp.Hide
    Me.Show
    Modo_Tray = False
    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Public Sub Mostrar_Faixa_Musica_Formulario_Popup()
    'Procedimento para carregar o idioma selecionado
    On Error Resume Next
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    Idioma_Janela_Oculta = ReadINI("PopUp", "State_Window_Over", Localizacao_Ficheiro_Lingua)
    
    'Procedimento para ver o formulário popup
    With Form_PopUp
        .Label_Contador.Caption = Grelha_Reproduzida.Row & " / " & Grelha_Reproduzida.Rows - 1
        .Label_Faixa.Caption = Label_Faixa.Caption
        .Label_Artista.Caption = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 2)
    End With
End Sub

Private Sub Slide_Som_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Captar a posição do slider do som
    DNa_Som = True
    Txa_Som = X
    Slide_Som.Picture = Form_Skin.Slide_Som_Down.Picture
End Sub

Private Sub Slide_Som_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o slider do som
    If DNa_Som Then
        NewLeft_Som = Slide_Som.left + X - Txa_Som '- 6
        If NewLeft_Som < 0 Then
            NewLeft_Som = 0
        End If
        If NewLeft_Som > Picture_Slide_Som.Width - Slide_Som.Width Then
            NewLeft_Som = Picture_Slide_Som.Width - Slide_Som.Width
        End If
        Form_Principal.Slide_Som.left = NewLeft_Som
        Form_Principal.Slide_Som_Mini.left = NewLeft_Som
        Form_Mini_Player.Slide_Som.left = NewLeft_Som
        Form_PopUp.Slide_Som.left = NewLeft_Som
    End If
    Verificar_Volume
End Sub

Private Sub Slide_Som_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Posicionar o slider do som na posição largada
    On Error Resume Next
'''    Dim offseti As Single
    DNa_Som = False
    Verificar_Volume
    Form_Principal.Text_Slide_Som.Text = Slide_Som.left
    Form_Mini_Player.Text_Slide_Som.Text = Slide_Som.left
        
    'Caso o player esteja sem som
    If Mudo = True Then
        Form_Principal.Wmp.settings.mute = False
        Form_Wmp.Wmp.settings.mute = True
        Mudo = False
        
        Form_Principal.Menu_Controlos(4).Caption = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_Principal.Botao_Mudo_Mini.ToolTipText = Idioma_Mudo_On
        Form_Mini_Player.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        Form_PopUp.Botao_Mudo.ToolTipText = Idioma_Mudo_On
        
        Form_Principal.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_Principal.Botao_Mudo_Mini.Picture = Form_Skin.Som_On_Normal_Mini.Picture
        Form_Mini_Player.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
        Form_PopUp.Botao_Mudo.Picture = Form_Skin.Som_On_Normal.Picture
    End If
    Slide_Som.Picture = Form_Skin.Slide_Som_Normal.Picture
End Sub

Public Sub Verificar_Volume()
    'Procedimento para verificar o estado do volume/ slider do player
    Wmp.settings.Volume = Slide_Som.left
    If Mudo = True Then Wmp.settings.mute = True: Form_Wmp.Wmp.settings.mute = True
End Sub

Public Sub Verificar_Classificacao()
    'Procedimento para vverificar a classificacao dos ficheiros
    Select Case Text_Classificacao.Text
        Case ""
            Classificacao False, False, False, False, False
        Case "0"
            Classificacao False, False, False, False, False
        Case "1"
            Classificacao True, False, False, False, False
        Case "2"
            Classificacao True, True, False, False, False
        Case "3"
            Classificacao True, True, True, False, False
        Case "4"
            Classificacao True, True, True, True, False
        Case "5"
            Classificacao True, True, True, True, True
    End Select
End Sub

Public Sub Ocultar_Frame_Central()
    'Procedimento para ocultar as grelhas e frames que ficam ao centro
    Grelha_Musica.Visible = False
    Grelha_Filmes.Visible = False
    Grelha_Radio.Visible = False
    Grelha_Contactos.Visible = False
    Grelha_Eventos.Visible = False
    Grelha_Mensagens.Visible = False
    Grelha_Ficheiros.Visible = False
    Grelha_Recentes.Visible = False
    Grelha_Favoritos.Visible = False
    Grelha_Comunidade.Visible = False
    Grelha_Amigos.Visible = False
    Grelha_Minha_Musica.Visible = False
    Grelha_Loja.Visible = False
    Grelha_Listas.Visible = False
    Frame_Programas.Visible = False
    
'    Frame_Programas_Home.Visible = False
'    Frame_Lista.Visible = False
'    Frame_Informacoes.Visible = False
    Frame_Perfil.Visible = False
    Frame_My_Drive.Visible = False
    Frame_Music_Link.Visible = False
    
    
    Dim xpto As Integer: For xpto = 0 To Label_Botao.count - 1
        Label_Botao(xpto).Visible = False
    Next xpto
    Imagem_Votar.Visible = False
End Sub

Public Sub Ocultar_Objectos()
    'Procedimento para os objectos não pretendidos serem vistos
    Ocultar_Frame_Central
    
    Barra_Drive.Visible = False
    Botao_Legendas.Visible = False
    Barra_Mini_Player.Visible = False
    Barra_Botoes_Musica.Visible = False
    Barra_Lateral.Visible = False
    Frame_Album.Visible = False
    Frame_Wmp.Visible = False
    Barra_Playlist.Visible = False
    Frame_My_Drive.Visible = False
    
    Dim xpto As Integer: For xpto = 0 To Label_Botao.count - 1
        Label_Botao(xpto).Visible = False
    Next xpto
    Botao_Actualizar_Programa.Visible = False
    
    Dim star As Integer: For star = 0 To Estrela.count - 1
        Estrela(star).Visible = False
    Next star
    
    Dim vista As Integer: For vista = 0 To Icon_Visao.count - 1
        Icon_Visao(vista).Visible = False
    Next vista
    
    Dim h As Integer: For h = 0 To Botao_Barra_Drive.count - 1
        Botao_Barra_Drive(h).Visible = False
    Next
    Dim j As Integer: For j = 3 To Label_Barra_Drive.count + 2 '- 1 -> porque os index 0 to 2 não existem
        Label_Barra_Drive(j).Visible = False
    Next
End Sub

Public Sub Conectar_a_Base_de_Dados()
    On erro GoTo Corrige_Erro
    'Procedimento para conectar á BD e respectivas tabelas
    Cnn_Biblioteca.CursorLocation = adUseClient
    
    'On Error GoTo Corrige_Erro
    Cnn_Biblioteca.Open "provider=microsoft.jet.oledb.4.0;persist security info = false; data source = " & App.Path & "\Library\Library.mdb"
    Verifica_Rs_Musica
    Rs_Musica.Open "select * from Tabela_Musica order by Titulo", Cnn_Biblioteca
    Verifica_Rs_Filmes
    Rs_Filmes.Open "select * from Tabela_Filmes order by Titulo", Cnn_Biblioteca
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Public Sub Carregar_Grelhas()
    'Procedimento para carregar as grelhas com os dados da bd
    Carregar_Grelha_Musica
    Carregar_Grelha_Filmes
    Carregar_Grelha_Lista_Em_Reproducao
End Sub

Public Sub Verificar_Opcoes_do_Programa()
    'Verificar as propriedades do formulário de opções
    With Form_Preferencias
        If .Check_Ver_Playlist.Value = 0 Then
            Barra_Playlist.Visible = False
            Icon_Barra_Informacoes(5).ToolTipText = Idioma_Ver_Lista
            Icon_Barra_Informacoes(5).Picture = Form_Skin.Button_Playlist_View_Normal.Picture
            Menu_Check(1).Visible = False
        Else
            Barra_Playlist.Visible = True
            Icon_Barra_Informacoes(5).ToolTipText = Idioma_Ocultar_Lista
            Icon_Barra_Informacoes(5).Picture = Form_Skin.Button_Playlist_Hide_Normal.Picture
            Menu_Check(1).Visible = True
        End If
        
        If .Check_Ver_Capa.Value = 0 Then
            Frame_Capa.Visible = False
            Icon_Barra_Informacoes(3).ToolTipText = Idioma_Ver_Capa
            Icon_Barra_Informacoes(3).Picture = Form_Skin.Button_Cover_View_Normal.Picture
            Menu_Check(0).Visible = False
        Else
            Frame_Capa.Visible = True
            Icon_Barra_Informacoes(3).ToolTipText = Idioma_Ocultar_Capa
            Icon_Barra_Informacoes(3).Picture = Form_Skin.Button_Cover_Hide_Normal.Picture
            Menu_Check(0).Visible = True
        End If
    End With
End Sub

Public Sub Carregar_Grelha_Lista_Em_Reproducao()
    'Procedimento para carregar a grelha biblioteca com os dados dos ficheiro
    If Form_Preferencias.Check_Guardar_Lista.Value = 1 Then
        DoEvents
        Dim cFlexSettings As clsFlexSettings
        Set cFlexSettings = New clsFlexSettings
        Set cFlexSettings.FlexGrid = Grelha_Lista_Em_Reproducao
        cFlexSettings.LoadSettings App.Path & "\Library\Standard.ini", True, True, True, True
        Set cFlexSettings = Nothing
        
        Grelha_Lista_Em_Reproducao.Visible = True
        Label_Carregar_Favoritos.Visible = False

    Else
        Grelha_Lista_Em_Reproducao.Visible = False
        Label_Carregar_Favoritos.Visible = True
    End If

    Formatar_Grelha_Musica Grelha_Lista_Em_Reproducao
    Personalizar_Grid Grelha_Lista_Em_Reproducao
End Sub

Public Sub Actualiza_Dados_da_Tabela()
    'Procedimento para actualizar os dados referente á classificação da música
    Dim Chave As String
    If Grelha_Visivel = Grelha_Musica Then
        Chave = Grelha_Musica.TextMatrix(Grelha_Musica.Row, 9)
        Cnn_Biblioteca.Execute "Update Tabela_Musica Set Classificacao = '" & Text_Classificacao.Text & "'   where Id = '" & Chave & "'"
        Rs_Musica.Requery 1
        Grelha_Musica.TextMatrix(Grelha_Musica.Row, 8) = Text_Classificacao.Text
    End If
    
    If Grelha_Visivel = Grelha_Filmes Then
        Chave = Grelha_Filmes.TextMatrix(Grelha_Filmes.Row, 7)
        Cnn_Biblioteca.Execute "Update Tabela_Filmes Set Classificacao = '" & Text_Classificacao.Text & "'   where Id = '" & Chave & "'"
        Rs_Filmes.Requery 1
        Grelha_Filmes.TextMatrix(Grelha_Filmes.Row, 6) = Text_Classificacao.Text
    End If
    
    If Grelha_Visivel = Grelha_Listas Then
        'Chave = Grelha_Filmes.TextMatrix(Grelha_Filmes.Row, 9)
        'Cnn_Biblioteca.Execute "Update Tabela_Filmes Set Classificacao = '" & Text_Classificacao.text & "'   where Id = '" & Chave & "'"
        'Rs_Filmes.Requery 1
        Grelha_Listas.TextMatrix(Grelha_Listas.Row, 8) = Text_Classificacao.Text
    End If
End Sub

Public Sub Ver_Capa_Album()
    'Procedimento para ver a capa do album
    If Grelha_Reproduzida = Grelha_Musica Or Grelha_Reproduzida = Grelha_Lista_Em_Reproducao Or Grelha_Reproduzida = Grelha_Listas Then
        Dim k As Long
        MP3Path = Grelha_Reproduzida.TextMatrix(Musica_Linha_Pressionada, 0)
        Pic_Capa_Album.Cls
        Form_PopUp.Pic_Capa.Cls
        If ID3Exist(MP3Path) Then
            MaxIndex = GetAlbumArtCount(MP3Path)
            If MaxIndex > 0 Then
                CurIndex = 1
                If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
                    ResizePic Pic_Capa_Album
                    ResizePic Form_PopUp.Pic_Capa
                End If
            Else
                CurIndex = 0
                'Pic_Capa_Album.Picture = Form_Skin.Image_Sem_Capa.Picture
            End If
        Else
            CurIndex = 0
            MaxIndex = 0
            'Pic_Capa_Album.Picture = Form_Skin.Image_Sem_Capa.Picture
        End If
        lblIndex = CurIndex & " / " & MaxIndex
    End If
End Sub

Public Sub Ocultar_menus()
    'Procedimento para ocultar os menus
    On Error Resume Next
    Text_Nome_Lista.Visible = False
    Menu_Activo = False
    
    Frame_Menu(0).Visible = False
    Frame_Menu(1).Visible = False
    Frame_Menu(2).Visible = False
    Frame_Menu(3).Visible = False
    Frame_Menu(4).Visible = False
    Frame_Menu(5).Visible = False
    
    Dim a As Integer: For a = 0 To List_Menu(0).ListCount  'Menu_Ficheiro.count - 1
        Menu_Ficheiro(a).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Next a
    Sombra_Ficheiro(Linha_Selecionada_Ficheiro).Visible = False
    Linha_Selecionada_Ficheiro = 0
    Sombra_Ficheiro(0).Visible = True
    Menu_Ficheiro(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    
    Dim b As Integer: For b = 0 To List_Menu(1).ListCount  ' Menu_Editar.count - 1
        Menu_Editar(b).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Next b
    Sombra_Editar(Linha_Selecionada_Editar).Visible = False
    Linha_Selecionada_Editar = 0
    Sombra_Editar(0).Visible = True
    Menu_Editar(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    
    Dim c As Integer: For c = 0 To List_Menu(2).ListCount  'Menu_Ver.count - 1
        Menu_Ver(c).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Next c
    Sombra_Ver(Linha_Selecionada_Ver).Visible = False
    Linha_Selecionada_Ver = 0
    Sombra_Ver(0).Visible = True
    Menu_Ver(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    
    Dim d As Integer: For d = 0 To List_Menu(3).ListCount  'Menu_Controlos.count - 1
        Menu_Controlos(d).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Next d
    Sombra_Controlos(Linha_Selecionada_Controlos).Visible = False
    Linha_Selecionada_Controlos = 0
    Sombra_Controlos(0).Visible = True
    Menu_Controlos(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    
    Dim e As Integer: For e = 0 To List_Menu(4).ListCount  'Menu_Ferramentas.count - 1
        Menu_Ferramentas(e).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Next e
    Sombra_Ferramentas(Linha_Selecionada_Ferramentas).Visible = False
    Linha_Selecionada_Ferramentas = 0
    Sombra_Ferramentas(0).Visible = True
    Menu_Ferramentas(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    
    Dim F As Integer: For F = 0 To List_Menu(5).ListCount  'Menu_Ajuda.count - 1
        Menu_Ajuda(F).ForeColor = Form_Skin.Cor_Menu_ForeColor.backcolor
    Next F
    Sombra_Ajuda(Linha_Selecionada_Ajuda).Visible = False
    Linha_Selecionada_Ajuda = 0
    Sombra_Ajuda(0).Visible = True
    Menu_Ajuda(0).ForeColor = Form_Skin.Cor_Menu_ForeColorSel.backcolor
    
    'Remover o backstyle dos menus
    Shape_Menu_Activo 0, 0, 0, 0, 0, 0
    Dim i As Integer: For i = 0 To Label_Menu.count - 1
        Label_Menu(i).ForeColor = vbWhite 'Form_Skin.Cor_Menu_ForeColor.backcolor
    Next i
End Sub

Public Sub Ajustar_Objectos_Na_Vertical()
    On Error GoTo Corrige_Erro
    'Procedimento para ajustar os objectos na vertical
    With Barra_Actualizar
        .top = Barra_Botoes_Musica.top - .ScaleHeight
        .Height = Form_Skin.Fundo_Barra_Actualizar.Height
    End With
    
    With Frame_Album
        .top = Barra_Lateral.top
        .Height = Form_Skin.Fundo_Frame_Album.Height
    End With
    
    With Barra_Playlist
        .top = Barra_Lateral.top
        If Barra_Actualizar.Visible = False Then
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height
        Else
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Actualizar.ScaleHeight
        End If
    End With
    
    With Grelha_Lista_Em_Reproducao
        If Barra_Actualizar.Visible = True Then
            .Height = Barra_Playlist.ScaleHeight
        Else
            .Height = Barra_Playlist.ScaleHeight
        End If
        .top = 0
    End With
    
    With Linha_Barra_Playlist
        .Height = Barra_Playlist.ScaleHeight
        .top = 0
    End With
    
    With Grelha_Musica
        If visao_actual_da_biblioteca = "1" Or visao_actual_da_biblioteca = "2" Then
            .Height = Barra_Playlist.ScaleHeight - Frame_Album.ScaleHeight
            .top = Frame_Album.top + Frame_Album.ScaleHeight
        Else
            .Height = Barra_Playlist.ScaleHeight
            .top = Frame_Album.top
        End If
    End With
    
    With Grelha_Filmes
        .Height = Barra_Playlist.ScaleHeight
        .top = Barra_Playlist.top
    End With
    
    With Grelha_Radio
        If Barra_Actualizar.Visible = True Then
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Actualizar.ScaleHeight
        Else
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height
        End If
        .top = Barra_Lateral.top
    End With
    
    With Barra_Drive
        .top = Barra_Lateral.top
    End With
    
    With Frame_My_Drive
        If Barra_Actualizar.Visible = True Then
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Actualizar.ScaleHeight - Barra_Drive.Height
        Else
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Drive.Height
        End If
        .top = Barra_Drive.top + Barra_Drive.Height
    End With
    
    With Grelha_Loja
        If Barra_Actualizar.Visible = True Then
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Actualizar.ScaleHeight - Barra_Drive.Height
        Else
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Drive.Height
        End If
        .top = Barra_Drive.top + Barra_Drive.Height
    End With
    
    With Grelha_Minha_Musica
        If Barra_Actualizar.Visible = True Then
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Actualizar.ScaleHeight - Barra_Drive.Height
        Else
            .Height = Barra_Lateral.ScaleHeight - Form_Skin.Fundo_Barra_Botoes_Musica.Height - Barra_Drive.Height
        End If
        .top = Barra_Drive.top + Barra_Drive.Height
    End With

    
    '======================================================================================
    With Grelha_Listas
        .Height = Barra_Playlist.ScaleHeight
        .top = Barra_Playlist.top
    End With

    With Barra_Informacoes
        .Height = Fundo_Barra_Informacoes.Height
        .top = Me.ScaleHeight - .ScaleHeight
    End With
    
    With Fundo_Barra_Informacoes
        .top = 0
    End With
        
    Dim X As Integer: For X = 0 To Icon_Barra_Informacoes.count - 1
        Icon_Barra_Informacoes(X).top = (Barra_Informacoes.ScaleHeight - Icon_Barra_Informacoes(X).Height) / 2
    Next X
    
    With Botao_Legendas
        .Height = Form_Skin.Button_Menu_Normal.Height
        .top = (Barra_Informacoes.ScaleHeight - .ScaleHeight) / 2
    End With
    
    With Icon_Legendas
        .top = (Botao_Legendas.ScaleHeight - .Height) / 2
    End With
    
    With Label_Legendas
        .top = (Botao_Legendas.ScaleHeight - .Height) / 2
    End With
        
    With Barra_Botoes_Musica
        .Height = Form_Skin.Fundo_Barra_Botoes_Musica.Height
        .top = Barra_Lateral.top + Barra_Lateral.ScaleHeight - .ScaleHeight
    End With
    
    With Linha_Barra_Botoes_Musica
        .top = 0
    End With
    
    Dim a As Integer: For a = 0 To Label_Botao.count - 1
        Label_Botao(a).top = (Barra_Botoes_Musica.ScaleHeight - Label_Botao(a).Height) / 2
    Next a
    
    With Barra_Conexao
        .Height = Form_Skin.Fundo_Barra_Botoes_Musica.Height
        .top = Barra_Lateral.ScaleHeight - .ScaleHeight
    End With
    
    With Linha_Barra_Conexao
        .top = 0
    End With
    
    With Label_Conexao
        .top = (Barra_Conexao.ScaleHeight - .Height) / 2
    End With
    
    
Exit Sub
Corrige_Erro:
End Sub

Public Sub Ajustar_Objectos_Na_Horizontal()
    'Procedimento para ajustar a barra dos botoes
    On Error GoTo Corrige_Erro
    Estrela(4).left = Barra_Faixa.ScaleWidth - Estrela(4).Width - 10
    Estrela(3).left = Estrela(4).left - Estrela(3).Width - 2
    Estrela(2).left = Estrela(3).left - Estrela(2).Width - 2
    Estrela(1).left = Estrela(2).left - Estrela(1).Width - 2
    Estrela(0).left = Estrela(1).left - Estrela(0).Width - 2
    Dim star As Integer: For star = 0 To Estrela.count - 1
        Estrela(star).top = Label_Faixa.top
    Next star
    
    Dim h As Integer: For h = 3 To 6
        Botao_Barra_Drive(h).Stretch = True
        Botao_Barra_Drive(h).Width = Label_Barra_Drive(h).Width + 40
        Botao_Barra_Drive(h).top = Barra_Drive.top
        Botao_Barra_Drive(h).left = Botao_Barra_Drive(h - 1).left + Botao_Barra_Drive(h - 1).Width
    Next
    
    Dim k As Integer: For k = 7 To 14
        Botao_Barra_Drive(k).Stretch = True
        Botao_Barra_Drive(k).Width = Label_Barra_Drive(k).Width + 40
        Botao_Barra_Drive(k).top = Barra_Drive.top
    Next
    Botao_Barra_Drive(7).left = Botao_Barra_Drive(2).left + Botao_Barra_Drive(2).Width
    Botao_Barra_Drive(8).left = Botao_Barra_Drive(7).left + Botao_Barra_Drive(7).Width
    Botao_Barra_Drive(9).left = Botao_Barra_Drive(8).left + Botao_Barra_Drive(8).Width
    Botao_Barra_Drive(10).left = Botao_Barra_Drive(9).left + Botao_Barra_Drive(9).Width
    Botao_Barra_Drive(11).left = Botao_Barra_Drive(10).left + Botao_Barra_Drive(10).Width
    Botao_Barra_Drive(12).left = Botao_Barra_Drive(11).left + Botao_Barra_Drive(11).Width
    Botao_Barra_Drive(13).left = Botao_Barra_Drive(12).left + Botao_Barra_Drive(12).Width
    Botao_Barra_Drive(14).left = Botao_Barra_Drive(13).left + Botao_Barra_Drive(13).Width
    
    Label_Botao(0).left = 10
    Label_Botao(6).left = Label_Botao(0).left
    Label_Botao(3).left = Label_Botao(6).left + Label_Botao(6).Width + 20
    Label_Botao(5).left = Label_Botao(3).left + Label_Botao(3).Width + 20
    Label_Botao(1).left = Label_Botao(5).left + Label_Botao(5).Width + 20
    Label_Botao(2).left = Label_Botao(1).left + Label_Botao(1).Width + 20
    
    If Grelha_Minha_Musica.Visible = True Then
        Label_Botao(4).left = Label_Botao(1).left + Label_Botao(1).Width + 20
    ElseIf Grelha_Listas.Visible = True Then
        Label_Botao(4).left = Label_Botao(7).left + Label_Botao(7).Width + 20
    ElseIf Grelha_Amigos.Visible = True Then
        Label_Botao(4).left = Label_Botao(11).left + Label_Botao(11).Width + 20
    ElseIf Grelha_Musica.Visible = True Or Grelha_Filmes.Visible = True Then
        Label_Botao(4).left = Label_Botao(0).left + Label_Botao(0).Width + 20
    End If
    
    Label_Botao(7).left = 10
    Label_Botao(8).left = Label_Botao(4).left + Label_Botao(4).Width + 20
    Label_Botao(9).left = Label_Botao(5).left + Label_Botao(5).Width + 20
    
    If Grelha_Comunidade.Visible = True Then
        Label_Botao(10).left = Label_Botao(9).left + Label_Botao(9).Width + 20
    ElseIf Frame_Perfil.Visible = True Then
        Label_Botao(10).left = 10
    End If
    
    If Grelha_Amigos.Visible = True Then
        Label_Botao(12).left = Label_Botao(9).left + Label_Botao(9).Width + 20
    ElseIf Frame_Perfil.Visible = True Then
        Label_Botao(12).left = 10
    End If
    
    If Grelha_Comunidade.Visible = True Then
        Label_Botao(11).left = Label_Botao(10).left + Label_Botao(10).Width + 20
    ElseIf Grelha_Amigos.Visible = True Then
        Label_Botao(11).left = Label_Botao(12).left + Label_Botao(12).Width + 20
    ElseIf Frame_Perfil.Visible = True Then
        Label_Botao(11).left = Label_Botao(12).left + Label_Botao(12).Width + 20
    End If
    
    Label_Botao(13).left = Label_Botao(3).left + Label_Botao(3).Width + 20
    Label_Botao(14).left = Label_Botao(13).left + Label_Botao(13).Width + 20
    Label_Botao(15).left = Label_Botao(14).left + Label_Botao(14).Width + 20
    Label_Botao(16).left = Label_Botao(15).left + Label_Botao(15).Width + 20
    
    Label_Botao(17).left = Label_Botao(3).left + Label_Botao(3).Width + 20
    Label_Botao(18).left = Label_Botao(17).left + Label_Botao(17).Width + 20
    Label_Botao(19).left = Label_Botao(18).left + Label_Botao(18).Width + 20
    Label_Botao(20).left = 10
    Imagem_Votar.left = Label_Botao(20).left + Label_Botao(20).Width + 10
    Imagem_Votar.top = (Barra_Botoes_Musica.ScaleHeight - Imagem_Votar.Height) / 2
    
    'Barra_Informacoes-----------------------------------------------------------------------
    With Barra_Informacoes
        .Width = Me.ScaleWidth - 2
        .left = 1
    End With
    
    With Fundo_Barra_Informacoes
        .Stretch = True
        .Width = Barra_Informacoes.ScaleWidth
        .left = 0
    End With
    
    With Label_Contador
        .top = (Barra_Informacoes.ScaleHeight - .Height) / 2
        .Width = Barra_Informacoes.ScaleWidth
        .left = (Barra_Informacoes.ScaleWidth - .Width) / 2 '10
    End With
    
    'Esquerda
    With Icon_Barra_Informacoes(0)
        .left = 10
    End With
    
    With Icon_Barra_Informacoes(1)
        .left = Icon_Barra_Informacoes(0).left + Icon_Barra_Informacoes(0).Width
    End With
    
    With Icon_Barra_Informacoes(2)
        .left = Icon_Barra_Informacoes(1).left + Icon_Barra_Informacoes(1).Width
    End With
    
    With Icon_Barra_Informacoes(3)
        .left = Icon_Barra_Informacoes(2).left + Icon_Barra_Informacoes(2).Width
    End With
    
    With Botao_Legendas
        .Width = Form_Skin.Button_Menu_Normal.Width
        .left = Icon_Barra_Informacoes(3).left + Icon_Barra_Informacoes(3).Width + 10
    End With
    
    With Icon_Legendas
        .left = 4
    End With
    
    With Label_Legendas
        .Width = Botao_Legendas.ScaleWidth
    End With
    
    'Direita
    With Icon_Barra_Informacoes(5)
        .left = Barra_Informacoes.ScaleWidth - .Width - 10
    End With
    
    With Icon_Barra_Informacoes(4)
        .left = Icon_Barra_Informacoes(5).left - Icon_Barra_Informacoes(4).Width
    End With
    
    With Botao_Mensagens
        .top = (Barra_Informacoes.ScaleHeight - .ScaleHeight) / 2
        .Height = Form_Skin.Button_Menu_Standard_Normal.Height
        .Width = Form_Skin.Button_Menu_Standard_Normal.Width
        .left = Icon_Barra_Informacoes(4).left - .ScaleWidth - 10
    End With
    
    With Icon_Mensagens
        .top = (Botao_Legendas.ScaleHeight - .Height) / 2
        .left = 4
    End With
    
    With Label_Mensagens
        .top = (Botao_Legendas.ScaleHeight - .Height) / 2
        .Width = Botao_Mensagens.ScaleWidth
    End With
    
    With Barra_Botoes_Musica
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 3
        .left = Barra_Lateral.left + Barra_Lateral.ScaleWidth + 1
    End With
    
    With Linha_Barra_Botoes_Musica
        .Width = Barra_Botoes_Musica.ScaleWidth
        .left = 0
    End With
    
    With Barra_Conexao
        .Width = Barra_Lateral.ScaleWidth
        .left = 0
    End With
    
    With Linha_Barra_Conexao
        .Width = Barra_Conexao.ScaleWidth
        .left = 0
    End With
    
    With Label_Conexao
        .left = 10
    End With
        

    
Exit Sub
Corrige_Erro:
End Sub

Public Sub Activar_Linha_em_Reproducao(Grelha_Selecionada As MSFlexGrid)
    'Procedimento para indicar qual é a música que está a ser reproduzida
    Dim Linha_que_foi_Clicada, Total_Colunas As Integer
    With Grelha_Selecionada
        Linha_que_foi_Clicada = Linha_Activa '.Row
        'Caso seja a primeira música da grelha a ser reproduzida
        If Linha_Activa = -1 Then
            For Total_Colunas = 0 To .Cols - 1
            .Col = Total_Colunas
            .CellBackColor = &HD3D7DA 'Cinzento
            .CellForeColor = vbBlack
            Next
        
        'Caso o player já tenha sido activo
        ElseIf Linha_Activa > -1 Then
            'Limpa a cor da música que está a ser reproduzida actualmente
            .Row = Linha_Activa
            For Total_Colunas = 0 To .Cols - 1
            .Col = Total_Colunas
            .CellBackColor = vbWhite
            .CellForeColor = vbBlack
            Next
            
            'Mostra a cor da música que foi escolhida para ser reproduzida
            For Total_Colunas = 0 To .Cols - 1
            .Col = Total_Colunas
            .Row = Linha_que_foi_Clicada
                .CellBackColor = &HD3D7DA 'Cinzento
                .CellForeColor = vbBlack
            Next
        End If
        Linha_Activa = Linha_que_foi_Clicada
        
        'Seleciona a linha inteira da grelha
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Public Sub Classificacao(Estrela1 As Boolean, Estrela2 As Boolean, Estrela3 As Boolean, Estrela4 As Boolean, Estrela5 As Boolean)
    'Procedimento para limpar todas as estrelas
    If Estrela1 = False Then Estrela(0).Picture = Form_Skin.Estrela_Normal.Picture Else Estrela(0).Picture = Form_Skin.Estrela_Over.Picture
    If Estrela2 = False Then Estrela(1).Picture = Form_Skin.Estrela_Normal.Picture Else Estrela(1).Picture = Form_Skin.Estrela_Over.Picture
    If Estrela3 = False Then Estrela(2).Picture = Form_Skin.Estrela_Normal.Picture Else Estrela(2).Picture = Form_Skin.Estrela_Over.Picture
    If Estrela4 = False Then Estrela(3).Picture = Form_Skin.Estrela_Normal.Picture Else Estrela(3).Picture = Form_Skin.Estrela_Over.Picture
    If Estrela5 = False Then Estrela(4).Picture = Form_Skin.Estrela_Normal.Picture Else Estrela(4).Picture = Form_Skin.Estrela_Over.Picture
End Sub

Private Sub Wmp_Buffering(ByVal Start As Boolean)
    'Indicar a faixa que está a ser reproduzida
    If Start Then
        Label_Faixa.Caption = Grelha_Reproduzida.TextMatrix(Grelha_Reproduzida.Row, 1) & " (" & Idioma_Conectando & "...)"
    Else
        Label_Faixa.Caption = Grelha_Reproduzida.TextMatrix(Grelha_Reproduzida.Row, 1) & " (" & Idioma_Reproduzindo & ")"
    End If
    
    Form_Mini_Player.Label_Faixa.Caption = Label_Faixa.Caption
End Sub

Public Sub Ajustar_Botoes(Botao_Selecionado As PictureBox, Label_Selecionada As Label)
    'Procedimento para ajustar os botoes da barra botoes
    With Label_Selecionada
        .top = (Barra_Botoes_Musica.ScaleHeight - .Height) / 2
    End With
End Sub

Public Sub Carregar_Estacoes_Radio()
    'Procedimento para carregar as estações de rádio
    With Grelha_Radio
        .Rows = 47
        .TextMatrix(0, 1) = Idioma_Grid_Radio_Col_1
        .Cols = 3
        .RowHeight(0) = 270
        .ColWidth(0) = 0
        .ColWidth(1) = 22000
        .ColWidth(2) = 0
        .TextMatrix(1, 0) = "12": .TextMatrix(1, 1) = "Adult Alternative"
        .TextMatrix(2, 0) = "14": .TextMatrix(2, 1) = "Adult Contemporary"
        .TextMatrix(3, 0) = "12": .TextMatrix(3, 1) = "Alternative Rock"
        .TextMatrix(4, 0) = "54": .TextMatrix(4, 1) = "Big Band and Swing"
        .TextMatrix(5, 0) = "15": .TextMatrix(5, 1) = "Bluegrass"
        .TextMatrix(6, 0) = "16": .TextMatrix(6, 1) = "Blues"
        .TextMatrix(7, 0) = "46": .TextMatrix(7, 1) = "Celtic"
        .TextMatrix(8, 0) = "17": .TextMatrix(8, 1) = "Christian Contemporary"
        .TextMatrix(9, 0) = "61": .TextMatrix(9, 1) = "Christmas Celebration"
        .TextMatrix(10, 0) = "71": .TextMatrix(10, 1) = "Classic 60s"
        .TextMatrix(11, 0) = "70": .TextMatrix(11, 1) = "Classic Country"
        .TextMatrix(12, 0) = "19": .TextMatrix(12, 1) = "Classic Hits"
        .TextMatrix(13, 0) = "22": .TextMatrix(13, 1) = "Classic Rock"
        .TextMatrix(14, 0) = "21": .TextMatrix(14, 1) = "Classical"
        .TextMatrix(15, 0) = "23": .TextMatrix(15, 1) = "Country"
        .TextMatrix(16, 0) = "24": .TextMatrix(16, 1) = "Dance"
        .TextMatrix(17, 0) = "25": .TextMatrix(17, 1) = "Disco"
        .TextMatrix(18, 0) = "26": .TextMatrix(18, 1) = "Electronica"
        .TextMatrix(19, 0) = "27": .TextMatrix(19, 1) = "Folk"
        .TextMatrix(20, 0) = "53": .TextMatrix(20, 1) = "Forever Fifties"
        .TextMatrix(21, 0) = "59": .TextMatrix(21, 1) = "Halloween Rock"
        .TextMatrix(22, 0) = "28": .TextMatrix(22, 1) = "Hip Hop"
        .TextMatrix(23, 0) = "73": .TextMatrix(23, 1) = "Hot Hits"
        .TextMatrix(24, 0) = "30": .TextMatrix(24, 1) = "Indie Rock"
        .TextMatrix(25, 0) = "31": .TextMatrix(25, 1) = "Jazz"
        .TextMatrix(26, 0) = "63": .TextMatrix(26, 1) = "Mash-Ups"
        .TextMatrix(27, 0) = "34": .TextMatrix(27, 1) = "Metal Rock"
        .TextMatrix(28, 0) = "55": .TextMatrix(28, 1) = "Musical Magic"
        .TextMatrix(29, 0) = "36": .TextMatrix(29, 1) = "Native American"
        .TextMatrix(30, 0) = "35": .TextMatrix(30, 1) = "New Age"
        .TextMatrix(31, 0) = "37": .TextMatrix(31, 1) = "R&B Classics"
        .TextMatrix(32, 0) = "38": .TextMatrix(32, 1) = "Reggae"
        .TextMatrix(33, 0) = "29": .TextMatrix(33, 1) = "Retro Radio"
        .TextMatrix(34, 0) = "39": .TextMatrix(34, 1) = "Rock"
        .TextMatrix(35, 0) = "40": .TextMatrix(35, 1) = "Rockin 80's"
        .TextMatrix(36, 0) = "41": .TextMatrix(36, 1) = "Smooth Jazz"
        .TextMatrix(37, 0) = "64": .TextMatrix(37, 1) = "Soundtracks"
        .TextMatrix(38, 0) = "20": .TextMatrix(38, 1) = "Top 40"
        .TextMatrix(39, 0) = "52": .TextMatrix(39, 1) = "Top Alternative 2003"
        .TextMatrix(40, 0) = "51": .TextMatrix(40, 1) = "Top Hits 2003"
        .TextMatrix(41, 0) = "62": .TextMatrix(41, 1) = "Top Hits 2004"
        .TextMatrix(42, 0) = "72": .TextMatrix(42, 1) = "Top Hits 2005"
        .TextMatrix(43, 0) = "42": .TextMatrix(43, 1) = "Urban"
        .TextMatrix(44, 0) = "57": .TextMatrix(44, 1) = "Vintage Vault"
        .TextMatrix(45, 0) = "43": .TextMatrix(45, 1) = "Women's Alternative"
        .TextMatrix(46, 0) = "44": .TextMatrix(46, 1) = "World"
    End With
End Sub

Public Sub Formatar_Grelha_Contactos()
    'Procedimento para carregar o cabeçalho da grelha da lista de reprodução
    With Grelha_Contactos
        .RowHeight(0) = 270
        .AllowUserResizing = flexResizeColumns
        .Cols = 12
        .Rows = 1
        .RowHeight(0) = 270
        .ColWidth(0) = 0
        .ColAlignment(0) = vbleft
        .ColWidth(1) = 3000
        .ColAlignment(1) = vbleft
        .ColWidth(2) = 2700
        .ColAlignment(2) = vbleft
        .ColWidth(3) = 2000
        .ColAlignment(3) = vbleft
        .ColWidth(4) = 1500
        .ColAlignment(4) = vbleft
        .ColWidth(5) = 2000
        .ColAlignment(5) = vbleft
        .ColWidth(6) = 1000
        .ColAlignment(6) = vbleft
        .ColWidth(7) = 1000
        .ColAlignment(7) = vbleft
        .ColWidth(8) = 1000
        .ColAlignment(8) = vbleft
        .ColWidth(9) = 1000
        .ColAlignment(9) = vbleft
        .ColWidth(10) = 1000
        .ColAlignment(10) = vbleft
        .ColWidth(11) = 1000
        .ColAlignment(11) = vbleft
        
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = ReadINI("Main", "Grid_Contacts_Col_1", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 2) = ReadINI("Main", "Grid_Contacts_Col_2", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 3) = ReadINI("Main", "Grid_Contacts_Col_3", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 4) = ReadINI("Main", "Grid_Contacts_Col_4", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 5) = ReadINI("Main", "Grid_Contacts_Col_5", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 6) = ReadINI("Main", "Grid_Contacts_Col_6", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 7) = ReadINI("Main", "Grid_Contacts_Col_7", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 8) = ReadINI("Main", "Grid_Contacts_Col_8", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 9) = ReadINI("Main", "Grid_Contacts_Col_9", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 10) = ReadINI("Main", "Grid_Contacts_Col_10", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 11) = ReadINI("Main", "Grid_Contacts_Col_11", Localizacao_Ficheiro_Lingua)
    End With
End Sub

Public Sub Formatar_Grelha(nome_grelha As MSFlexGrid)
    'Procedimento para carregar o cabeçalho da grelha da lista de reprodução
    With nome_grelha
        .RowHeight(0) = 270
        .AllowUserResizing = flexResizeColumns
        .Cols = 6
        .Rows = 1
        .RowHeight(0) = 270
        .ColWidth(0) = 0
        .ColAlignment(0) = vbleft
        .ColWidth(1) = 8000
        .ColAlignment(1) = vbleft
        .ColWidth(2) = 3000
        .ColAlignment(2) = vbleft
        .ColWidth(3) = 3000
        .ColAlignment(3) = vbleft
        .ColWidth(4) = 10000
        .ColAlignment(4) = vbleft
        .ColWidth(5) = 0
        .ColAlignment(5) = vbleft
        
        .TextMatrix(0, 0) = "Hyperlink"
        .TextMatrix(0, 1) = Idioma_Grid_Loja_Col_1
        .TextMatrix(0, 2) = Idioma_Grid_Loja_Col_2
        .TextMatrix(0, 3) = Idioma_Grid_Loja_Col_3
        .TextMatrix(0, 4) = Idioma_Grid_Loja_Col_4
        .TextMatrix(0, 5) = "ID"
    End With
End Sub

Public Sub Carregar_Minha_Musica()
    'Efectuar pesquisa na base de dados consuante os dados introduzidos
    'On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "carregarlistas.asp?Recebe_Utilizador=" & Form_Perfil.Label_Utilizador.Caption, False
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If servidor.responseText = "false" Then
        Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Me.MousePointer = 11
            Grelha_Minha_Musica.Clear
            Formatar_Grelha Grelha_Minha_Musica
            Dados_Servidor_Minha_Musica servidor.responseText
            Me.MousePointer = 0
        End If
    End If
    Set servidor = Nothing
    
    
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

Private Sub Dados_Servidor_Minha_Musica(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/minhamusica/resultado")
        Dim i As Integer: i = Grelha_Minha_Musica.Rows
        
        For Each node In nodeList
            DoEvents
            Grelha_Minha_Musica.Rows = Grelha_Minha_Musica.Rows + 1

            If Not IsEmpty(node.selectSingleNode("servidor")) Then Grelha_Minha_Musica.TextMatrix(i, 0) = node.selectSingleNode("servidor").Text
            If Not IsEmpty(node.selectSingleNode("titulo")) Then Grelha_Minha_Musica.TextMatrix(i, 1) = node.selectSingleNode("titulo").Text
            If Not IsEmpty(node.selectSingleNode("artista")) Then Grelha_Minha_Musica.TextMatrix(i, 2) = node.selectSingleNode("artista").Text
            If Not IsEmpty(node.selectSingleNode("data")) Then Grelha_Minha_Musica.TextMatrix(i, 3) = node.selectSingleNode("data").Text
            If Not IsEmpty(node.selectSingleNode("adicionado")) Then Grelha_Minha_Musica.TextMatrix(i, 4) = node.selectSingleNode("adicionado").Text
            If Not IsEmpty(node.selectSingleNode("id")) Then Grelha_Minha_Musica.TextMatrix(i, 5) = node.selectSingleNode("id").Text
            i = i + 1
        Next
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Private Sub Wmp_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
    'Chamar o procedimento
    Ocultar_menus
End Sub

Private Sub Wmp_CurrentMediaItemAvailable(ByVal bstrItemName As String)
    'Carregar as capas online
    If Grelha_Reproduzida = Grelha_Radio Then
        GetInfo
    End If
End Sub

Private Sub Wmp_Error()
    On Error GoTo Corrige_Erro
Corrige_Erro:
End Sub

Private Sub Wmp_MediaError(ByVal pMediaObject As Object)
    'Mostrar possiveis erros que possam vir a surgir
    'On Error GoTo Corrige_Erro
    If Musica_Linha_Pressionada = 0 Then
        Exit Sub
    Else
        Mensagem_de_Aviso "Error", ReadINI("Main", "Error_Playing_File", Localizacao_Ficheiro_Lingua) & vbNewLine & Grelha_Reproduzida.TextMatrix(Grelha_Reproduzida.Row, 1)
    End If
    
Exit Sub
Corrige_Erro:
End Sub

Private Sub Wmp_MouseMove(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
    'Ver a barra mini player
    Barra_Mini_Player.Visible = True
End Sub

Private Sub Wmp_PlayStateChange(ByVal NewState As Long)
'    If NewState = wmppsPlaying Then '3
'        'if the player is playing, fill in the song info
'        'GetInfo
'    Else
'    If NewState = wmppsTransitioning Then '9
'        'clear the song info if the song is done
'        'ClearInfo
'        Botao_Seguinte_Click
'    End If

    Select Case Wmp.playState
        Case 0
            'MsgBox ("0")
            
        Case 1
            'MsgBox ("Player is stopped")
            Timer_Slider_Video.Enabled = False
            Botao_Play.Visible = True: Form_Mini_Player.Botao_Play.Visible = True: Form_PopUp.Botao_Play.Visible = True: Botao_Player_Mini(1).Visible = True
            Botao_Pausa.Visible = False: Form_Mini_Player.Botao_Pausa.Visible = False: Form_PopUp.Botao_Pausa.Visible = False: Botao_Player_Mini(2).Visible = False
            Slide.left = 0: Form_Mini_Player.Slide.left = 0
            Image_Progresso.Width = 1: Form_Mini_Player.Image_Progresso.Width = 1
            Image_Progresso.left = 0: Form_Mini_Player.Image_Progresso.left = 0
            VideoDuration = 0
            Posicao_do_Player = 0
            Label_Duracao.Caption = "00:00"
            Form_Mini_Player.Label_Duracao.Caption = Label_Duracao.Caption
            Tempo_Estimado.Caption = "00:00"
            Form_Mini_Player.Tempo_Estimado.Caption = Tempo_Estimado.Caption
            
        Case 2
            'MsgBox ("Player is in paused mode")
            Botao_Pausa_Click
        Case 3
'            MsgBox ("Player is playing")
        Case 4
           ' MsgBox ("4")
        Case 5
            'MsgBox ("5")
        Case 6
            'MsgBox ("6")
        Case 7
            'MsgBox ("7")
        Case 8
            'MsgBox (Wmp.playState)
'            Botao_Play.Visible = True: Form_Mini_Player.Botao_Play.Visible = True: Form_PopUp.Botao_Play.Visible = True:Botao_Player_Mini(1).Visible = True
'            Botao_Pausa.Visible = False: Form_Mini_Player.Botao_Pausa.Visible = False: Form_PopUp.Botao_Pausa.Visible = False:Botao_Player_Mini(2).Visible = False
'            Botao_Seguinte_Click
        
        Case 9
            'MsgBox ("Player is in transition mode")
    End Select
End Sub

Private Sub Wmp_Warning(ByVal WarningType As Long, ByVal Param As Long, ByVal Description As String)
    On Error GoTo Corrige_Erro
Corrige_Erro:
End Sub

Public Sub Ajustar_Menus()
    'Procedimento para ajustar as labels e frames dos menus
    On Error Resume Next
    Dim ctd As Integer: For ctd = 0 To Label_Menu.count - 1
        Label_Menu(ctd).top = (Barra_ControlBox.ScaleHeight - Label_Menu(ctd).Height) / 2
        If ctd = 0 Then
            Label_Menu(ctd).left = Label_Titulo.left + Label_Titulo.Width + 40
        Else
            Label_Menu(ctd).left = Label_Menu(ctd - 1).left + Label_Menu(ctd - 1).Width + 16 '+16 é o espaço da shape_menu em relação á largura da label_menu
        End If
        With Shape_Menu(ctd)
            .Height = Label_Menu(ctd).Height + 10
            .top = Label_Menu(ctd).top - 5
            .Width = Label_Menu(ctd).Width + 16
            .left = Label_Menu(ctd).left - 8
        End With
    Next ctd
    
    Dim i As Integer: For i = 0 To Label_Menu.count - 1
        Frame_Menu(i).top = Shape_Menu(i).top + Shape_Menu(i).Height
        Frame_Menu(i).left = Shape_Menu(i).left
    Next i
    
    'Calcular a largura do menu ficheiro
    Menu_Ficheiro(0).AutoSize = True
    Dim Largura_Menu_Ficheiro As Integer:  Largura_Menu_Ficheiro = Menu_Ficheiro(0).Width
    Dim i_Ficheiro, j_Ficheiro As Integer
    For i_Ficheiro = 0 To List_Menu(0).ListCount - 1
        If List_Menu(0).List(i_Ficheiro) <> "-" Then
            Menu_Ficheiro(i_Ficheiro).AutoSize = True
            If Menu_Ficheiro(i_Ficheiro).Width > Largura_Menu_Ficheiro Then Largura_Menu_Ficheiro = Menu_Ficheiro(i_Ficheiro).Width
            Menu_Ficheiro(i_Ficheiro).AutoSize = False
        End If
    Next i_Ficheiro
    With Frame_Menu(0)
        .Width = Largura_Menu_Ficheiro + 60
        .Height = Sombra_Ficheiro(0).Height * (Sombra_Ficheiro.count) + (2 * Sombra_Ficheiro(0).top)
    End With
    For j_Ficheiro = 0 To List_Menu(0).ListCount - 1
        If List_Menu(0).List(j_Ficheiro) <> "-" Then
            Menu_Ficheiro(j_Ficheiro).AutoSize = False
            Menu_Ficheiro(j_Ficheiro).Width = Frame_Menu(0).ScaleWidth
            Sombra_Ficheiro(j_Ficheiro).Width = Frame_Menu(0).ScaleWidth - 4
        End If
    Next j_Ficheiro
    
    'Calcular a largura do menu editar
    Menu_Editar(0).AutoSize = True
    Dim Largura_Menu_editar As Integer:  Largura_Menu_editar = Menu_Editar(0).Width
    Dim i_editar, j_editar As Integer
    For i_editar = 0 To List_Menu(1).ListCount - 1
        If List_Menu(1).List(i_editar) <> "-" Then
            Menu_Editar(i_editar).AutoSize = True
            If Menu_Editar(i_editar).Width > Largura_Menu_editar Then Largura_Menu_editar = Menu_Editar(i_editar).Width
            Menu_Editar(i_editar).AutoSize = False
        End If
    Next i_editar
    With Frame_Menu(1)
        .Width = Largura_Menu_editar + 60
        .Height = Sombra_Editar(0).Height * (Sombra_Editar.count) + (2 * Sombra_Editar(0).top)
    End With
    For j_editar = 0 To List_Menu(1).ListCount - 1
        If List_Menu(1).List(j_editar) <> "-" Then
            Menu_Editar(j_editar).AutoSize = False
            Menu_Editar(j_editar).Width = Frame_Menu(1).ScaleWidth
            Sombra_Editar(j_editar).Width = Frame_Menu(1).ScaleWidth - 4
        End If
    Next j_editar
    
    'Calcular a largura do menu ver
    Menu_Ver(0).AutoSize = True
    Dim Largura_Menu_ver As Integer:  Largura_Menu_ver = Menu_Ver(0).Width
    Dim i_ver, j_ver As Integer
    For i_ver = 0 To List_Menu(2).ListCount - 1
        If List_Menu(2).List(i_ver) <> "-" Then
            Menu_Ver(i_ver).AutoSize = True
            If Menu_Ver(i_ver).Width > Largura_Menu_ver Then Largura_Menu_ver = Menu_Ver(i_ver).Width
            Menu_Ver(i_ver).AutoSize = False
        End If
    Next i_ver
    With Frame_Menu(2)
        .Width = Largura_Menu_ver + 60
        .Height = Sombra_Ver(0).Height * (Sombra_Ver.count) + (2 * Sombra_Ver(0).top)
    End With
    For j_ver = 0 To List_Menu(2).ListCount - 1
        If List_Menu(2).List(j_ver) <> "-" Then
            Menu_Ver(j_ver).AutoSize = False
            Menu_Ver(j_ver).Width = Frame_Menu(2).ScaleWidth
            Sombra_Ver(j_ver).Width = Frame_Menu(2).ScaleWidth - 4
        End If
    Next j_ver
    
    'Calcular a largura do menu controlos
    Menu_Controlos(0).AutoSize = True
    Dim Largura_Menu_controlos As Integer:  Largura_Menu_controlos = Menu_Controlos(0).Width
    Dim i_controlos, j_controlos As Integer
    For i_controlos = 0 To List_Menu(3).ListCount - 1
        If List_Menu(3).List(i_controlos) <> "-" Then
            Menu_Controlos(i_controlos).AutoSize = True
            If Menu_Controlos(i_controlos).Width > Largura_Menu_controlos Then Largura_Menu_controlos = Menu_Controlos(i_controlos).Width
            Menu_Controlos(i_controlos).AutoSize = False
        End If
    Next i_controlos
    With Frame_Menu(3)
        .Width = Largura_Menu_controlos + 60
        .Height = Sombra_Controlos(0).Height * (Sombra_Controlos.count) + (2 * Sombra_Controlos(0).top)
    End With
    For j_controlos = 0 To List_Menu(3).ListCount - 1
        If List_Menu(3).List(j_controlos) <> "-" Then
            Menu_Controlos(j_controlos).AutoSize = False
            Menu_Controlos(j_controlos).Width = Frame_Menu(3).ScaleWidth
            Sombra_Controlos(j_controlos).Width = Frame_Menu(3).ScaleWidth - 4
        End If
    Next j_controlos
    
    'Calcular a largura do menu ferramentas
    Menu_Ferramentas(0).AutoSize = True
    Dim Largura_Menu_ferramentas As Integer:  Largura_Menu_ferramentas = Menu_Ferramentas(0).Width
    Dim i_ferramentas, j_ferramentas As Integer
    For i_ferramentas = 0 To List_Menu(4).ListCount - 1
        If List_Menu(4).List(i_ferramentas) <> "-" Then
            Menu_Ferramentas(i_ferramentas).AutoSize = True
            If Menu_Ferramentas(i_ferramentas).Width > Largura_Menu_ferramentas Then Largura_Menu_ferramentas = Menu_Ferramentas(i_ferramentas).Width
            Menu_Ferramentas(i_ferramentas).AutoSize = False
        End If
    Next i_ferramentas
    With Frame_Menu(4)
        .Width = Largura_Menu_ferramentas + 60
        .Height = Sombra_Ferramentas(0).Height * (Sombra_Ferramentas.count) + (2 * Sombra_Ferramentas(0).top)
    End With
    For j_ferramentas = 0 To List_Menu(4).ListCount - 1
        If List_Menu(4).List(j_ferramentas) <> "-" Then
            Menu_Ferramentas(j_ferramentas).AutoSize = False
            Menu_Ferramentas(j_ferramentas).Width = Frame_Menu(4).ScaleWidth
            Sombra_Ferramentas(j_ferramentas).Width = Frame_Menu(4).ScaleWidth - 4
        End If
    Next j_ferramentas
    
    'Calcular a largura do menu ajuda
    Menu_Ajuda(0).AutoSize = True
    Dim Largura_Menu_ajuda As Integer:  Largura_Menu_ajuda = Menu_Ajuda(0).Width
    Dim i_ajuda, j_ajuda As Integer
    For i_ajuda = 0 To List_Menu(5).ListCount - 1
        If List_Menu(5).List(i_ajuda) <> "-" Then
            Menu_Ajuda(i_ajuda).AutoSize = True
            If Menu_Ajuda(i_ajuda).Width > Largura_Menu_ajuda Then Largura_Menu_ajuda = Menu_Ajuda(i_ajuda).Width
            Menu_Ajuda(i_ajuda).AutoSize = False
        End If
    Next i_ajuda
    With Frame_Menu(5)
        .Width = Largura_Menu_ajuda + 60
        .Height = Sombra_Ajuda(0).Height * (Sombra_Ajuda.count) + (2 * Sombra_Ajuda(0).top)
    End With
    For j_ajuda = 0 To List_Menu(5).ListCount - 1
        If List_Menu(5).List(j_ajuda) <> "-" Then
            Menu_Ajuda(j_ajuda).AutoSize = False
            Menu_Ajuda(j_ajuda).Width = Frame_Menu(5).ScaleWidth
            Sombra_Ajuda(j_ajuda).Width = Frame_Menu(5).ScaleWidth - 4
        End If
    Next j_ajuda
    
    'Shape vertical dos menus para colocar os icons
    Dim barra_icon As Integer: For barra_icon = 0 To Shape_Vertical.count - 1
        With Shape_Vertical(barra_icon)
            .top = 2
            .Height = Frame_Menu(barra_icon).ScaleHeight - 4
            .left = 2
            .Width = 22
            .ZOrder 1
        End With
    Next
    
    'Linhas dos menus
    Dim a As Integer: For a = 0 To 10 'Linha_Ficheiro.count - 1
        Linha_Ficheiro(a).left = Shape_Vertical(0).left + Shape_Vertical(0).Width + 4
        Linha_Ficheiro(a).Width = Sombra_Ficheiro(0).Width - Shape_Vertical(0).Width - 8
    Next
    Dim b As Integer: For b = 0 To 10 'Linha_Editar.count - 1
        Linha_Editar(b).left = Shape_Vertical(1).left + Shape_Vertical(1).Width + 4
        Linha_Editar(b).Width = Sombra_Editar(0).Width - Shape_Vertical(1).Width - 8
    Next
    Dim c As Integer: For c = 0 To 10 'Linha_Ver.count - 1
        Linha_Ver(c).left = Shape_Vertical(2).left + Shape_Vertical(2).Width + 4
        Linha_Ver(c).Width = Sombra_Ver(0).Width - Shape_Vertical(2).Width - 8
    Next
    Dim d As Integer: For d = 0 To 10 'Linha_Controlos.count - 1
        Linha_Controlos(d).left = Shape_Vertical(3).left + Shape_Vertical(3).Width + 4
        Linha_Controlos(d).Width = Sombra_Controlos(0).Width - Shape_Vertical(3).Width - 8
    Next
    Dim e As Integer: For e = 0 To 10 'Linha_Ferramentas.count - 1
        Linha_Ferramentas(e).left = Shape_Vertical(4).left + Shape_Vertical(4).Width + 4
        Linha_Ferramentas(e).Width = Sombra_Ferramentas(0).Width - Shape_Vertical(4).Width - 8
    Next
    Dim F As Integer: For F = 0 To 10 'Linha_Ajuda.count - 1
        Linha_Ajuda(F).left = Shape_Vertical(5).left + Shape_Vertical(5).Width + 4
        Linha_Ajuda(F).Width = Sombra_Ajuda(0).Width - Shape_Vertical(5).Width - 8
    Next
    
    'Imagens checks dos botoes
    Dim Z As Integer: For Z = 0 To Menu_Check.count - 1
        Menu_Check(Z).left = Frame_Menu(2).ScaleWidth - Menu_Check(Z).Width - 10
    Next
    Menu_Check(0).top = Menu_Ver(2).top
    Menu_Check(1).top = Menu_Ver(3).top
    Menu_Check(2).top = Menu_Ver(5).top
    Menu_Check(3).top = Menu_Ver(6).top
    Menu_Check(4).top = Menu_Ver(7).top
End Sub

Private Sub ListSubDirs(Path)
    'Procedimento para listar as pastas/albuns das músicas
    On Error Resume Next
    Dim count, d(), i, DirName
    DirName = Dir(Path, 16)
    Do While DirName <> ""
        If DirName <> "." And DirName <> ".." Then
            If GetAttr(Path + DirName) = 16 Then
                If (count Mod 10) = 0 Then
                    ReDim Preserve d(count + 10)
                End If
                count = count + 1
                d(count) = DirName
            End If
        End If
        DirName = Dir
    Loop
    For i = 1 To count
        Lista_Pastas.AddItem Path & d(i)
        'ListSubDirs Path & d(i) & "\"
    Next i
    DoEvents
End Sub

Public Sub Formatar_Grelha_Artista()
    'Procedimento para carregar o cabeçalho da grelha da lista de reprodução
    With Grelha_Artista
        .Clear
        .Rows = 1
        .Cols = 2
        .RowHeight(0) = 270
        .ColWidth(0) = 0
        .ColWidth(1) = 10000
        .ColAlignment(1) = vbleft
        .TextMatrix(0, 1) = Idioma_Grid_Music_Col_2
    End With
End Sub

Public Sub Formatar_Grelha_Genero()
    'Procedimento para carregar o cabeçalho da grelha da lista de reprodução
    With Grelha_Genero
        .Clear
        .Rows = 1
        .Cols = 2
        .RowHeight(0) = 270
        .ColWidth(0) = 0
        .ColWidth(1) = 10000
        .ColAlignment(1) = vbleft
        .TextMatrix(0, 1) = Idioma_Grid_Music_Col_5
    End With
End Sub

Public Sub Formatar_Grelha_Album()
    'Procedimento para carregar o cabeçalho da grelha da lista de reprodução
    With Grelha_Album
        .Clear
        .Rows = 1
        .Cols = 2
        .RowHeight(0) = 270
        .ColWidth(0) = 0
        .ColWidth(1) = 10000
        .ColAlignment(1) = vbleft
        .TextMatrix(0, 1) = Idioma_Grid_Music_Col_3
    End With
End Sub

Public Sub Carregar_Grelha_Artista()
    'Procedimento para carregar a grelha
    Verifica_Rs_Musica
    Rs_Musica.Open "select Distinct Artista from Tabela_Musica where len(Artista) > 0 Order By Artista Asc", Cnn_Biblioteca
    Formatar_Grelha_Artista
    
    Dim i As Integer
    With Grelha_Artista
        If Rs_Musica.RecordCount <> 0 Then
            Rs_Musica.MoveFirst
            i = 1
            Do While Not Rs_Musica.EOF
                .Rows = .Rows + 1
                If Rs_Musica("Artista").Value <> "" Then .TextMatrix(i, 1) = Rs_Musica("Artista").Value
                i = i + 1
                Rs_Musica.MoveNext
            Loop
        End If
    End With
End Sub

Public Sub Carregar_Grelha_Genero()
    'Procedimento para carregar a grelha
    Verifica_Rs_Musica
    Rs_Musica.Open "select Distinct Genero from Tabela_Musica where len(Genero) > 0  Order By Genero Asc", Cnn_Biblioteca
    Formatar_Grelha_Genero
    
    Dim i As Integer
    With Grelha_Genero
        If Rs_Musica.RecordCount <> 0 Then
            Rs_Musica.MoveFirst
            i = 1
            Do While Not Rs_Musica.EOF
                .Rows = .Rows + 1
                If Rs_Musica("Genero").Value <> "" Then .TextMatrix(i, 1) = Rs_Musica("Genero").Value
                i = i + 1
                Rs_Musica.MoveNext
            Loop
        End If
    End With
End Sub

Public Sub Carregar_Grelha_Album()
    'Procedimento para carregar a grelha
    Verifica_Rs_Musica
    Rs_Musica.Open "select Distinct Album from Tabela_Musica where len(Album) > 0  order by Album Asc", Cnn_Biblioteca
    Formatar_Grelha_Album
    
    Dim i As Integer
    With Grelha_Album
        If Rs_Musica.RecordCount <> 0 Then
            Rs_Musica.MoveFirst
            i = 1
            Do While Not Rs_Musica.EOF
                .Rows = .Rows + 1
                If Rs_Musica("Album").Value <> "" Then .TextMatrix(i, 1) = Rs_Musica("Album").Value
                i = i + 1
                Rs_Musica.MoveNext
            Loop
        End If
    End With
End Sub

Sub GetInfo()
    'Permite carregar as capas do albuns online
    On Error GoTo Oops
    Me.MousePointer = 11
    Pic_Capa_Album.Picture = Nothing
    Pic_Capa_Album.Picture = Form_Skin.Image_Sem_Capa.Picture
    
    With Wmp.currentMedia
        'Obter as informações da música
        AID = Val(.getItemInfo("AID"))
        Author$ = .getItemInfo("AUTHOR")
        Artist$ = .getItemInfo("Artist")
        Label$ = .getItemInfo("COPYRIGHT")
        Title$ = .getItemInfo("TITLE")
        adType$ = .getItemInfo("ADTYPE")
    
        If LCase(adType$) <> "none" Then
            Wmp.Controls.Next
        End If
    
        If oldTitle$ <> Title$ Then
            albumID = Val(.getItemInfo("ALBUMID"))
            Pic_Capa_Album.Picture = LoadPicture("")
            album$ = ""
            SmallCover$ = ""
            MedCover$ = ""
            LargeCover$ = ""
            
            'Visualizar a capa caso ela esteja disponivel
            If albumID <> 0 Then
                album$ = .getItemInfo("ALBUM")
                SmallCover$ = .getItemInfo("SCOVER")
                MedCover$ = Replace(.getItemInfo("MCOVER"), "LZ", "MZ", , , vbTextCompare)
                LargeCover$ = .getItemInfo("LCOVER")
                Pic_Capa_Album.Picture = OLELoadPicture(LargeCover$)
                ResizePic Pic_Capa_Album
            End If
        End If
    End With
GoTo Exit_GetInfo
Oops:
    If mError = vbRetry Then Resume
    If mError = vbIgnore Then Resume Next
Exit_GetInfo:
    Me.MousePointer = 0
End Sub

Public Function OLELoadPicture(ByVal strFileName As String) As Picture
    'Função para carregar a capa online
    On Error GoTo Oops
    Dim myTGUID As TGUID
    myTGUID.Data1 = &H7BF80980
    myTGUID.Data2 = &HBF32
    myTGUID.Data3 = &H101A
    myTGUID.Data4(0) = &H8B
    myTGUID.Data4(1) = &HBB
    myTGUID.Data4(2) = &H0
    myTGUID.Data4(3) = &HAA
    myTGUID.Data4(4) = &H0
    myTGUID.Data4(5) = &H30
    myTGUID.Data4(6) = &HC
    myTGUID.Data4(7) = &HAB
    OleLoadPicturePath StrPtr(strFileName), 0, 0, 0, myTGUID, OLELoadPicture
    GoTo Exit_LoadPicture
LblError:
    Set OLELoadPicture = VB.LoadPicture(strFileName)
    GoTo Exit_LoadPicture
Oops:
    If mError = vbRetry Then Resume
    If mError = vbIgnore Then Resume Next
Exit_LoadPicture:
End Function

Private Sub Slide_Mini_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o slider na posição pretendida
    DNa = True
    Txa = X
End Sub

Private Sub Slide_Mini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Possicionar a música na posição pretendida
    If DNa Then
        NewLeft = Slide_Mini.left + X - Txa
        If NewLeft < Form_Principal.Image_Barra_Slide_Mini.left + 3 Then
            NewLeft = Form_Principal.Image_Barra_Slide_Mini.left + 3
        End If
        If NewLeft > Form_Principal.Image_Barra_Slide_Mini.Width + Form_Principal.Image_Barra_Slide_Mini.left - 7 - Slide.Width Then
            NewLeft = Form_Principal.Image_Barra_Slide_Mini.Width + Form_Principal.Image_Barra_Slide_Mini.left - 7 - Slide.Width
        End If
        Slide.left = NewLeft
        Slide_Mini.left = NewLeft
        Form_Mini_Player.Slide.left = NewLeft
        Image_Progresso.Width = Slide.left
        Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
    End If
End Sub

Private Sub Slide_Mini_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Colocar o slide na posição largada
    On Error Resume Next
    Dim offseti As Single
    DNa = False
    offseti = (Slide_Mini.left - Form_Principal.Image_Barra_Slide_Mini.left - 3) / (Form_Principal.Image_Barra_Slide_Mini.Width - 10 - Slide_Mini.Width)
    Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    Form_Wmp.Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    Image_Progresso.Width = Slide.left
    Form_Mini_Player.Image_Progresso.Width = Form_Mini_Player.Slide.left
End Sub

Public Sub Parar_o_Player()
    'Procedimento para parar o player
    Musica_Linha_Pressionada = Grelha_Reproduzida.Row
    Timer_Slider_Video.Enabled = False
    'Form_PopUp.Hide
    
    Slide.left = 0: Form_Mini_Player.Slide.left = 0
    Image_Progresso.Width = 1: Form_Mini_Player.Image_Progresso.Width = 1
    Image_Progresso.left = 0: Form_Mini_Player.Image_Progresso.left = 0
    VideoDuration = 0
    Posicao_do_Player = 0
    Wmp.Controls.stop: Form_Wmp.Wmp.Controls.stop
    Wmp.URL = ""
    Faixa_em_Reproducao = ""

    'Reproduzir o som
    Label_Duracao.Caption = "00:00"
    Form_Mini_Player.Label_Duracao.Caption = Label_Duracao.Caption
    Tempo_Estimado.Caption = "00:00"
    Form_Mini_Player.Tempo_Estimado.Caption = Tempo_Estimado.Caption
End Sub

Public Sub Actualizar_Lista()
    'Procedimento para actualizar o ficheiro da lista correspondente
    Dim nome_lista As String: nome_lista = Replace(Label_Topico_Lista(index_lista_selecionada).Caption, Espaco, "")

    If Grelha_Listas.Rows > 1 Then
        Personalizar_Grid Grelha_Listas
        Dim nova_classe As clsFlexSettings
        Set nova_classe = New clsFlexSettings
        Set nova_classe.FlexGrid = Grelha_Listas
        nova_classe.SaveSettings App.Path & "\Library\Playlist\" & nome_lista & ".ini", True, True, True, True
        Set nova_classe = Nothing
    End If
End Sub

Public Sub Repor_Objectos()
    'Procedimento para repor os objectos ao estado normal após a animação do mesmo
    Dim i As Integer: For i = 0 To Label_Botao.count - 1
        Label_Botao(i).FontUnderline = False
    Next
    
    Label_Titulo_Frame_Programas(2).Caption = ""
End Sub

Private Function Aleatorio(Minimo As Long, Maximo As Long) As Long
    'Função para reproduzir as músicas aleatóriamente
    Randomize ' inicializar la semilla
    Aleatorio = CLng((Minimo - Maximo) * Rnd + Maximo)
End Function

Public Sub Formatar_Grelha_Comunidade()
    'Procedimento para carregar o cabeçalho da grelha da lista de reprodução
    With Grelha_Comunidade
        .RowHeight(0) = 270
        .AllowUserResizing = flexResizeColumns
        .Cols = 9
        .Rows = 1
        .ColWidth(0) = 0
        .ColAlignment(0) = vbleft
        .ColWidth(1) = 4000
        .ColAlignment(1) = vbleft
        .ColWidth(2) = 5000
        .ColAlignment(2) = vbleft
        .ColWidth(3) = 1000
        .ColAlignment(3) = vbleft
        .ColWidth(4) = 2000
        .ColAlignment(4) = vbleft
        .ColWidth(5) = 2200
        .ColAlignment(5) = vbleft
        .ColWidth(6) = 0
        .ColAlignment(6) = vbleft
        .ColWidth(7) = 0
        .ColAlignment(7) = vbleft
        .ColWidth(8) = 10000
        .ColAlignment(8) = vbleft
        
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = Idioma_Grid_Community_Col_1
        .TextMatrix(0, 2) = Idioma_Grid_Community_Col_2
        .TextMatrix(0, 3) = Idioma_Grid_Community_Col_3
        .TextMatrix(0, 4) = Idioma_Grid_Community_Col_4
        .TextMatrix(0, 5) = Idioma_Grid_Community_Col_5
        .TextMatrix(0, 6) = Idioma_Grid_Community_Col_6
        .TextMatrix(0, 7) = Idioma_Grid_Community_Col_7
        .TextMatrix(0, 8) = Idioma_Grid_Community_Col_8
    End With
End Sub

Public Sub Formatar_Grelha_Amigos()
    'Procedimento para formatar a grelha amigos
    With Grelha_Amigos
        .RowHeight(0) = 270
        .AllowUserResizing = flexResizeColumns
        .Cols = 9
        .Rows = 1
        .ColWidth(0) = 0
        .ColAlignment(0) = vbleft
        .ColWidth(1) = 3000
        .ColAlignment(1) = vbleft
        .ColWidth(2) = 3000
        .ColAlignment(2) = vbleft
        .ColWidth(3) = 1500
        .ColAlignment(3) = vbleft
        .ColWidth(4) = 2000
        .ColAlignment(4) = vbleft
        .ColWidth(5) = 2500
        .ColAlignment(5) = vbleft
        .ColWidth(6) = 0 '2000 'foto
        .ColAlignment(6) = vbleft
        .ColWidth(7) = 3000
        .ColAlignment(7) = vbleft
        .ColWidth(8) = 10000
        .ColAlignment(8) = vbleft
        
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = Idioma_Grid_Community_Col_1
        .TextMatrix(0, 2) = Idioma_Grid_Community_Col_2
        .TextMatrix(0, 3) = Idioma_Grid_Community_Col_3
        .TextMatrix(0, 4) = Idioma_Grid_Community_Col_4
        .TextMatrix(0, 5) = Idioma_Grid_Community_Col_5
        .TextMatrix(0, 6) = Idioma_Grid_Community_Col_6
        .TextMatrix(0, 7) = Idioma_Grid_Community_Col_7
        .TextMatrix(0, 8) = Idioma_Grid_Community_Col_8
    End With
End Sub

Public Sub Formatar_Grelha_Mensagens()
    'Procedimento para formatar a grelha amigos
    With Grelha_Mensagens
        .RowHeight(0) = 270
        .AllowUserResizing = flexResizeColumns
        .Cols = 7
        .Rows = 1
        .ColWidth(0) = 0
        .ColAlignment(0) = vbleft
        .ColWidth(1) = 3000
        .ColAlignment(1) = vbleft
        .ColWidth(2) = 3000
        .ColAlignment(2) = vbleft
        .ColWidth(3) = 6000
        .ColAlignment(3) = vbleft
        .ColWidth(4) = 2000
        .ColAlignment(4) = vbleft
        .ColWidth(5) = 3000
        .ColAlignment(5) = vbleft
        .ColWidth(6) = 10000
        .ColAlignment(6) = vbleft
        
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = ReadINI("Main", "Grid_Messages_Col_1", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 2) = ReadINI("Main", "Grid_Messages_Col_2", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 3) = ReadINI("Main", "Grid_Messages_Col_3", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 4) = ReadINI("Main", "Grid_Messages_Col_4", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 5) = ReadINI("Main", "Grid_Messages_Col_5", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 6) = ReadINI("Main", "Grid_Messages_Col_6", Localizacao_Ficheiro_Lingua)
    End With
End Sub

Public Sub Carregar_Meus_Amigos()
    'Efectuar pesquisa na base de dados consuante os dados introduzidos
    'On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "carregaramigos.asp?Recebe_Utilizador=" & Form_Perfil.Label_Utilizador.Caption, False
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If servidor.responseText = "false" Then
        Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Me.MousePointer = 11
            Grelha_Amigos.Clear
            Formatar_Grelha_Amigos
            Dados_Servidor_Meus_Amigos servidor.responseText
            Me.MousePointer = 0
        End If
    End If
    Set servidor = Nothing
    
    
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

Private Sub Dados_Servidor_Meus_Amigos(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument

    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/pesquisa/resultado")
        Dim i As Integer: i = Grelha_Amigos.Rows

        For Each node In nodeList
            DoEvents
            Grelha_Amigos.Rows = Grelha_Amigos.Rows + 1

            Dim Data_Dia, Data_Mes, Dia_Ano As String
            If Not IsEmpty(node.selectSingleNode("utilizador")) Then Grelha_Amigos.TextMatrix(i, 1) = node.selectSingleNode("utilizador").Text
            If Not IsEmpty(node.selectSingleNode("nome")) Then Grelha_Amigos.TextMatrix(i, 2) = node.selectSingleNode("nome").Text
            If Not IsEmpty(node.selectSingleNode("genero")) Then Grelha_Amigos.TextMatrix(i, 3) = node.selectSingleNode("genero").Text
            If Not IsEmpty(node.selectSingleNode("dia")) Then Data_Dia = node.selectSingleNode("dia").Text
            If Not IsEmpty(node.selectSingleNode("mes")) Then Data_Mes = node.selectSingleNode("mes").Text
            If Not IsEmpty(node.selectSingleNode("ano")) Then Dia_Ano = node.selectSingleNode("ano").Text
            If Data_Dia <> Empty Then Grelha_Amigos.TextMatrix(i, 4) = Data_Dia & "-" & Data_Mes & "-" & Dia_Ano
            If Not IsEmpty(node.selectSingleNode("pais")) Then Grelha_Amigos.TextMatrix(i, 5) = node.selectSingleNode("pais").Text
            If Not IsEmpty(node.selectSingleNode("foto")) Then Grelha_Amigos.TextMatrix(i, 6) = node.selectSingleNode("foto").Text
            If Not IsEmpty(node.selectSingleNode("email")) Then Grelha_Amigos.TextMatrix(i, 7) = node.selectSingleNode("email").Text
            i = i + 1
        Next
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Public Sub Carregar_Minhas_Mensagens()
    'Efectuar pesquisa na base de dados consuante os dados introduzidos
    'On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "carregarminhasmensagens.asp?Recebe_Utilizador=" & Form_Perfil.Label_Utilizador.Caption, False
    servidor.send
    
    'Verificar os dados acesso
    If servidor.responseText = "false" Then
        Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Me.MousePointer = 11
            Grelha_Mensagens.Clear
            Formatar_Grelha_Mensagens
            Dados_Servidor_Minhas_Mensagens servidor.responseText
            Me.MousePointer = 0
        End If
    End If
    Set servidor = Nothing
    
    
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

Private Sub Dados_Servidor_Minhas_Mensagens(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    Dim total_mensagens As Integer: total_mensagens = 0
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/pesquisa/resultado")
        Dim i As Integer: i = Grelha_Mensagens.Rows
        
        For Each node In nodeList
            DoEvents
            Grelha_Mensagens.Rows = Grelha_Mensagens.Rows + 1

            If Not IsEmpty(node.selectSingleNode("utilizador")) Then Grelha_Mensagens.TextMatrix(i, 1) = node.selectSingleNode("utilizador").Text
            If Not IsEmpty(node.selectSingleNode("assunto")) Then Grelha_Mensagens.TextMatrix(i, 2) = node.selectSingleNode("assunto").Text
            If Not IsEmpty(node.selectSingleNode("mensagem")) Then Grelha_Mensagens.TextMatrix(i, 3) = node.selectSingleNode("mensagem").Text
            If Not IsEmpty(node.selectSingleNode("data")) Then Grelha_Mensagens.TextMatrix(i, 4) = node.selectSingleNode("data").Text
            If Not IsEmpty(node.selectSingleNode("anexo")) Then Grelha_Mensagens.TextMatrix(i, 5) = node.selectSingleNode("anexo").Text
            If Not IsEmpty(node.selectSingleNode("visualizada")) Then
                Grelha_Mensagens.TextMatrix(i, 6) = node.selectSingleNode("visualizada").Text
                If node.selectSingleNode("visualizada").Text = "0" Then
                    total_mensagens = total_mensagens + 1
                End If
            End If
            i = i + 1
        Next
        'Verificar se existem novas mensagens/ mensagens não lindas
        If total_mensagens > 0 Then
            'Label_Barra_Drive(13).Caption = ReadINI("Main", "Label_Bar_Messages", Localizacao_Ficheiro_Lingua) & " (" & total_mensagens & ")"
            'Ajustar_Objectos_Na_Horizontal
            Label_Mensagens.Caption = total_mensagens
            Botao_Mensagens.Visible = True
        End If
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Public Sub Carregar_Meus_Contactos()
    'Efectuar pesquisa na base de dados consuante os dados introduzidos
    'On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "carregarcontactos.asp?Recebe_Utilizador=" & Form_Perfil.Label_Utilizador.Caption, False
    servidor.send
    
    'Verificar os dados acesso
    If servidor.responseText = "false" Then
        Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Me.MousePointer = 11
            Grelha_Contactos.Clear
            Formatar_Grelha_Contactos
            Dados_Servidor_Meus_Contactos servidor.responseText
            Me.MousePointer = 0
        End If
    End If
    Set servidor = Nothing
    
    
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

Private Sub Dados_Servidor_Meus_Contactos(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    Dim nova_mensagem As Integer: nova_mensagem = 0
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/pesquisa/resultado")
        Dim i As Integer: i = Grelha_Contactos.Rows
        
        For Each node In nodeList
            DoEvents
            Grelha_Contactos.Rows = Grelha_Contactos.Rows + 1

            If Not IsEmpty(node.selectSingleNode("id_contacto")) Then Grelha_Contactos.TextMatrix(i, 0) = node.selectSingleNode("id_contacto").Text
            If Not IsEmpty(node.selectSingleNode("nome")) Then Grelha_Contactos.TextMatrix(i, 1) = node.selectSingleNode("nome").Text
            If Not IsEmpty(node.selectSingleNode("morada")) Then Grelha_Contactos.TextMatrix(i, 2) = node.selectSingleNode("morada").Text
            If Not IsEmpty(node.selectSingleNode("pais")) Then Grelha_Contactos.TextMatrix(i, 3) = node.selectSingleNode("pais").Text
            If Not IsEmpty(node.selectSingleNode("genero")) Then Grelha_Contactos.TextMatrix(i, 4) = node.selectSingleNode("genero").Text
            If Not IsEmpty(node.selectSingleNode("data")) Then Grelha_Contactos.TextMatrix(i, 5) = node.selectSingleNode("data").Text
            If Not IsEmpty(node.selectSingleNode("telefone")) Then Grelha_Contactos.TextMatrix(i, 6) = node.selectSingleNode("telefone").Text
            If Not IsEmpty(node.selectSingleNode("telemovel")) Then Grelha_Contactos.TextMatrix(i, 7) = node.selectSingleNode("telemovel").Text
            If Not IsEmpty(node.selectSingleNode("email")) Then Grelha_Contactos.TextMatrix(i, 8) = node.selectSingleNode("email").Text
            If Not IsEmpty(node.selectSingleNode("site")) Then Grelha_Contactos.TextMatrix(i, 9) = node.selectSingleNode("site").Text
            If Not IsEmpty(node.selectSingleNode("observacoes")) Then Grelha_Contactos.TextMatrix(i, 10) = node.selectSingleNode("observacoes").Text
            If Not IsEmpty(node.selectSingleNode("foto")) Then Grelha_Contactos.TextMatrix(i, 11) = node.selectSingleNode("foto").Text
            i = i + 1
        Next
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Public Sub Formatar_Grelha_Eventos()
    'Procedimento para formatar a grelha amigos
    With Grelha_Eventos
        .RowHeight(0) = 270
        .AllowUserResizing = flexResizeColumns
        .Cols = 6
        .Rows = 1
        .ColWidth(0) = 0
        .ColAlignment(0) = vbleft
        .ColWidth(1) = 3000
        .ColAlignment(1) = vbleft
        .ColWidth(2) = 6000
        .ColAlignment(2) = vbleft
        .ColWidth(3) = 2000
        .ColAlignment(3) = vbleft
        .ColWidth(4) = 2000
        .ColAlignment(4) = vbleft
        .ColWidth(5) = 10000
        .ColAlignment(5) = vbleft
        
        .TextMatrix(0, 0) = "ID_Evento"
        .TextMatrix(0, 1) = ReadINI("Main", "Grid_Events_Col_1", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 2) = ReadINI("Main", "Grid_Events_Col_2", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 3) = ReadINI("Main", "Grid_Events_Col_3", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 4) = ReadINI("Main", "Grid_Events_Col_4", Localizacao_Ficheiro_Lingua)
        .TextMatrix(0, 5) = ReadINI("Main", "Grid_Events_Col_5", Localizacao_Ficheiro_Lingua)
    End With
End Sub

Public Sub Carregar_Meus_eventos()
    'Efectuar pesquisa na base de dados consuante os dados introduzidos
    'On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "carregareventos.asp?Recebe_Utilizador=" & Form_Perfil.Label_Utilizador.Caption, False
    servidor.send
    
    'Verificar os dados acesso
    If servidor.responseText = "false" Then
        Mensagem_de_Aviso "Error", ReadINI("Message", "Error_DB_Server_Not_Found", Localizacao_Ficheiro_Lingua)
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Me.MousePointer = 11
            Grelha_Eventos.Clear
            Formatar_Grelha_Eventos
            Dados_Servidor_Meus_eventos servidor.responseText
            Me.MousePointer = 0
        End If
    End If
    Set servidor = Nothing
    
    
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

Private Sub Dados_Servidor_Meus_eventos(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    Dim nova_mensagem As Integer: nova_mensagem = 0
    Dim eventos_marcados As Integer: eventos_marcados = 0
    Dim data_actual As String
    
    'Verificar a data actual
    Dim dia, mes, ano As String: dia = Day(Now): mes = Month(Now): ano = Year(Now)
    If Len(dia) < 2 Then dia = "0" & dia
    If Len(mes) < 2 Then mes = "0" & mes
    data_actual = dia & "-" & mes & "-" & ano
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/pesquisa/resultado")
        Dim i As Integer: i = Grelha_Eventos.Rows
        
        For Each node In nodeList
            DoEvents
            Grelha_Eventos.Rows = Grelha_Eventos.Rows + 1

            If Not IsEmpty(node.selectSingleNode("id_evento")) Then Grelha_Eventos.TextMatrix(i, 0) = node.selectSingleNode("id_evento").Text
            If Not IsEmpty(node.selectSingleNode("titulo")) Then Grelha_Eventos.TextMatrix(i, 1) = node.selectSingleNode("titulo").Text
            If Not IsEmpty(node.selectSingleNode("descricao")) Then Grelha_Eventos.TextMatrix(i, 2) = node.selectSingleNode("descricao").Text
            If Not IsEmpty(node.selectSingleNode("data")) Then Grelha_Eventos.TextMatrix(i, 3) = node.selectSingleNode("data").Text
            If Not IsEmpty(node.selectSingleNode("hora")) Then Grelha_Eventos.TextMatrix(i, 4) = node.selectSingleNode("hora").Text
            If Not IsEmpty(node.selectSingleNode("lembrar")) Then Grelha_Eventos.TextMatrix(i, 5) = node.selectSingleNode("lembrar").Text
            'Verificar se existe eventos para hoje
            If node.selectSingleNode("data").Text = data_actual Then
                If node.selectSingleNode("lembrar").Text = "1" Then eventos_marcados = eventos_marcados + 1
            End If
            i = i + 1
        Next
        'Se existirem eventos para hoje informa ao utilizador
        If eventos_marcados > 0 Then
            Frame_Evento.Visible = True
        End If
    End If
    Set xml = Nothing
    Set nodeList = Nothing
End Sub

Public Sub Formatar_Grelha_Ficheiros()
    'Procedimento para formatar a grelha amigos
    With Grelha_Ficheiros
        .RowHeight(0) = 270
        .AllowUserResizing = flexResizeColumns
        .Cols = 6
        .Rows = 1
        .ColWidth(0) = 0
        .ColAlignment(0) = vbleft
        .ColWidth(1) = 3000
        .ColAlignment(1) = vbleft
        .ColWidth(2) = 6000
        .ColAlignment(2) = vbleft
        .ColWidth(3) = 2000
        .ColAlignment(3) = vbleft
        .ColWidth(4) = 2000
        .ColAlignment(4) = vbleft
        .ColWidth(5) = 10000
        .ColAlignment(5) = vbleft
        
'        .TextMatrix(0, 0) = "ID_Ficheiro"
'        .TextMatrix(0, 1) = ReadINI("Main", "Grid_Events_Col_1", Localizacao_Ficheiro_Lingua)
'        .TextMatrix(0, 2) = ReadINI("Main", "Grid_Events_Col_2", Localizacao_Ficheiro_Lingua)
'        .TextMatrix(0, 3) = ReadINI("Main", "Grid_Events_Col_3", Localizacao_Ficheiro_Lingua)
'        .TextMatrix(0, 4) = ReadINI("Main", "Grid_Events_Col_4", Localizacao_Ficheiro_Lingua)
'        .TextMatrix(0, 5) = ReadINI("Main", "Grid_Events_Col_5", Localizacao_Ficheiro_Lingua)
    End With
End Sub

Public Sub Selecionar_Categoria(Categoria_Selecionada As String, texto_categoria As String)
    'Procedimento para escolher a categoria do programa
    Formatar_Lista_Programas
    
    Label_Frame_Programas(2).Caption = texto_categoria
    Label_Frame_Programas(2).Visible = True
    
    With Separador_Frame_Programas(2)
        Separador_Frame_Programas(2).Stretch = True
        Separador_Frame_Programas(2).Width = Label_Frame_Programas(2).Width + 40
        Separador_Frame_Programas(2).left = Separador_Frame_Programas(1).left + Separador_Frame_Programas(1).Width
        Label_Frame_Programas(2).left = Separador_Frame_Programas(2).left + 20
        .Visible = True
    End With
    
    Categoria_a_ser_Pesquisada = Categoria_Selecionada
    Carregar_Programas Categoria_Selecionada
End Sub

Public Sub Verificar_Pastas()
    'Procedimento para verificar se as pastas utilizadas pelo programa existem
    If Not ArquivoExiste(App.Path & "\Programs", True) Then
        MkDir App.Path & "\Programs\"
    End If
End Sub

Public Sub Carregar_Programas(Categoria_Selecionada As String)
    'Efectuar pesquisa na base de dados consuante os dados introduzidos
    'On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/nplayer/applibrary/" & "carregarprogramas.asp?Recebe_Categoria=" & Categoria_a_ser_Pesquisada
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If servidor.responseText = "NaoExiste" Then
        Label_Nenum_Resultado.Visible = True
        'Exit Sub
        
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Me.MousePointer = 11
            Formatar_Lista_Programas
            Ajustar_Linha_Lista_Programas
            Formatar_Lista_Programas
            Servidor_Carregar_Programas servidor.responseText
            Me.MousePointer = 0
            Frame_Lista.Visible = True
        End If
    End If
    Set servidor = Nothing
    
    Frame_Programas_Home.Visible = False
    Frame_Lista.Visible = True
    
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

Private Sub Servidor_Carregar_Programas(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados da galeria do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/meusprogramas/resultado")
        Dim i As Integer: i = 0
        Dim j As Integer: j = 0
        
        For Each node In nodeList
            DoEvents
            j = j + 1
                        
            If Not IsEmpty(node.selectSingleNode("programa")) Then Label_Programa(i).Caption = node.selectSingleNode("programa").Text
            If Not IsEmpty(node.selectSingleNode("titulo")) Then Label_Nome(i).Caption = node.selectSingleNode("titulo").Text
            If Not IsEmpty(node.selectSingleNode("descricao")) Then Label_Descricao(i).Caption = node.selectSingleNode("descricao").Text
            If Not IsEmpty(node.selectSingleNode("downloads")) Then Label_Downloads(i).Caption = node.selectSingleNode("downloads").Text
            If Not IsEmpty(node.selectSingleNode("observacoes")) Then Label_Observacoes(i).Caption = node.selectSingleNode("observacoes").Text
            If Not IsEmpty(node.selectSingleNode("icon")) Then Label_Icon(i).Caption = node.selectSingleNode("icon").Text
            If Not IsEmpty(node.selectSingleNode("logotipo")) Then Label_Logotipo(i).Caption = node.selectSingleNode("logotipo").Text
            If Not IsEmpty(node.selectSingleNode("tela")) Then Label_Tela(i).Caption = node.selectSingleNode("tela").Text
            If Not IsEmpty(node.selectSingleNode("avaliacao")) Then Label_Avaliacao(i).Caption = node.selectSingleNode("avaliacao").Text
            If Not IsEmpty(node.selectSingleNode("id")) Then Label_Id(i).Caption = node.selectSingleNode("id").Text
            If Not IsEmpty(node.selectSingleNode("site")) Then Label_Site(i).Caption = node.selectSingleNode("site").Text
            
            'Carregar o icon, logotipo e tela do programa
            If Label_Icon(i).Caption <> Empty Then Set Icon_Programa(i).Picture = LoadPicture("http://www.nikyts.com/nplayer/applibrary/imagens/" & Label_Icon(i).Caption)
            If Label_Logotipo(i).Caption <> Empty Then Set Logotipo_Programa(i).Picture = LoadPicture("http://www.nikyts.com/nplayer/applibrary/imagens/" & Label_Logotipo(i).Caption)
            If Label_Tela(i).Caption <> Empty Then Set Tela_Programa(i).Picture = LoadPicture("http://www.nikyts.com/nplayer/applibrary/imagens/" & Label_Tela(i).Caption)
            Pic_Linha(i).Visible = True
            
            '------------------------------------------------------------------------------------------------------------------------
            'Verificar se as pastas utilizadas pelo programa existem
            ficheiro = App.Path & "\Programs\" & Label_Nome(i).Caption & "\" '& Label_Frame_Informacoes(0).Caption & ".exe"
            If ArquivoExiste(ficheiro, True) Then
'                Label_Remover_Transferencia(i).Caption = Idioma_Button_Remove_Program
                sVar = DataArq(ficheiro & Label_Nome(i).Caption & ".exe")
                If sVar <> "ERRO" Then
                    'Label_Frame_Informacoes(3).Caption = ReadINI("Main", "Label_Installed_In", Localizacao_Ficheiro_Lingua) & ": " & sVar
                    Botao_Executar_Programa(i).Enabled = True
                    Label_Executar_Programa(i).Enabled = True
                    'Botao_Remover_Transferencia(i).Enabled = True
                    'Label_Remover_Transferencia(i).Enabled = True
                    Label_Remover_Transferencia(i).Caption = Idioma_Button_Remove_Program
                Else
                    'Label_Frame_Informacoes(3).Caption = Label_Nome(i).Caption & ".zip"
                    Botao_Executar_Programa(i).Enabled = False
                    Label_Executar_Programa(i).Enabled = False
                    'Botao_Remover_Transferencia(i).Enabled = False
                    'Label_Remover_Transferencia(i).Enabled = False
                    Label_Remover_Transferencia(i).Caption = Idioma_Button_Transfer_Program
                End If
            Else
                Botao_Executar_Programa(i).Enabled = False
                Label_Executar_Programa(i).Enabled = False
                'Botao_Remover_Transferencia(i).Enabled = False
                'Label_Remover_Transferencia(i).Enabled = False
                Label_Remover_Transferencia(i).Caption = Idioma_Button_Transfer_Program
            End If
            '------------------------------------------------------------------------------------------------------------------------
            
            i = i + 1
        Next
    End If
End Sub

Public Sub Formatar_Lista_Programas()
    'Procedimento para limpar todos os campos da lista de programas
    Dim i As Integer: For i = 0 To Pic_Linha.count - 1
        Pic_Linha(i).Visible = False
        Pic_Linha(i).Height = Form_Skin.Linha_Normal.Height
        
        Icon_Programa(i).Picture = LoadPicture("")
        Logotipo_Programa(i).Picture = LoadPicture("")
        Tela_Programa(i).Picture = LoadPicture("")
        
        Label_Programa(i).Caption = Empty
        Label_Nome(i).Caption = Empty
        Label_Descricao(i).Caption = Empty
        Label_Downloads(i).Caption = Empty
        Label_Observacoes(i).Caption = Empty
        Label_Icon(i).Caption = Empty
        Label_Logotipo(i).Caption = Empty
        Label_Tela(i).Caption = Empty
        Label_Avaliacao(i).Caption = Empty
        Label_Id(i).Caption = Empty
        Label_Site(i).Caption = Empty
        Label_Nome(i).ForeColor = vbBlack
        Label_Descricao(i).ForeColor = &H808080
        Progresso(i).Visible = False
        Pic_Linha(i).Height = Form_Skin.Linha_Normal.Height
        Pic_Linha(i).backcolor = vbWhite
        Botao_Mais_Informacoes(i).Picture = Form_Skin.Botao_Linha_Normal.Picture
        Botao_Mais_Informacoes(i).Visible = False
        Botao_Remover_Transferencia(i).Visible = False
        Botao_Executar_Programa(i).Visible = False
    Next i
    
    Label_Nenum_Resultado.Visible = False
End Sub

Private Sub txtZip_Change()
    'Indicação de progresso da compactação/descompactação por arquivo
    '----------------------------------------------------------------
    'Tipo de ação que esta sendo feita no momento
    lblProgresso = TipoAção(Val(GetAction(txtZip.Text))) & " "
    'Nome do arquivo que esta sendo compactado
    lblProgresso = lblProgresso & GetFileName(txtZip.Text) & " -> "
    'Porcentagem de compactação do arquivo
    lblProgresso = lblProgresso & GetPercentComplete(txtZip.Text) & "%"
    'Força a atualização da tela
    DoEvents
End Sub

Public Sub Verificar_Downloads()
    'Procedimento para verificar o total de downloads de cada programa
    'On Error GoTo Corrige_Erro
    If Label_Transferencias.Caption = "" Then Exit Sub
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    
    'Adicionar um voto á avaliação do programa
    Label_Transferencias.Caption = Val(Label_Transferencias.Caption) + 1
    servidor.Open "GET", "http://www.nikyts.com/nplayer/applibrary/" & "actualizardownloads.asp?id_programa=" & Label_Id_Programa.Caption & "&downloads=" & Label_Transferencias.Caption, False
    servidor.send

    '"http://www.nikyts.com/nplayer/applibrary/actualizardownloads.asp?id_programa=" + idPrograma
    
'    'Actualizar a senha
'    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
'        With Form_Principal
'            If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
'                'Adicionar + 1 download ao total de dpwnloads do programa
'                If Val(Label_Transferencias.Caption) = 1 Then
'                    Label_Frame_Informacoes(5).Caption = "(" & Label_Transferencias.Caption & " download)"
'                Else
'                    Label_Frame_Informacoes(5).Caption = "(" & Label_Transferencias.Caption & " downloads)"
'                End If
'            End If
'        End With
'    End If
    
    
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

Public Sub Verificar_Se_Programa_Existe()
    On Error Resume Next
    'Procedimento para verificar se o programa em questão já foi instalado
    Dim ficheiro, sVar As String
    ficheiro = App.Path & "\Programs\" & Label_Frame_Informacoes(0).Caption & "\" '& Label_Frame_Informacoes(0).Caption & ".exe"
    
    'Procedimento para verificar se as pastas utilizadas pelo programa existem
    If ArquivoExiste(ficheiro, True) Then
        Label_Botao_Frame_Informacoes(0).Caption = Idioma_Button_Remove_Program
        sVar = DataArq(ficheiro & Label_Frame_Informacoes(0).Caption & ".exe")
        If sVar <> "ERRO" Then
            Label_Frame_Informacoes(3).Caption = ReadINI("Main", "Label_Installed_In", Localizacao_Ficheiro_Lingua) & ": " & sVar
            Barra_Estado_Visivel True
            Botao_Frame_Informacoes(1).Enabled = False
            Label_Botao_Frame_Informacoes(1).Enabled = False
            
        Else
            Label_Frame_Informacoes(3).Caption = Label_Frame_Informacoes(0).Caption & ".zip"
            Label_Botao_Frame_Informacoes(0).Caption = Idioma_Button_Transfer_Program
            Barra_Estado_Visivel False
        End If
    
    Else
        Label_Botao_Frame_Informacoes(0).Caption = Idioma_Button_Transfer_Program
        Barra_Estado_Visivel False
    End If
End Sub

Private Sub Barra_Estado_Visivel(visivel As Boolean)
    'Procedimento para ver/ocultar objectos da barra de estado
    Shape_Estado.Visible = visivel
    Image_Download.Visible = visivel
    Label_Frame_Informacoes(6).Visible = visivel
    Botao_Frame_Informacoes(1).Visible = visivel
    Botao_Frame_Informacoes(2).Visible = visivel
End Sub

Public Sub DeleteFolderTree(ByVal vFolder As String)
    'Procedimento para eliminar a pasta, sub-pastas e respectivos ficheiros referentes ao programa
    Dim FSO As FileSystemObject
    Dim FoldersObj As Folders
    Dim FolderObj As Folder
    Set FSO = New FileSystemObject
    
    If Not FSO.FolderExists(vFolder) Then
    Set FSO = Nothing
    Exit Sub
    End If
    
    Set FolderObj = FSO.GetFolder(vFolder)
    Set FoldersObj = FolderObj.SubFolders
    For Each FolderObj In FoldersObj
    DeleteFolderTree FolderObj.Path
    Next FolderObj
    On Error Resume Next
    
    Kill vFolder & "\*.*"
    RmDir vFolder
    
    err.Clear
    On Error GoTo 0
    
    Set FolderObj = Nothing
    Set FoldersObj = Nothing
    Set FSO = Nothing
End Sub

Private Sub Download_Programa_DowloadComplete()
    'Transferência concluida
    GetFileName (Text_Servidor.Text)
    Progresso(0).Value = 0
    GetFileName (Text_Servidor.Text)
    
    Botao_Remover_Transferencia(Linha_Programa_Selecionado).Visible = True
    Botao_Executar_Programa(Linha_Programa_Selecionado).Enabled = True
    Label_Executar_Programa(Linha_Programa_Selecionado).Enabled = True
    Progresso(progress).Visible = False
    
    'Actualiza no servidor nº de downloads do programa
    Label_Transferencias.Caption = Val(Label_Downloads(Linha_Programa_Selecionado).Caption) + 1
    Label_Id_Programa.Caption = Label_Downloads(Linha_Programa_Selecionado).Caption
    Verificar_Downloads
    
    'Iniciar a decompactação do programa zipado
    DesCompacta App.Path & "\Programs\" & Label_Programa(Linha_Programa_Selecionado).Caption, "*.*", App.Path & "\Programs\", True
    Kill App.Path & "\Programs\" & Label_Programa(Linha_Programa_Selecionado).Caption
    
    'Ao terminar a transferência do ficheiro a Idioma_Button_Transfer_Program passa a ser Idioma_Button_Remove_Program
    Label_Remover_Transferencia(Linha_Programa_Selecionado).Caption = Idioma_Button_Remove_Program
    
    'Actualizar a data e hora de criação do programa
    Dim Ficheiro_Para_Actualizar As String
    Ficheiro_Para_Actualizar = App.Path & "\Programs\" & Label_Nome(Linha_Programa_Selecionado).Caption & "\" & Label_Nome(Linha_Programa_Selecionado).Caption & ".exe"
    
    'Set the creation time
    FileSetDate Ficheiro_Para_Actualizar, Now, True
    'Set the last accessed time
    FileSetDate Ficheiro_Para_Actualizar, Now, , True
    'Set the last write time
    FileSetDate Ficheiro_Para_Actualizar, Now, , , True
    
    Me.MousePointer = 0
End Sub

Private Sub Download_Programa_DownloadErrors(strError As String)
    'Caso ocorra um erro durante o download
    Label_Remover_Transferencia(0).Caption = Idioma_Button_Transfer_Program
    Botao_Remover_Transferencia(Linha_Programa_Selecionado).Visible = True
    Botao_Executar_Programa(Linha_Programa_Selecionado).Enabled = True
    Label_Executar_Programa(Linha_Programa_Selecionado).Enabled = True
    Progresso(progress_activo).Visible = False
    
    Mensagem_de_Aviso "Error", ReadINI("Main", "Error_Transfer_Program", Localizacao_Ficheiro_Lingua)
    Me.MousePointer = 0
End Sub

Private Sub Download_Programa_DownloadProgress(intPercent As String)
    'Mostrar o progresso do download
    Progresso(0).Value = intPercent
    GetFileName (Text_Servidor.Text)
    Text_Servidor.Text = ""
End Sub

Public Sub Ajustar_Linha_Lista_Programas()
    'Procedimento para ajustar as linhas da lista dos programas
    Dim Linha, Altura As Integer
    Linha = 0: Altura = 0
    For Linha = 0 To Pic_Linha.count - 1
        With Pic_Linha(Linha)
            .top = Altura
        End With
        Altura = Altura + Pic_Linha(Linha).ScaleHeight
    Next Linha
End Sub

Private Sub Pic_Linha_Click(Index As Integer)
    'Selecionar a linha
    If Pic_Linha(Index).Height = Form_Skin.Linha_Normal.Height Then
        Repor_Altura_das_Linhas
        Pic_Linha(Index).backcolor = Azul
        Botao_Mais_Informacoes(Index).Picture = Form_Skin.Botao_Linha_Over.Picture
        Botao_Remover_Transferencia(Index).Picture = Form_Skin.Botao_Linha_2_Over.Picture
        Botao_Executar_Programa(Index).Picture = Form_Skin.Botao_Linha_2_Over.Picture
        Pic_Linha(Index).Height = Form_Skin.Linha_Over.Height
        Botao_Mais_Informacoes(Index).Visible = True
        Botao_Remover_Transferencia(Index).Visible = True
        Botao_Executar_Programa(Index).Visible = True
        Label_Nome(Index).ForeColor = vbWhite
        Label_Descricao(Index).ForeColor = vbWhite
        
    Else
        Pic_Linha(Index).Height = Form_Skin.Linha_Normal.Height
        Pic_Linha(Index).backcolor = vbWhite
        Botao_Mais_Informacoes(Index).Picture = Form_Skin.Botao_Linha_Normal.Picture
        Botao_Remover_Transferencia(Index).Picture = Form_Skin.Botao_Linha_2_Normal.Picture
        Botao_Executar_Programa(Index).Picture = Form_Skin.Botao_Linha_2_Normal.Picture
        Botao_Mais_Informacoes(Index).Visible = False
        Botao_Remover_Transferencia(Index).Visible = False
        Botao_Executar_Programa(Index).Visible = False
        Label_Nome(Index).ForeColor = vbBlack
        Label_Descricao(Index).ForeColor = &H808080
    End If
    
    Ajustar_Linha_Lista_Programas
End Sub

Public Sub Repor_Altura_das_Linhas()
    'Procedimento para repor a altura de todas as linhas da lista de programas
    Dim Linha As Integer
    Linha = 0
    For Linha = 0 To Pic_Linha.count - 1
        With Pic_Linha(Linha)
            .Height = Form_Skin.Linha_Normal.Height
            .backcolor = vbWhite
            Botao_Mais_Informacoes(Linha).Visible = False
            Botao_Remover_Transferencia(Linha).Visible = False
            Botao_Executar_Programa(Linha).Visible = False
            Label_Nome(Linha).ForeColor = vbBlack
            Label_Descricao(Linha).ForeColor = &H808080
        End With
    Next Linha
End Sub

