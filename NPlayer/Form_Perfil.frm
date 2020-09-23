VERSION 5.00
Begin VB.Form Form_Perfil 
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15225
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
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00222222&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   6615
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "A minha conta"
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
         TabIndex        =   15
         Top             =   120
         Width           =   1395
      End
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   6120
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
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   921
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   465
      Width           =   13815
      Begin VB.PictureBox Separador_Perfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   56
         Top             =   180
         Width           =   975
         Begin VB.Label Label_Perfil 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Perfil"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   58
            Top             =   120
            Width           =   435
         End
         Begin VB.Shape Shape_Perfil 
            Height          =   255
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox Separador_Senha 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   57
         Top             =   180
         Width           =   975
         Begin VB.Label Label_Senha 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Senha"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   540
         End
         Begin VB.Shape Shape_Senha 
            Height          =   255
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox Text_Password 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7080
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox Text_Usuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   6000
         TabIndex        =   53
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox Text_Servidor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4920
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox Text_Perfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3840
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.PictureBox Frame_Senha 
         Appearance      =   0  'Flat
         BackColor       =   &H00313131&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   9000
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   385
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   5775
         Begin VB.PictureBox Pic_Visualizar 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   2760
            Width           =   195
         End
         Begin VB.PictureBox Barra_Text_Nova_Senha 
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
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1380
            Width           =   5475
            Begin VB.TextBox Text_Nova_Senha 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   600
               PasswordChar    =   "*"
               TabIndex        =   9
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Nova_Senha 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.PictureBox Barra_Text_Confirmar 
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
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   2100
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
               TabIndex        =   10
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
         Begin VB.PictureBox Barra_Text_Senha_Actual 
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
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   660
            Width           =   5475
            Begin VB.TextBox Text_Senha_Actual 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   600
               PasswordChar    =   "*"
               TabIndex        =   8
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Senha_Actual 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.CheckBox Check_Visualizar 
            BackColor       =   &H00313131&
            Caption         =   "Visualizar o conteúdo dos campos"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   49
            Top             =   2760
            Width           =   4575
         End
         Begin VB.Label Label_Esqueceu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Esqueceu-se dos seus dados de acesso?"
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
            Left            =   0
            TabIndex        =   50
            Top             =   3240
            Width           =   3465
         End
         Begin VB.Label Label_Nova_Senha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nova senha"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   47
            Top             =   1140
            Width           =   1005
         End
         Begin VB.Label Label_Confirmar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirmar a nova senha"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   46
            Top             =   1860
            Width           =   2085
         End
         Begin VB.Image Image_Erro2 
            Enabled         =   0   'False
            Height          =   210
            Left            =   60
            Picture         =   "Form_Perfil.frx":0000
            Top             =   60
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label_Erro2 
            AutoSize        =   -1  'True
            BackColor       =   &H00F5F5F5&
            BackStyle       =   0  'Transparent
            Caption         =   "Senha inválida."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   45
            Top             =   60
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label_Senha_Actual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Senha actual"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   44
            Top             =   420
            Width           =   1110
         End
         Begin VB.Shape Shape_Erro2 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H008080FF&
            Height          =   315
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   5475
         End
      End
      Begin VB.PictureBox Frame_Perfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00313131&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   240
         ScaleHeight     =   281
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   561
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   8415
         Begin VB.TextBox Text_Foto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   6000
            TabIndex        =   60
            Top             =   360
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.PictureBox Moldura_Foto 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2070
            Left            =   6000
            ScaleHeight     =   138
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   139
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   720
            Width           =   2085
            Begin VB.Image Image_Foto 
               Height          =   1920
               Left            =   75
               Top             =   75
               Width           =   1920
            End
         End
         Begin VB.PictureBox Barra_Text_Nome 
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
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   660
            Width           =   5475
            Begin VB.TextBox Text_Nome 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   600
               TabIndex        =   0
               Top             =   0
               Width           =   1380
            End
            Begin VB.Shape Contorno_Nome 
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
            Left            =   0
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1380
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
         Begin VB.PictureBox Pic_Feminino 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2535
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   3420
            Width           =   210
         End
         Begin VB.PictureBox Pic_Masculino 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   960
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   3420
            Width           =   210
         End
         Begin VB.PictureBox Barra_Text_Pais 
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
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2880
            Width           =   5475
            Begin VB.TextBox Text_Pais 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   600
               TabIndex        =   5
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Pais 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.PictureBox Barra_Text_Dia 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   0
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   93
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   2130
            Width           =   1395
            Begin VB.TextBox Text_Dia 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   600
               MaxLength       =   2
               TabIndex        =   2
               Top             =   30
               Width           =   300
            End
            Begin VB.Shape Contorno_Dia 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.PictureBox Barra_Text_Mes 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   1440
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   93
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   2130
            Width           =   1395
            Begin VB.TextBox Text_Mes 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   600
               MaxLength       =   2
               TabIndex        =   3
               Top             =   30
               Width           =   300
            End
            Begin VB.Shape Contorno_Mes 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.PictureBox Barra_Text_Ano 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   2880
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   85
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2130
            Width           =   1275
            Begin VB.TextBox Text_Ano 
               Appearance      =   0  'Flat
               BackColor       =   &H00101010&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   600
               MaxLength       =   4
               TabIndex        =   4
               Top             =   0
               Width           =   300
            End
            Begin VB.Shape Contorno_Ano 
               BorderColor     =   &H00D88316&
               Height          =   375
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.OptionButton Opcao_Feminino 
            BackColor       =   &H00313131&
            Caption         =   "Feminino"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   2520
            TabIndex        =   7
            Top             =   3420
            Width           =   1335
         End
         Begin VB.OptionButton Opcao_Masculino 
            BackColor       =   &H00313131&
            Caption         =   "Masculino"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   960
            TabIndex        =   6
            Top             =   3420
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label Label_Pais 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pais"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   39
            Top             =   2640
            Width           =   345
         End
         Begin VB.Image Image_Erro 
            Enabled         =   0   'False
            Height          =   210
            Left            =   60
            Picture         =   "Form_Perfil.frx":02AA
            Top             =   60
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label_Erro 
            AutoSize        =   -1  'True
            BackColor       =   &H00F5F5F5&
            BackStyle       =   0  'Transparent
            Caption         =   "Indique um endereço de email válido."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   38
            Top             =   60
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Label Label_Nome_Completo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome completo"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   37
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label_Email 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   36
            Top             =   1140
            Width           =   465
         End
         Begin VB.Label Label_Genero 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Género"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   35
            Top             =   3435
            Width           =   630
         End
         Begin VB.Label Label_Data_Nascimento 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data de nascimento (dd/mm/aaaa)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   34
            Top             =   1890
            Width           =   3030
         End
         Begin VB.Label Label_Utilizador 
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6000
            TabIndex        =   33
            Top             =   2880
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label_Minha_Senha 
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6000
            TabIndex        =   32
            Top             =   3240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Shape Shape_Erro 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H008080FF&
            Height          =   315
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   5475
         End
      End
      Begin VB.Shape Shape_Frame 
         Height          =   495
         Left            =   3120
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1080
         TabIndex        =   22
         Top             =   2940
         Width           =   75
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   21
         Top             =   2940
         Width           =   75
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
      AutoRedraw      =   -1  'True
      BackColor       =   &H002A2A2A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   561
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5400
      Width           =   8415
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
         TabIndex        =   12
         Top             =   120
         Width           =   1740
         Begin VB.Label Label_Ok 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ok"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   750
            TabIndex        =   19
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
         TabIndex        =   13
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
            TabIndex        =   18
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
         TabIndex        =   11
         Top             =   120
         Width           =   1740
         Begin VB.Label Label_Actualizar 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actualizar"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   450
            TabIndex        =   17
            Top             =   45
            Width           =   840
         End
         Begin VB.Shape Contorno_Actualizar 
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
Attribute VB_Name = "Form_Perfil"
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

'Variáveis do idioma
Dim Idioma_Erro As String
Dim Idioma_Descricao As String
Dim Idioma_Erro_Execucao As String
Dim Idioma_Conectar_Servidor As String
Dim Idioma_Internet_Desligada As String

Dim Idioma_Error_Email_Invalid As String
Dim Idioma_Error_TextBox_Required As String
Dim Idioma_Info_Profile_Update As String
Dim Idioma_Error_Current_Password As String
Dim Idioma_Error_Password_Characters As String
Dim Idioma_Error_Confirm_Password As String
Dim Idioma_Info_Password_Changed As String

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

Private Sub Check_Visualizar_Click()
    'Des/Activar a opcção
    If Check_Visualizar.Value = 1 Then
        Pic_Visualizar.Picture = Form_Skin.Check_Over.Picture
        Text_Senha_Actual.PasswordChar = ""
        Text_Nova_Senha.PasswordChar = ""
        Text_Confirmar.PasswordChar = ""
        
    Else
        Pic_Visualizar.Picture = Form_Skin.Check_Normal.Picture
        Text_Senha_Actual.PasswordChar = "*"
        Text_Nova_Senha.PasswordChar = "*"
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
    
    'Estado inicial dos separadores
    Separador_Perfil.Height = Form_Skin.Separador_Mini_Normal.Height + 1
    Separador_Senha.Height = Form_Skin.Separador_Mini_Normal.Height
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Botao_Cancelar_Click
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Frame_Centro.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Frame_Perfil.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Frame_Senha.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Centro.BorderColor = .Cor_Contorno_Frame_Centro.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Contorno_Nome.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Email.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Dia.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Mes.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Ano.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Pais.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Senha_Actual.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Nova_Senha.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Confirmar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Nome_Completo.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Email.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Data_Nascimento.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Pais.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Genero.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Opcao_Masculino.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Opcao_Masculino.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Opcao_Feminino.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Opcao_Feminino.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Label_Senha_Actual.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Nova_Senha.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Confirmar.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Check_Visualizar.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Check_Visualizar.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Barra_Text_Nome.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Nome.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Nome.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Nome.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Nome.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Email.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Email.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Dia.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Dia.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Dia.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Dia.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Dia.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Mes.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Mes.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Mes.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Mes.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Mes.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Ano.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Ano.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Ano.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Ano.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Ano.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Pais.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Pais.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Pais.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Pais.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Pais.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Senha_Actual.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Senha_Actual.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Senha_Actual.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Senha_Actual.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Senha_Actual.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Nova_Senha.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Nova_Senha.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Nova_Senha.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Nova_Senha.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Nova_Senha.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Barra_Text_Confirmar.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Confirmar.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Confirmar.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Confirmar.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Confirmar.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Pic_Masculino.Picture = .Opcao_Over.Picture
        Pic_Masculino.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Pic_Feminino.Picture = .Opcao_Normal.Picture
        Pic_Feminino.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Pic_Visualizar.Picture = .Check_Normal.Picture
        Image_Foto.Picture = .Foto_Masculino.Picture
        Label_Esqueceu.ForeColor = .Cor_Contorno_Caixas.backcolor
        Text_Nome.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Nome.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Email.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Email.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Dia.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Dia.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Mes.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Mes.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Ano.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Ano.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Pais.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Pais.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Senha_Actual.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Senha_Actual.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Nova_Senha.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Nova_Senha.ForeColor = .Cor_Letra_Textbox.backcolor
        Text_Confirmar.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Confirmar.ForeColor = .Cor_Letra_Textbox.backcolor
        Moldura_Foto.Picture = .Moldura_Foto.Picture
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Actualizar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Actualizar.Picture = .Pic_Button.Picture
        Botao_Ok.Picture = .Pic_Button.Picture
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Actualizar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Perfil.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Label_Senha.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Moldura_Foto.Picture = Nothing
        Moldura_Foto.backcolor = .Cor_Fundo_Textbox.backcolor
        Moldura_Foto.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 10, 0, 0, 10, 10
        Moldura_Foto.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Moldura_Foto.ScaleWidth, 10, 10, 0, 40, 10
        Moldura_Foto.PaintPicture Form_Skin.Pic_TextBox.Picture, (Moldura_Foto.ScaleWidth - 10), 0, 10, 10, 51, 0, 10, 10
        Moldura_Foto.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 10, 10, (Moldura_Foto.ScaleHeight - 20), 0, 10, 10, 10
        Moldura_Foto.PaintPicture Form_Skin.Pic_TextBox.Picture, (Moldura_Foto.ScaleWidth - 10), 10, 10, (Moldura_Foto.ScaleHeight - 20), 51, 10, 10, 10
        Moldura_Foto.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, (Moldura_Foto.ScaleHeight - 10), 10, 10, 0, 17, 10, 10
        Moldura_Foto.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, (Moldura_Foto.ScaleHeight - 10), (Moldura_Foto.ScaleWidth - 20), 10, 10, 17, 40, 10
        Moldura_Foto.PaintPicture Form_Skin.Pic_TextBox.Picture, (Moldura_Foto.ScaleWidth - 10), (Moldura_Foto.ScaleHeight - 10), 10, 10, 51, 17, 10, 10
        Shape_Frame.BorderColor = .Cor_Line_Border_Frames.backcolor
        Separador_Perfil.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Separador_Senha.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Perfil.BorderColor = .Cor_Line_Border_Frames.backcolor
        Shape_Senha.BorderColor = .Cor_Line_Border_Frames.backcolor
    End With
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Profile", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Profile", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Perfil.Caption = ReadINI("Profile", "Label_Profile", Localizacao_Ficheiro_Lingua)
    Label_Senha.Caption = ReadINI("Profile", "Label_Password", Localizacao_Ficheiro_Lingua)
    Label_Nome_Completo.Caption = ReadINI("Profile", "Label_Name", Localizacao_Ficheiro_Lingua)
    Label_Email.Caption = ReadINI("Profile", "Label_Email", Localizacao_Ficheiro_Lingua)
    Label_Data_Nascimento.Caption = ReadINI("Profile", "Label_Date", Localizacao_Ficheiro_Lingua)
    Label_Pais.Caption = ReadINI("Profile", "Label_Country", Localizacao_Ficheiro_Lingua)
    Label_Genero.Caption = ReadINI("Profile", "Label_Gender", Localizacao_Ficheiro_Lingua)
    Opcao_Masculino.Caption = ReadINI("Profile", "Option_Male", Localizacao_Ficheiro_Lingua)
    Opcao_Feminino.Caption = ReadINI("Profile", "Option_Female", Localizacao_Ficheiro_Lingua)
    Label_Senha_Actual.Caption = ReadINI("Profile", "Label_Current_Password", Localizacao_Ficheiro_Lingua)
    Label_Nova_Senha.Caption = ReadINI("Profile", "Label_New_Password", Localizacao_Ficheiro_Lingua)
    Label_Confirmar.Caption = ReadINI("Profile", "Label_Confirm_Password", Localizacao_Ficheiro_Lingua)
    Check_Visualizar.Caption = ReadINI("Profile", "Check_View", Localizacao_Ficheiro_Lingua)
    Label_Esqueceu.Caption = ReadINI("Profile", "Label_Forgot", Localizacao_Ficheiro_Lingua)
    Label_Actualizar.Caption = ReadINI("Profile", "Button_Update", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Profile", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Profile", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    
    Idioma_Erro = ReadINI("Message", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Message", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Message", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Conectar_Servidor = ReadINI("Message", "Error_Connect", Localizacao_Ficheiro_Lingua)
    Idioma_Internet_Desligada = ReadINI("Message", "Error_Internet", Localizacao_Ficheiro_Lingua)
    
    Idioma_Error_Email_Invalid = ReadINI("Message", "Error_Email_Invalid", Localizacao_Ficheiro_Lingua)
    Idioma_Error_TextBox_Required = ReadINI("Message", "Error_TextBox_Required", Localizacao_Ficheiro_Lingua)
    Idioma_Info_Profile_Update = ReadINI("Message", "Info_Profile_Update", Localizacao_Ficheiro_Lingua)
    Idioma_Error_Current_Password = ReadINI("Message", "Error_Current_Password", Localizacao_Ficheiro_Lingua)
    Idioma_Error_Password_Characters = ReadINI("Message", "Error_Password_Characters", Localizacao_Ficheiro_Lingua)
    Idioma_Error_Confirm_Password = ReadINI("Message", "Error_Confirm_Password", Localizacao_Ficheiro_Lingua)
    Idioma_Info_Password_Changed = ReadINI("Message", "Info_Password_Changed", Localizacao_Ficheiro_Lingua)
End Sub

Private Sub Label_Cancelar_Click()
    'Fechar formulário
    Me.Hide
End Sub

Private Sub Label_Esqueceu_Click()
    'Recuperar os dados de acesso
    Form_Recuperar_Conta.Show vbModal
End Sub

Private Sub Label_Ok_Click()
    'Fechar o formulário
    Me.Hide
End Sub

Private Sub Label_Perfil_Click()
    'Ver frame perfil
    Frame_Perfil.Visible = True
    Frame_Senha.Visible = False
    Separador_Perfil.Height = Form_Skin.Separador_Mini_Normal.Height + 1
    Separador_Senha.Height = Form_Skin.Separador_Mini_Normal.Height
    Text_Nome.SetFocus
End Sub

Private Sub Label_Actualizar_Click()
    'Concluir a operação
    On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    
    'Actualizar o perfil do utilizador------------------------------------------------------------------------------------------------
    If Frame_Perfil.Visible = True Then
        Shape_Erro.Visible = False
        Label_Erro.Visible = False
        Image_Erro.Visible = False
        
        'Verificar se a senha actual é fálida
        If Text_Email.Text = Empty Then
            Label_Erro.Caption = Idioma_Error_TextBox_Required
            Shape_Erro.Visible = True
            Label_Erro.Visible = True
            Image_Erro.Visible = True
            Text_Email.SetFocus
            Exit Sub
        End If
    
        'Verificar se o email é válido
        If Not IsEmail(Text_Email.Text) Then
            Label_Erro.Caption = Idioma_Error_Email_Invalid
            Shape_Erro.Visible = True
            Label_Erro.Visible = True
            Image_Erro.Visible = True
            Text_Email.SetFocus
            Exit Sub
        End If
        
        'Verificar qual o sexo do utilizador
        Dim Sexo As String
        If Opcao_Masculino.Value = True Then
            Sexo = "M"
        Else
            Sexo = "F"
        End If
        
        'Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
        servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "actualizarconta.asp?Utilizador=" & Form_Perfil.Label_Utilizador.Caption & "&Email=" & Text_Email.Text & "&Genero=" & Sexo & "&Dia=" & Text_Dia.Text & "&Mes=" & Text_Mes.Text & "&Ano=" & Text_Ano.Text & "&Pais=" & Text_Pais.Text & "&Nome=" & Text_Nome.Text & "&Foto=" & Text_Foto.Text & "&Servidor=" & Text_Servidor.Text & "&Usuario=" & Text_Usuario.Text & "&Senha2=" & Text_Password.Text & "&Perfil=" & Text_Perfil.Text, False
        servidor.send
        'Actualizar a senha
        If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
            With Form_Principal
                If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
                    Mensagem_de_Aviso "Information", Idioma_Info_Profile_Update
                End If
            End With
        End If
    
    
    'Actualizar a senha do utilizador------------------------------------------------------------------------------------------------
    Else
        If Text_Senha_Actual.Text = "" Or Text_Nova_Senha.Text = "" Or Text_Confirmar.Text = "" Then Exit Sub
        Shape_Erro2.Visible = False
        Label_Erro2.Visible = False
        Image_Erro2.Visible = False
        
        'Verificar se a senha actual é fálida
        If Text_Senha_Actual.Text <> Label_Minha_Senha.Caption Then
            Label_Erro.Caption = Idioma_Error_Current_Password
            Shape_Erro2.Visible = True
            Label_Erro2.Visible = True
            Image_Erro2.Visible = True
            Text_Senha_Actual.Text = ""
            Text_Nova_Senha.Text = ""
            Text_Confirmar.Text = ""
            Text_Senha_Actual.SetFocus
            Exit Sub
        End If
        
        'Verificar o tamanho da senha
        If Len(Text_Nova_Senha.Text) < 6 Or Len(Text_Confirmar.Text) < 6 Then
            Label_Erro.Caption = Idioma_Error_Password_Characters
            Shape_Erro2.Visible = True
            Label_Erro2.Visible = True
            Image_Erro2.Visible = True
            Text_Nova_Senha.Text = ""
            Text_Confirmar.Text = ""
            Text_Nova_Senha.SetFocus
            Exit Sub
        End If
        
        'Verificar se a nova senha é = á confirmação de senha
        If Text_Nova_Senha.Text <> Text_Confirmar.Text Then
            Label_Erro.Caption = Idioma_Error_Confirm_Password
            Shape_Erro2.Visible = True
            Label_Erro2.Visible = True
            Image_Erro2.Visible = True
            Text_Nova_Senha.Text = ""
            Text_Confirmar.Text = ""
            Text_Nova_Senha.SetFocus
            Exit Sub
        End If
        
        'Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
        servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "actualizarsenha.asp?utilizador=" & Label_Utilizador.Caption & "&Senha=" & Text_Nova_Senha.Text, False
        servidor.send 'envia o pedido para o servidor
    
        'Actualizar a senha
        If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
            With Form_Principal
            
                If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
                    Mensagem_de_Aviso "Information", Idioma_Info_Password_Changed
                End If
            End With
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

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Nome.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * (Barra_Text_Nome.ScaleWidth + (2 * 10) + (2 * 16) + (Moldura_Foto.Width + 16 + 16)) '2*10 -> Left da frame centro | 2 * 16 -> Left das frames perfil e senha | +16 espaço entre a text e a moldura
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Separador_Perfil.Height + 16 + Shape_Erro.Height _
                + Label_Nome_Completo.Height + 3 + Form_Skin.Caixa_de_Texto.Height + 6 + Label_Email.Height + 3 + Form_Skin.Caixa_de_Texto.Height + 6 _
                + Label_Data_Nascimento.Height + 3 + Form_Skin.Caixa_de_Texto.Height + 6 + Label_Pais.Height + 3 + Form_Skin.Caixa_de_Texto.Height _
                + 16 + Label_Genero.Height + Fundo_Frame_Botoes.Height + (3 * 16))
    End With
    
    Ajustar_Formulario Form_Perfil, False, False, True, True
    
    Ajustar_Botao Form_Perfil, Botao_Actualizar, Label_Actualizar, True, Contorno_Actualizar
    Ajustar_Botao Form_Perfil, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Perfil, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With
    With Botao_Actualizar
        .left = Botao_Ok.left - .ScaleWidth - .top
    End With
    
    Ajustar_Caixa_Texto Barra_Text_Nome, Text_Nome, Contorno_Nome, False
    Ajustar_Caixa_Texto Barra_Text_Email, Text_Email, Contorno_Email, False
    Ajustar_Caixa_Texto Barra_Text_Pais, Text_Pais, Contorno_Pais, False
    Ajustar_Caixa_Texto Barra_Text_Senha_Actual, Text_Senha_Actual, Contorno_Senha_Actual, False
    Ajustar_Caixa_Texto Barra_Text_Nova_Senha, Text_Nova_Senha, Contorno_Nova_Senha, False
    Ajustar_Caixa_Texto Barra_Text_Confirmar, Text_Confirmar, Contorno_Confirmar, False
    Ajustar_Caixa_Texto_Mini Barra_Text_Dia, Text_Dia, Contorno_Dia
    Ajustar_Caixa_Texto_Mini Barra_Text_Mes, Text_Mes, Contorno_Mes
    Ajustar_Caixa_Texto_Mini Barra_Text_Ano, Text_Ano, Contorno_Ano
    
    With Separador_Perfil
        .Height = Form_Skin.Separador_Mini_Normal.Height
        .top = 10
        .Width = Form_Skin.Separador_Mini_Normal.Width
        .left = 0
    End With
    
    With Shape_Perfil
        .Height = Separador_Perfil.ScaleHeight + 6
        .top = 0
        .Width = Separador_Perfil.ScaleWidth
        .left = 0
    End With
    
    With Separador_Senha
        .Height = Separador_Perfil.ScaleHeight
        .top = Separador_Perfil.top
        .Width = Separador_Perfil.ScaleWidth
        .left = Separador_Perfil.left + Separador_Perfil.Width - 1
    End With
    
    With Shape_Senha
        .Height = Shape_Perfil.Height
        .top = Shape_Perfil.top
        .Width = Shape_Perfil.Width
        .left = Shape_Perfil.left
    End With
    
    With Label_Perfil
        .top = (Separador_Perfil.ScaleHeight - .Height) / 2
        .Width = Separador_Perfil.ScaleWidth
        .left = 0
    End With
    
    With Label_Senha
        .top = Label_Perfil.top
        .Width = Label_Perfil.Width
        .left = Label_Perfil.left
    End With
        
    With Shape_Frame
        .top = Separador_Perfil.top + Separador_Perfil.Height
        .Height = Frame_Centro.ScaleHeight - Separador_Perfil.top - Separador_Perfil.Height - Frame_Centro.left
        .left = 0
        .Width = Frame_Centro.ScaleWidth
    End With
    
    With Frame_Perfil
        .Height = Label_Genero.top + Label_Genero.Height + 16
        .top = Separador_Perfil.top + Separador_Perfil.Height + 16
        .Width = Barra_Text_Nome.ScaleWidth + Moldura_Foto.Width + 16 + 16
        .left = 16
    End With
    
    With Frame_Senha
        .top = Frame_Perfil.top
        .Height = Frame_Perfil.Height
        .Width = Frame_Perfil.Width
        .left = Frame_Perfil.left
    End With
    
    Ajustar_ChecBox Pic_Visualizar, Check_Visualizar
    Ajustar_Option Pic_Masculino
    Ajustar_Option Pic_Feminino
    
    With Shape_Erro
        .top = .left
        .Width = Form_Skin.Caixa_de_Texto.Width
    End With
    
    With Image_Erro
        .top = ((Shape_Erro.top + Shape_Erro.Height) - .Height) / 2
    End With
    
    With Label_Erro
        .top = Image_Erro.top
    End With
    
    With Label_Nome_Completo
        .top = Shape_Erro.top + Shape_Erro.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Nome
        .top = Label_Nome_Completo.top + Label_Nome_Completo.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Email
        .top = Barra_Text_Nome.top + Barra_Text_Nome.ScaleHeight + 6
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Email
        .top = Label_Email.top + Label_Email.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Data_Nascimento
        .top = Barra_Text_Email.top + Barra_Text_Email.ScaleHeight + 6
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Dia
        .top = Label_Data_Nascimento.top + Label_Data_Nascimento.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Mes
        .top = Barra_Text_Dia.top
        .left = Barra_Text_Dia.left + Barra_Text_Dia.Width + 6
    End With
    
    With Barra_Text_Ano
        .top = Barra_Text_Dia.top
        .left = Barra_Text_Mes.left + Barra_Text_Mes.Width + 6
    End With
    
    With Label_Pais
        .top = Barra_Text_Dia.top + Barra_Text_Dia.ScaleHeight + 6
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Pais
        .top = Label_Pais.top + Label_Pais.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Genero
        .top = Barra_Text_Pais.top + Barra_Text_Pais.ScaleHeight + 16
        .left = Shape_Erro.left
    End With
    
    With Opcao_Masculino
        .top = Label_Genero.top
    End With
    
    With Pic_Masculino
        .top = Opcao_Masculino.top
        .left = Opcao_Masculino.left
        .Width = Form_Skin.Opcao_Normal.Width
    End With
    
    With Opcao_Feminino
        .top = Opcao_Masculino.top
    End With
    
    With Pic_Feminino
        .top = Opcao_Feminino.top
        .left = Opcao_Feminino.left
    End With
    
    With Moldura_Foto
        .Height = Form_Skin.Moldura_Foto.Height
        .top = Barra_Text_Nome.top
        .Width = Form_Skin.Moldura_Foto.Width
        .left = Barra_Text_Nome.left + Barra_Text_Nome.ScaleWidth + Barra_Text_Nome.Height
    End With
    
    With Image_Foto
        .top = ((Moldura_Foto.ScaleHeight - .Height) / 2)
        .left = ((Moldura_Foto.ScaleWidth - .Width) / 2)
    End With
    
    'Frame_Senha----------------------------------------------------------------------------
    With Shape_Erro2
        .top = .left
        .Width = Form_Skin.Caixa_de_Texto.Width
    End With
    
    With Image_Erro2
        .top = ((Shape_Erro.top + Shape_Erro.Height) - .Height) / 2
    End With
    
    With Label_Erro2
        .top = Image_Erro2.top
    End With
    
    With Label_Senha_Actual
        .top = Shape_Erro2.top + Shape_Erro2.Height + 3
        .left = Shape_Erro2.left
    End With
    
    With Barra_Text_Senha_Actual
        .top = Label_Senha_Actual.top + Label_Senha_Actual.Height + 3
        .left = Shape_Erro2.left
    End With
    
    With Label_Nova_Senha
        .top = Barra_Text_Senha_Actual.top + Barra_Text_Senha_Actual.ScaleHeight + 6
        .left = Shape_Erro2.left
    End With
    
    With Barra_Text_Nova_Senha
        .top = Label_Nova_Senha.top + Label_Nova_Senha.Height + 3
        .left = Shape_Erro2.left
    End With
    
    With Label_Confirmar
        .top = Barra_Text_Nova_Senha.top + Barra_Text_Nova_Senha.ScaleHeight + 6
        .left = Shape_Erro2.left
    End With
    
    With Barra_Text_Confirmar
        .top = Label_Confirmar.top + Label_Confirmar.Height + 3
        .left = Shape_Erro2.left
    End With
    
    With Check_Visualizar
        .top = Barra_Text_Confirmar.top + Barra_Text_Confirmar.ScaleHeight + 6
        .left = Shape_Erro2.left
    End With
    
    With Pic_Visualizar
        .top = Check_Visualizar.top
        .left = Check_Visualizar.left
    End With
    
    With Label_Esqueceu
        .top = Check_Visualizar.top + Check_Visualizar.Height + 16
        .left = Shape_Erro2.left
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Label_Senha_Click()
    'Ver frame senha
    Frame_Senha.Visible = True
    Frame_Perfil.Visible = False
    Separador_Perfil.Height = Form_Skin.Separador_Mini_Normal.Height
    Separador_Senha.Height = Form_Skin.Separador_Mini_Normal.Height + 1
    Text_Senha_Actual.SetFocus
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Perfil
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Perfil
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Perfil
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Perfil
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Perfil
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Perfil
End Sub

Private Sub Opcao_Feminino_Click()
    'Activar opcao
    Pic_Masculino.Picture = Form_Skin.Opcao_Normal.Picture
    Pic_Feminino.Picture = Form_Skin.Opcao_Over.Picture
    Image_Foto.Picture = Form_Skin.Foto_Feminino.Picture
End Sub

Private Sub Opcao_Masculino_Click()
    'Activar opcao
    Pic_Masculino.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Feminino.Picture = Form_Skin.Opcao_Normal.Picture
    Image_Foto.Picture = Form_Skin.Foto_Masculino.Picture
End Sub

Private Sub Opcao_Masculino_GotFocus()
    'Activar opcao
    Pic_Masculino.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Feminino.Picture = Form_Skin.Opcao_Normal.Picture
    Image_Foto.Picture = Form_Skin.Foto_Masculino.Picture
End Sub

Private Sub Pic_Feminino_Click()
    'Activar opcao
    Opcao_Feminino.Value = True
    Pic_Masculino.Picture = Form_Skin.Opcao_Normal.Picture
    Pic_Feminino.Picture = Form_Skin.Opcao_Over.Picture
    Image_Foto.Picture = Form_Skin.Foto_Feminino.Picture
End Sub

Private Sub Pic_Visualizar_Click()
    'Des/Activar a opcção
    If Check_Visualizar.Value = 0 Then
        Check_Visualizar.Value = 1
        Pic_Visualizar.Picture = Form_Skin.Check_Over.Picture
        Text_Senha_Actual.PasswordChar = ""
        Text_Nova_Senha.PasswordChar = ""
        Text_Confirmar.PasswordChar = ""
        
    Else
        Check_Visualizar.Value = 0
        Pic_Visualizar.Picture = Form_Skin.Check_Normal.Picture
        Text_Senha_Actual.PasswordChar = "*"
        Text_Nova_Senha.PasswordChar = "*"
        Text_Confirmar.PasswordChar = "*"
    End If
End Sub

Private Sub Pic_Masculino_Click()
    'Activar opcao
    Opcao_Masculino.Value = True
    Pic_Masculino.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Feminino.Picture = Form_Skin.Opcao_Normal.Picture
    Image_Foto.Picture = Form_Skin.Foto_Masculino.Picture
End Sub

Private Sub Separador_Perfil_Click()
    'Atalho para
    Label_Perfil_Click
End Sub

Private Sub Separador_Senha_Click()
    'Atalho para
    Label_Senha_Click
End Sub

Private Sub Text_Dia_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Dia.Visible = True
End Sub

Private Sub Text_Dia_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Dia.Visible = False
End Sub

Private Sub Text_Email_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Email.Visible = True
End Sub

Private Sub Text_Email_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Email.Visible = False
End Sub

Private Sub Text_Mes_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Mes.Visible = True
End Sub

Private Sub Text_Mes_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Mes.Visible = False
End Sub

Private Sub Text_Ano_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Ano.Visible = True
End Sub

Private Sub Text_Ano_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Ano.Visible = False
End Sub

Private Sub Text_Nome_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Nome.Visible = True
End Sub

Private Sub Text_Nome_LostFocus()
    'Contorno da text box ao perder o focus
    Text_Nome.Text = StrConv(Text_Nome.Text, vbProperCase)
    Contorno_Nome.Visible = False
End Sub

Private Sub Text_Pais_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Pais.Visible = True
End Sub

Private Sub Text_Pais_LostFocus()
    'Contorno da text box ao perder o focus
    Text_Pais.Text = StrConv(Text_Pais.Text, vbProperCase)
    Contorno_Pais.Visible = False
End Sub

Private Sub Text_Nova_Senha_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Nova_Senha.Visible = False
End Sub

Private Sub Text_Nova_Senha_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Nova_Senha.Visible = True
End Sub

Private Sub Text_Confirmar_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Confirmar.Visible = False
End Sub

Private Sub Text_Confirmar_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Confirmar.Visible = True
End Sub

Private Sub Text_Senha_Actual_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Senha_Actual.Visible = False
End Sub

Private Sub Text_Senha_Actual_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Senha_Actual.Visible = True
End Sub
