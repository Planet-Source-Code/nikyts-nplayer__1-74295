VERSION 5.00
Begin VB.Form Form_Recuperar_Conta 
   BackColor       =   &H00313131&
   BorderStyle     =   0  'None
   ClientHeight    =   3585
   ClientLeft      =   13095
   ClientTop       =   1725
   ClientWidth     =   16410
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
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1094
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text_BCC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6720
      TabIndex        =   27
      Top             =   2880
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.TextBox Text_CC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6720
      TabIndex        =   26
      Top             =   2280
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.ListBox Lista_Anexos 
      Appearance      =   0  'Flat
      Height          =   615
      ItemData        =   "Form_Recuperar_Conta.frx":0000
      Left            =   6720
      List            =   "Form_Recuperar_Conta.frx":0007
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.TextBox Text_Email_Conta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6720
      TabIndex        =   17
      Text            =   "nikyts@gmail.com"
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text_Nome_Destinatario 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6720
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text_Email_Destinatario 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   10440
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text_Email_Remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   10440
      TabIndex        =   14
      Text            =   "nikyts@gmail.com"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text_Nome_Remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6720
      TabIndex        =   13
      Text            =   "NPlayer"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text_Senha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   10440
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text_Mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   6720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.TextBox Text_Assunto 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6720
      TabIndex        =   10
      Text            =   "Recuperar dados de acesso"
      Top             =   3480
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   465
      Width           =   6375
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5400
         Top             =   1320
      End
      Begin VB.PictureBox Barra_Text_Email 
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Width           =   5475
         Begin VB.TextBox Text_Email 
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
      Begin NPlayer.NProgressBar ProgressBar1 
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1800
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
      End
      Begin VB.Label Label_Conectando 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting to smtp.gmail.com on port 465 ..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   3900
      End
      Begin VB.Label Label_Erro 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         BackStyle       =   0  'Transparent
         Caption         =   "Indique um endereço de email válido."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Image Image_Erro 
         Enabled         =   0   'False
         Height          =   210
         Left            =   300
         Picture         =   "Form_Recuperar_Conta.frx":0019
         Top             =   240
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label_Email 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   540
      End
      Begin VB.Shape Shape_Erro 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H008080FF&
         Height          =   315
         Left            =   240
         Top             =   180
         Visible         =   0   'False
         Width           =   5475
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
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2760
      Width           =   6135
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
         TabIndex        =   1
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
            TabIndex        =   5
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
      ScaleWidth      =   337
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   5055
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Recuperar dados de acesso"
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
         TabIndex        =   3
         Top             =   120
         Width           =   2685
      End
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   4440
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   32
      Top             =   5040
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anexos"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BCC"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   30
      Top             =   2640
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CC"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   29
      Top             =   2040
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label_Anexos 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   4080
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha de acesso"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10440
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do destinatário"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do remetente"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email do destinatário"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10440
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email do remetente"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10440
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email da conta do gmail"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00212121&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Recuperar_Conta"
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

'Variáveis que vão receber os dados enviados pelo servidor referentes á conta do utilizador
Dim Utilizador, Senha As String

'Variável para o idioma
Dim Idioma_Erro As String
Dim Idioma_Descricao As String
Dim Idioma_Erro_Execucao As String
Dim Idioma_Conectar_Servidor As String
Dim Idioma_Internet_Desligada As String
Dim Idioma_Error_Email_Invalid As String
Dim Idioma_Error_Email_No_Exist As String

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
    If Text_Email.Text = "" Then Exit Sub
    Shape_Erro.Visible = False
    Label_Erro.Visible = False
    Image_Erro.Visible = False
    
    'Verifica se o campo email está no formato correcto
    If Not IsEmail(Text_Email.Text) Then
        Label_Erro.Caption = Idioma_Error_Email_Invalid
        Text_Email.Text = ""
        Shape_Erro.Visible = True
        Label_Erro.Visible = True
        Image_Erro.Visible = True
        Text_Email.SetFocus
        Exit Sub
    End If
    
    Label_Ok.Enabled = False
    Botao_Ok.Enabled = False
    Me.MousePointer = 11
    Label_Conectando.Visible = True
    ProgressBar1.Visible = True
    
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60 'Mensagens: http://www.nikyts.com/suporte/vermensagens.asp
    servidor.Open "GET", "http://www.nikyts.com/nplayer/" & "recuperarsenha.asp?email=" & Text_Email.Text, False
    servidor.send

    'Verifica se existe o email indicado
    If servidor.responseText = "NaoExiste" Then
        Me.MousePointer = 0
        Text_Email.Text = ""
        Label_Erro.Caption = Idioma_Error_Email_No_Exist
        Shape_Erro.Visible = True
        Label_Erro.Visible = True
        Image_Erro.Visible = True
        
        'Reactiva os botões
        Label_Ok.Enabled = True
        Botao_Ok.Enabled = True
        Me.MousePointer = 1
        Text_Email.SetFocus
        Exit Sub
        
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        With Form_Principal
            If servidor.readyState = 4 And servidor.Status = 200 Then ' 4 - deu resposta e 200 validou -> And servidor.responseText = "sucesso"
                Me.MousePointer = 11
                'Vai ler os dados
                Servidor_Dados_da_Conta servidor.responseText
                Text_Nome_Destinatario.Text = Text_Email.Text
                Text_Email_Destinatario.Text = Text_Email.Text
                Text_Assunto.Text = ReadINI("Recover_Account", "Label_Subject", Localizacao_Ficheiro_Lingua)
                Text_Mensagem.Text = ReadINI("Recover_Account", "Info_Send_Access_Data", Localizacao_Ficheiro_Lingua) & vbNewLine _
                    & ReadINI("Recover_Account", "Info_Send_Access_User", Localizacao_Ficheiro_Lingua) & " " & Utilizador & vbNewLine _
                    & ReadINI("Recover_Account", "Info_Send_Access_Password", Localizacao_Ficheiro_Lingua) & " " & Senha & vbNewLine & vbNewLine _
                    & ReadINI("Recover_Account", "Info_Send_Yours_Sincerely", Localizacao_Ficheiro_Lingua) & " Nikyts (Nelson do Carmo)."
                    
                'Enviar os dados de acesso para a conta do utilizador
                If ProgressBar1.Value = 100 Then
                    ProgressBar1.Value = 0
                End If
                Trim_Functions
                Me.MousePointer = 11
                Botao_Ok.Enabled = False
                Label_Ok.Enabled = False
                Label_Conectando.Visible = True
                ProgressBar1.Visible = True
                Timer1.Enabled = True
                
'                servidor.Open "GET", "http://www.nikyts.com/suporte/" & "enviarmensagem.asp?Email=" & Text_Email.Text & "&Assunto=" & App.ProductName & " - " & "Recuperar dados de acesso" & "&Mensagem=" & "O utilizador solicitou a recuperação dos seus dados de acesso.", False
'                servidor.send
'                'Verificar os dados acesso
'                If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
'                    If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
'                        Me.MousePointer = 0
'                        'Informa o utilizador que os dados foram enviados com sucesso
'                        Mensagem_de_Aviso "Information", ReadINI("Message", "Info_Request_Success1", Localizacao_Ficheiro_Lingua) & vbNewLine & ReadINI("Message", "Info_Request_Success2", Localizacao_Ficheiro_Lingua)
'                        Unload Me
'                    End If
'                End If
                Me.MousePointer = 0
            End If
        End With
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
Botao_Ok.Enabled = True
Label_Ok.Enabled = True
Label_Conectando.Visible = False
ProgressBar1.Visible = False
End Sub

Private Sub Servidor_Dados_da_Conta(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados do perfil do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument

    If xml.loadXML(responseText) Then
        Dim dados: Set dados = xml.selectSingleNode("/dados")
        Utilizador = dados.selectSingleNode("utilizador").Text
        Senha = dados.selectSingleNode("senha").Text
    End If
    Set xml = Nothing
End Sub

Private Sub Botao_Ok_GotFocus()
    'Colocar o focus no botao
    Contorno_Ok.Visible = True
End Sub

Private Sub Botao_Ok_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Ok_Click
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
        
    'Alterar cores do progreesbar
    ProgressBar1.backcolor = Form_Skin.Cor_Contorno_Caixas.backcolor
    
    Label_Anexos.Caption = Lista_Anexos.ListCount & " Ficheiro(s) adicionados."
    Label_Conectando.Caption = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Preferencias.Text_Lingua.Text & ".lng"
    
    Label_Titulo.Caption = ReadINI("Recover_Account", "Title", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Recover_Account", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Email.Caption = ReadINI("Recover_Account", "Label_Email", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Recover_Account", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Text_Assunto.Text = ReadINI("Recover_Account", "Label_Subject", Localizacao_Ficheiro_Lingua)
    
    Idioma_Erro = ReadINI("Message", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Message", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Message", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Conectar_Servidor = ReadINI("Message", "Error_Connect", Localizacao_Ficheiro_Lingua)
    Idioma_Internet_Desligada = ReadINI("Message", "Error_Internet", Localizacao_Ficheiro_Lingua)
    Idioma_Error_Email_Invalid = ReadINI("Message", "Error_Email_Invalid", Localizacao_Ficheiro_Lingua)
    Idioma_Error_Email_No_Exist = ReadINI("Message", "Error_Email_No_Exist", Localizacao_Ficheiro_Lingua)
    
    Label_Erro.Caption = Idioma_Error_Email_Invalid
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
        Contorno_Email.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Email.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        'Barra_Text_Email.Picture = .Caixa_de_Texto.Picture
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Email.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Email.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Email.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Email.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Email.ForeColor = .Cor_Letra_Textbox.backcolor
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Ok.Picture = .Pic_Button.Picture
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Conectando.ForeColor = .Cor_Letra_Label_Formulario.backcolor
    End With
End Sub

Private Sub Label_Cancelar_Click()
    'Atalho para
    Botao_Cancelar_Click
End Sub

Private Sub Label_Contacto_Click()
End Sub

Private Sub Label_Ok_Click()
    'Atalho para
    Botao_Ok_Click
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Barra_Text_Email.Width = Form_Skin.Caixa_de_Texto.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Barra_Text_Email.ScaleWidth) + (2 * Barra_Text_Email.left) + 20)
        '.Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Fundo_Frame_Botoes.Height + Frame_Centro.ScaleHeight)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Shape_Erro.left + Shape_Erro.Height + 3 + Label_Email.Height + 3 _
                + Barra_Text_Email.ScaleHeight + Label_Conectando.Height + 6 + ProgressBar1.Height + Fundo_Frame_Botoes.Height + (3 * Shape_Erro.left))
    End With
    
    Ajustar_Formulario Form_Recuperar_Conta, False, False, True, True
    
    Ajustar_Botao Form_Recuperar_Conta, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Ok
        .left = (Frame_Botoes.ScaleWidth - .ScaleWidth) / 2
    End With

    Ajustar_Caixa_Texto Barra_Text_Email, Text_Email, Contorno_Email, False
        
    With Shape_Erro
        .top = .left
        .Width = Barra_Text_Email.ScaleWidth
    End With
    
    With Image_Erro
        .top = (Shape_Erro.top + Shape_Erro.Height) / 2
    End With
    
    With Label_Erro
        .top = Image_Erro.top
    End With
    
    With Label_Email
        .top = Shape_Erro.top + Shape_Erro.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Barra_Text_Email
        .top = Label_Email.top + Label_Email.Height + 3
        .left = Shape_Erro.left
    End With
    
    With Label_Conectando
        .top = Barra_Text_Email.top + Barra_Text_Email.ScaleHeight + 16
        .left = Barra_Text_Email.left
    End With
    
    With ProgressBar1
        .top = Label_Conectando.top + Label_Conectando.Height + 6
        .left = Barra_Text_Email.left
        .Width = Barra_Text_Email.Width
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Recuperar_Conta
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Recuperar_Conta
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Recuperar_Conta
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Recuperar_Conta
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Recuperar_Conta
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Recuperar_Conta
End Sub

Private Sub Text_Email_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Email.Visible = True
End Sub

Private Sub Text_Email_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Email.Visible = False
End Sub

Private Sub Text_Email_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Ok_Click
End Sub

Private Sub Trim_Functions()
    'Setting veriables to trim text of textboxes
    Dim Gmail_Address As String
    Dim Receiver_email As String
    Dim Subject As String
    Dim Sender_name As String
    Dim Receiver_name As String
    Dim CC As String
    Dim Bcc As String
    
    'Trimming all
    Gmail_Address = Trim$(Text_Email_Conta.Text)
    Receiver_email = Trim$(Text_Email_Destinatario.Text)
    Subject = Trim$(Text_Assunto.Text)
    Sender_name = Trim$(Text_Nome_Remetente.Text)
    Receiver_name = Trim$(Text_Nome_Destinatario.Text)
    CC = Trim$(Text_CC.Text)
    Bcc = Trim$(Text_BCC.Text)
    
    'Putting trimed text back to textboxes
    Form_Recuperar_Conta.Text_Email_Conta.Text = Gmail_Address
    Form_Recuperar_Conta.Text_Email_Destinatario.Text = Receiver_email
    Form_Recuperar_Conta.Text_Assunto.Text = Subject
    Form_Recuperar_Conta.Text_Nome_Remetente.Text = Sender_name
    Form_Recuperar_Conta.Text_Nome_Destinatario.Text = Receiver_name
    Form_Recuperar_Conta.Text_CC.Text = CC
    Form_Recuperar_Conta.Text_BCC.Text = Bcc
End Sub

Private Sub Timer1_Timer()
    'Attached files counter veriables
    Dim i As Integer
    Dim Index
    
    'Email Sending veriables
    Dim iMsg
    Dim iConf
    Dim Flds
    Dim Schema
    
    'Setting Progressbar & Sending Status
    ProgressBar1.Value = 20
    Label_Conectando.Caption = ReadINI("Recover_Account", "Label_Connect1", Localizacao_Ficheiro_Lingua) 'Conectando-se ao servidor...
    
    On Error GoTo SendMail_Error:
    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    Set Flds = iConf.Fields
    Schema = "http://schemas.microsoft.com/cdo/configuration/"
    
    'Configurar o serviço de email
    With Flds
        .Item(Schema & "sendusing") = 2
        .Item(Schema & "smtpserver") = "SMTP.Gmail.Com"
        .Item(Schema & "smtpserverport") = "465"
        ProgressBar1.Value = 30
        .Item(Schema & "smtpauthenticate") = 1 'autenticação
        .Item(Schema & "sendusername") = Form_Recuperar_Conta.Text_Email_Conta.Text
        .Item(Schema & "sendpassword") = Form_Recuperar_Conta.Text_Senha.Text
        .Item(Schema & "smtpConnectionTimeout") = 40 'Conexão timeout
        Flds.Item(Schema & "smtpusessl") = 1 'SSL setting
        .Update
    End With
    
    'Show progress of sending
    ProgressBar1.Value = 50
    Label_Conectando.Caption = ReadINI("Recover_Account", "Label_Connect2", Localizacao_Ficheiro_Lingua) 'Enviando o email, por favor aguarde...
    
    'Setting-up email perameters
    With iMsg
       .To = Form_Recuperar_Conta.Text_Nome_Destinatario.Text & "<" & Form_Recuperar_Conta.Text_Email_Destinatario.Text & ">"
       .from = Form_Recuperar_Conta.Text_Nome_Remetente.Text & "<" & Form_Recuperar_Conta.Text_Email_Remetente.Text & ">"
       .CC = Form_Recuperar_Conta.Text_CC.Text
       .Bcc = Form_Recuperar_Conta.Text_BCC.Text
       .Subject = Form_Recuperar_Conta.Text_Assunto.Text
       
        'E-mail Text-body
        ProgressBar1.Value = 60
       .TextBody = Form_Recuperar_Conta.Text_Mensagem.Text
    
Leave_Attachents:
    
    'Send all
    Set .Configuration = iConf
       .send
    
    End With
    
    'Clear veriables if needed
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    Set Schema = Nothing
    
    'Email sent, then progress
    ProgressBar1.Value = 90
    
    Me.MousePointer = 0
    Botao_Ok.Enabled = True
    Label_Ok.Enabled = True
    
    'Email sent with all attachments, then progress bar value
    ProgressBar1.Value = 100
    
    'Show sucess message
    Mensagem_de_Aviso "Information", ReadINI("Message", "Info_Request_Success1", Localizacao_Ficheiro_Lingua) & vbNewLine & ReadINI("Message", "Info_Request_Success2", Localizacao_Ficheiro_Lingua)
    
    'Disable sending timer
    Timer1.Enabled = False
    
    'Call Clear fields function
    Clear_Data
    Unload Me
    
    GoTo End_Sub:
    
SendMail_Error:
    Me.MousePointer = 0
    Botao_Ok.Enabled = True
    Label_Ok.Enabled = True
    Label_Conectando.Visible = False
    ProgressBar1.Visible = False

    Timer1.Enabled = False
    Mensagem_de_Aviso "Information", ReadINI("Message", "Error_Send_Mail1", Localizacao_Ficheiro_Lingua) & vbNewLine & ReadINI("Message", "Error_Send_Mail2", Localizacao_Ficheiro_Lingua)
End_Sub:
End Sub

Private Sub Clear_Data()
    'Clear Main form's data fields
    Form_Recuperar_Conta.Text_Mensagem.Text = ""
    Form_Recuperar_Conta.Text_Nome_Destinatario.Text = ""
    Form_Recuperar_Conta.Text_Email_Destinatario.Text = ""
    Form_Recuperar_Conta.Text_CC.Text = ""
    Form_Recuperar_Conta.Text_BCC.Text = ""
    'Form_Recuperar_Conta.Text_Assunto.Text = ""
    Form_Recuperar_Conta.Label_Anexos.Caption = Form_Recuperar_Conta.Lista_Anexos.ListCount & " Ficheiro(s) adicionados."
End Sub

