Attribute VB_Name = "Module_Geral"
'Api's para mover o formulário
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'Api par obter a cor do pixel selecionado
Public Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Const SRCCOPY = &HCC0020

'Posição x e y
Global iTPPY As Long
Global iTPPX As Long

'API para o procedimento alway's on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Para ver o relat´rio através do browser
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Variável das msgboxs
Public Resposta As String

'Variaveis para verificar a resposta de remover ficheiro da biblioteca/ do meu computador
Public Remover_da_Biblioteca As Boolean

'Variáveis para poder mover o formulário
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'Variável para indicar o caminho do ficheiro das opções do programa
Public Localizacao_Ficheiro_Preferencias As String
Public Localizacao_Ficheiro_Lingua As String
'Public Localizacao_Ficheiro_Skin As String
    
'Variável para verificar se o utilizador está logado
Public Utilizador_Logado As Boolean

'Variável para saber se o utilizador ao iniciar o programa pela 1ª vez se solicitou importar musicas, pois caso contrário não carrega os albuns
'desse directório ao iniciar o form principal
Public importar_media As Boolean

Sub Main()
    'Verificar se já existe algum instância da aplicação
    'If App.PrevInstance = True Then End
    Localizacao_Ficheiro_Preferencias = App.Path & "\Options\Properties.ini"
    'Localizacao_Ficheiro_Skin = App.Path & "\Skins\" & Form_Preferencias.Text_Skin.Text & "\Style.ini"
    
    importar_media = True
    
    On Error Resume Next
    'Verificar se o programa já foi instalado
    Dim Programa_Instalado As String: Programa_Instalado = ReadINI("Settings", "Installed_Program", Localizacao_Ficheiro_Preferencias)
    If Programa_Instalado = "False" Then
        Form_Setup.Show
    Else
        Form_Instalar.Show
    End If
    
'Para testes
'Form_Mini_Player.Show
End Sub

Public Sub Ajustar_ChecBox(Pic_CheckBox As PictureBox, CheckBox As CheckBox)
    'Procedimento para ajustar as checboxs do formulário
    With Pic_CheckBox
        .Height = CheckBox.Height
        .Width = CheckBox.Height
    End With
End Sub

Public Sub Ajustar_Option(Pic_Option As PictureBox)
    'Procedimento para ajustar as checboxs do formulário
    With Pic_Option
        .Height = Form_Skin.Opcao_Normal.Height
        .Width = Form_Skin.Opcao_Normal.Width
    End With
End Sub

'Colocar o formulário por cima dos outros
Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    If OnTop = -1 Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    End If
End Sub

Public Sub Mensagem_de_Aviso(Aviso As String, Mensagem As String)
    'Procedimento para mostrar uma mensagem de aviso
    With Form_Mensagem
        If Aviso = "Information" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Info.Picture
            .Botao_Ok.Visible = True
        
        ElseIf Aviso = "Error" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
            .Botao_Ok.Visible = True
        
        ElseIf Aviso = "Question" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Quest.Picture
            .Botao_Sim.Visible = True
            .Botao_Nao.Visible = True
            
        ElseIf Aviso = "Invitation" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Invitation.Picture
            .Botao_Ok.Visible = True
            
        ElseIf Aviso = "Hyperlink" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Link.Picture
            .Botao_Ok.Visible = True
            .Barra_Text_Servidor.Visible = True
        End If
        
        .Label_Mensagem.Caption = Mensagem
        .Show vbModal
    End With
End Sub

Public Function ArquivoExiste(ByVal Caminho As String, Optional ByVal SomenteDiretorio As Boolean = False) As Boolean
    'Função para verificar se a pasta existe
    On Error Resume Next
    If SomenteDiretorio Then
        ArquivoExiste = GetAttr(Mid(Caminho, 1, InStrRev(Caminho, ""))) And vbDirectory
    Else
        ArquivoExiste = GetAttr(Caminho)
    End If
    On Error GoTo 0
End Function

Public Sub Ajustar_Formulario(Form As Form, Icon_Visivel As Boolean, Form_Ajustavel As Boolean, Frame_Centro_Visivel As Boolean, _
                                Frame_Botoes_Visivel As Boolean)
    'Procedimento para ajustar os componentes dos formulários
    If Form.WindowState = 1 Then Exit Sub
    With Form.Shape_Contorno
        .Height = Form.ScaleHeight
        .Top = 0
        .Width = Form.ScaleWidth
        .Left = 0
    End With
    
    With Form.Barra_ControlBox
        .Height = Form_Skin.Fundo_Barra_ControlBox.Height
        .Top = 0 ' 1
        .Width = Form.ScaleWidth '- 2
        .Left = 0 ' 1
    End With
    
    With Form.Fundo_Barra_ControlBox
        .Stretch = True
        .Top = 0
        .Width = Form.Barra_ControlBox.ScaleWidth
        .Left = 0
    End With

    With Form.Label_Titulo
        .Top = (Form.Barra_ControlBox.ScaleHeight - .Height) / 2
        If Icon_Visivel = False Then .Left = 10 Else: .Left = 26
    End With
    
    'Botões do controlbox
    Dim Ajustar_Botoes As String
    Ajustar_Botoes = "False" 'ReadINI("Dimensions", "Adjust_Button_ControlBox", Localizacao_Ficheiro_Skin)
    
    With Form.Botao_Fechar
        .Height = Form_Skin.Botao_Fechar.Height
        If Ajustar_Botoes = "False" Then
            .Top = (Form.Barra_ControlBox.ScaleHeight - .Height) / 2
        Else
            .Top = 0
        End If
        .Width = Form_Skin.Botao_Fechar.Width
        .Left = Form.Barra_ControlBox.Width - .Width - 6
    End With
    
    If Form_Ajustavel = True Then
        With Form.Botao_Maximizar
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Fechar.Left - .Width - 8
            Else
                .Left = Form.Botao_Fechar.Left - .Width
            End If
        End With
        
        With Form.Botao_Restaurar
            .Top = Form.Botao_Fechar.Top
            .Left = Form.Botao_Maximizar.Left
        End With
        
        With Form.Botao_Minimizar
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Maximizar.Left - .Width - 8
            Else
                .Left = Form.Botao_Maximizar.Left - .Width
            End If
        End With
    End If
    
    If Frame_Botoes_Visivel = True Then
        With Form.Frame_Botoes
            .Height = Form_Skin.Fundo_Frame_Botoes.Height
            .Top = Form.ScaleHeight - .ScaleHeight - 1
            .Width = Form.ScaleWidth - 2
            .Left = 1
        End With
        
        With Form.Fundo_Frame_Botoes
            .Stretch = True
            .Top = 0
            .Width = Form.Frame_Botoes.ScaleWidth
            .Left = 0
        End With
    End If
    
    If Frame_Centro_Visivel = True Then
        With Form.Frame_Centro
            .Height = Form.ScaleHeight - Form.Barra_ControlBox.ScaleHeight - Form.Frame_Botoes.ScaleHeight - 2
            .Top = Form.Barra_ControlBox.Top + Form.Barra_ControlBox.ScaleHeight
            .Width = Form.ScaleWidth - 20
            .Left = 10
        End With
        
        With Form.Shape_Centro
            .Top = 0
            .Height = Form.Frame_Centro.Height
            .Left = 0
            .Width = Form.Frame_Centro.Width
            .Visible = True
        End With
    End If
End Sub

Public Sub Ajustar_Botao(Form As Form, Nome_Botao As PictureBox, Nome_Label As Label, Botao_Esta_Na_Frame_Botoes As Boolean, Nome_Shape As Shape)
    'Procedimento para ajustar os botoes e respectivas labels
    With Nome_Botao
        .Height = Form_Skin.Botao_Form.Height
        .Width = Form_Skin.Botao_Form.Width
        If Botao_Esta_Na_Frame_Botoes = True Then
            .Top = (Form.Frame_Botoes.ScaleHeight - .ScaleHeight) / 2
        End If
    End With
    
    With Nome_Label
        .AutoSize = False
        .Alignment = vbCenter
        .Top = (Nome_Botao.ScaleHeight - .Height) / 2
        .Width = Nome_Botao.ScaleWidth
        .Left = 0
    End With
    
    With Nome_Shape
        .Height = Nome_Botao.ScaleHeight
        .Top = 0
        .Width = Nome_Botao.ScaleWidth
        .Left = 0
    End With
End Sub

Public Sub Ajustar_Caixa_Texto(Barra_TextBox As PictureBox, Nome_TextBox As TextBox, Nome_Shape As Shape, Caixa_Observacoes As Boolean)
    'Procedimento para ajustar as caixas de texto
    If Caixa_Observacoes = False Then
        With Barra_TextBox
            .Height = Form_Skin.Caixa_de_Texto.Height
            .Width = Form_Skin.Caixa_de_Texto.Width
        End With
    Else
        With Barra_TextBox
            .Height = Form_Skin.Caixa_de_Observacoes.Height
            .Width = Form_Skin.Caixa_de_Observacoes.Width
        End With
    End If
    
    With Nome_TextBox
        .Height = Barra_TextBox.ScaleHeight - 8 - 8
        .Top = (Barra_TextBox.ScaleHeight - .Height) / 2
        .Width = Barra_TextBox.ScaleWidth - 8 - 8
        .Left = 8
    End With
    
    With Nome_Shape
        .Height = Barra_TextBox.ScaleHeight
        .Top = 0
        .Width = Barra_TextBox.ScaleWidth
        .Left = 0
    End With
End Sub

Public Sub Ajustar_Caixa_Texto_Mini(Caixa_Texto As PictureBox, Nome_TextBox As TextBox, Nome_Shape As Shape)
    'Procedimento para ajustar as caixas de texto
    With Caixa_Texto
        .Height = Form_Skin.Caixa_de_Texto_Mini.Height
        .Width = Form_Skin.Caixa_de_Texto_Mini.Width
    End With
    
    With Nome_TextBox
        .Height = Caixa_Texto.ScaleHeight - 8 - 8
        .Top = (Caixa_Texto.ScaleHeight - .Height) / 2
        .Width = Caixa_Texto.ScaleWidth - 8 - 8
        .Left = 8
    End With
    
    With Nome_Shape
        .Height = Caixa_Texto.ScaleHeight
        .Top = 0
        .Width = Caixa_Texto.ScaleWidth
        .Left = 0
    End With
End Sub

Public Sub Ajustar_Caixa_Texto_Media(Caixa_Texto As PictureBox, Nome_TextBox As TextBox, Nome_Shape As Shape)
    'Procedimento para ajustar as caixas de texto
    With Caixa_Texto
        .Height = Form_Skin.TextBox_Intermediate.Height
        .Width = Form_Skin.TextBox_Intermediate.Width
    End With
    
    With Nome_TextBox
        .Height = Caixa_Texto.ScaleHeight - 8 - 8
        .Top = (Caixa_Texto.ScaleHeight - .Height) / 2
        .Width = Caixa_Texto.ScaleWidth - 8 - 8
        .Left = 8
    End With
    
    With Nome_Shape
        .Height = Caixa_Texto.ScaleHeight
        .Top = 0
        .Width = Caixa_Texto.ScaleWidth
        .Left = 0
    End With
End Sub

Public Sub Mover_Formulario(Form As Form)
    'Procedimento para poder mover o formulário
    If Form.WindowState = 0 Then
        Dim iDX As Long, iDY As Long
        Dim POINT As POINTAPI
        If Not bMoveFrom Then Exit Sub
        GetCursorPos POINT
        iDX& = (POINT.X - LastPoint.X) * iTPPX&
        iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
        LastPoint.X = POINT.X
        LastPoint.Y = POINT.Y
        Form.Move Form.Left + iDX&, Form.Top + iDY&
    End If
End Sub

Public Sub Capturar_Posicao_Formulario(Form As Form)
    'Capturar a posição de x e y
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    bMoveFrom = True
End Sub

Public Sub Largar_Formulario(Form As Form)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Public Sub Ajustar_Formulario_com_Menu(Form As Form, Icon_Visivel As Boolean, Form_Ajustavel As Boolean, Frame_Centro_Visivel As Boolean, _
                                Frame_Botoes_Visivel As Boolean)
    'Procedimento para ajustar os componentes dos formulários
    If Form.WindowState = 1 Then Exit Sub
    With Form.Shape_Contorno
        .Height = Form.ScaleHeight
        .Top = 0
        .Width = Form.ScaleWidth
        .Left = 0
    End With
    
    With Form.Barra_ControlBox
        .Height = Form_Skin.Fundo_Barra_ControlBox.Height
        .Top = 0 ' 1
        .Width = Form.ScaleWidth '- 2
        .Left = 0 ' 1
    End With
    
    With Form.Fundo_Barra_ControlBox
        .Stretch = True
        .Top = 0
        .Width = Form.Barra_ControlBox.ScaleWidth
        .Left = 0
    End With

    With Form.Label_Titulo
        .Top = (Form.Barra_ControlBox.ScaleHeight - .Height) / 2
        If Icon_Visivel = False Then .Left = 10 Else: .Left = 26
    End With
    
    'Botões do controlbox
    Dim Ajustar_Botoes As String
    Ajustar_Botoes = "False" 'ReadINI("Dimensions", "Adjust_Button_ControlBox", Localizacao_Ficheiro_Skin)
    
    With Form.Botao_Fechar
        .Height = Form_Skin.Botao_Fechar.Height
        If Ajustar_Botoes = "False" Then
            .Top = (Form.Barra_ControlBox.ScaleHeight - .Height) / 2
        Else
            .Top = 0
        End If
        .Width = Form_Skin.Botao_Fechar.Width
        .Left = Form.Barra_ControlBox.Width - .Width - 6
    End With
    
    If Form_Ajustavel = True Then
        With Form.Botao_Maximizar
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Fechar.Left - .Width - 8
            Else
                .Left = Form.Botao_Fechar.Left - .Width
            End If
        End With
        
        With Form.Botao_Restaurar
            .Top = Form.Botao_Fechar.Top
            .Left = Form.Botao_Maximizar.Left
        End With
        
        With Form.Botao_Minimizar
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Maximizar.Left - .Width - 8
            Else
                .Left = Form.Botao_Maximizar.Left - .Width
            End If
        End With
        
        With Form.Botao_Tray
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Minimizar.Left - .Width - 8
            Else
                .Left = Form.Botao_Minimizar.Left - .Width
            End If
        End With
    End If
    
    If Frame_Botoes_Visivel = True Then
        With Form.Frame_Botoes
            .Height = Form_Skin.Fundo_Frame_Botoes.Height
            .Top = Form.ScaleHeight - .ScaleHeight - 1
            .Width = Form.ScaleWidth - 2
            .Left = 1
        End With
        
        With Form.Fundo_Frame_Botoes
            .Stretch = True
            .Top = 0
            .Width = Form.Frame_Botoes.ScaleWidth
            .Left = 0
        End With
    End If
    
    If Frame_Centro_Visivel = True Then
        With Form.Frame_Centro
            .Height = Form.ScaleHeight - Form.Barra_ControlBox.ScaleHeight - Form.Frame_Botoes.ScaleHeight - 2 - Form_Skin.Bar_Menu.Height
            .Top = Form.Barra_ControlBox.Top + Form.Barra_ControlBox.ScaleHeight + Form_Skin.Bar_Menu.Height
            .Width = Form.ScaleWidth - 20
            .Left = 10
        End With
        
        With Form.Shape_Centro
            .Top = 0
            .Height = Form.Frame_Centro.Height
            .Left = 0
            .Width = Form.Frame_Centro.Width
            .Visible = False
        End With
    End If
End Sub

Public Sub Aplicar_Novo_Skin_do_Programa(nome_do_skin As String)
    'Procedimento para aplicar todos o skins dos formulários
    With Form_Preferencias
        .Text_Skin.Text = nome_do_skin
        .Salvar_Valores
    End With
    
    Form_Actualizar_Biblioteca.Carregar_Skin
    Form_AddOns.Carregar_Skin
    Form_Adicionar.Carregar_Skin
    Form_Atributos.Carregar_Skin
    Form_Criar.Carregar_Skin
    Form_Download.Carregar_Skin
    Form_Legendas.Carregar_Skin
    Form_Lista.Carregar_Skin
    Form_Login.Carregar_Skin
    Form_Mensagem.Carregar_Skin
    Form_Mini_Player.Carregar_Skin
    Form_Perfil.Carregar_Skin
    Form_PopUp.Carregar_Skin
    Form_Preferencias.Carregar_Skin
    Form_Principal.Carregar_Skin
    Form_Recuperar_Conta.Carregar_Skin
    Form_Reportar_Erro.Carregar_Skin
    Form_Sobre.Carregar_Skin
    Form_Tag.Carregar_Skin
    Form_Wmp.Carregar_Skin
End Sub

Public Sub Aplicar_Idioma_do_Programa()
    'Procedimento para aplicar todos o skins dos formulários
    Form_Actualizar_Biblioteca.Carregar_Idioma
    Form_AddOns.Carregar_Idioma
    Form_Adicionar.Carregar_Idioma
    Form_Atributos.Carregar_Idioma
    Form_Criar.Carregar_Idioma
    Form_Download.Carregar_Idioma
    Form_Legendas.Carregar_Idioma
    Form_Lista.Carregar_Idioma
    Form_Login.Carregar_Idioma
    Form_Mensagem.Carregar_Idioma
    Form_Mini_Player.Carregar_Idioma
    Form_Perfil.Carregar_Idioma
    Form_PopUp.Carregar_Idioma
    Form_Preferencias.Carregar_Idioma
    Form_Principal.Carregar_Idioma
    Form_Recuperar_Conta.Carregar_Idioma
    Form_Reportar_Erro.Carregar_Idioma
    Form_Sobre.Carregar_Idioma
    Form_Tag.Carregar_Idioma
    Form_Wmp.Carregar_Idioma
    
    With Form_Principal
        .Ajustar_Menus
        .Ajustar_Objectos_Na_Horizontal
        
        .Grelha_Musica.TextMatrix(0, 0) = "Dir"
        .Grelha_Musica.TextMatrix(0, 1) = .Idioma_Grid_Music_Col_1
        .Grelha_Musica.TextMatrix(0, 2) = .Idioma_Grid_Music_Col_2
        .Grelha_Musica.TextMatrix(0, 3) = .Idioma_Grid_Music_Col_3
        .Grelha_Musica.TextMatrix(0, 4) = .Idioma_Grid_Music_Col_4
        .Grelha_Musica.TextMatrix(0, 5) = .Idioma_Grid_Music_Col_5
        .Grelha_Musica.TextMatrix(0, 6) = .Idioma_Grid_Music_Col_6
        .Grelha_Musica.TextMatrix(0, 7) = .Idioma_Grid_Music_Col_7
        .Grelha_Musica.TextMatrix(0, 8) = .Idioma_Grid_Music_Col_8
        .Grelha_Musica.TextMatrix(0, 9) = "Id"
        
        .Grelha_Filmes.TextMatrix(0, 0) = "Dir"
        .Grelha_Filmes.TextMatrix(0, 1) = .Idioma_Grid_Movies_Col_1
        .Grelha_Filmes.TextMatrix(0, 2) = .Idioma_Grid_Movies_Col_2
        .Grelha_Filmes.TextMatrix(0, 3) = .Idioma_Grid_Movies_Col_3
        .Grelha_Filmes.TextMatrix(0, 4) = .Idioma_Grid_Movies_Col_4
        .Grelha_Filmes.TextMatrix(0, 5) = .Idioma_Grid_Movies_Col_5
        .Grelha_Filmes.TextMatrix(0, 6) = .Idioma_Grid_Movies_Col_6
        .Grelha_Filmes.TextMatrix(0, 7) = "Id"
                
        .Grelha_Radio.TextMatrix(0, 1) = Idioma_Grid_Radio_Col_1
        
        .Grelha_Listas.TextMatrix(0, 0) = "Dir"
        .Grelha_Listas.TextMatrix(0, 1) = .Idioma_Grid_Music_Col_1
        .Grelha_Listas.TextMatrix(0, 2) = .Idioma_Grid_Music_Col_2
        .Grelha_Listas.TextMatrix(0, 3) = .Idioma_Grid_Music_Col_3
        .Grelha_Listas.TextMatrix(0, 4) = .Idioma_Grid_Music_Col_4
        .Grelha_Listas.TextMatrix(0, 5) = .Idioma_Grid_Music_Col_5
        .Grelha_Listas.TextMatrix(0, 6) = .Idioma_Grid_Music_Col_6
        .Grelha_Listas.TextMatrix(0, 7) = .Idioma_Grid_Music_Col_7
        .Grelha_Listas.TextMatrix(0, 8) = .Idioma_Grid_Music_Col_8
        .Grelha_Listas.TextMatrix(0, 9) = "Id"
        
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 0) = "Dir"
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 1) = .Idioma_Grid_Music_Col_1
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 2) = .Idioma_Grid_Music_Col_2
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 3) = .Idioma_Grid_Music_Col_3
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 4) = .Idioma_Grid_Music_Col_4
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 5) = .Idioma_Grid_Music_Col_5
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 6) = .Idioma_Grid_Music_Col_6
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 7) = .Idioma_Grid_Music_Col_7
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 8) = .Idioma_Grid_Music_Col_8
        .Grelha_Lista_Em_Reproducao.TextMatrix(0, 9) = "Id"
        
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 0) = "Dir"
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 1) = .Idioma_Grid_Music_Col_1
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 2) = .Idioma_Grid_Music_Col_2
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 3) = .Idioma_Grid_Music_Col_3
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 4) = .Idioma_Grid_Music_Col_4
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 5) = .Idioma_Grid_Music_Col_5
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 6) = .Idioma_Grid_Music_Col_6
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 7) = .Idioma_Grid_Music_Col_7
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 8) = .Idioma_Grid_Music_Col_8
        Form_Lista.Grelha_Lista_Em_Reproducao.TextMatrix(0, 9) = "Id"
        
        .Grelha_Loja.TextMatrix(0, 0) = "Hyperlink"
        .Grelha_Loja.TextMatrix(0, 1) = .Idioma_Grid_Loja_Col_1
        .Grelha_Loja.TextMatrix(0, 2) = .Idioma_Grid_Loja_Col_2
        .Grelha_Loja.TextMatrix(0, 3) = .Idioma_Grid_Loja_Col_3
        .Grelha_Loja.TextMatrix(0, 4) = "ID"
        
        .Grelha_Minha_Musica.TextMatrix(0, 0) = "Hyperlink"
        .Grelha_Minha_Musica.TextMatrix(0, 1) = .Idioma_Grid_Loja_Col_1
        .Grelha_Minha_Musica.TextMatrix(0, 2) = .Idioma_Grid_Loja_Col_2
        .Grelha_Minha_Musica.TextMatrix(0, 3) = .Idioma_Grid_Loja_Col_3
        .Grelha_Minha_Musica.TextMatrix(0, 4) = "ID"
        
        .Verificar_Contador
        
        .Menu_Ficheiro(0).Caption = .List_Menu(0).List(0)
        .Menu_Ficheiro(1).Caption = .List_Menu(0).List(1)
        .Menu_Ficheiro(3).Caption = .List_Menu(0).List(3)
        .Menu_Ficheiro(4).Caption = .List_Menu(0).List(4)
        .Menu_Ficheiro(5).Caption = .List_Menu(0).List(5)
        .Menu_Ficheiro(7).Caption = .List_Menu(0).List(7)
        .Menu_Ficheiro(9).Caption = .List_Menu(0).List(9)
        
        .Menu_Editar(0).Caption = .List_Menu(1).List(0)
        .Menu_Editar(1).Caption = .List_Menu(1).List(1)
        .Menu_Editar(3).Caption = .List_Menu(1).List(3)
        .Menu_Editar(5).Caption = .List_Menu(1).List(5)
        
        .Menu_Ver(0).Caption = .List_Menu(2).List(0)
        .Menu_Ver(2).Caption = .List_Menu(2).List(2)
        .Menu_Ver(3).Caption = .List_Menu(2).List(3)
        .Menu_Ver(5).Caption = .List_Menu(2).List(5)
        .Menu_Ver(6).Caption = .List_Menu(2).List(6)
        .Menu_Ver(7).Caption = .List_Menu(2).List(7)
        .Menu_Ver(9).Caption = .List_Menu(2).List(9)
        
        .Menu_Controlos(0).Caption = .List_Menu(3).List(0)
        .Menu_Controlos(1).Caption = .List_Menu(3).List(1)
        .Menu_Controlos(2).Caption = .List_Menu(3).List(2)
        .Menu_Controlos(4).Caption = .List_Menu(3).List(4)
        
        .Menu_Ferramentas(0).Caption = .List_Menu(4).List(0)
        .Menu_Ferramentas(1).Caption = .List_Menu(4).List(1)
        .Menu_Ferramentas(3).Caption = .List_Menu(4).List(3)
        .Menu_Ferramentas(5).Caption = .List_Menu(4).List(5)
        
        .Menu_Ajuda(0).Caption = .List_Menu(5).List(0)
        .Menu_Ajuda(1).Caption = .List_Menu(5).List(1)
        .Menu_Ajuda(3).Caption = .List_Menu(5).List(3)
        .Menu_Ajuda(5).Caption = .List_Menu(5).List(5)
    End With
End Sub

Public Function DataArq(ByVal sArq As String) As String
    'Função para verificar a data de criação dos ficheiros
    If Dir$(sArq) <> "" Then
        DataArq = FileDateTime(sArq)
    Else
        DataArq = "ERRO" 'Não foi possivel identificar a data de criação do ficheiro
    End If
End Function
