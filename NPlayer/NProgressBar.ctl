VERSION 5.00
Begin VB.UserControl NProgressBar 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   ScaleHeight     =   17
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   Begin VB.PictureBox Fundo_Barra_Progresso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.PictureBox Barra_Progresso 
         Appearance      =   0  'Flat
         BackColor       =   &H00CBB534&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   1
         Top             =   -120
         Width           =   15
      End
   End
   Begin VB.Label Label_Percentagem 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "NProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Declaração das variáveis utilizadas pelo control
Public Max As Long
Attribute Max.VB_VarProcData = "PropertyPage1"
Private mvarvalue As Long
Private mvarpercent As String
Private mvarbackcolor As String

Public Property Let backcolor(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.percent = 5
    mvarbackcolor = vData
    Barra_Progresso.backcolor = vData
End Property

Public Property Get backcolor() As String
Attribute backcolor.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.percent
    backcolor = mvarbackcolor
End Property
Public Property Let percent(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.percent = 5
    mvarpercent = vData
End Property

Public Property Get percent() As String
Attribute percent.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.percent
    percent = mvarpercent
End Property

Private Sub Timer3_Timer()
    'Mostar na label a percentagem actual do progressBar
    Label_Percentagem = myval
End Sub

Private Sub Fundo_Barra_Progresso_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Mostar a percentagem do progressbar ao passar com o rato
    Fundo_Barra_Progresso.ToolTipText = percent
End Sub

Private Sub Barra_Progresso_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Mostar a percentagem do progressbar ao passar com o rato
    Barra_Progresso.ToolTipText = percent
End Sub

Private Sub UserControl_Initialize()
    'start with a 0 sized picture to represnt 0%
    Barra_Progresso.Width = 0
    Value = 0
    ' set the progressbar max value like in the common controls version
    Max = 100
    percent = ("00.00")
    'value 1
End Sub

Public Function bcolor(color As Long)
    'Carregar o progress com a cor selecionada
    On Error Resume Next
    Barra_Progresso.backcolor = color
End Function

Public Property Let Value(ByVal vData As Long)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.value = 5
    On Error Resume Next
        
     ''set the percentage using a simple math equasion
     
    percent = Format(Value / (Max / 100), "00.00")
    
    '' check if value is higher than the max value so as not to go over 100%
    If vData < Max Then
    ''if its lower than 100% set our value
     mvarvalue = vData
    Else
    '' our value is higher than it should be so setting it to 100%
    mvarvalue = Max
    percent = Format(100, "00.00")
    End If
    
    If vData > 0 Then
    ''set the picture width to the level of the percent for visual purpose
    Barra_Progresso.Width = (Fundo_Barra_Progresso.ScaleWidth / 100) * Value / (Max / 100)
    Caption = (Fundo_Barra_Progresso.ScaleWidth / 100) * Value / (Max / 100)
    If Value < Max Then
    percent = Format(Value / (Max / 100), "00.00")
    Else
    percent = Format(100, "00.00")
    
    End If
    Else
    Exit Property
    End If
err:
End Property

Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.value
    Value = mvarvalue
End Property

Private Sub UserControl_Paint()
    'Largura inicial do progressbar
    Barra_Progresso.Width = 0
End Sub

Private Sub UserControl_Resize()
    'Desenhando o control, ajustando os objectos
    On Error Resume Next
    With Fundo_Barra_Progresso
        .Height = UserControl.ScaleHeight - 2
        .top = 1
        .Width = UserControl.ScaleWidth - 2
        .left = 1
    End With
    
    With Barra_Progresso
        .Height = Fundo_Barra_Progresso.ScaleHeight
        .top = 0
        .Width = Fundo_Barra_Progresso.ScaleWidth
        .left = 0
    End With
End Sub
