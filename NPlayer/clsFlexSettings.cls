VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlexSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFlexSettings
' DateTime  : 4-7-2004 00:10
' Author    : Flyguy
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Private m_FG As MSFlexGrid  ' The FlexGrid
Private m_cINI As clsINI    ' The INI file in which we store our settings

Private Type tpCELL
  Alignment As String
  BackColor As String
  ForeColor As String
  FontBold As String
  FontItalic As String
  FontSize As String
  FontName As String
  TextStyle As String
End Type

Public Property Set FlexGrid(theGrid As MSFlexGrid)
  Set m_FG = theGrid
End Property

Public Property Get FlexGrid() As MSFlexGrid
  Set FlexGrid = m_FG
End Property

Public Function LoadSettings(sFileName As String, Optional bLoadGeneral As Boolean, Optional bLoadLayout As Boolean, Optional bLoadData As Boolean, Optional bLoadCellSettings As Boolean) As Boolean
  
10      On Error GoTo LoadSettings_Error

20      If Not Exists(sFileName) Then Exit Function
30      If m_FG Is Nothing Then Exit Function
  
40      Set m_cINI = New clsINI
50      m_cINI.Load sFileName
  
60      m_FG.Redraw = False
  
        ' The general grid settings
70      If bLoadGeneral Then _
          If m_cINI.SectionExists("General") Then LoadGeneral
  
        ' Layout
80      If bLoadLayout Then _
          If m_cINI.SectionExists("Layout") Then LoadLayout
  
        ' Cell definitions
90      If bLoadCellSettings Then _
          If m_cINI.SectionExists("CellLayout") Then LoadCellLayout
  
        ' The actual data
100     If bLoadData Then _
          If m_cINI.SectionExists("Data") Then LoadData
  
110     m_FG.Redraw = True

120     On Error GoTo 0
130     Exit Function

LoadSettings_Error:

140     'MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbLf & "On line: " & Erl & vbLf & "in procedure LoadSettings of Class Module clsFlexSettings", vbCritical
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = "Ocorreu um erro ao carregar os dados da biblioteca."
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
End Function

Public Function SaveSettings(sFileName As String, Optional bSaveGeneral As Boolean, Optional bSaveLayout As Boolean, Optional bSaveData As Boolean, Optional bSaveCellSettings As Boolean) As Boolean
        Dim lRow As Long, lCol As Long, lRowSel As Long, lColSel As Long
  
10      On Error Resume Next 'On Error GoTo SaveSettings_Error

20      If m_FG Is Nothing Then Exit Function
30      If Len(sFileName) = 0 Then Exit Function
  
40      Set m_cINI = New clsINI
50      m_cINI.FileName = sFileName
  
60      With m_FG
70        lRow = .Row
80        lRowSel = .RowSel
90        lCol = .Col
100       lColSel = .ColSel
110       .Redraw = False
120     End With
  
130     If bSaveGeneral Then SaveGeneral
140     If bSaveLayout Then SaveLayout
150     If bSaveCellSettings Then SaveCellLayout
160     If bSaveData Then SaveData
  
170     With m_FG
180       .Row = lRow
190       .Col = lCol
200       .RowSel = lRowSel
210       .ColSel = lColSel
220       .Redraw = True
230     End With
  
240     m_cINI.Save
  
250     SaveSettings = True

260     On Error GoTo 0
270     Exit Function

SaveSettings_Error:

280
''    With Form_Mensagem
''        .Label_Titulo.Caption = "Erro"
''        .Pic_Mensagem.Picture = Form_Skin.Icon_info.Picture
''        .Label_Mensagem.Caption = "Nenhum dados a ser guardado."
''        .Botao_Ok.Visible = True
''        .Show vbModal
''    End With
End Function

Private Function Exists(FileName As String) As Boolean
  Exists = Len(Dir$(FileName)) > 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : LoadGeneral
' DateTime  : 4-7-2004 00:00
' Author    : Flyguy
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LoadGeneral()
        Dim cSection As clsSection
    
10      On Error GoTo LoadGeneral_Error

20      Set cSection = m_cINI.Section("General")
  
30      With m_FG
40        .AllowBigSelection = CBool(cSection.Value("AllowBigSelection", CBool(.AllowBigSelection)))
50        .AllowUserResizing = CInt(cSection.Value("AllowUserResizing", CInt(.AllowUserResizing)))
60        .Appearance = CInt(cSection.Value("Appearance", CInt(.Appearance)))
70        .BackColor = CLng(cSection.Value("BackColor", CLng(.BackColor)))
80        .BackColorBkg = CLng(cSection.Value("BackColorBkg", CLng(.BackColorBkg)))
90        .BackColorFixed = CLng(cSection.Value("BackColorFixed", CLng(.BackColorFixed)))
100       .BackColorSel = CLng(cSection.Value("BackColorSel", CLng(.BackColorSel)))
110       .BorderStyle = CInt(cSection.Value("BorderStyle", CInt(.BorderStyle)))
120       .CausesValidation = CBool(cSection.Value("CausesValidation", CBool(.CausesValidation)))
130       .DragMode = CInt(cSection.Value("DragMode", CInt(.DragMode)))
140       .Enabled = CBool(cSection.Value("Enabled", CBool(.Enabled)))
150       .FillStyle = CInt(cSection.Value("FillStyle", CInt(.FillStyle)))
160       .FocusRect = CInt(cSection.Value("FocusRect", CInt(.FocusRect)))
170       .Font.Bold = CBool(cSection.Value("Font.Bold", CBool(.Font.Bold)))
180       .Font.Charset = cSection.Value("Font.Charset", .Font.Charset)
190       .Font.Italic = CBool(cSection.Value("Font.Italic", CBool(.Font.Italic)))
200       .Font.Name = cSection.Value("Font.Name", .Font.Name)
210       .Font.Size = CInt(cSection.Value("Font.Size", CInt(.Font.Size)))
220       .Font.Strikethrough = CBool(cSection.Value("Font.Strikethrough", CBool(.Font.Strikethrough)))
230       .Font.Underline = CBool(cSection.Value("Font.Underline", CBool(.Font.Underline)))
240       .Font.Weight = CInt(cSection.Value("Font.Weight", CInt(.Font.Weight)))
250       .FontWidth = CLng(cSection.Value("FontWidth", CLng(.FontWidth)))
260       .ForeColor = CLng(cSection.Value("ForeColor", CLng(.ForeColor)))
270       .ForeColorFixed = CLng(cSection.Value("ForeColorFixed", CLng(.ForeColorFixed)))
280       .ForeColorSel = CLng(cSection.Value("ForeColorSel", CLng(.ForeColorSel)))
290       .FormatString = cSection.Value("FormatString", .FormatString)
300       .GridColor = CLng(cSection.Value("GridColor", CLng(.GridColor)))
310       .GridColorFixed = CLng(cSection.Value("GridColorFixed", CLng(.GridColorFixed)))
320       .GridLines = CLng(cSection.Value("GridLines", CLng(.GridLines)))
330       .GridLinesFixed = CLng(cSection.Value("GridLinesFixed", CLng(.GridLinesFixed)))
340       .GridLineWidth = CLng(cSection.Value("GridLineWidth", CLng(.GridLineWidth)))
350       .HighLight = CInt(cSection.Value("HighLight", CInt(.HighLight)))
360       .MergeCells = CInt(cSection.Value("MergeCells", CInt(.MergeCells)))
370       .MousePointer = CInt(cSection.Value("MousePointer", CInt(.MousePointer)))
380       .RowHeightMin = CLng(cSection.Value("RowHeightMin", CLng(.RowHeightMin)))
390       .ScrollBars = CInt(cSection.Value("ScrollBars", CInt(.ScrollBars)))
400       .ScrollTrack = CInt(cSection.Value("ScrollTrack", CInt(.ScrollTrack)))
410       .SelectionMode = CInt(cSection.Value("SelectionMode", CInt(.SelectionMode)))
420       .Tag = cSection.Value("Tag", CStr(.Tag))
430       .TextStyle = CInt(cSection.Value("TextStyle", CInt(.TextStyle)))
440       .TextStyleFixed = CInt(cSection.Value("TextStyleFixed", CInt(.TextStyleFixed)))
450       .WordWrap = CBool(cSection.Value("WordWrap", CBool(.WordWrap)))
460     End With

470     On Error GoTo 0
480     Exit Sub

LoadGeneral_Error:

490     'MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbLf & "On line: " & Erl & vbLf & "in procedure LoadGeneral of Class Module clsFlexSettings", vbCritical
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = "Ocorreu um erro ao carregar os dados da biblioteca."
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SaveGeneral
' DateTime  : 4-7-2004 00:00
' Author    : Flyguy
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SaveGeneral()
        Dim cSection As clsSection
    
10      On Error GoTo SaveGeneral_Error

20      Set cSection = New clsSection
30      cSection.Key = "General"
  
40      With m_FG
50        cSection.Item("AllowBigSelection") = CBool(.AllowBigSelection)
60        cSection.Item("AllowUserResizing") = CInt(.AllowUserResizing)
70        cSection.Item("Appearance") = CInt(.Appearance)
80        cSection.Item("BackColor") = CLng(.BackColor)
90        cSection.Item("BackColorBkg") = CLng(.BackColorBkg)
100       cSection.Item("BackColorFixed") = CLng(.BackColorFixed)
110       cSection.Item("BackColorSel") = CLng(.BackColorSel)
120       cSection.Item("BorderStyle") = CInt(.BorderStyle)
130       cSection.Item("CausesValidation") = CBool(.CausesValidation)
140       cSection.Item("DragMode") = CInt(.DragMode)
150       cSection.Item("Enabled") = CBool(.Enabled)
160       cSection.Item("FillStyle") = CInt(.FillStyle)
170       cSection.Item("FocusRect") = CInt(.FocusRect)
180       cSection.Item("Font.Bold") = CBool(.Font.Bold)
190       cSection.Item("Font.Charset") = .Font.Charset
200       cSection.Item("Font.Italic") = CBool(.Font.Italic)
210       cSection.Item("Font.Name") = .Font.Name
220       cSection.Item("Font.Size") = CInt(.Font.Size)
230       cSection.Item("Font.Strikethrough") = CBool(.Font.Strikethrough)
240       cSection.Item("Font.Underline") = CBool(.Font.Underline)
250       cSection.Item("Font.Weight") = CInt(.Font.Weight)
260       cSection.Item("FontWidth") = CLng(.FontWidth)
270       cSection.Item("ForeColor") = CLng(.ForeColor)
280       cSection.Item("ForeColorFixed") = CLng(.ForeColorFixed)
290       cSection.Item("ForeColorSel") = CLng(.ForeColorSel)
300       cSection.Item("FormatString") = .FormatString
310       cSection.Item("GridColor") = CLng(.GridColor)
320       cSection.Item("GridColorFixed") = CLng(.GridColorFixed)
330       cSection.Item("GridLines") = CLng(.GridLines)
340       cSection.Item("GridLinesFixed") = CLng(.GridLinesFixed)
350       cSection.Item("GridLineWidth") = CLng(.GridLineWidth)
360       cSection.Item("HighLight") = CInt(.HighLight)
370       cSection.Item("MergeCells") = CInt(.MergeCells)
380       cSection.Item("MousePointer") = CInt(.MousePointer)
390       cSection.Item("RowHeightMin") = CLng(.RowHeightMin)
400       cSection.Item("ScrollBars") = CInt(.ScrollBars)
410       cSection.Item("ScrollTrack") = CInt(.ScrollTrack)
420       cSection.Item("SelectionMode") = CInt(.SelectionMode)
430       cSection.Item("Tag") = CStr(.Tag)
440       cSection.Item("TextStyle") = CInt(.TextStyle)
450       cSection.Item("TextStyleFixed") = CInt(.TextStyleFixed)
460       cSection.Item("WordWrap") = CBool(.WordWrap)
470     End With

480     m_cINI.Remove cSection.Key
490     m_cINI.AddSection cSection.Key, cSection

500     On Error GoTo 0
510     Exit Sub

SaveGeneral_Error:

520     'MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbLf & "On line: " & Erl & vbLf & "in procedure SaveGeneral of Class Module clsFlexSettings", vbCritical
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = "Ocorreu um erro ao guardar os dados da biblioteca."
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadLayout
' DateTime  : 4-7-2004 00:09
' Author    : Flyguy
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LoadLayout()
        Dim cSection As clsSection
        Dim sData As String, aData() As String
        Dim i As Long
  
10      On Error Resume Next 'On Error GoTo LoadLayout_Error

20      Set cSection = m_cINI.Section("Layout")
  
30      With m_FG
40        .Cols = CLng(cSection.Value("Cols", CLng(.Cols)))
50        .Rows = CLng(cSection.Value("Rows", CLng(.Rows)))
60        .FixedCols = CLng(cSection.Value("FixedCols", CLng(.FixedCols)))
70        .FixedRows = CLng(cSection.Value("FixedRows", CLng(.FixedRows)))
80        .LeftCol = CLng(cSection.Value("LeftCol", CLng(.LeftCol)))
90        .TopRow = CLng(cSection.Value("TopRow", CLng(.TopRow)))
    
100       sData = cSection.Value("ColAlignment", "")
110       If Len(sData) > 0 Then
120         aData = Split(sData, vbTab)
130         For i = 0 To UBound(aData)
140           If Len(aData(i)) > 0 And i < .Cols Then .ColAlignment(i) = CInt(aData(i))
150         Next i
160       End If
    
170       sData = cSection.Value("ColWidth", "")
180       If Len(sData) > 0 Then
190         aData = Split(sData, vbTab)
200         For i = 0 To UBound(aData)
210           If Len(aData(i)) > 0 And i < .Cols Then .ColWidth(i) = CLng(aData(i))
220         Next i
230       End If
    
240       sData = cSection.Value("ColData", "")
250       If Len(sData) > 0 Then
260         aData = Split(sData, vbTab)
270         For i = 0 To UBound(aData)
280           If Len(aData(i)) > 0 And i < .Cols Then .ColData(i) = CLng(aData(i))
290         Next i
300       End If
    
310       sData = cSection.Value("RowData", "")
320       If Len(sData) > 0 Then
330         aData = Split(sData, vbTab)
340         For i = 0 To UBound(aData)
350           If Len(aData(i)) > 0 And i < .Rows Then .RowData(i) = CLng(aData(i))
360         Next i
370       End If
    
380       sData = cSection.Value("RowHeight", "")
390       If Len(sData) > 0 Then
400         aData = Split(sData, vbTab)
410         For i = 0 To UBound(aData)
420           If Len(aData(i)) > 0 And i < .Rows Then .RowHeight(i) = CLng(aData(i))
430         Next i
440       End If
    
450     End With

460     On Error GoTo 0
470     Exit Sub

LoadLayout_Error:

480
'    With Form_Mensagem
'        .Label_Titulo.Caption = "Biblioteca vazia"
'        .Pic_Mensagem.Picture = Form_Skin.Icon_Info.Picture
'        .Label_Mensagem.Caption = "Não foram encontrados dados na para carregar."
'        .Botao_Ok.Visible = True
'        .Show vbModal
'    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SaveLayout
' DateTime  : 4-7-2004 00:09
' Author    : Flyguy
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SaveLayout()
        Dim cSection As clsSection
        Dim sData() As String, aData() As String
        Dim i As Long
  
10      On Error GoTo SaveLayout_Error

20      Set cSection = New clsSection
30      cSection.Key = "layout"
  
40      With m_FG
50        cSection.Item("Cols") = CLng(.Cols)
60        cSection.Item("Rows") = CLng(.Rows)
70        cSection.Item("FixedCols") = CLng(.FixedCols)
80        cSection.Item("FixedRows") = CLng(.FixedRows)
90        cSection.Item("LeftCol") = CLng(.LeftCol)
100       cSection.Item("TopRow") = CLng(.TopRow)
    
110       ReDim sData(2)
120       For i = 0 To .Cols - 1
130         sData(0) = sData(0) & vbTab & CStr(.ColAlignment(i))
140         sData(1) = sData(1) & vbTab & CStr(.ColWidth(i))
150         sData(2) = sData(2) & vbTab & CStr(.ColData(i))
            '.ColPosition
160       Next i
170       cSection.Item("ColAlignment") = Mid$(sData(0), 2)
180       cSection.Item("ColWidth") = Mid$(sData(1), 2)
190       cSection.Item("ColData") = Mid$(sData(2), 2)
    
200       ReDim sData(1)
210       For i = 0 To .Rows - 1
220         sData(0) = sData(0) & vbTab & CStr(.RowData(i))
230         sData(1) = sData(1) & vbTab & CStr(.RowHeight(i))
            '.RowPosition
240       Next i
250       cSection.Item("RowData") = Mid$(sData(0), 2)
260       cSection.Item("RowHeight") = Mid$(sData(1), 2)
    
270     End With
  
280     m_cINI.Remove cSection.Key
290     m_cINI.AddSection cSection.Key, cSection

300     On Error GoTo 0
310     Exit Sub

SaveLayout_Error:

320     'MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbLf & "On line: " & Erl & vbLf & "in procedure SaveLayout of Class Module clsFlexSettings", vbCritical
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = "Ocorreu um erro ao guardar os dados da biblioteca."
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
End Sub

Private Sub LoadCellLayout()
        Dim i As Long
        Dim lRow As Long, lCol As Long
        Dim aData() As String, sData As String
        Dim aXYD() As String
        Dim cSection As clsSection
  
10      On Error GoTo LoadCellLayout_Error

20      Set cSection = m_cINI.Section("CellLayout")
  
30      With m_FG
40        sData = cSection.Value("Alignment", "")
50        If Len(sData) > 0 Then
60          aData = Split(sData, vbTab)
70          For i = 0 To UBound(aData)
80            aXYD = Split(aData(i), ":")
90            .Row = CLng(aXYD(0))
100           .Col = CLng(aXYD(1))
110           .CellAlignment = CInt(aXYD(2))
120         Next i
130       End If
  
140       sData = cSection.Value("BackColor", "")
150       If Len(sData) > 0 Then
160         aData = Split(sData, vbTab)
170         For i = 0 To UBound(aData)
180           aXYD = Split(aData(i), ":")
190           .Row = CLng(aXYD(0))
200           .Col = CLng(aXYD(1))
210           .CellBackColor = CLng(aXYD(2))
220         Next i
230       End If
  
240       sData = cSection.Value("FontBold", "")
250       If Len(sData) > 0 Then
260         aData = Split(sData, vbTab)
270         For i = 0 To UBound(aData)
280           aXYD = Split(aData(i), ":")
290           .Row = CLng(aXYD(0))
300           .Col = CLng(aXYD(1))
310           .CellFontBold = CBool(aXYD(2))
320         Next i
330       End If
  
340       sData = cSection.Value("FontItalic", "")
350       If Len(sData) > 0 Then
360         aData = Split(sData, vbTab)
370         For i = 0 To UBound(aData)
380           aXYD = Split(aData(i), ":")
390           .Row = CLng(aXYD(0))
400           .Col = CLng(aXYD(1))
410           .CellFontItalic = CBool(aXYD(2))
420         Next i
430       End If
  
440       sData = cSection.Value("FontName", "")
450       If Len(sData) > 0 Then
460         aData = Split(sData, vbTab)
470         For i = 0 To UBound(aData)
480           aXYD = Split(aData(i), ":")
490           .Row = CLng(aXYD(0))
500           .Col = CLng(aXYD(1))
510           .CellFontName = aXYD(2)
520         Next i
530       End If

540       sData = cSection.Value("FontSize", "")
550       If Len(sData) > 0 Then
560         aData = Split(sData, vbTab)
570         For i = 0 To UBound(aData)
580           aXYD = Split(aData(i), ":")
590           .Row = CLng(aXYD(0))
600           .Col = CLng(aXYD(1))
610           .CellFontSize = CLng(aXYD(2))
620         Next i
630       End If
  
640       sData = cSection.Value("ForeColor", "")
650       If Len(sData) > 0 Then
660         aData = Split(sData, vbTab)
670         For i = 0 To UBound(aData)
680           aXYD = Split(aData(i), ":")
690           .Row = CLng(aXYD(0))
700           .Col = CLng(aXYD(1))
710           .CellForeColor = CLng(aXYD(2))
720         Next i
730       End If
    
740       sData = cSection.Value("TextStyle", "")
750       If Len(sData) > 0 Then
760         aData = Split(sData, vbTab)
770         For i = 0 To UBound(aData)
780           aXYD = Split(aData(i), ":")
790           .Row = CLng(aXYD(0))
800           .Col = CLng(aXYD(1))
810           .CellTextStyle = CInt(aXYD(2))
820         Next i
830       End If
840     End With

850     On Error GoTo 0
860     Exit Sub

LoadCellLayout_Error:

870     'MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbLf & "On line: " & Erl & vbLf & "in procedure LoadCellLayout of Class Module clsFlexSettings", vbCritical
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = "Ocorreu um erro ao carregar os dados da biblioteca."
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
End Sub

Private Sub SaveCellLayout()
        Dim lCol As Long, lRow As Long
        Dim xx As tpCELL
        Dim cSection As clsSection
  
10      On Error GoTo SaveCellLayout_Error

20      With m_FG
30        For lRow = 0 To .Rows - 1
40          .Row = lRow
50          For lCol = 0 To .Cols - 1
60            .Col = lCol
70            If .CellAlignment <> 0 Then _
                xx.Alignment = xx.Alignment & vbTab & lRow & ":" & lCol & ":" & CStr(.CellAlignment)
        
80            If .CellBackColor <> 0 And .BackColor <> .CellBackColor Then _
                xx.BackColor = xx.BackColor & vbTab & lRow & ":" & lCol & ":" & CStr(.CellBackColor)
          
90            If .CellForeColor <> 0 And .ForeColor <> .CellForeColor Then _
                xx.ForeColor = xx.ForeColor & vbTab & lRow & ":" & lCol & ":" & CStr(.CellForeColor)
          
100           If .CellFontBold <> .Font.Bold Then _
                xx.FontBold = xx.FontBold & vbTab & lRow & ":" & lCol & ":" & CStr(.CellFontBold)
        
110           If .CellFontItalic <> .Font.Italic Then _
                xx.FontItalic = xx.FontItalic & vbTab & lRow & ":" & lCol & ":" & CStr(.CellFontItalic)
          
120           If .CellFontName <> .Font.Name Then _
                xx.FontName = xx.FontName & vbTab & lRow & ":" & lCol & ":" & CStr(.CellFontName)
        
130           If .CellFontSize <> .Font.Size Then _
                xx.FontSize = xx.FontSize & vbTab & lRow & ":" & lCol & ":" & CStr(.CellFontSize)
          
140           If .CellTextStyle <> .TextStyle Then _
                xx.TextStyle = xx.TextStyle & vbTab & lRow & ":" & lCol & ":" & CStr(.CellTextStyle)
          
150         Next lCol
160       Next lRow
170     End With
  
180     Set cSection = New clsSection
190     cSection.Key = "CellLayout"
  
200     cSection.Item("Alignment") = Mid$(xx.Alignment, 2)
210     cSection.Item("BackColor") = Mid$(xx.BackColor, 2)
220     cSection.Item("FontBold") = Mid$(xx.FontBold, 2)
230     cSection.Item("FontItalic") = Mid$(xx.FontItalic, 2)
240     cSection.Item("FontName") = Mid$(xx.FontName, 2)
250     cSection.Item("FontSize") = Mid$(xx.FontSize, 2)
260     cSection.Item("ForeColor") = Mid$(xx.ForeColor, 2)
270     cSection.Item("TextStyle") = Mid$(xx.TextStyle, 2)

280     m_cINI.Remove cSection.Key
290     m_cINI.AddSection cSection.Key, cSection

300     On Error GoTo 0
310     Exit Sub

SaveCellLayout_Error:

320     'MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbLf & "On line: " & Erl & vbLf & "in procedure SaveCellLayout of Class Module clsFlexSettings", vbCritical
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = "Ocorreu um erro ao guardar os dados da biblioteca."
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
End Sub

Private Function LoadData() As Boolean
        Dim cSection As clsSection
        Dim lRows As Long, lCols As Long
        Dim lRow As Long, lCol As Long
        Dim aData() As String, sData As String
  
10      On Error GoTo LoadData_Error

20      Set cSection = m_cINI.Section("Data")
  
30      lRows = CLng(cSection.Value("rows", "-1"))
40      lCols = CLng(cSection.Value("cols", "-1"))
  
50      If lCols = -1 Or lRows = -1 Then Exit Function
  
60      For lRow = 0 To lRows - 1
70        sData = cSection.Value("row:" & lRow, "")
80        If Len(sData) > 0 Then
90          aData = Split(sData, vbTab)
100         If UBound(aData) + 1 = lCols Then
110           For lCol = 0 To UBound(aData)
120             m_FG.TextMatrix(lRow, lCol) = aData(lCol)
130           Next lCol
140         End If
150       End If
160     Next lRow
  
170     LoadData = True

180     On Error GoTo 0
190     Exit Function

LoadData_Error:

200     'MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbLf & "On line: " & Erl & vbLf & "in procedure LoadData of Class Module clsFlexSettings", vbCritical
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = "Ocorreu um erro ao carregar os dados da biblioteca."
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
End Function

Private Function SaveData() As Boolean
        Dim lRow As Long, lCol As Long
        Dim cSection As clsSection
        Dim sData As String
  
10      On Error GoTo SaveData_Error

20      Set cSection = New clsSection
30      cSection.Key = "Data"
40      cSection.Item("rows") = CStr(m_FG.Rows)
50      cSection.Item("cols") = CStr(m_FG.Cols)
  
60      With m_FG
70        For lRow = 0 To .Rows - 1
80          sData = ""
90          For lCol = 0 To .Cols - 1
100           sData = sData & vbTab & .TextMatrix(lRow, lCol)
110         Next lCol
120         cSection.Item("row:" & lRow) = Mid$(sData, 2)
130       Next lRow
140     End With
  
150     m_cINI.Remove cSection.Key
160     m_cINI.AddSection cSection.Key, cSection
  
170     SaveData = True

180     On Error GoTo 0
190     Exit Function

SaveData_Error:

200     'MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbLf & "On line: " & Erl & vbLf & "in procedure SaveData of Class Module clsFlexSettings", vbCritical
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = "Ocorreu um erro ao guardar os dados da biblioteca."
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
End Function

Private Sub Class_Terminate()
  Set m_FG = Nothing
  Set m_cINI = Nothing
End Sub
