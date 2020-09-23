VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsSection"
Attribute VB_Ext_KEY = "Member0" ,"clsSection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : clsINI
' DateTime  : 22-8-2003 09:48
' Author    : FlyGuy
' Purpose   : A collection of clsSection to manipulate INI files
'---------------------------------------------------------------------------------------
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private m_sFileName As String

Private Function NewSection(SectionKey As String) As clsSection
  'create a new object
  Dim objNewMember As clsSection
  Set objNewMember = New clsSection

  'set the properties passed into the method
  objNewMember.Key = SectionKey
  mCol.Add objNewMember, SectionKey

  'return the object created
  Set NewSection = objNewMember
  Set objNewMember = Nothing
End Function

Public Function AddSection(SectionKey As String, cSection As clsSection) As Boolean
  ' return False if section already exists
  If Not SectionExists(SectionKey) Then
    mCol.Add cSection, SectionKey
    AddSection = True
  End If
End Function

Public Property Get Section(vntIndexKey As Variant) As clsSection
Attribute Section.VB_UserMemId = 0
  'used when referencing an element in the collection
  'vntIndexKey contains either the Index or Key to the collection,
  'this is why it is declared as a Variant
  'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  
  If IsNumeric(vntIndexKey) Then
    ' get section based on Index
    If vntIndexKey > 0 And vntIndexKey <= mCol.count Then
      Set Section = mCol(vntIndexKey)
    End If
  Else
    ' get section based on key
    ' if section doesn't exist, create new one
    If SectionExists(CStr(vntIndexKey)) Then
      Set Section = mCol(vntIndexKey)
    Else
      Set Section = New clsSection
      Section.Key = CStr(vntIndexKey)
      mCol.Add Section, CStr(vntIndexKey)
    End If
  End If
End Property

Public Function SectionExists(sKey As String) As Boolean
  Dim aKeys() As String, i As Long
  
  ' Instead of error trapping
  If mCol.count > 0 Then
    aKeys = SectionKeys
    For i = 0 To UBound(aKeys)
      If LCase$(aKeys(i)) = LCase$(sKey) Then
        SectionExists = True
      End If
    Next i
  End If

End Function

Public Function SectionKeys() As String()
  Dim i As Long
  Dim aTemp() As String
  Dim cSection As clsSection
  
  If mCol.count > 0 Then
    ReDim aTemp(mCol.count - 1)
    For i = 1 To mCol.count
      Set cSection = mCol(i)
      aTemp(i - 1) = cSection.Key
    Next i
  End If
  SectionKeys = aTemp
End Function

Public Property Let Value(sSection As String, sItem As String, sValue As String)
  If Not SectionExists(sSection) Then
    NewSection sSection
  End If
  mCol.Item(sSection).Item(sItem) = sValue
End Property

Public Property Get Value(sSection As String, sItem As String) As String
  If SectionExists(sSection) Then
    Value = mCol.Item(sSection).Item(sItem)
  End If
End Property

Public Property Get count() As Long
  'used when retrieving the number of elements in the
  'collection. Syntax: Debug.Print x.Count
  count = mCol.count
End Property

Public Property Let FileName(ByVal sValue As String)
  m_sFileName = sValue
End Property

Public Property Get FileName() As String
  FileName = m_sFileName
End Property

Public Sub Remove(vntIndexKey As Variant)
  'used when removing an element from the collection
  'vntIndexKey contains either the Index or Key, which is why
  'it is declared as a Variant
  'Syntax: x.Remove(xyz)
  If IsNumeric(vntIndexKey) Then
    If vntIndexKey > 0 And vntIndexKey <= mCol.count Then
      mCol.Remove vntIndexKey
    End If
  Else
    If SectionExists(CStr(vntIndexKey)) Then
      mCol.Remove vntIndexKey
    End If
  End If
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  'this property allows you to enumerate
  'this collection with the For...Each syntax
  Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
  'creates the collection when this class is created
  Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
  'destroys collection when this class is terminated
  Set mCol = Nothing
End Sub

Public Function FromString(ByVal TheData As String, Optional bAppend As Boolean = False) As String
        Dim mcolSection As clsSection
        Dim i As Long
        Dim sKeySection As String
        Dim sKey As String, sData As String, sLine As String
        Dim lPos As Long
        Dim aData() As String
  
10      On Error GoTo FromString_Error

20      If Not bAppend Then
30        Class_Initialize
40      End If
  
50      Set mcolSection = New clsSection
  
60      If Len(TheData) > 0 Then
70        aData = Split(TheData, vbCrLf)
80        For i = 0 To UBound(aData)
90          sLine = aData(i)
100         If Len(sLine) > 0 And left$(sLine, 1) <> ";" Then
110           If left$(sLine, 1) = "[" Then
    
                ' eventueel vorige sectie toevoegen
120             If Len(mcolSection.Key) > 0 Then
130               If Not AddSection(mcolSection.Key, mcolSection) Then
140                 FromString = "Section already exists!" & vbCrLf & mcolSection.Key & vbCrLf
150               End If
160             End If
    
                ' section key
170             sKeySection = Mid$(sLine, 2, Len(sLine) - 2)
                ' toevoegen
180             If SectionExists(sKeySection) Then
190               Set mcolSection = mCol(sKeySection)
200             Else
210               Set mcolSection = New clsSection
220               mcolSection.Key = sKeySection
230             End If
  
240           Else
250             lPos = InStr(1, sLine, "=", vbTextCompare)
260             If lPos > 1 Then
270               sKey = left$(sLine, lPos - 1)
280               sData = Mid$(sLine, lPos + 1)
                  'Add/update with new values
290               mcolSection.Item(sKey) = sData
300             End If
310           End If
320         End If
330       Next i
    
340       If Len(mcolSection.Key) > 0 Then
350         If Not AddSection(mcolSection.Key, mcolSection) Then
360           FromString = "Secção já existe!" & vbCrLf & mcolSection.Key & vbCrLf
370         End If
380       End If
    
390     End If

400     On Error GoTo 0
410     Exit Function

FromString_Error:
420     'MsgBox "Error: " & Err.Number & " on line: " & Erl & vbLf & _
               "Description: " & Err.Description & vbLf & _
               "in procedure FromString of Class Module clsINI", vbCritical
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = "Ocorreu um erro ao guardar os dados da biblioteca."
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
430     On Error GoTo 0

End Function

Public Function Load(ByVal sName As String, Optional bAppend As Boolean = False) As Long
  Dim iID As Integer
  Dim sData As String
  Dim sMessage As String
  
  On Error GoTo errHandler
  Load = -1
  m_sFileName = sName
  
  If Len(Dir(sName)) = 0 Then
    Load = 53
    Exit Function ' file does not exist
  End If
  
  ' Read from file
  iID = FreeFile
  Open sName For Input As iID
    sData = Input(LOF(iID), iID)
  Close iID
  
  ' Build data
  sMessage = FromString(sData, bAppend)
  If Len(sMessage) > 0 Then 'MsgBox sMessage & vbCrLf & sName, vbExclamation
    With Form_Mensagem
        .Label_Titulo.Caption = "Erro"
        .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
        .Label_Mensagem.Caption = sMessage
        .Botao_Ok.Visible = True
        .Show vbModal
    End With
  End If
  
  Load = 0
  On Error GoTo 0
  Exit Function
  
errHandler:
  Close iID
  Load = err.Number
  On Error GoTo 0
End Function

Public Function ToString() As String
  Dim i As Long, j As Long
  Dim cSection As clsSection
  
  For i = 1 To mCol.count
    Set cSection = mCol(i)
    
    If i = 1 Then
      ToString = vbCrLf & "[" & cSection.Key & "]"
    Else
      ToString = ToString & vbCrLf & vbCrLf & "[" & cSection.Key & "]"
    End If
    
    For j = 1 To cSection.count
      ToString = ToString & vbCrLf & cSection.DataLine(j)
    Next j
  Next i
End Function

Public Function Save(Optional ByVal sName As String = "", Optional sApplication As String = "") As Long
  Dim iID As Integer
  Dim sData As String
  
  On Error GoTo errHandler
  If Len(sName) > 0 Then m_sFileName = sName
  If Len(m_sFileName) = 0 Then
    Save = 53 ' File not found
    Exit Function
  End If
  
  iID = FreeFile
  
  sData = ToString
  
  If Len(sData) > 0 Then
    Open m_sFileName For Output As iID
      Print #iID, "; This file is generated by software, do not modify it by hand"
      If Len(sApplication) > 0 Then
        Print #iID, "; application = " & sApplication
      End If
      Print #iID, "; date = " & Format(Date, "yyyymmdd")
      Print #iID, "; time = " & Format(Time, "HHnnss")
      Print #iID, sData
    Close iID
  End If
  
  Save = 0
  
  On Error GoTo 0
  Exit Function
  
errHandler:
  Close iID
  Save = err.Number
  On Error GoTo 0
End Function

Public Sub Clone(cINI As clsINI)
  Set mCol = New Collection
  
  FromString cINI.ToString
  m_sFileName = cINI.FileName
End Sub
