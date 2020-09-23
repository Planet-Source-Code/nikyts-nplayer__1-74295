VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : clsSection
' DateTime  : 22-8-2003 09:49
' Author    : Flyguy
' Purpose   : To manipulate section within a INI file
'---------------------------------------------------------------------------------------
Option Explicit

Private m_sKey As String
Private m_cValues As Collection
Private m_cValueKeys As Collection

Public Property Let Key(sNewValue As String)
  m_sKey = sNewValue
End Property

Public Property Get Key() As String
  Key = m_sKey
End Property

Public Function Add(sKey As String, sValue As String) As Boolean
  If Len(sKey) > 0 Then
    ' If replacing then first remove
    Remove sKey
    
    ' Add the actual data
    m_cValueKeys.Add sKey, sKey
    m_cValues.Add sValue, sKey
    
    Add = True
  End If
  
End Function

Public Function Count() As Long
  Count = m_cValueKeys.Count
End Function

Public Function DataLine(vntIndex As Variant) As String
  DataLine = m_cValueKeys(vntIndex) & "=" & m_cValues(vntIndex)
End Function

Public Function Exists(sKey As String)
  Dim i As Long
  
  If Len(Trim$(sKey)) = 0 Then Exit Function
  
  For i = 1 To m_cValueKeys.Count
    If LCase$(sKey) = LCase$(m_cValueKeys(i)) Then
      Exists = True
      Exit For
    End If
  Next i
End Function

Public Function Remove(sKey As String)
  If Exists(sKey) Then
    m_cValueKeys.Remove sKey
    m_cValues.Remove sKey
  End If
End Function

Public Function Clear()
  Set m_cValues = New Collection
  Set m_cValueKeys = New Collection
End Function

Public Property Let Item(sKey As String, sValue As String)
  
  If Exists(sKey) Then
    m_cValueKeys.Remove sKey
    m_cValues.Remove sKey
  End If
  
  m_cValues.Add sValue, sKey
  m_cValueKeys.Add sKey, sKey
End Property

Public Property Get Item(sKey As String) As String
  If Exists(sKey) Then Item = m_cValues(sKey)
End Property

Public Function ItemKeys() As String()
  Dim i As Long
  Dim aTemp() As String
  
  If m_cValueKeys.Count > 0 Then
    ReDim aTemp(m_cValueKeys.Count - 1)
    For i = 1 To m_cValueKeys.Count
      aTemp(i - 1) = m_cValueKeys(i)
    Next i
  End If
  ItemKeys = aTemp
End Function

Public Function Value(sKey As String, Optional sDefault As String = "") As String
  If Exists(sKey) Then
    Value = m_cValues(sKey)
  Else
    Value = sDefault
  End If
End Function

Private Sub Class_Initialize()
  Set m_cValues = New Collection
  Set m_cValueKeys = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_cValues = Nothing
  Set m_cValueKeys = Nothing
End Sub
