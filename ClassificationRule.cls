VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassificationRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_folder As String

Public RetentionText As String
Public Addresses As String
Public Sequence As Integer

Public Property Let Folder(v As String)
    Dim chunks As Variant
    chunks = VBA.Split(v, "|")
    m_folder = chunks(0)
    If UBound(chunks) > 0 Then
        RetentionText = chunks(1)
        If UBound(chunks) > 1 Then
            Sequence = CInt(chunks(2))
        End If
    End If
End Property
Public Property Get Folder() As String
    Folder = m_folder
End Property


Public Property Get FolderName() As String
    If RetentionText <> "" Then
        FolderName = m_folder & "|" & RetentionText
    Else
        FolderName = m_folder
    End If
End Property


Public Function ToJson() As String
    ToJson = "{""name"":""" & Replace(Me.Folder, """", "\""") & """,""seq"":" & CStr(Me.Sequence) & ",""ret"":""" & Me.RetentionText & """,""criterion"":""" & Me.Addresses & """}"
End Function

