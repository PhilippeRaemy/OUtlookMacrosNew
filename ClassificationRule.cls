VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassificationRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public folder As String
Public addresses As String

Public Function ToJson() As String
    ToJson = "{""name"":""" & Replace(Me.folder, """", "\""") & """,""criterion"":""" & Me.addresses & """}"
End Function

