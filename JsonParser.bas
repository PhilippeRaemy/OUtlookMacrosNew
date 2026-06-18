Attribute VB_Name = "JsonParser"
Sub testParseJson()
Dim results As Scripting.Dictionary
Dim key As Variant
    Set results = ParseJson("{""name"": ""John Doe"", ""age"": 30, ""isStudent"": false}")
    For Each key In results.Keys
        Debug.Print "key:"; key, "value:"; results(key)
    Next key
End Sub

Function ParseJson(jsonString As String) As Scripting.Dictionary

    Dim matches As Object
    Dim match As Object
    Set ParseJson = New Scripting.Dictionary
    Dim regEx As New VBScript_RegExp_55.RegExp
    
    ' Create a RegExp object
    With regEx
        .Pattern = """([^""]+)""\s*:\s*(""([^""]*)""|\d+|true|false|null)"
        .Global = True
    End With

    ' Execute the regex search
    Set matches = regEx.Execute(jsonString)

    ' Iterate through the matches and print key-value pairs
    For Each match In matches
        ParseJson.Add match.SubMatches(0), IIf(IsEmpty(match.SubMatches(2)), match.SubMatches(1), match.SubMatches(2))
    Next match
End Function


