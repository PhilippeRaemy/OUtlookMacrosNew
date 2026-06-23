Attribute VB_Name = "ClassificationRuleStatic"
Public RulesDic As scripting.Dictionary

Private inited As Boolean



' This function returns 2 if an actual new rule has been added, 1 if the rule has been merged into a simlar one (same folder, same retention) or 0 if the new rule was already covered
Public Function NewClassificationRule(Folder As String, Addresses As String, Optional seq As Integer = 0, Optional Retention As String = "") As Boolean
    If RulesDic Is Nothing Then Set RulesDic = New scripting.Dictionary
    Dim rule As New ClassificationRule
    Dim existing As Variant
    Dim newad As Variant
    rule.Folder = Folder
    Dim FolderName As String: FolderName = rule.FolderName
    Dim required As Boolean
    If RulesDic.Exists(FolderName) Then
        Set rule = RulesDic.Item(FolderName)
        For Each newad In VBA.Split(Addresses)
            required = False
            For Each existing In VBA.Split(rule.Addresses)
                If Not Replace(newad, "%", "*") Like Replace(existing, "%", "*") Then
                    required = True
                    Exit For
                End If
            Next existing
            If required Then
                NewClassificationRule = 1
                rule.Addresses = rule.Addresses & ":" & newad
            End If
        Next newad
        
        rule.Addresses = RulesDic.Item(FolderName).Addresses & ";" & Addresses
        NewClassificationRule = False
    Else
        rule.Addresses = Addresses
        ' in case when setting the folder name has parsed it into these props, we're only overriding if they're not default
        If Retention <> "" Then rule.RetentionText = Retention
        If seq <> 0 Then rule.Sequence = seq
        RulesDic.Add FolderName, rule
        NewClassificationRule = 2
    End If
End Function

Private Function dicGet(dic As scripting.Dictionary, Item As Variant, default As Variant) As Variant
    If dic.Exists(Item) Then
        dicGet = dic(Item)
    Else
        dicGet = default
    End If
End Function


Public Sub ReadRules()
Dim fso As New FileSystemObject
Dim line As String
Dim parsedLine As scripting.Dictionary
Dim count As Integer: count = -1
Dim AnyMerge As Boolean
Dim r As Variant
Dim originalFilePath As String: originalFilePath = ThisOutlookSession.GetSearchesFileName()

    If RulesDic Is Nothing Then
        Set RulesDic = New scripting.Dictionary
        RulesDic.CompareMode = TextCompare
    Else
        RulesDic.RemoveAll
    End If
    
    With fso.OpenTextFile(originalFilePath, ForReading)
        While Not .AtEndOfStream
            line = .ReadLine
            Set parsedLine = ParseJson(line)
            AnyMerge = AnyMerge Or _
                (0 <> NewClassificationRule( _
                        parsedLine("name"), _
                        parsedLine("criterion"), _
                        dicGet(parsedLine, "seq", 0), _
                        dicGet(parsedLine, "ret", "")))
        Wend
    End With
    
    If AnyMerge Then
' Build precise ISO file naming layout
        fileExtension = fso.GetExtensionName(originalFilePath)
        baseFilePath = Left(originalFilePath, Len(originalFilePath) - Len(fileExtension) - 1)
        backupFilePath = baseFilePath & "_" & Format(Now, "yyyymmdd-hhmmss") & "." & fileExtension
        
        ' Create physical copy tracking point
        fso.CopyFile originalFilePath, backupFilePath, True
        
        ' Rewrite pristine, consolidated data sets to the original source path
        With fso.OpenTextFile(originalFilePath, ForWriting, True)
            For Each r In RulesDic
                If Not IsEmpty(RulesDic.Item(r)) Then .WriteLine RulesDic.Item(r).ToJson
            Next r
        End With
    End If
    inited = True
End Sub

Sub test_read_rules()

    ReadRules
    
    Stop
    
End Sub

Function CheckRuleExists(Addresses) As Boolean
Dim i As Integer

    If UBound(rules) < 0 Then ReadRules
    CheckRuleExists = True
    For i = LBound(rules) To UBound(rules)
        If rules(i).Addresses = Addresses Then Exit Function
    Next i

    CheckRuleExists = False
    
End Function

Function AnyRuleMatch(ParamArray Addresses() As Variant) As Boolean
Dim aList As Variant
Dim addr As Variant
Dim rule As Variant ' ClassificationRule
Dim addrMask As Variant
    
On Error GoTo err_proc
GoTo proc

err_proc:
Debug.Print Err.Description
Exit Function
Resume

proc:
    
    If Not inited Then ReadRules

    For Each aList In Addresses
        For Each addr In VBA.Split(aList, ";")
            For Each rule In rules
                For Each addrMask In VBA.Split(rule.Addresses, ";")
                    If addr Like Replace(addrMask, "%", "*") Then
                        AnyRuleMatch = True
                        Exit Function
                    End If
                Next
            Next
        Next
    Next
End Function
