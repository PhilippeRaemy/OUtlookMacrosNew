Attribute VB_Name = "ClassificationRuleStatic"
Private rules() As ClassificationRule
Private inited As Boolean

Public Function NewClassificationRule(folder As String, addresses As String) As ClassificationRule
    Set NewClassificationRule = New ClassificationRule
    NewClassificationRule.folder = folder
    NewClassificationRule.addresses = addresses
End Function


Public Sub ReadRules()
Dim fso As New FileSystemObject
Dim line As String
Dim parsedLine As Scripting.Dictionary
Dim count As Integer: count = -1

    ReDim rules(10)
    
    With fso.OpenTextFile(ThisOutlookSession.GetSearchesFileName(), ForReading)
        While Not .AtEndOfStream
            line = .ReadLine
            Set parsedLine = ParseJson(line)
            count = count + 1
            If count > UBound(rules) Then
                ReDim Preserve rules(count + 10)
            End If
            Set rules(count) = NewClassificationRule(parsedLine("name"), parsedLine("criterion"))
        Wend
    End With
    ReDim Preserve rules(count)
    inited = True
End Sub

Sub test_read_rules()

    ReadRules
    
    Stop
    
End Sub

Sub AddRule(rule As ClassificationRule)
    ReDim Preserve rules(LBound(rules) To UBound(rules) + 1)
    Set rules(UBound(rules)) = rule
End Sub

Function CheckRuleExists(addresses) As Boolean
Dim i As Integer

    CheckRuleExists = True
    For i = LBound(rules) To UBound(rules)
        If rules(i).addresses = addresses Then Exit Function
    Next i

    CheckRuleExists = False
    
End Function

Function AnyRuleMatch(ParamArray addresses() As Variant) As Boolean
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

    For Each aList In addresses
        For Each addr In VBA.Split(aList, ";")
            For Each rule In rules
                For Each addrMask In VBA.Split(rule.addresses, ";")
                    If addr Like Replace(addrMask, "%", "*") Then
                        AnyRuleMatch = True
                        Exit Function
                    End If
                Next
            Next
        Next
    Next
End Function
