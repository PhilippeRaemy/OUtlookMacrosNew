Attribute VB_Name = "SearchHelper"
Public BlnSearchComp As Boolean
  
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Outlook.Search)
    Debug.Print "The AdvancedSearchComplete Event fired"
    If SearchObject.tag = "Test" Then
        m_SearchComplete = True
    End If
  
End Sub
  
Sub TestAdvancedSearchComplete()
    Dim sch As Outlook.Search
    Dim rsts As Outlook.results
    Dim i As Integer
    BlnSearchComp = False
    Const strF As String = "urn:schemas:mailheader:subject = 'SPAM'"
    Const strS As String = "Inbox"
    Set sch = Application.AdvancedSearch(strS, strF, True, "Test")
    While BlnSearchComp = False
        DoEvents
    Wend
    Set rsts = sch.results
    For i = 1 To rsts.count
        Debug.Print rsts.Item(i).SenderName
    Next
End Sub
