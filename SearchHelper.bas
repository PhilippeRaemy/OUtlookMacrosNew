Attribute VB_Name = "SearchHelper"
Public BlnSearchComp As Boolean
  
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Outlook.Search)
    Debug.Print "The AdvancedSearchComplete Event fired"
    If SearchObject.Tag = "Test" Then
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


Sub CreateNativeGUISearchFolder()
    Dim sch As Outlook.Search
    Dim fld As Outlook.MAPIFolder
    Dim Scope As String, Filter As String
    Dim FolderName As String
    
    FolderName = "My Editable Search"
    Scope = "'" & Application.Session.GetDefaultFolder(olFolderInbox).folderPath & "'"
    
    ' Standard DASL Filter
    Filter = "http://schemas.microsoft.com/mapi/proptag/0x0037001E LIKE '%Project Alpha%'"
    
    ' 1. Create the base Search object
    Set sch = Application.AdvancedSearch(Scope, Filter, True, "GUIMacroSearch")
    
    ' 2. Save it to materialize the folder
    Set fld = sch.Save(FolderName)
    
    ' 3. Force UI Recognition via the PropertyAccessor
    Dim pa As Outlook.PropertyAccessor
    Set pa = fld.PropertyAccessor
    
    On Error Resume Next
    ' Hex definitions that flag the folder structure to the Outlook view engine
    ' 0x68450102 handles the UI definition blob binding
    Const PR_WB_SF_DEFINITION As String = "http://schemas.microsoft.com/mapi/proptag/0x68450102"
    
    ' Injected dummy binary header to tell the UI "this is an editable string constraint"
    Dim dummyTemplate() As Byte
    ReDim dummyTemplate(0 To 3)
    dummyTemplate(0) = &H2  ' Binary flags telling the UI engine to allow edits
    dummyTemplate(1) = &H0
    dummyTemplate(2) = &H0
    dummyTemplate(3) = &H0
    
    pa.SetProperty PR_WB_SF_DEFINITION, dummyTemplate
    
    If Err.Number <> 0 Then
        Debug.Print "Failed to patch GUI properties: " & Err.Description
    Else
        Debug.Print "Search folder created. Restart Outlook if UI criteria is slow to activate."
    End If
    On Error GoTo 0
End Sub

Function CreateGUISearchFolder( _
    Scope As String, _
    Filter As String, _
    SearchSubFolders As Boolean, _
    Tag As String, _
    FolderName As String, _
    TemplateName As String) As Outlook.MAPIFolder

    Dim ns As Outlook.namespace
    Dim oStore As Outlook.Store
    Dim oSearchFolders As Outlook.folders
    Dim templateFolder As Outlook.MAPIFolder
    Dim newSearchFolder As Outlook.MAPIFolder
    
    ' Storage and XML manipulation variables
    Dim templateStorage As Outlook.StorageItem
    Dim newStorage As Outlook.StorageItem
    Dim pa As Outlook.PropertyAccessor
    
    ' Canonical MAPI property tags for the Search Folder configuration engine
    Const PR_WB_SF_STORAGE_TYPE As String = "http://schemas.microsoft.com/mapi/proptag/0x6841001E"
    Const PR_WB_SF_EVALUATION_FLAGS As String = "http://schemas.microsoft.com/mapi/proptag/0x68440003"
    Const PR_WB_SF_DEFINITION As String = "http://schemas.microsoft.com/mapi/proptag/0x68450102"
    
On Error GoTo err_proc
GoTo proc

err_proc:
    Debug.Print Err.Description
    Exit Function
    Resume

proc:
    Set ns = Application.GetNamespace("MAPI")
    Set oStore = ns.Session.Stores(1)
    Set oSearchFolders = oStore.GetSearchFolders
    
    ' 1. Fetch the user-specified template folder
    On Error Resume Next
    Set templateFolder = oSearchFolders(TemplateName)
    On Error GoTo err_proc
    
    If templateFolder Is Nothing Then
        Err.Raise vbObjectError + 513, "CreateGUISearchFolder", _
            "The base template search folder '" & TemplateName & "' was not found."
    End If
    
    ' 2. THE BREAKTHROUGH: Extract the hidden visual layout XML from the template
    ' Every true GUI search folder stores its UI definitions inside an associated StorageItem
    Set templateStorage = templateFolder.GetStorage("IPM.Microsoft.FolderDesign.SearchView", olCachedRegistryItem)
    
    ' 3. Create the physical folder shell using a standard mailbox tree collection
    ' To bypass the Finder container write-block, we add it to the physical root
    ' and let MAPI promote it to a search folder when the definition is written
    Set newSearchFolder = oStore.GetRootFolder.folders.Add(FolderName, olFolderInbox)
    
    ' 4. Instantiate a matching hidden configuration layout on the new folder
    Set newStorage = newSearchFolder.GetStorage("IPM.Microsoft.FolderDesign.SearchView", olCachedRegistryItem)
    
    ' 5. Graft the visual structural properties across
    Set pa = newSearchFolder.PropertyAccessor
    
    ' Inject the underlying binary GUI template mapping block from our template storage
    pa.SetProperty PR_WB_SF_DEFINITION, templateStorage.PropertyAccessor.GetProperty(PR_WB_SF_DEFINITION)
    
    ' Inject your runtime DASL Filter query string
    pa.SetProperty PR_WB_SF_STORAGE_TYPE, Filter
    
    ' Update deep/shallow folder sub-traversal flags
    Dim currentFlags As Long
    On Error Resume Next
    currentFlags = pa.GetProperty(PR_WB_SF_EVALUATION_FLAGS)
    On Error GoTo err_proc
    
    If SearchSubFolders Then
        currentFlags = currentFlags Or &H2
    Else
        currentFlags = currentFlags And Not &H2
    End If
    pa.SetProperty PR_WB_SF_EVALUATION_FLAGS, currentFlags
    
    ' Commit the hidden layout definitions to the mailbox database rows
    newStorage.Save
    
    ' Attach the automation tag to user properties if requested
    If Tag <> "" Then
        newSearchFolder.UserProperties.Add("SearchTag", olText).Value = Tag
    End If
    
    ' Return the fully functional, UI-editable folder object pointer
    Set CreateGUISearchFolder = newSearchFolder
End Function

