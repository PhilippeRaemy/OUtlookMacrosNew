Attribute VB_Name = "ImportThunderbird"
Sub ImportThunderbird()
Dim roodfld As Outlook.Folder
Dim fld As Outlook.Folder
Dim fs As New Scripting.FileSystemObject
Dim ffld As Scripting.Folder
Dim sfld As Scripting.Folder
Dim ffile As Scripting.File
Const Rootpath = "C:\Users\philippe\Documents\backupmail"
Dim path As String
Dim mi As MailItem

    
    Set roodfld = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set ffld = fs.GetFolder(Rootpath)
    
    For Each sfld In ffld.SubFolders
        For Each ffile In sfld.Files
        
        
        
            Set mi = Application.GetNamespace("MAPI").OpenSharedItem(ffile.path)
            Debug.Print sfld.Name, ffile.Name
        Next ffile
    Next sfld
    
    Shell


End Sub


