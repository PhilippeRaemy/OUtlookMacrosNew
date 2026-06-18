Attribute VB_Name = "ExtractAttachements"
Option Explicit
Dim fso As New Scripting.FileSystemObject
Sub main()
  Dim stArchive As Store
  Set stArchive = openStore("C:\Users\Philippe\Documents\Outlook Files\Personal Archives.pst")
  ExtractFiles "d:\Users\Philippe\Documents\CargillMailArchives", stArchive.GetRootFolder()
End Sub

Function GetFolder(path As String) As Scripting.Folder
    Dim subFolderName As String
    If fso.FolderExists(path) Then
        Set GetFolder = fso.GetFolder(path)
    Else
        subFolderName = Mid(path, Len(fso.GetParentFolderName(path)) + 1)
        If Left(subFolderName, 1) = "\" Then subFolderName = Mid(subFolderName, 2)
        Set GetFolder = GetFolder(fso.GetParentFolderName(path)).SubFolders.Add(subFolderName)
    End If
End Function
Function openStore(archiveFileName As String) As Store
  Dim myNameSpace As NameSpace, st As Store
  Set myNameSpace = Application.GetNamespace("MAPI")
  myNameSpace.AddStore archiveFileName
  Set openStore = myNameSpace.Stores(myNameSpace.Stores.Count)
  Debug.Print "Store " & openStore.filepath & " is open."
End Function

Sub ExtractFiles(path As String, oFolder As Outlook.Folder, Optional SaveAttachements As Boolean = True)
  
  Dim mi As MailItem, subfld As Outlook.Folder, obj As Object
  Dim ai As AppointmentItem
  Dim att As Attachment, atts As Attachments
  Dim FileName As String, fileRoot As String
  Dim fFolder As Scripting.Folder
  Set fFolder = GetFolder(path)
  Dim i As Integer, j As Integer
  On Error GoTo err_proc
  GoTo proc
err_proc:
  Debug.Print "Error " & Err.Number & ", " & Err.Description & vbCrLf & TypeName(obj) & vbCrLf & mi.Parent.folderPath & "\" & mi.subject & " - " & mi.SentOn
  Resume Next
proc:
  For Each obj In oFolder.Items
    Set atts = Nothing
    Select Case TypeName(obj)
      Case "MailItem"
        Set mi = obj
        fileRoot = fFolder.path & "\" & Format(mi.SentOn, "yyyymmdd_hhmmss") & "_" & CleanName(mi.subject)
        If SaveAttachements Then
          For Each att In mi.Attachments
              SaveAttachment att, fileRoot
          Next att
        End If
        On Error GoTo err_proc
        mi.SaveAs Left(fileRoot, 251) & ".msg", olMSG
      Case "AppointmentItem"
        Set ai = obj
        fileRoot = fFolder.path & "\" & Format(ai.Start, "yyyymmdd_hhmmss") & "_" & CleanName(ai.subject)
        If SaveAttachements Then
          For Each att In ai.Attachments
              SaveAttachment att, fileRoot
          Next att
        End If
        ai.SaveAs fileRoot & ".msg", olMSG
    End Select
    
  Next obj
  For Each subfld In oFolder.folders
    ExtractFiles path & "\" & CleanName(subfld.name), subfld
  Next subfld
End Sub
Function CleanName(name As String) As String
  Const undesired = ":\/?*<>|&""”“+%!"
  Dim i As Integer
  CleanName = name
  For i = 1 To Len(undesired)
    CleanName = Replace(CleanName, Mid(undesired, i, 1), "_")
  Next i
  For i = 1 To 31
    CleanName = Replace(CleanName, Chr(i), "_")
  Next i
End Function
Sub SaveAttachment(att As Attachment, fileRoot As String)
    Dim FileName As String
    Dim FileSuffix As String
    Dim NameParts As Variant, Extension As String
    On Error Resume Next
    FileSuffix = "_" & CleanName(att.FileName)
    If Err.Number <> 0 Then
      Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    Dim j As Integer
    FileName = fileRoot & FileSuffix
    NameParts = Split(FileName, ".")
    If UBound(NameParts) = 0 Then
      Extension = ""
    Else
      Extension = "." & NameParts(UBound(NameParts))
      FileName = Left(FileName, Len(FileName) - Len(Extension))
    End If
    If Len(FileName) + Len(Extension) > 250 Then
      FileName = Left(FileName, 250 - Len(Extension))
    End If
    j = 1
    fileRoot = FileName
    While fso.FileExists(FileName + Extension)
        j = j + 1
        FileName = fileRoot & "(" & j & ")"
    Wend
    ' Debug.Print " ==> " & FileName + Extension
    att.SaveAsFile FileName + Extension
End Sub
