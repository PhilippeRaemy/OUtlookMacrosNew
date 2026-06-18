VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private SearchComplete As Scripting.Dictionary
Private WithEvents items As Outlook.items
Attribute items.VB_VarHelpID = -1
Private bannedAddresses As Variant
Private MainAccount As String

Const ConfigFolder = "C:\Users\Philippe\OneDrive - MSFT\outlook"
  
Private Type Period
    duration As Integer
    periodType As String
End Type

Private Sub Application_Startup()
Dim namespace As Outlook.namespace:    Set namespace = Outlook.Application.GetNamespace("MAPI")
    
    Set items = namespace.GetDefaultFolder(olFolderInbox).items
    MainAccount = namespace.Folders(1).name
    
    ReadBannedAddresses True

    ClassificationRuleStatic.ReadRules
End Sub

Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Outlook.Search)
    Debug.Print "The AdvancedSearchComplete Event fired: " & SearchObject.tag
    If Not SearchComplete Is Nothing Then
        If SearchComplete.Exists(SearchObject.tag) Then
            SearchComplete(SearchObject.tag) = True
        End If
    End If
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
Dim banned As Variant
  ReadBannedAddresses
  On Error GoTo ErrorHandler
  Dim Msg As Outlook.MailItem
  If TypeName(Item) = "MailItem" Then
    Set Msg = Item
    For Each banned In bannedAddresses
        If Msg.SenderEmailAddress Like banned Then
            SetCategory Msg, "Banned"
            Exit For
        End If
    Next banned
    If Not MainInRecipients(Msg.recipients) Then
        SetCategory Msg, "NotForMe"
    End If
  End If
ProgramExit:
  Exit Sub
ErrorHandler:
  MsgBox Err.Number & " - " & Err.Description
  Resume ProgramExit
End Sub
Function MainInRecipients(recipients As Outlook.recipients) As Boolean
    Dim rec As Recipient
    For Each rec In recipients
        If rec.address = MainAccount Then
            MainInRecipients = True
            Exit Function
        End If
    Next rec
End Function

Private Function ParseDurationString(testString As String) As Period
    Static regEx As RegExp
    Dim matches As Variant
    
    If regEx Is Nothing Then
        Set regEx = New RegExp
        With regEx
            .Global = False
            .IgnoreCase = True
            .Pattern = "(\d+)([dwmy])"
        End With
    End If
    
    If regEx.Test(testString) Then
        Set matches = regEx.Execute(testString)
        ParseDurationString.duration = CInt(matches(0).SubMatches(0))
        ParseDurationString.periodType = matches(0).SubMatches(1)
    End If
End Function

Sub PurgeOldMails()

Dim colStores As Outlook.Stores: Set colStores = Application.Session.Stores
Dim oStore As Outlook.Store
Dim ns As Outlook.namespace: Set ns = Application.GetNamespace("MAPI")
    
On Error GoTo err_proc
GoTo proc

err_proc:
Debug.Print Err.Description
Exit Sub
Resume

proc:
    
    For Each oStore In colStores
        PurgeOldMailsOne ns, oStore.GetRootFolder.Folders("Archives")
    Next oStore
        
End Sub

Sub PurgeOldMailsOne(ns As Outlook.namespace, TimedFolder As Outlook.MAPIFolder)
Dim retentionPeriod As Period
Dim retentionDate As Date
Dim fld As Outlook.MAPIFolder
Dim mi As MailItem
Dim mis() As String
Dim j As Integer
    
On Error GoTo err_proc
GoTo proc

err_proc:
Debug.Print Err.Description
Exit Sub
Resume

proc:
        
    retentionPeriod = ParseDurationString(TimedFolder.name)
    If retentionPeriod.periodType <> "" Then
        retentionDate = VBA.DateAdd(retentionPeriod.periodType, -retentionPeriod.duration, Now)
        Debug.Print TimedFolder.folderPath, retentionDate
        mis = MaterializeEntryIDs(TimedFolder)
        For j = UBound(mis) To LBound(mis) Step -1
            Set mi = ns.GetItemFromID(mis(j))
            If mi.SentOn < retentionDate Then
                On Error Resume Next
                mi.Delete
                If Err.Number <> 0 Then
                    Debug.Print , mi.SentOn, mi.subject, "->", Err.Description
                End If
                On Error GoTo 0
            End If
        Next j

    End If
    For Each fld In TimedFolder.Folders
        PurgeOldMailsOne ns, fld
    Next fld

End Sub

Function MaterializeEntryIDs(oFolder As Outlook.MAPIFolder) As String()
    ' Returns an array of EntryID strings instead of live MailItem objects
    Dim mobj As Object, mi As Outlook.MailItem
    Dim ids() As String
    Dim i As Integer
    
    ' Initialize the array based on folder count
    If oFolder.items.count = 0 Then
        MaterializeEntryIDs = Split(vbNullString)
        Exit Function
    End If
    ReDim ids(1 To oFolder.items.count)
    
    i = 0
    For Each mobj In oFolder.items
        ' Ensure we are only grabbing MailItems (skips ReportItems, MeetingRequests, etc.)
        If TypeOf mobj Is Outlook.MailItem Then
            Set mi = mobj
            i = i + 1
            ids(i) = mi.EntryID
        End If
    Next
    
    ' Resize to the actual number of MailItems found
    If i > 0 Then
        ReDim Preserve ids(1 To i)
        MaterializeEntryIDs = ids
    End If
End Function

Function MaterializePhysicalMails(oFolder As Outlook.MAPIFolder) As Collection
    Dim mobj As Object
    Dim physicalItem As Object
    Dim coll As New Collection
    Dim pa As Outlook.PropertyAccessor
    Dim trueEntryID As String
    
On Error GoTo err_proc
GoTo proc

err_proc:
Debug.Print Err.Description
Exit Function
Resume

proc:
        
' The MAPI property tag for the true, canonical PR_ENTRYID
    Const PR_ENTRYID As String = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
    
    If oFolder.items.count = 0 Then
        Set MaterializePhysicalMails = coll
        Exit Function
    End If
    
    For Each mobj In oFolder.items
        If TypeOf mobj Is Outlook.MailItem Then
            On Error Resume Next
            
            ' Use the PropertyAccessor to dig past the search folder's virtual wrapper
            Set pa = mobj.PropertyAccessor
            
            ' This returns the true, long-term binary EntryID converted to a hex string
            trueEntryID = pa.BinaryToString(pa.GetProperty(PR_ENTRYID))
            
            If Err.Number = 0 And trueEntryID <> "" Then
                ' Now, look up the item using its absolute, unchangeable ID
                Set physicalItem = mobj.Application.Session.GetItemFromID(trueEntryID)
                
                If Not physicalItem Is Nothing Then
                    coll.Add physicalItem
                End If
            End If
            
            ' Reset tracking variables
            Set physicalItem = Nothing
            Set pa = Nothing
            Err.Clear
            On Error GoTo 0
        End If
    Next mobj
    
    Set MaterializePhysicalMails = coll
    Set MaterializePhysicalMails = coll
End Function

Sub ClassifyMails()
    Dim colStores As Outlook.Stores
    Dim oStore As Outlook.Store
    Dim oSearchFolders As Outlook.Folders
    Dim oFolder As Outlook.MAPIFolder
    Dim FolderName As String
    Dim Retention As String
    Dim chunks As Variant
    Dim Archive As Outlook.MAPIFolder
    Dim Target As Outlook.MAPIFolder
    Dim mi As MailItem
    Dim mis() As String
    Dim i As Integer, j As Integer
    Dim mobj As Variant
    Dim ns As Outlook.namespace: Set ns = Application.GetNamespace("MAPI")
    
On Error GoTo err_proc
GoTo proc

err_proc:
Debug.Print Err.Description
Exit Sub
Resume

proc:
    Set colStores = Application.Session.Stores
    For Each oStore In colStores
        Set oSearchFolders = oStore.GetSearchFolders
        Set Archive = oStore.GetRootFolder.Folders("Archives")
        For Each oFolder In oSearchFolders
            If Not oFolder.name Like "*MS-OLK-*'" _
            And oFolder.items.count > 0 _
            Then
                Debug.Print oFolder.folderPath
                chunks = Split(oFolder.name, "|")
                FolderName = chunks(0)
                If UBound(chunks) > 1 Then
                    FolderName = Join(Array("ShortTerm", FolderName, chunks(2)), "\")
                Else
                    Retention = ""
                End If
                Set Target = EnsureFolderExists(Archive, FolderName)
                Debug.Print Target.folderPath
                Dim mailColl As Collection
                Dim mailIdx As Long
                
                ' Get the true physical items
                Set mailColl = MaterializePhysicalMails(oFolder)
                
                ' Loop backwards through our disconnected VBA collection
                For mailIdx = mailColl.count To 1 Step -1
                    Set mi = mailColl(mailIdx)
                    
                    On Error Resume Next
                    mi.Move Target
                    
                    If Err.Number <> 0 Then
                        Debug.Print Err.Number, Err.Description, mi.subject
                        Err.Clear
                    End If
                    
                    Set mi = Nothing
                    On Error GoTo err_proc
                    DoEvents
                Next mailIdx
                
                Set mailColl = Nothing
            End If
        Next oFolder
    Next oStore
    
    PurgeOldMails

End Sub

Sub CreateSearchFolderFromArchive()
Dim sch As Outlook.Search
Dim rsts As Outlook.results
Dim i As Integer
Dim tag As String
Dim fso As New FileSystemObject
Dim line As Scripting.Dictionary
Dim existingFolder As Outlook.MAPIFolder
    
    If SearchComplete Is Nothing Then Set SearchComplete = New Scripting.Dictionary
    With fso.OpenTextFile(GetSearchesFileName(), ForReading)
        While Not .AtEndOfStream
            Set line = ParseJson(.ReadLine)
            Set existingFolder = GetSearchFolder(line("name"))
            If Not existingFolder Is Nothing Then
                existingFolder.Delete
            End If
            tag = Format(Now, "yyyymmdd-hhmmss")
            SearchComplete(tag) = False
            Set sch = Application.AdvancedSearch(line("scope"), line("criterion"), True, tag)
            While SearchComplete(tag) = False
                DoEvents
            Wend
            SearchComplete.Remove tag
            Set rsts = sch.results
            Debug.Print line("name"), line("criterion"), rsts.count; " items found"
            For i = 1 To rsts.count
                Debug.Print rsts.Item(i).SenderName
            Next
            sch.Save line("name")
        Wend
    End With
End Sub

Sub CreateSearchFolderFromSender()
    Dim mi As Outlook.MailItem
    For Each mi In Application.ActiveExplorer.Selection
        CreateSearchFolderFromSenderImpl False, mi
    Next mi
End Sub
Function CreateSearchFolderFromSenderImpl(showAbort As Boolean, mi As Outlook.MailItem) As Boolean
    Dim searchName As String
    Dim finalSearchName As String
    Dim fso As New Scripting.FileSystemObject
    Dim SenderAddress As String
    Dim scope As String
    Static form As New CreateSearchForm
    Dim classRule As ClassificationRule

    On Error GoTo err_proc

    scope = "'" & Application.Session.GetDefaultFolder(olFolderInbox).folderPath & "','" & Application.Session.GetDefaultFolder(olFolderSentMail).folderPath & "'"
    
    SenderAddress = IIf(mi.SenderEmailAddress = "philippe_raemy@swissonline.ch", mi.recipients.Item(1).address, mi.SenderEmailAddress)
    
    If Not vbOK = form.OpenDialog(SenderAddress, mi.subject, showAbort) Then
        CreateSearchFolderFromSenderImpl = form.abort
        Exit Function
    End If
    If form.ToBan Then
        BanEmailAddressImpl mi
    Else
        searchName = IIf(form.ArchiveGroup.Text = "", "", form.ArchiveGroup.Text & "\") & form.FolderName
        
        If CheckRuleExists(form.SenderAddress.Text) Then
            MsgBox form.SenderAddress.Text, vbOKOnly And vbCritical, "A rule with same address exists"
        Else
    
            Set classRule = NewClassificationRule( _
                CreateAdvancedSearch(scope, form.SenderAddress.Text, searchName, form.Retention.Text), _
                form.SenderAddress.Text)
                
            With fso.OpenTextFile(GetSearchesFileName(), ForAppending, True)
                .WriteLine classRule.ToJson()
            End With
            
            ClassificationRuleStatic.AddRule classRule
        End If
    End If
    
    Exit Function

err_proc:
    Debug.Print Err.Description
End Function

Function CreateAdvancedSearch(ByVal scope As String, ByVal senderAddressText As String, ByVal initialSearchName As String, ByVal retentionText As String) As String
    Dim sch As Outlook.Search
    Dim tag As String
    Dim addresses As Variant
    Dim address As Variant
    Dim criterions() As String, c As Integer
    Dim criterion As String
    Dim finalSearchName As String

    addresses = Split(senderAddressText, ";")
    ReDim criterions(3 * UBound(addresses) + 2)
    c = 0
    For Each address In addresses
        address = Replace(address, "*", "%")
        criterions(c) = "urn:schemas:httpmail:fromemail like '%" & address & "%'"
        criterions(c + 1) = "urn:schemas:httpmail:displayto like '%" & address & "%'"
        criterions(c + 2) = "urn:schemas:mailheader:to like '%" & address & "%'"
        c = c + 3
    Next address
    criterion = Join(criterions, " or ")

    tag = Format(Now, "yyyymmdd-hhmmss")
    If SearchComplete Is Nothing Then Set SearchComplete = New Scripting.Dictionary
    SearchComplete(tag) = False
    
    Set sch = Application.AdvancedSearch(scope, criterion, True, tag)
    
    While SearchComplete(tag) = False
        DoEvents
    Wend
    SearchComplete.Remove tag

    finalSearchName = GetSearchFoldersUniqueName(initialSearchName, IIf(retentionText = "Forever", "", retentionText))
    sch.Save finalSearchName
    
    CreateAdvancedSearch = finalSearchName
End Function


Function GetSearchFoldersUniqueName(name As String, Retention As String) As String
    Dim colStores As Outlook.Stores
    Dim oStore As Outlook.Store
    Dim oSearchFolders As Outlook.Folders
    Dim oFolder As Outlook.MAPIFolder
    Dim i As Integer
    Dim chunks As Variant
    Dim found As Boolean
    
  
    Set colStores = Application.Session.Stores
    For Each oStore In colStores
    Set oSearchFolders = oStore.GetSearchFolders
        For Each oFolder In oSearchFolders
            chunks = Split(oFolder.name, "|")
            If chunks(0) = name Then
                found = True
                If UBound(chunks) > 0 Then
                    If IsNumeric(chunks(1)) Then
                        If CInt(chunks(1)) >= i Then
                            i = CInt(chunks(1)) + 1
                        End If
                    End If
                End If
            End If
        Next
    Next
    If found And i = 0 Then i = 1 ' found unnumbered folder name
    If Retention = "" Then
        GetSearchFoldersUniqueName = VBA.Join(Array(name, CStr(i)), "|")
    Else
        GetSearchFoldersUniqueName = VBA.Join(Array(name, CStr(i), Retention), "|")
    End If
End Function
Function GetSearchFolder(name As String) As Outlook.MAPIFolder
    Dim colStores As Outlook.Stores
    Dim oStore As Outlook.Store
    Dim oSearchFolders As Outlook.Folders
    Dim oFolder As Outlook.MAPIFolder
  
    Set colStores = Application.Session.Stores
    For Each oStore In colStores
        Set oSearchFolders = oStore.GetSearchFolders
        For Each oFolder In oSearchFolders
            Debug.Print oFolder.name, name
            If oFolder.name = name Then
                Set GetSearchFolder = oFolder
                Exit Function
            End If
        Next
    Next
    Set GetSearchFolder = Nothing
End Function

Public Function GetBannedFileName() As String
    Dim fso As New Scripting.FileSystemObject
    GetBannedFileName = fso.BuildPath(ConfigFolder, "OutlookBanned.txt")
End Function

Public Function GetSearchesFileName() As String
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    GetSearchesFileName = fso.BuildPath(ConfigFolder, "OutlookSearches.jsonl")
End Function


Public Sub ReadBannedAddresses(Optional force As Boolean)

    If Not force And Not IsEmpty(bannedAddresses) Then Exit Sub
    
    Dim fso As New Scripting.FileSystemObject
    Dim a(0) As String
    Dim li As Integer
    Dim home As String: home = GetBannedFileName()
    If fso.FileExists(home) Then
        With fso.OpenTextFile(home, ForReading, False)
            bannedAddresses = Split(.ReadAll, vbCrLf)
        End With
        Exit Sub
    End If
    bannedAddresses = Array()
End Sub

Public Sub AppendBannedAddress(address As String)
    Dim fso As New Scripting.FileSystemObject
    With fso.OpenTextFile(GetBannedFileName(), ForAppending, True)
        .WriteLine address
    End With
    bannedAddresses = Array()
End Sub

Public Sub BanEmailAddress()
Dim mi As Outlook.MailItem
    
    For Each mi In Application.ActiveExplorer.Selection
        BanEmailAddressImpl mi
    Next mi
End Sub

Public Sub BanEmailAddressImpl(mi As MailItem)
Dim address As String
Dim bannedAddress As Variant
Dim known As Boolean
Dim obj As Variant
Dim allMi As Outlook.MailItem

    
    address = InputBox("Addresses like:", "Ban this guy!", mi.SenderEmailAddress)
    If address <> "" Then
        ReadBannedAddresses
        For Each bannedAddress In bannedAddresses
            If bannedAddress = address Then
                known = True
                Exit For
            End If
        Next bannedAddress
        If Not known Then
            AppendBannedAddress address
            ReadBannedAddresses True
        End If
    
        For Each obj In Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).items
            If TypeName(obj) = "MailItem" Then
                Set allMi = obj
                If allMi.SenderEmailAddress Like address Then
                    Debug.Print allMi.ReceivedTime, allMi.subject
                    SetCategory allMi, "Banned"
                    allMi.Save
                End If
            End If
        Next obj
    End If

End Sub

Public Sub DefineFromInbox()
Dim obj As Variant
Dim allMi As Outlook.MailItem
Dim bannedAddress As Variant
Dim covered As Boolean
Dim abort As Boolean
Dim address As String
        
    ReadBannedAddresses True
    
    
    For Each obj In Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).items
        If TypeName(obj) = "MailItem" Then
            Set allMi = obj
            address = allMi.SenderEmailAddress
            covered = False
            SelectMailItem allMi
            If AnyRuleMatch(address) Then
                Debug.Print address & " is covered by a rule"
                covered = True
            Else
                For Each bannedAddress In bannedAddresses
                    If address Like bannedAddress Then
                        Debug.Print address & " is already banned"
                        covered = True
                        Exit For
                    End If
                Next bannedAddress
            End If
            If Not covered Then
                Debug.Print "Making a decision for " & address
                allMi.Display
                DoEvents
                abort = CreateSearchFolderFromSenderImpl(True, allMi)
                allMi.Close olDiscard
            End If
        End If
        If abort Then Exit Sub
    Next obj
End Sub


Public Sub SetCategory(mi As MailItem, category As String)
Dim categories As Variant, cat As Variant
    categories = Split(mi.categories, ";")
    If UBound(categories) < 0 Then
        mi.categories = category
    Else
        For Each cat In categories
            If cat = category Then Exit Sub
        Next cat
        mi.categories = mi.categories & ";" & category
        mi.Save
    End If
End Sub

Public Sub ScanInbox()
Dim obj As Variant
Dim mi As Outlook.MailItem
    ReadBannedAddresses
    For Each obj In Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).items
        If TypeName(obj) = "MailItem" Then
            Items_ItemAdd obj
        End If
    Next obj
End Sub

Public Function EnsureFolderExists(ByVal parentFolder As Outlook.folder, ByVal folderPath As String) As Outlook.folder
Dim a As Variant
Dim i As Integer
Dim parentPath As String

Set EnsureFolderExists = parentFolder

a = Split(folderPath, "\")
For i = 0 To UBound(a)
    If a(i) <> "" Then
        a(i) = Replace(a(i), "/", "_")
        On Error Resume Next
        Set EnsureFolderExists = EnsureFolderExists.Folders(a(i))
        If Err.Number <> 0 Then
            Set EnsureFolderExists = EnsureFolderExists.Folders.Add(a(i))
            Err.Clear
            DoEvents
        End If
        On Error GoTo 0
    End If
Next i
End Function
Sub ArchiveOldMails()
Dim fld As Outlook.folder
Dim mi As Outlook.MailItem
    For Each fld In Application.GetNamespace("MAPI").Folders
        ArchiveOldMails_sub fld, "C:\Users\Philippe\OneDrive - MSFT\Mails", CDate("2025-01-01")
    Next fld

End Sub

Sub ArchiveOldMails_sub(rootFolder As Outlook.folder, saveFolder As String, threshold As Date)
Dim fld As Outlook.folder
Dim obj As Variant
Dim mi As Outlook.MailItem
Dim ci As Outlook.MeetingItem
Dim items As Scripting.Dictionary
Dim i As Integer
    Debug.Print "Exploring folder " & rootFolder.folderPath
    For Each obj In rootFolder.items
        Select Case TypeName(obj)
            Case "MailItem":
                Set mi = obj
                If mi.SentOn < threshold Then
                    Debug.Print "Found mail item from " & mi.SenderEmailAddress & " : " & mi.subject
                    If mi.Attachments.count > 0 Then
                    Else
                    End If
                End If
            Case "MeetingItem":
                Set ci = obj
                If ci.SentOn < threshold Then
                    Debug.Print "Found Meeting item from " & ci.SenderEmailAddress & " : " & ci.subject
                Else
                    Debug.Print "Unknown item time found: " & TypeName(obj)
                End If
        End Select
    Next obj
    For Each fld In rootFolder.Folders
        ArchiveOldMails_sub fld, saveFolder & "\" & fld.name, threshold
    Next fld

        
End Sub

Function EnsureOsFolderExists(path As String) As Scripting.folder
Dim chunk As Variant
Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
Dim partial As String

    For Each chunk In VBA.Split(path, "\")
        partial = IIf(partial = "", "", partial & "\") & chunk
        If Not fso.FolderExists(partial) Then
            Set EnsureOsFolderExists = fso.CreateFolder(partial)
        End If
    Next chunk
End Function

Sub SelectMailItem(oMail As Outlook.MailItem)
    Dim oExplorer As Outlook.Explorer
    Dim oFolder As Outlook.MAPIFolder
    
    ' 1. Get the explorer
    Set oExplorer = Application.ActiveExplorer
    
    ' 2. Navigate to the folder containing the item
    ' Important: The item must be in the active view to be selected
    Set oFolder = oMail.Parent
    If oExplorer.CurrentFolder <> oFolder Then
        Set oExplorer.CurrentFolder = oFolder
    End If
    
    ' 3. Select the item
    ' Note: To "select" only one item, you focus on it.
    ' There is no direct "Selection.Add" method.
    ' The workaround is to use the CurrentFolder.Display or simply
    ' set the CurrentItem (if the inspector is open).
    ' For the Explorer, we focus the view:
    oExplorer.ClearSelection
    oExplorer.AddToSelection oMail
End Sub

Sub DefineFromInbox_old()
Dim colStores As Outlook.Stores: Set colStores = Application.Session.Stores
Dim oStore As Outlook.Store
Dim mobj As Object
Dim mi As MailItem

On Error GoTo err_proc
GoTo proc

err_proc:
Debug.Print Err.Description
Exit Sub
Resume

proc:
    For Each oStore In colStores
        For Each mobj In oStore.GetRootFolder.Folders("Inbox").items
            If TypeOf mobj Is Outlook.MailItem Then
                Set mi = mobj
                If Not AnyRuleMatch(mi.SenderEmailAddress) Then
                    If InStr(1, mi.categories, "banned", vbTextCompare) = 0 Then
                        SelectMailItem mi
                        CreateSearchFolderFromSenderImpl True
                    End If
                End If
            End If
        Next mobj
    Next oStore

End Sub
