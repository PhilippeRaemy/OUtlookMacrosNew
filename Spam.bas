Attribute VB_Name = "Spam"
Dim spammers As Scripting.Dictionary
Const JunkFolderRootName = "Inbox\shortTerm\Junk E-mail{%1}"

Private Sub initSpammersList()
Dim fs As FileSystemObject
Dim ts As TextStream
On Error GoTo proc_err
GoTo proc
proc_err:
  MsgBox Err.Number & " " & Err.Description & " in initSpammersList", vbCritical
  Exit Sub
  Resume
proc:

  If spammers Is Nothing Then
    Set spammers = New Scripting.Dictionary
    spammers.CompareMode = TextCompare
    Set fs = New FileSystemObject
    Set ts = fs.OpenTextFile(Environ("USERPROFILE") & "\Local Settings\Application Data\Microsoft\Outlook\spammers.txt", ForReading)
    While Not ts.AtEndOfStream
      spammers(ts.ReadLine) = Empty
    Wend
    ts.Close
    Set ts = Nothing
    Set fs = Nothing
  End If
End Sub
Private Sub SaveSpammersList()
Dim fs As FileSystemObject
Dim ts As TextStream
Dim k As Variant
  If spammers Is Nothing Then Exit Sub
  Set fs = New FileSystemObject
  Set ts = fs.OpenTextFile(Environ("USERPROFILE") & "\Local Settings\Application Data\Microsoft\Outlook\spammers.txt", ForWriting, Create:=True)
  ts.WriteLine "/* file created on " & Format(Now, "yyyy-mm-dd hh:mm:ss") & " */"
  For Each k In spammers.Keys
    ts.WriteLine k
  Next k
  ts.WriteLine "/* end of file */"
  ts.Close
  Set ts = Nothing
  Set fs = Nothing
End Sub

Public Sub MakeSpamMail(item As MailItem)
On Error GoTo proc_err
GoTo proc
proc_err:
  MsgBox Err.Number & " " & Err.Description & " in MakeSpamMail", vbCritical
  Exit Sub
  Resume
proc:
AddJunkSearchFolder item.SenderEmailAddress
'Dim AlreadyFound As Boolean
'  initSpammersList
'  For i = 0 To spammers.Count - 1
'    If InStr(1, Item.SenderEmailAddress, spammers.Keys(i), vbTextCompare) > 0 Then
'      AlreadyFound = True
'      Exit For
'    End If
'  Next i
'  If Not AlreadyFound Then
'    spammers.Add Item.SenderEmailAddress, Empty
'    SaveSpammersList
'  End If
'  HandleIncomingMails Item
End Sub

Public Sub HandleIncomingMails(item As MailItem)
Dim obj As Object
Dim rObj As ReportItem
Dim mObj As MailItem
Dim i As Integer
Dim fld As Outlook.Folder
Dim fs As New FileSystemObject
Dim ts As TextStream
Dim Msg As String

On Error GoTo proc_err
GoTo proc
proc_err:
  If Err.Number = -2147352567 Then
    Resume Next
  Else
    MsgBox Err.Number & " " & Err.Description & " in HandleIncomingMails", vbCritical
    Exit Sub
  End If
  Resume
proc:


  trace.trace "From «" & item.SenderEmailAddress & "»: " & item.subject
  initSpammersList
  For i = 0 To spammers.Count - 1
    If InStr(1, item.SenderEmailAddress, spammers.Keys(i), vbTextCompare) > 0 Then
      If InStr(1, item.SenderEmailAddress, "cargill.com", vbTextCompare) > 0 Then
        Set fld = Application.Session.GetDefaultFolder(olFolderInbox).folders("Various")
      Else
        Set fld = Application.Session.GetDefaultFolder(olFolderJunk)
      End If
      item.Move fld
      Msg = Format(Now, "YYYY-MM-DD hh:mm:ss") & " Move «" & item.subject & "»" & _
        " from «" & item.Parent.folderPath & "»" & _
        " to " & fld.folderPath
      Set ts = fs.OpenTextFile(Environ("USERPROFILE") & "\Local Settings\Application Data\Microsoft\Outlook\spammers.log", ForAppending, True)
      ts.WriteLine Msg
      ts.Close
      trace.trace Msg
      Exit For
    End If
  Next i
End Sub
Public Function AddJunkSearchFolder(mailAddress As String) As Boolean
Dim oStore As Store
Dim primaryStore As Store
Dim mailDomain As String
Dim JunkFolderName As String
Dim SearchFld As Folder
Dim scope As String
Dim searchresult As Search
On Error GoTo proc_err
GoTo proc
proc_err:
  MsgBox Err.Number & " " & Err.Description & " in AddJunkSearchFolder", vbCritical
  Exit Function
  Resume
proc:

  If InStr(mailAddress, "@") > 0 Then
    mailDomain = Mid(mailAddress, InStr(mailAddress, "@") + 1)
  Else
    mailDomain = mailAddress
  End If
  JunkFolderName = Replace(JunkFolderRootName, "%1", mailDomain)

  For Each oStore In Application.Session.Stores
  
    If oStore.ExchangeStoreType = olPrimaryExchangeMailbox Then
      Set primaryStore = oStore
      Set oSearchFolders = oStore.GetSearchFolders
      For Each SearchFld In oSearchFolders
        If SearchFld.name Like JunkFolderName & "*" Then
          AddJunkSearchFolder = False
          Exit Function
        End If
      Next
    End If
  Next
  'If arrived there, we've not found the search folder: create on the main store
  scope = "'" & Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) & "'"
  Set searchresult = Application.AdvancedSearch(scope, """urn:schemas:httpmail:from"" like '%@" & mailDomain & "'", False)
  searchresult.save JunkFolderName
  AddJunkSearchFolder = True
  
End Function
Sub DisplayAvailableScopes()

    'Declare a variable that references a
    'SearchScope object.
    Dim ss As SearchScope
    Dim sss As SearchScopes

        'Loop through the SearchScopes collection.
        For Each ss In sss
            Select Case ss.Type
                Case msoSearchInMyComputer
                    MsgBox "My Computer is an available search scope."
                Case msoSearchInMyNetworkPlaces
                    MsgBox "My Network Places is an available search scope."
                Case msoSearchInOutlook
                    MsgBox "Outlook is an available search scope."
                Case msoSearchInCustom
                    MsgBox "A custom search scope is available."
                Case Else
                    MsgBox "Can't determine search scope."
            End Select
        Next ss

End Sub
Sub initJunkSearchFolders()
AddJunkSearchFolder "actuate.com"
AddJunkSearchFolder "adobe.com"
AddJunkSearchFolder "airdefense.net"
AddJunkSearchFolder "altigenweb-mail.info"
AddJunkSearchFolder "altigenwebmail.info"
AddJunkSearchFolder "angel.com"
AddJunkSearchFolder "angel.com"
AddJunkSearchFolder "announcements.informatica-news.com"
AddJunkSearchFolder "ArchitectureSummit.net"
AddJunkSearchFolder "asaaaa.com"
AddJunkSearchFolder "ashley.taylor@shunra.com"
AddJunkSearchFolder "castsoftware.com"
AddJunkSearchFolder "cavisualdesign-mail.info"
AddJunkSearchFolder "ccpguides-mails.info"
AddJunkSearchFolder "centrifugesystems.com"
AddJunkSearchFolder "communicatevisually.com"
AddJunkSearchFolder "communicatevisually.com"
AddJunkSearchFolder "connect.vmware.com"
AddJunkSearchFolder "creditcardprocessguides.info"
AddJunkSearchFolder "cybercartes-mail.com"
AddJunkSearchFolder "db.nl00.net"
AddJunkSearchFolder "defensepactom.com"
AddJunkSearchFolder "dkpromo-mail.info"
AddJunkSearchFolder "docucrunch.com"
AddJunkSearchFolder "DocuCrunch.com"
AddJunkSearchFolder "eiqnetworks.com"
AddJunkSearchFolder "elastra.com"
AddJunkSearchFolder "en25.com"
AddJunkSearchFolder "FinanceTechNews.com"
AddJunkSearchFolder "FinanceTechNews.com"
AddJunkSearchFolder "FinanceTechNews.com"
AddJunkSearchFolder "focus-erpmail.info"
AddJunkSearchFolder "focuscrmmail.info"
AddJunkSearchFolder "focusvoipguides.info"
AddJunkSearchFolder "hardwarecity-mail.info"
AddJunkSearchFolder "i-speak-mail.info"
AddJunkSearchFolder "info.newscale.com"
AddJunkSearchFolder "infosys.com"
AddJunkSearchFolder "interwoven.com"
AddJunkSearchFolder "jgs-dom-notification.com"
AddJunkSearchFolder "mail.communications.sun.com"
AddJunkSearchFolder "mail.vresp.com"
AddJunkSearchFolder "mail.vresp.com"
AddJunkSearchFolder "messagelabs.com"
AddJunkSearchFolder "mindtree.com"
AddJunkSearchFolder "mindtree.com"
AddJunkSearchFolder "morecrm-mails.info"
AddJunkSearchFolder "netapp.com"
AddJunkSearchFolder "nonewsletter.resaplus.ch"
AddJunkSearchFolder "nosonicwall.com"
AddJunkSearchFolder "noverizonwireless.com"
AddJunkSearchFolder "offers.ztfsg.com"
AddJunkSearchFolder "omniture.com"
AddJunkSearchFolder "omniture.com"
AddJunkSearchFolder "omniture.com"
AddJunkSearchFolder "onhold-companymail.info"
AddJunkSearchFolder "onholdco-mail.info"
AddJunkSearchFolder "optier.marketbright.com"
AddJunkSearchFolder "osibusinessmail.info"
AddJunkSearchFolder "owireless-mails.info"
AddJunkSearchFolder "pbp-executivereports.net"
AddJunkSearchFolder "pbpmedia.com"
AddJunkSearchFolder "pbpmedia.com"
AddJunkSearchFolder "pbtechnologytraining.com"
AddJunkSearchFolder "pbtechnologytraining.com"
AddJunkSearchFolder "pdb33.info"
AddJunkSearchFolder "polaris.co.in"
AddJunkSearchFolder "polaris.com"
AddJunkSearchFolder "progressivebusinesstechnologytraining.com"
AddJunkSearchFolder "rapidresponsemarketinginc.com"
AddJunkSearchFolder "reply.informatica-news.com"
AddJunkSearchFolder "reply.mb00.net"
AddJunkSearchFolder "sgi.com"
AddJunkSearchFolder "shunra.com"
AddJunkSearchFolder "smartdraw.com"
AddJunkSearchFolder "smartdrawcommunity.com"
AddJunkSearchFolder "smartdrawcommunity.com"
AddJunkSearchFolder "smartdrawinfo.com"
AddJunkSearchFolder "smartdrawinfo.com"
AddJunkSearchFolder "spl03.net"
AddJunkSearchFolder "ssimpson@layer7tech.com"
AddJunkSearchFolder "systemsinmotion.com"
AddJunkSearchFolder "targetedconferences.com"
AddJunkSearchFolder "targetedconferences.com"
AddJunkSearchFolder "tp.omnichannel.net"
AddJunkSearchFolder "trendmicro.rsys1.com"
AddJunkSearchFolder "trythenewsilktest@microfocus.com"
AddJunkSearchFolder "verizonwireless.com"
AddJunkSearchFolder "vietnamam.com"
AddJunkSearchFolder "vinmails.info"
AddJunkSearchFolder "voipguidemail.info"
End Sub
