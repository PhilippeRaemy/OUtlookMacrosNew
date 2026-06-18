Attribute VB_Name = "OneDriveLocation"
Option Explicit



Public Type OneDriveLocationType
    account As String
    url As String
    localroot As String
End Type
    
Function NewOneDriveLocation(account As String, url As String, localroot As String) As OneDriveLocationType
    NewOneDriveLocation.account = account
    NewOneDriveLocation.url = VBA.LCase(Replace(Replace(url, "/", "\"), "%20", ""))
    NewOneDriveLocation.localroot = localroot
End Function


Function DiscoverLocalPathByServiceURI() As OneDriveLocationType()
    Dim wsh As IWshRuntimeLibrary.WshShell: Set wsh = New IWshRuntimeLibrary.WshShell
    Dim i As Integer
    Dim regBase As String: regBase = "HKCU\Software\Microsoft\OneDrive\Accounts\"
    
    ' Constants for Registry
    Const HKEY_CURRENT_USER = &H80000001
    Dim strKeyPath As String: strKeyPath = "Software\Microsoft\OneDrive\Accounts"
    Dim subKey As Variant
    Dim subKeys As Variant
    
    Dim serviceURI As String
    Dim localroot As String
    Dim cleanedURI As String
    Dim reg As Variant
    Dim objReg As Object
    Dim locations() As OneDriveLocationType
    
    ' 3. Loop through potential Business accounts (and Personal)
    Dim accountFolders As Variant
    
   ' 2. Connect to WMI Registry Provider
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
    ' Enumerate all subkeys under ...\OneDrive\Accounts
    objReg.EnumKey HKEY_CURRENT_USER, strKeyPath, subKeys
    If Not IsNull(subKeys) Then
        For Each subKey In subKeys:
            serviceURI = ""
            localroot = ""
            On Error Resume Next
            serviceURI = wsh.RegRead(regBase & subKey & "\ServiceEndpointURI")
            localroot = wsh.RegRead(regBase & subKey & "\UserFolder")
            On Error GoTo 0
            
            If serviceURI <> "" And localroot <> "" Then
                ' Clean the URI: Remove "/_api", swap slashes to backslashes
                cleanedURI = Replace(serviceURI, "/_api", "")
                cleanedURI = Replace(cleanedURI, "/", "\")
                ReDim Preserve locations(i)
                locations(i) = NewOneDriveLocation(CStr(subKey), cleanedURI, localroot)
                i = i + 1
                'Debug.Print subKey, cleanedURI, localroot
            End If
        Next subKey
    End If
    DiscoverLocalPathByServiceURI = locations
    
End Function

Sub printLocations()
Dim locations() As OneDriveLocationType
Dim loc As OneDriveLocationType
Dim i As Integer

    locations = DiscoverLocalPathByServiceURI()
    For i = LBound(locations) To UBound(locations)
        loc = locations(i)
        Debug.Print loc.account, loc.localroot, loc.url
    Next i
End Sub
