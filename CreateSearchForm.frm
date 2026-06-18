VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateSearchForm 
   Caption         =   "Create search folder"
   ClientHeight    =   2595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7350
   OleObjectBlob   =   "CreateSearchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateSearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim results As Integer
Dim m_ban As Boolean
Dim m_abort As Boolean

Public Function OpenDialog(SenderAddress As String, mailSubject As String, showAbort As Boolean) As Integer
Dim chunks As Variant
Dim chunk As Variant

On Error GoTo err_proc
GoTo proc

err_proc:
    Debug.Print Err.Description
    Exit Function
    Resume

proc:
    Me.SenderAddress.Clear
    Me.SenderAddress.AddItem SenderAddress
    Me.SenderAddress.Text = SenderAddress
    Me.FolderName.Clear
    chunks = Split(SenderAddress, "@")
    Me.FolderName.AddItem chunks(0)
    Me.SenderAddress.AddItem chunks(0) & "@%"
    Me.cmdAbort.Enabled = showAbort
    Me.subject.Caption = mailSubject
    If UBound(chunks) > 0 Then
        Me.SenderAddress.AddItem "%@" & chunks(1)
        Me.FolderName.Text = ""
        chunks = Split(chunks(1), ".")
        For Each chunk In chunks
            Me.SenderAddress.AddItem "%" & chunk & "%"
            Me.FolderName.AddItem chunk
        Next chunk
    Else
        Me.FolderName.Text = SenderAddress
    End If
    
    Me.Show
    OpenDialog = results
End Function

Public Property Get ToBan() As Boolean
    ToBan = m_ban
End Property

Public Property Get abort() As Boolean
    abort = m_abort
End Property

Private Sub cmdAbort_Click()
    m_abort = True
    results = vbCancel
    Me.Hide

End Sub

Private Sub cmdBan_Click()
    m_ban = True
    results = vbOK
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    m_ban = False
    results = vbOK
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    results = vbCancel
    Me.Hide
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub SenderAddress_Change()
    SenderAddress.Text = Replace(SenderAddress.Text, "*", "%")
End Sub

Private Sub UserForm_Initialize()
Dim Item As Variant

    For Each Item In Array("Forever", "1m", "3m", "6m")
        Me.Retention.AddItem Item
    Next Item
    For Each Item In Array("", "eBusiness", "ShortTerm", "Famille", "Various")
        Me.ArchiveGroup.AddItem Item
    Next Item


End Sub
