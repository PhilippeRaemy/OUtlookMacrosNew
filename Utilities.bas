Attribute VB_Name = "Utilities"
Global trace As New Tracer
Sub moveItem(miv As Variant, fld As Outlook.Folder, context As String)
On Error GoTo proc_err
GoTo proc
proc_err:
  Dim Msg As String
  Dim errNb As Long
  errNb = Err.Number
  Msg = Err.Number & " " & Err.Description & " in Utilities.moveItem"
  trace.trace "ERROR", Msg
  If errNb <> -2147221233 Then MsgBox Msg, vbCritical
  Exit Sub
  Resume
proc:
  If fld Is Nothing Then
    trace.trace context, "DELETED", miv.subject & "(" & miv.CreationTime & ")"
    trace.trace context, "FROM-->", miv.Parent.folderPath
    miv.Delete
  ElseIf miv.Parent.folderPath = fld.folderPath Then
    trace.trace context, "NOT MOVED ", miv.subject & "(" & miv.CreationTime & ")"
    trace.trace context, "ON SAME FOLDER->", miv.Parent.folderPath
  Else
    trace.trace context, "MOVED ", miv.subject & "(" & miv.CreationTime & ")"
    trace.trace context, "FROM->", miv.Parent.folderPath
    trace.trace context, "--->TO", fld.folderPath
    miv.Move fld
  End If
End Sub
