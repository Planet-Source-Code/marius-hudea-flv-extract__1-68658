Attribute VB_Name = "modTimeStampWriter"
Private txtHandle As Long

Public Sub InitializeTS()
txtHandle = INVALID_HANDLE_VALUE
End Sub

Public Sub OpenTS()
On Error Resume Next
Dim ret As Long
Dim s As String
If timestampDummyMode Then Exit Sub
If txtHandle <> INVALID_HANDLE_VALUE Then Close #txtHandle
txtHandle = FreeFile
Err.Clear
txtHandle = CreateFile(frmMain.txtFolder.Text & "timestamps.txt", GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, CREATE_NEW, 0, 0)
If txtHandle = INVALID_HANDLE_VALUE Then
 message = MSG_CONFIRM_OVERWRITE
 message = Replace(message, "%FORMAT%", "timestamp")
 message = Replace(message, "%FILE%", frmMain.txtFolder.Text & "timestamp.txt")
 ret = MsgBox(message, vbYesNo Or vbExclamation, STR_WARNING)
 If ret = vbNo Then
  timestampDummyMode = True
  Exit Sub
 Else
  txtHandle = CreateFile(frmMain.txtFolder.Text & "timestamps.txt", GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, CREATE_ALWAYS, 0, 0)
  If txtHandle = INVALID_HANDLE_VALUE Then
     message = MSG_NOTOPEN
     message = Replace(message, "%FORMAT%", "timestamp")
     message = Replace(message, "%FILE%", frmMain.txtFolder.Text & "timestamp.txt")
     MsgBox message, vbOKOnly Or vbExclamation, STR_ERROR
     timestampDummyMode = True
     Exit Sub
  End If
 End If
 s = "# timecode format v2" & vbCrLf
 WriteFile txtHandle, ByVal s, Len(s), ret, ByVal 0&
End If
  
End Sub

Public Sub WriteTS(value As Long)
If timestampDummyMode Then Exit Sub
If txtHandle = INVALID_HANDLE_VALUE Then Exit Sub
Dim s As String
Dim bWrote As Long
s = CStr(value) & vbCrLf
WriteFile txtHandle, ByVal s, Len(s), bWrote, ByVal 0&

'Print #txtHandle, Value
End Sub


Public Sub closeTS()
If timestampDummyMode Then Exit Sub
CloseHandle txtHandle
txtHandle = INVALID_HANDLE_VALUE
End Sub
