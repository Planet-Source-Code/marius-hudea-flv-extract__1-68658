Attribute VB_Name = "modMP3Writer"
Public mp3Handle As Long

Public Sub InitializeMP3()
mp3Handle = INVALID_HANDLE_VALUE

End Sub

Public Function OpenMP3() As Boolean
Dim ret As Long
Dim message As String

If audioDummyMode = True Then
 OpenMP3 = True
 Exit Function
End If
If mp3Handle <> INVALID_HANDLE_VALUE Then CloseHandle mp3Handle
mp3Handle = CreateFile(frmMain.txtFolder.Text & "audio.mp3", GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, CREATE_NEW, 0, 0)
If mp3Handle = INVALID_HANDLE_VALUE Then
  message = MSG_CONFIRM_OVERWRITE
  message = Replace(message, "%FORMAT%", "audio")
  message = Replace(message, "%FILE%", frmMain.txtFolder.Text & "audio.mp3")
  ret = MsgBox(message, vbYesNo Or vbExclamation, STR_WARNING)
  If ret = vbNo Then
   audioDummyMode = True
   OpenMP3 = False
   Exit Function
  Else
   mp3Handle = CreateFile(frmMain.txtFolder.Text & "audio.mp3", GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, CREATE_ALWAYS, 0, 0)
   If mp3Handle = INVALID_HANDLE_VALUE Then
    message = MSG_NOTOPEN
    message = Replace(message, "%FORMAT%", "audio")
    message = Replace(message, "%FILE%", frmMain.txtFolder.Text & "audio.mp3")
    MsgBox message, vbOKOnly Or vbExclamation, STR_ERROR
    audioDummyMode = True
    OpenMP3 = False
    Exit Function
   End If
 End If
End If
OpenMP3 = True
End Function

Public Function CloseMP3()
If mp3Handle <> INVALID_HANDLE_VALUE Then CloseHandle mp3Handle
mp3Handle = INVALID_HANDLE_VALUE
End Function

Public Sub WriteAudioChunk()
Dim bWrote As Long
If audioDummyMode = True Then Exit Sub
'The first byte in the buffer is used by FLV format, so we don't write it to file
WriteFile mp3Handle, Buffer(1), tagDataSize - 1, bWrote, ByVal 0&
End Sub
