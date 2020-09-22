Attribute VB_Name = "modAVIWriter"
Public aviHandle As Long
Public VideoTimeStamps As New clsDynamicArray
Public videoIndex As New clsDynamicArray

Public Type Fraction
 D As Long
 N As Long
End Type

Public VideoFourCC As String
Public VideoWidth As Long
Public VideoHeight As Long
Public VideoFrameCount As Long
Public VideoDataSize As Long
Public VideoIndexSize As Long

Public VideoFrameRate As Fraction

Public Sub InitializeAVI()
If aviHandle <> INVALID_HANDLE_VALUE Then CloseHandle aviHandle
aviHandle = INVALID_HANDLE_VALUE
VideoWidth = 0
VideoHeight = 0
VideoFrameCount = 0
VideoDataSize = 0
Set VideoTimeStamps = New clsDynamicArray
Set videoIndex = New clsDynamicArray
End Sub

Public Function OpenAVI()
If videoDummyMode = True Then
 OpenAVI = True
 Exit Function
End If
If aviHandle <> INVALID_HANDLE_VALUE Then CloseHandle aviHandle
aviHandle = CreateFile(frmMain.txtFolder.Text & "video.avi", GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, CREATE_NEW, 0, 0)
If aviHandle = INVALID_HANDLE_VALUE Then
  message = MSG_CONFIRM_OVERWRITE
  message = Replace(message, "%FORMAT%", "video")
  message = Replace(message, "%FILE%", frmMain.txtFolder.Text & "video.avi")
  ret = MsgBox(message, vbYesNo Or vbExclamation, STR_WARNING)
  If ret = vbNo Then
   videoDummyMode = True
   OpenAVI = False
   Exit Function
  Else
   aviHandle = CreateFile(frmMain.txtFolder.Text & "video.avi", GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, CREATE_ALWAYS, 0, 0)
   If aviHandle = INVALID_HANDLE_VALUE Then
    message = MSG_NOTOPEN
    message = Replace(message, "%FORMAT%", "audio")
    message = Replace(message, "%FILE%", frmMain.txtFolder.Text & "audio.mp3")
    MsgBox message, vbOKOnly Or vbExclamation, STR_ERROR
    videoDummyMode = True
    OpenAVI = False
    Exit Function
   End If
  End If
End If
OpenAVI = True

End Function

Public Sub WriteHeaderAVI()
If videoDummyMode Then Exit Sub
Dim bWrote As Long

Dim h(0 To 55) As Long  ' AVI Header
h(0) = &H46464952       ' "RIFF"
h(1) = &H0              ' (uint)0   chunk size
h(2) = &H20495641       ' "AVI "
h(3) = &H5453494C       ' "LIST"
h(4) = &HC0             ' (uint)192
h(5) = &H6C726468       ' "hdrl"
h(6) = &H68697661       ' "avih"
h(7) = &H38             ' (uint)56
h(8) = &H0              ' (uint)0
h(9) = &H0              ' (uint)0
h(10) = &H0             ' (uint)0
h(11) = &H10            ' (uint)10
h(12) = &H0             ' (uint)0   frame count
h(13) = &H0             ' (uint)0
h(14) = &H1              ' (uint)1
h(15) = &H0             ' (uint)0
h(16) = &H0             ' (uint)0   width
h(17) = &H0             ' (uint)0   height
h(18) = &H0             ' (uint)0
h(19) = &H0             ' (uint)0
h(20) = &H0             ' (uint)0
h(21) = &H0             ' (uint)0
h(22) = &H5453494C      ' "LIST"
h(23) = &H74            ' (uint)116
h(24) = &H6C727473      ' "strl"
h(25) = &H68727473      ' "strh"
h(26) = &H38            ' (uint)56
h(27) = &H73646976      ' "vids"
If VideoFourCC = "FLV1" Then
 h(28) = &H31564C46     ' "FLV1"
Else
 h(28) = &H34564C46     ' "FLV4"
End If
h(29) = &H0             ' (uint)0
h(30) = &H0             ' (uint)0
h(31) = &H0             ' (uint)0
h(32) = &H0             ' (uint)0   frame rate denominator
h(33) = &H0             ' (uint)0   frame rate numerator
h(34) = &H0             ' (uint)0
h(35) = &H0             ' (uint)0   frame count
h(36) = &H0             ' (uint)0
h(37) = &HFFFFFFFF      ' (int)-1
h(38) = &H0             ' (int) 0
h(39) = &H0             ' 2 x (ushort)0
h(40) = &H0             ' 2 x (ushort)0 width, height
h(41) = &H66727473      ' "strf"
h(42) = &H28            ' (uint)40
h(43) = &H28            ' (uint)40
h(44) = &H0             ' (uint)0 width
h(45) = &H0             ' (uint)0 height
h(46) = &H180001         ' (ushort)1, (ushort)24
If VideoFourCC = "FLV1" Then
 h(47) = &H31564C46     ' "FLV1"
Else
 h(47) = &H34564C46     ' "FLV4"
End If
h(48) = &H0             ' (uint) 0  biSizeImage
h(49) = &H0             ' (uint) 0
h(50) = &H0             ' (uint) 0
h(51) = &H0             ' (uint) 0
h(52) = &H0             ' (uint) 0
h(53) = &H5453494C      ' "LIST"
h(54) = &H0             ' (uint) 0  chunk size
h(55) = &H69766F6D      ' "movi"
WriteFile aviHandle, h(0), 224, bWrote, ByVal 0&  '224 = 56 longs x 4 bytes each
End Sub

Public Sub CloseAVI()
CloseHandle aviHandle
aviHandle = INVALID_HANDLE_VALUE
End Sub

Public Sub WriteVideoChunk()
Dim offset As Long
Dim length As Long
Dim bWrote As Long

Dim templ As Long
Dim tempb As Byte
Dim ByteBuffer(0 To 3) As Byte
If videoDummyMode Then Exit Sub

offset = 1
If tagVideoCodec = 4 Then offset = 2
If tagVideoCodec = 5 Then offset = 5
length = tagDataSize - offset
If length < 0 Then length = 0
' add frame to frame index
videoIndex.Add (IIf(tagVideoFrameType = 1, 16, 0))
videoIndex.Add VideoDataSize + 4
videoIndex.Add length

templ = &H63643030  '00dc
WriteFile aviHandle, templ, 4, bWrote, ByVal 0&
WriteUINT (length)
WriteFile aviHandle, Buffer(offset), length, bWrote, ByVal 0&
tempb = 0
If length Mod 2 = 1 Then
 WriteFile aviHandle, tempb, 1, bWrote, ByVal 0&
 length = length + 1
End If
VideoDataSize = VideoDataSize + 8 + length
VideoFrameCount = VideoFrameCount + 1
End Sub


Public Sub WriteIndexAVI()
'On Error Resume Next
Dim templ As Long
Dim bWrote As Long

Dim indexCount As Long
Dim indexDataSize As Long
If videoDummyMode Then Exit Sub

indexDataSize = VideoFrameCount * 16
VideoIndexSize = indexDataSize + 8

templ = &H31786469      ' "idx1"
WriteFile aviHandle, templ, 4, bWrote, ByVal 0&
templ = &H63643030      '00dc
WriteUINT indexDataSize
indexCount = videoIndex.GetCount() \ 3 + 1
For i = 0 To indexCount - 1
 WriteFile aviHandle, templ, 4, bWrote, ByVal 0&
 WriteUINT videoIndex.GetValue(i * 3 + 0)
 WriteUINT videoIndex.GetValue(i * 3 + 1)
 WriteUINT videoIndex.GetValue(i * 3 + 2)
Next i

End Sub

Public Sub FinalizeAVI()
Dim templ As Long
Dim tempb1 As Byte
Dim tempb2 As Byte

If videoDummyMode Then Exit Sub
SetFilePointer aviHandle, 4, 0, FILE_BEGIN
templ = 224 + VideoDataSize + VideoIndexSize - 8
WriteUINT templ
SetFilePointer aviHandle, 32, 0, FILE_BEGIN
WriteUINT 0
SetFilePointer aviHandle, 12, 0, FILE_CURRENT
WriteUINT VideoFrameCount
SetFilePointer aviHandle, 12, 0, FILE_CURRENT
WriteUINT VideoWidth
WriteUINT VideoHeight
SetFilePointer aviHandle, 128, 0, FILE_BEGIN
' to do: determine frame rate
VideoFrameRate.D = 25
VideoFrameRate.N = 1
DetermineFrameRate

WriteUINT VideoFrameRate.N
WriteUINT VideoFrameRate.D
SetFilePointer aviHandle, 4, 0, FILE_CURRENT
WriteUINT VideoFrameCount
SetFilePointer aviHandle, 16, 0, FILE_CURRENT
tempb2 = VideoWidth \ 256
tempb1 = VideoWidth - (tempb2 * 256)
WriteFile aviHandle, tempb1, 1, bWrote, ByVal 0&
WriteFile aviHandle, tempb2, 1, bWrote, ByVal 0&
tempb2 = VideoHeight \ 256
tempb1 = VideoHeight - (tempb2 * 256)
WriteFile aviHandle, tempb1, 1, bWrote, ByVal 0&
WriteFile aviHandle, tempb2, 1, bWrote, ByVal 0&
SetFilePointer aviHandle, 176, 0, FILE_BEGIN
WriteUINT VideoWidth
WriteUINT VideoHeight
SetFilePointer aviHandle, 8, 0, FILE_CURRENT
WriteUINT VideoWidth * VideoHeight * 6
SetFilePointer aviHandle, 216, 0, FILE_BEGIN
WriteUINT VideoDataSize + 4
End Sub
'
' Works fine but it's a bit dumb. A more precise algorithm should be used
'
Public Sub DetermineFrameRate()
Dim fcount As Long
Dim fps As Double
Dim retgcd As Long
fcount = VideoTimeStamps.GetCount()
If fcount < 0 Then Exit Sub
Dim i As Long
Dim value As Long
i = 0
value = 0
While value < 1000 And i <= fcount
 value = VideoTimeStamps.GetValue(i)
 If value < 1000 Then i = i + 1
Wend
fps = i / value * 1000
'Debug.Print "i=", i, "value=", value, "fps=", fps

VideoFrameRate.D = Round(fps * 1000)
VideoFrameRate.N = 1000
' now let's try to be more accurate than that.
With VideoFrameRate
 If fps >= 5.9 And fps <= 6.1 Then
  .D = 5994: .N = 1000
 End If
 If fps > 6.1 And fps <= 6.4 Then
  .D = 625: .N = 100
 End If
 If fps > 7.4 And fps <= 7.6 Then
  .D = 75: .N = 10
 End If
 If fps > 9.9 And fps <= 10.1 Then
  .D = 10: .N = 1
 End If
 If fps > 11.9 And fps <= 11.999 Then
  .D = 11998: .N = 1000
 End If
 If fps > 12.4 And fps <= 12.6 Then
  .D = 125: .N = 10
 End If
 If fps > 14.95 And fps <= 14.99 Then
  .D = 14985: .N = 1000
 End If
 If fps > 14.99 And fps <= 15.1 Then
  .D = 15: .N = 1
 End If
 If fps > 23.9 And fps <= 23.99 Then
  .D = 23976: .N = 1000
 End If
 If fps > 23.99 And fps <= 24.2 Then
  .D = 24: .N = 1
 End If
 If fps > 24.9 And fps <= 25.2 Then
  .D = 25: .N = 1
 End If
 If fps > 29.9 And fps <= 29.99 Then
  .D = 2997: .N = 100
 End If
 If fps > 29.99 And fps <= 30.1 Then
  .D = 30: .N = 1
 End If
End With
retgcd = gcd()
VideoFrameRate.D = VideoFrameRate.D \ retgcd
VideoFrameRate.N = VideoFrameRate.N \ retgcd
Debug.Print VideoFrameRate.D, VideoFrameRate.N

End Sub

Private Function gcd() As Long
Dim a As Long
Dim b As Long
Dim r As Long
a = VideoFrameRate.D
b = VideoFrameRate.N
While b <> 0
 r = a Mod b
 a = b
 b = r
Wend
gcd = a

End Function
Public Sub WriteUINT(value As Long)
Dim bWrote As Long
Dim templ As Long
Dim tempb As Byte
Dim ByteBuffer(0 To 3) As Byte

If videoDummyMode Then Exit Sub
templ = value
ByteBuffer(3) = templ \ 16777216
templ = templ - CLng(ByteBuffer(3)) * CLng(16777216)
ByteBuffer(2) = templ \ 65536
templ = templ - CLng(ByteBuffer(2)) * CLng(65536)
ByteBuffer(1) = templ \ 256
templ = templ - CLng(ByteBuffer(1)) * CLng(256)
ByteBuffer(0) = templ
WriteFile aviHandle, ByteBuffer(0), 4, bWrote, ByVal 0&
End Sub

