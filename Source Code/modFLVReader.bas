Attribute VB_Name = "modFLVReader"
Private flvHandle As Long

Public smallBuffer(0 To 1023) As Byte           'Buffer for various read needs
Public Buffer() As Byte                         'Resizable buffer
Public BufferLen As Long
Public BufferUseLen As Long

Private bRead As Long
Private flvOffset As Long
Private flvLength As Long

Private flvFlags As Byte
Private flvDataOffset As Long

Public flvFlagHasVideo As Boolean
Public flvFlagHasAudio As Boolean

Public flvTagsCount As Long
Public flvTagsVideo As Long
Public flvTagsAudio As Long
Public flvTagsScript As Long

Public flvTagVideoI As Long
Public flvTagVideoK As Long
Public flvTagVideoD As Long


Public videoDummyMode As Boolean
Public audioDummyMode As Boolean
Public timestampDummyMode As Boolean
Public videoInit As Boolean
Public audioInit As Boolean
Public timestampInit As Boolean

Public tagType As Long
Public tagDataSize As Long
Public tagTimeStamp As Long
Public tagStreamID As Long

Public tagAudioFormat As Long
Public tagAudioRate As Long
Public tagAudioBits As Long
Public tagAudioChannels As Long

Public tagVideoFrameType As Long
Public tagVideoCodec As Long
Public tagVideoPictureSize As Long  'h263 specific

Public VideoFrameRateCaption As String
Public AudioFormatCaption As String


Public Function OpenFLV(FileName As String) As Boolean
Dim Signature As String
Dim DataOffset As Long
Dim tempCurr As Currency

InitializeFLV
flvHandle = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)
' Could not open the FLV file. Abort...
If flvHandle = INVALID_HANDLE_VALUE Then
 OpenFLV = False: Exit Function
End If
flvOffset = 0
'Determine File Size
GetFileSizeEx flvHandle, tempCurr
flvLength = tempCurr * 10000
'Read Signature
ReadFile flvHandle, smallBuffer(0), 4, bRead, ByVal 0&
flvOffset = flvOffset + 4
Signature = Chr(smallBuffer(0)) + _
            Chr(smallBuffer(1)) + _
            Chr(smallBuffer(2))
If bRead <> 4 Or Signature <> "FLV" Then
' Couldn't even read the 4 byte signature or the signature is incorrect
 CloseHandle flvHandle: flvHandle = -1: OpenFLV = False: Exit Function
End If
If smallBuffer(3) <> 1 Then
' Only first version of FLV supported right now
 MsgBox STR_INVALIDVERSION, vbOKOnly Or vbCritical, STR_ERROR
 CloseHandle flvHandle: flvHandle = INVALID_HANDLE_VALUE: OpenFLV = False: Exit Function
End If
' Read Flags byte
ReadFile flvHandle, flvFlags, 1, bRead, ByVal 0&
flvOffset = flvOffset + 1
' Flag Format 00000101
' Bit 5: Has Audio Tags;
' Bit 7: Has Video Tags;
' Bit 0-4,6: Reserved
If flvFlags = 4 Or flvFlags = 5 Then flvFlagHasAudio = True
If flvFlags = 1 Or flvFlags = 5 Then flvFlagHasVideo = True
frmMain.chkFlagVideo.value = IIf(flvFlagHasVideo, vbChecked, vbUnchecked)
frmMain.chkFlagAudio.value = IIf(flvFlagHasAudio, vbChecked, vbUnchecked)
' Read Data Offset
ReadFile flvHandle, smallBuffer(0), 4, bRead, ByVal 0&
flvOffset = flvOffset + 4
'ReverseBuffer 4     ' change byte order from 3210 to 0123
flvDataOffset = smallBuffer(0) * 16777216 + _
                smallBuffer(1) * 65536 + _
                smallBuffer(2) * 256 + _
                smallBuffer(3)
SetFilePointer flvHandle, flvDataOffset, 0, FILE_BEGIN
flvOffset = flvDataOffset
OpenFLV = True
End Function

Public Sub CloseFLV()
If flvHandle <> INVALID_HANDLE_VALUE Then
 CloseHandle flvHandle
 flvHandle = INVALID_HANDLE_VALUE
End If

End Sub

' Initializes the FLV reader and resets all the variables used by the FLV reader
Public Sub InitializeFLV()
ReDim Buffer(0 To 4095)
BufferLen = 4096
CloseHandle flvHandle
flvHandle = INVALID_HANDLE_VALUE
flvOffset = 0
flvDataOffset = 0
flvLength = 0
flvFlags = 0
flvFlagHasVideo = False
flvFlagHasAudio = False
flvTagsCount = 0
flvTagsVideo = 0
flvTagsAudio = 0
flvTagsScript = 0
flvTagVideoI = 0
flvTagVideoK = 0
flvTagVideoD = 0

videoDummyMode = False
audioDummyMode = False
timestampDummyMode = False
videoInit = False
audioInit = False
timestampInit = False
End Sub

Public Sub ExtractFLV()
Dim Progress As Long
Dim count As Long

If flvHandle = INVALID_HANDLE_VALUE Then Exit Sub
' Read PreviousTagSize0 which is 4 bytes, always 0, so we don't even check
ReadFile flvHandle, smallBuffer(0), 4, bRead, ByVal 0&
flvOffset = flvOffset + 4
While flvOffset < flvLength
 ' Read FLVOBJECT structure, read all before Data in SmallBuffer, Data in Buffer()
 'TagType    UI8    Type of this tag. Values are: 8: audio 9: video 18: script Data All others: reserved
 'DataSize   UI24   Length of the data in the Data field
 'Timestamp  UI24   Time in milliseconds at which the data in this tag applies. This is relative to the first
 '                  tag in the FLV file, which always has a timestamp of 0.
 'TimestExt  UI8    Extension of the Timestamp field to form a UI32 value. This field represents the upper
 '                  8bits, while the previous Timestamp field represents the lower 24bits of the time in milliseconds.
 'StreamID   UI24   Always 0
 'Data       If TagType = 8     AUDIODATA           Body of the tag
 '           If TagType = 9     VIDEODATA
 '           If TagType = 18    SCRIPTDATAOBJECT
 
 ' Read 11 bytes, up until data
 ReadFile flvHandle, smallBuffer(0), 11, bRead, ByVal 0&
 tagType = smallBuffer(0)
 ' Determine data size
 
 tagDataSize = smallBuffer(1) * 65536 + _
               smallBuffer(2) * 256 + _
               smallBuffer(3)
 ' increase buffer size if required, for holding data
 If BufferLen < tagDataSize Then
  ReDim Buffer(0 To tagDataSize - 1)
  BufferLen = tagDataSize
 End If
 tagTimeStamp = CLng(smallBuffer(7) * 16777216)
 tagTimeStamp = tagTimeStamp + smallBuffer(4) * CLng(65536)
 tagTimeStamp = tagTimeStamp + smallBuffer(5) * CLng(256)
 tagTimeStamp = tagTimeStamp + smallBuffer(6)
 tagStreamID = 0 ' always, according to specs
 flvOffset = flvOffset + 11
 ' Read Data, tagDataSize bytes
 ReadFile flvHandle, Buffer(0), tagDataSize, bRead, ByVal 0&
 flvOffset = flvOffset + tagDataSize
 ' Read PreviousTagSizeN, not interested in the actual value
 ReadFile flvHandle, smallBuffer(0), 4, bRead, ByVal 0&
 If tagType = 8 Or tagType = 9 Or tagType = 18 Then
  flvTagsCount = flvTagsCount + 1
  
  Select Case tagType
  Case 8:  ' Audio Tag
          flvTagsAudio = flvTagsAudio + 1
          ProcessAudioTag
  Case 9:  ' video tag
          flvTagsVideo = flvTagsVideo + 1
          ProcessVideoTag
  Case 18: ' script tag
          flvTagsScript = flvTagsScript + 1
           ' script tags are not interesting for us at this time, we ignore them
  End Select
  'Timestamps
  If timestampInit = False Then
   timestampDummyMode = IIf(frmMain.chkTimeStamps.value = vbChecked, False, True)
   OpenTS
   timestampInit = True
  End If
  If tagType = 9 Then
   WriteTS tagTimeStamp
   VideoTimeStamps.Add tagTimeStamp
  End If
  
  
  count = count + 1
  If count > 5 Then ' refresh statistics, progress bar and other stuff
   count = 0
   frmMain.lblTagTotal.Caption = flvTagsCount
   frmMain.lblTagVideo.Caption = flvTagsVideo
   frmMain.lblTagAudio.Caption = flvTagsAudio
   frmMain.lblTagScript.Caption = flvTagsScript
   frmMain.lblVideoFrames.Caption = VideoFrameRateCaption
   frmMain.lblAudioFormat.Caption = AudioFormatCaption

   frmMain.lblProgressText.Caption = " Reading... " & CStr(Round(flvOffset / 1024, 0)) & "KB"
   Progress = Round(flvOffset * 100 / flvLength, 0)
   If Progress > 100 Then Progress = 100
   PBarSetPos 1, Progress
  End If
  NewDoEvents
  
 End If
Wend
CloseMP3
closeTS
WriteIndexAVI
FinalizeAVI
CloseAVI
'Show final data in labels, reset progress bar and progress text label
frmMain.lblTagTotal.Caption = flvTagsCount
frmMain.lblTagVideo.Caption = flvTagsVideo
frmMain.lblTagAudio.Caption = flvTagsAudio
frmMain.lblTagScript.Caption = flvTagsScript
frmMain.lblVideoFrames.Caption = VideoFrameRateCaption
frmMain.lblAudioFormat.Caption = AudioFormatCaption

PBarSetPos 1, 0
frmMain.lblProgressText.Caption = "Ready."
 
End Sub

Public Sub ProcessAudioTag()
Dim sl As Long      ' first 4 bits from a byte
Dim sr As Long      ' last 4 bits from a byte
Dim ssl As Long     ' first 2 bits from sr
Dim ssr As Long     ' last 2 bits from sr
Dim temps As String

'AUDIODATA Tag
'Field          Type            Comment
'SoundFormat    UB[4]           0 = uncompressed
'                               1 = ADPCM
'                               2 = MP3
'                               5 = Nellymoser 8kHz mono
'                               6 = Nellymoser
'SoundRate      UB[2]           0 = 5.5 kHz
'                               1 = 11 kHz
'                               2 = 22 kHz
'                               3 = 44 kHz
'SoundSize      UB[1]           0 = snd8Bit
'                               1 = snd16Bit
'SoundType      UB[1]           0 = sndMono         For Nellymoser: always 0
'                               1 = sndStereo
'SoundData      UI8[size]       Sound data—varies by format

' holder for audio processing code
If audioInit = False Then ' let's initialize audio system
 'switch to dummy mode if checkbox not checked on the window
 audioDummyMode = IIf(frmMain.chkAudio.value = vbChecked, False, True)
 'Split the first byte in the audio data into 4, 2 and 2 bits in order to get
 'information about the audio
 sl = Buffer(0) \ 16        ' first 4 bits
 sr = Buffer(0) - sl * 16   ' last 4 bits
 ssl = sr \ 4               ' first 2 bits from sr
 ssr = sr - ssl * 4         ' last 2 bits
 
 tagAudioFormat = sl
 tagAudioRate = ssl
 tagAudioBits = IIf(ssr = 2, 1, 0)
 tagAudioChannels = IIf(ssr = 2, 1, 0)
 Select Case tagAudioFormat
  Case 0: frmMain.lblAudioCodec.Caption = "Uncompressed [RAW]"
  Case 1: frmMain.lblAudioCodec.Caption = "ADPCM"
  Case 2: frmMain.lblAudioCodec.Caption = "MP3"
  Case 5: frmMain.lblAudioCodec.Caption = "Nellymoser 8kHz mono"
  Case 6: frmMain.lblAudioCodec.Caption = "Nellymoser"
  Case Else: frmMain.lblAudioCodec.Caption = "Unknown [" & CStr(tagAudioFormat) & "]"
 End Select
 temps = ""
 Select Case tagAudioRate
  Case 0: temps = "5.5 kHz"
  Case 1: temps = "11.025 Hz"
  Case 2: temps = "22.050 Hz"
  Case 3: temps = "44.100 Hz"
  Case Else: temps = "Unknown kHz"
 End Select
 Select Case tagAudioBits
  Case 0: temps = temps & ", 8 Bit"
  Case 1: temps = temps & ", 16 Bit"
 End Select
 Select Case tagAudioChannels
  Case 0: temps = temps & ", Mono"
  Case 1: temps = temps & ", Stereo"
 End Select
 AudioFormatCaption = temps
 If tagAudioFormat <> 2 And audioDummyMode = False Then
  MsgBox STR_AUDIO_INVALIDFORMAT, vbOKOnly Or vbInformation, STR_ERROR
  audioDummyMode = True
 End If
 If OpenMP3() = False Then
 ' will not be able to extract the audio track
 End If
audioInit = True
End If
If audioDummyMode Then Exit Sub
WriteAudioChunk
End Sub

Public Sub ProcessVideoTag()
Dim bitString As String
Dim cutWidth As Long
Dim cutHeight As Long
If videoInit = False Then
 videoDummyMode = IIf(frmMain.chkVideo.value = vbChecked, False, True)
 If videoDummyMode Then ' no point going into details if user does not want video
  videoInit = True
  Exit Sub
 End If
 tagVideoFrameType = Buffer(0) \ 16
 tagVideoCodec = Buffer(0) - tagVideoFrameType * 16
 Select Case tagVideoCodec
  Case 2: frmMain.lblVideoCodec.Caption = "Sorenson H.263"
          VideoFourCC = "FLV1"
  Case 3: frmMain.lblVideoCodec.Caption = "Screen Video"
          MsgBox "The video format of this FLV file is not supported. Switching to Dummy mode (skip extracting video)", vbOKOnly Or vbExclamation, "Warning"
          videoDummyMode = True
  Case 4: frmMain.lblVideoCodec.Caption = "On2 VP6"
          VideoFourCC = "FLV4"
  Case 5: frmMain.lblVideoCodec.Caption = "On2 VP6 with alpha channel"
          VideoFourCC = "FLV4"
  Case 6: frmMain.lblVideoCodec.Caption = "Screen Video version 2"
          MsgBox "The video format of this FLV file is not supported. Switching to Dummy mode (skip extracting video)", vbOKOnly Or vbExclamation, "Warning"
          videoDummyMode = True
 End Select
 ' we need to get the video resolution for the AVI header
 ' Debug.Print tagVideoPictureSize, Buffer(0), Buffer(1), Buffer(2), Buffer(3), Buffer(4), Buffer(5), Buffer(6), Buffer(7), Buffer(8), Buffer(9)
 If tagVideoCodec = 2 Then
 'H263VIDEOPACKET
 '
 'PictureStartCode     UB[17]   Similar to H.263 5.1.1 0000 0000 0000 0000 1
 'Version              UB[5]    Video format version Flash Player 6 supports 0 and 1
 'TemporalReference    UB[8]    See H.263 5.1.2
 'PictureSize       UB[3]   000: custom, 1 byte
 '                          001: custom, 2 bytes
 '                          010: CIF (352x288)
 '                          011: QCIF (176x144)
 '                          100: SQCIF (128x96)
 '                          101: 320x240
 '                          110: 160x120
 '                          111: reserved
 'CustomWidth       If PictureSize = 000 UB[8]
 '                  If PictureSize = 001 UB[16] Otherwise absent
 '                  Note: UB[16] is not the same as UI16; there is no byte swapping.
 'CustomHeight      If PictureSize = 000 UB[8]
 '                  If PictureSize = 001 UB[16] Otherwise absent
 '                  Note: UB[16] is not the same as UI16; there is no byte swapping.
 ' So we need the first 3 bits and up to 32 bits (4 bytes) depending on the choice.
 ' No point aiming for speed for an operation that is executed once per run, just
 ' get 9 bytes and convert to string
 
 bitString = convert_base2(Buffer(1)) & _
             convert_base2(Buffer(2)) & _
             convert_base2(Buffer(3)) & _
             convert_base2(Buffer(4)) & _
             convert_base2(Buffer(5)) & _
             convert_base2(Buffer(6)) & _
             convert_base2(Buffer(7)) & _
             convert_base2(Buffer(8)) & _
             convert_base2(Buffer(9))
             
 tagVideoPictureSize = convert_base10(Mid(bitString, 31, 3))
 Select Case tagVideoPictureSize
  Case 0: 'custom, 1 byte for width, 1 for height
  Case 1: 'custom, 2 bytes for width, 2 for height
  Case 2: 'CIF (352x288)
          VideoWidth = 352
          VideoHeight = 288
  Case 3: 'QCIF (176x144)
          VideoWidth = 176
          VideoHeight = 144
  Case 4: 'SQCIF (128x96)
          VideoWidth = 128
          VideoHeight = 96
  Case 5: ' 320 x 240
          VideoWidth = 320
          VideoHeight = 240
  Case 6: ' 160x120
          VideoWidth = 160
          VideoHeight = 120
 End Select
 If tagVideoPictureSize = 0 Then '1 byte for each
  VideoWidth = convert_base10(Mid(bitString, 34, 8))
  VideoHeight = convert_base10(Mid(bitString, 42, 8))
 End If
 If tagVideoPictureSize = 1 Then '2 byte for each
  VideoWidth = convert_base10(Mid(bitString, 34, 8)) * 256 + convert_base10(Mid(bitString, 42, 8))
  VideoHeight = convert_base10(Mid(bitString, 50, 8)) * 256 + convert_base10(Mid(bitString, 58, 8))
 End If
     
 End If
 
If tagVideoCodec = 4 Then 'VP6
  VideoWidth = Buffer(5) * 16
  VideoHeight = Buffer(6) * 16
  cutWidth = Buffer(1) \ 16
  cutHeight = Buffer(1) - cutWidth * 16
  'don't use the cutX variables for FLV4
  'VideoWidth = VideoWidth - cutWidth
  'VideoHeight = VideoHeight - cutHeight
End If

If tagVideoCodec = 5 Then  'VP6 with optional alpha channel
  VideoWidth = Buffer(8) * 16
  VideoHeight = Buffer(9) * 16
  cutWidth = Buffer(2) \ 16
  cutHeight = Buffer(2) - cutWidth * 16
  'don't use the cutX variables for FLV4
  'VideoWidth = VideoWidth - cutWidth
  'VideoHeight = VideoHeight - cutHeight
End If
frmMain.lblVideoFormat.Caption = CStr(VideoWidth + cutWidth) & "x" & CStr(VideoHeight + cutHeight)
If cutWidth <> 0 Or cutHeight <> 0 Then frmMain.lblVideoFormat.Caption = frmMain.lblVideoFormat.Caption & " (crop to " & CStr(VideoWidth) & "x" & CStr(VideoHeight) & ")"
OpenAVI
WriteHeaderAVI
videoInit = True
End If
If videoDummyMode Then Exit Sub
tagVideoFrameType = Buffer(0) \ 16
tagVideoCodec = Buffer(0) - tagVideoFrameType * 16
Select Case tagVideoFrameType
Case 1: flvTagVideoK = flvTagVideoK + 1              ' keyframe
Case 2: flvTagVideoI = flvTagVideoI + 1              ' interframe
Case 3: flvTagVideoD = flvTagVideoD + 1              ' disposable intraframe (h263 only)
End Select
WriteVideoChunk
VideoFrameRateCaption = CStr(flvTagVideoK) & " K, " & CStr(flvTagVideoI) & " I" & IIf(tagVideoCodec = 2, ", " & CStr(flvTagVideoD) & " DI", "")

End Sub

'Converts a number to a Base2 representation of it (ex. 103=01100111)
Private Function convert_base2(thebyte As Byte) As String
Dim s As String
Dim N As Byte
Dim r As Byte
s = ""
N = thebyte
While N > 0
r = N Mod 2
s = CStr(r) & s
N = N \ 2
Wend
If Len(s) < 8 Then s = String(8 - Len(s), "0") & s
convert_base2 = s
End Function

'Converts a Base2 string into a normal number
Private Function convert_base10(bitString As String) As Long
Dim i As Long
Dim value As Long
Dim slen As Long
slen = Len(bitString)
If slen = 0 Then
 convert_base10 = 0
 Exit Function
End If
value = 0
For i = slen To 1 Step -1
 value = value + IIf(Mid(bitString, i, 1) = "1", 1, 0) * 2 ^ (slen - i)
Next i
convert_base10 = value
End Function

Private Sub SeekFLV(offset As Long)
If flvHandle = INVALID_HANDLE_VALUE Then Exit Sub
SetFilePointer flvHandle, offset, 0, FILE_BEGIN
flvOffset = offset
End Sub

'Private Sub ReverseBuffer(length As Long, Optional offset As Long = 0)
'Dim X As Byte
'Select Case length
'Case 1: Exit Sub
'Case 2: X = smallBuffer(offset + 0): smallBuffer(offset + 0) = smallBuffer(offset + 1): smallBuffer(offset + 1) = X
'Case 3: X = smallBuffer(offset + 0): smallBuffer(offset + 0) = smallBuffer(offset + 2): smallBuffer(offset + 2) = X
'Case 4: X = smallBuffer(offset + 0): smallBuffer(offset + 0) = smallBuffer(offset + 3): smallBuffer(offset + 3) = X
'        X = smallBuffer(offset + 1): smallBuffer(offset + 1) = smallBuffer(offset + 2): smallBuffer(offset + 2) = X
'Case 8: X = smallBuffer(offset + 0): smallBuffer(offset + 0) = smallBuffer(offset + 7): smallBuffer(offset + 7) = X
'        X = smallBuffer(offset + 1): smallBuffer(offset + 1) = smallBuffer(offset + 6): smallBuffer(offset + 6) = X
'        X = smallBuffer(offset + 2): smallBuffer(offset + 2) = smallBuffer(offset + 5): smallBuffer(offset + 5) = X
'        X = smallBuffer(offset + 3): smallBuffer(offset + 3) = smallBuffer(offset + 4): smallBuffer(offset + 4) = X
'End Select
'End Sub


