Attribute VB_Name = "modStrings"
Public Const STR_NOINPUTFLV As String = "Please choose a FLV file first!"
Public Const STR_NOOUTPUTFOLDER As String = "Please choose a destination folder first!"
Public Const STR_NOSELECTION As String = "You must choose at least one of the parts needed to extract."
Public Const STR_FLVNOTOPEN As String = "There was a problem opening the FLV file. Process aborted."

Public Const STR_WARNING As String = "Warning"
Public Const STR_ERROR As String = "Error"

Public Const STR_ABOUT_TEXT As String = "FLV Extract version 1.0" & vbCrLf & vbCrLf & "(c)2007 Marius Hudea" & vbCrLf & vbCrLf & _
                                        "This program is free software and may be distributed" & vbCrLf & _
                                        "according to the terms listed on the dedicated web page." & vbCrLf & vbCrLf & _
                                        "More information, the complete source code and other" & vbCrLf & _
                                        "utilities can be found in the programming section of" & vbCrLf & _
                                        "Helpedia - http://www.helpedia.com"

Public Const STR_ABOUT_CAPTION As String = "About FLV Extract"

Public Const STR_INVALIDVERSION As String = "The FLV version of this file is not supported. Aborted."
Public Const STR_AUDIO_INVALIDFORMAT As String = "The audio track is not MP3. This version of the application can only extract MP3 data." & vbCrLf & "FLV Extract is switching audio to Dummy Mode (skips audio extract process)."

Public Const MSG_CONFIRM_OVERWRITE As String = "There was a problem creating the %FORMAT% file %FILE%." & vbCrLf & vbCrLf & _
                                               "Click Yes to overwrite the existing file or No to switch to Dummy mode  (skip extracting the %FORMAT% stream)"
               

Public Const MSG_NOTOPEN = "The %FORMAT% file %FILE% can not be created or opened in Write mode." & vbCrLf & _
                           "The application will switch to Dummy mode for the %FORMAT% section."
                           

