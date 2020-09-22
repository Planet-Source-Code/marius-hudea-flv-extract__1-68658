VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLV Extract"
   ClientHeight    =   3645
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9645
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   " FLV Information "
      Height          =   3615
      Left            =   5640
      TabIndex        =   15
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox chkFlagAudio 
         Caption         =   "Has Audio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkFlagVideo 
         Caption         =   "Has Video"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblAudioFormat 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label lblAudioCodec 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label lblVideoFormat 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   36
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label lblVideoFrames 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   35
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label lblVideoCodec 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   34
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label lblTagScript 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   33
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblTagAudio 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblTagVideo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblTagTotal 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label12 
         Caption         =   "Format:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Codec:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Audio:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "Format:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Frames:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Codec:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Video:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label16 
         Caption         =   "Script:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Audio:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Video:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Tags:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Flags: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox chkTimeStamps 
         Caption         =   "Timestamps"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkAudio 
         Caption         =   "Audio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkVideo 
         Caption         =   "Video"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txtInputFile 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   0
         Top             =   405
         Width           =   975
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   1
         Top             =   1095
         Width           =   975
      End
      Begin VB.TextBox txtFolder 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1170
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Parts to extract:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Input file:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label4 
         Caption         =   "Save selected parts in:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   5295
      End
   End
   Begin VB.Frame fraProgress 
      Caption         =   " Progress "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   2640
      Width           =   5535
      Begin VB.Label lblProgressText 
         Caption         =   "Ready."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label lblProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   570
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
Dim retb As Boolean

If Len(txtInputFile.Text) = 0 Then
 MsgBox STR_NOINPUTFLV, vbOKOnly Or vbCritical, STR_ERROR: Exit Sub
End If
If Len(txtFolder.Text) = 0 Then
 MsgBox STR_NOOUTPUTFOLDER, vbOKOnly Or vbCritical, STR_ERROR: Exit Sub
End If
If chkVideo.Value = vbUnchecked And _
   chkAudio.Value = vbUnchecked And _
   chkTimeStamps.Value = vbUnchecked Then
 MsgBox STR_NOSELECTION, vbOKOnly Or vbCritical, STR_ERROR: Exit Sub
End If
' Reset all the checkboxes and labels on the screen
ClearInterface
' Try to open the FLV file
retb = OpenFLV(txtInputFile.Text)
If retb = False Then
 MsgBox STR_FLVNOTOPEN, vbOKOnly Or vbCritical, STR_ERROR: Exit Sub
End If
mnuClose.Enabled = False
cmdStart.Enabled = False
InitializeMP3
InitializeAVI
InitializeTS
ExtractFLV
mnuClose.Enabled = True
cmdStart.Enabled = True
CloseFLV
End Sub

'Clears the labels, checkboxes and resets the progress bar
Private Sub ClearInterface()
PBarSetPos 1, 0
chkFlagVideo.Value = vbUnchecked
chkFlagAudio.Value = vbUnchecked
lblTagTotal.Caption = ""
lblTagVideo.Caption = ""
lblTagAudio.Caption = ""
lblTagScript.Caption = ""
lblVideoCodec.Caption = ""
lblVideoFormat.Caption = ""
lblVideoFrames.Caption = ""
lblAudioCodec.Caption = ""
lblAudioFormat.Caption = ""
End Sub

Private Sub Form_Load()
' Set the default values and initialize the progress bar
PBarLoad 1, fraProgress.hWnd, lblProgress.Left \ Screen.TwipsPerPixelX, lblProgress.Top \ Screen.TwipsPerPixelY, lblProgress.Width \ Screen.TwipsPerPixelX, lblProgress.Height \ Screen.TwipsPerPixelY
PBarSetRange 1, 0, 100
PBarSetPos 1, 0
ClearInterface
End Sub
' Shows the Open Folder window
Private Sub cmdBrowse_Click()
Dim s As String
s = OpenFolder()
If Len(s) <> 0 Then txtFolder.Text = s
End Sub

' Shows the Open dialog in order to select a file
Private Sub cmdSelect_Click()
Dim s As String
s = OpenFile()
If Len(s) <> 0 Then
 txtInputFile.Text = s: txtFolder.Text = ExtractFolder(s)
End If
End Sub

Private Sub mnuAbout_Click()
MsgBox STR_ABOUT_TEXT, vbOKOnly Or vbInformation, STR_ABOUT_CAPTION
End Sub

Private Sub mnuClose_Click()
End
End Sub
