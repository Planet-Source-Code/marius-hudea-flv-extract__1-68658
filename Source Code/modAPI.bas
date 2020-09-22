Attribute VB_Name = "modAPI"
Option Explicit
' Constants used by File I/O API functions
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const INVALID_HANDLE_VALUE = -1

' Constants used by SetFilePointer
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1

' Constants used by Open File dialog and Open Folder dialog
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_HIDEREADONLY = &H4
Private Const OF_SHARE_DENY_WRITE = &H20


' Constants used by GetQueueStatus API function
Private Const QS_HOTKEY = &H80
Private Const QS_KEY = &H1
Private Const QS_MOUSEBUTTON = &H4
Private Const QS_MOUSEMOVE = &H2
Private Const QS_PAINT = &H20
Private Const QS_POSTMESSAGE = &H8
Private Const QS_SENDMESSAGE = &H40
Private Const QS_TIMER = &H10
Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Const QS_MESSAGES = (QS_POSTMESSAGE Or QS_SENDMESSAGE)                      ' Not MS standard constant
Private Const QS_STANDARD = (QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT)     ' Not MS standard constant

'Used by Open File function
Private Type OpenFileName
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' Used by Select Folder function
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

' Enumerator to determine what messages are watched
Private Enum QueueMessagesUsed
    All_Inputs = QS_ALLINPUT
    All_Events = QS_ALLEVENTS
    Standard = QS_STANDARD
    Messages = QS_MESSAGES
    InputOnly = QS_INPUT
    Mouse = QS_MOUSE
    MouseMove = QS_MOUSEMOVE
    Timer = QS_TIMER
End Enum

' File I/O API functions
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
' DoEvents replacement
Private Declare Function GetQueueStatus Lib "User32" (ByVal fuFlags As Long) As Long

' API function used by the Select Folder window
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
' Open File dialog
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
' Self explanitory
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Public Function ExtractFolder(FileName As String) As String
Dim pos As Long
Dim pos2 As Long
Dim s As String

pos = 0
pos2 = InStr(1, FileName, "\")
While pos2 > 0
 pos = pos2: pos2 = InStr(pos2 + 1, FileName, "\")
Wend
If pos = 0 Then
  ExtractFolder = ""
Else
  ExtractFolder = Left(FileName, pos)
End If
End Function

Public Function OpenFile() As String
Dim s As String
Dim i As Long
Dim OFName As OpenFileName
Dim iNull As Long
OFName.lStructSize = Len(OFName) 'Set the length of the structure
OFName.hWndOwner = frmMain.hWnd  'Set the parent window
OFName.hInstance = App.hInstance 'Set the application's instance
' Select a filter
OFName.lpstrFilter = "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
'
OFName.lpstrFile = Space$(254) 'create a buffer for the file
OFName.nMaxFile = 255 'set the maximum length of a returned file
OFName.lpstrFileTitle = Space$(254) 'Create a buffer for the file title
OFName.nMaxFileTitle = 255 'Set the maximum length of a returned file title

'If FolderPath = "" Then FolderPath = "C:\"

OFName.lpstrInitialDir = "" 'Set the initial directory
'If frmMain.txtFolder.Text <> "" Then OFName.lpstrInitialDir = frmMain.txtFolder.Text
OFName.lpstrTitle = "Select file" 'Set the title
'Next we set the flags that will modify the way our window looks
OFName.flags = OFN_NONETWORKBUTTON Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
'Show the 'Open File' dialog
If GetOpenFileName(OFName) Then ' Everything is Ok
   iNull = InStr(OFName.lpstrFile, vbNullChar)
    If iNull > 0 Then
       OpenFile = Left$(OFName.lpstrFile, iNull - 1)
       Exit Function
    End If
   OpenFile = Trim$(OFName.lpstrFile)
Else 'There was an error or the user pressed Cancel
OpenFile = ""
End If
End Function

Public Function OpenFolder() As String
Dim iNull As Integer
Dim lpIDList As Long
Dim lResult As Long
Dim sPath As String
Dim BInfo As BrowseInfo

With BInfo
 'Set the owner window
 .hWndOwner = frmMain.hWnd
 'lstrcat appends the two strings and returns the memory address
 .lpszTitle = lstrcat("Please select a folder:", "")
 'Return only if the user selected a directory
 .ulFlags = BIF_RETURNONLYFSDIRS
 End With
 'Show the 'Browse for folder' dialog
 lpIDList = SHBrowseForFolder(BInfo)
 If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    'Get the path from the IDList
    SHGetPathFromIDList lpIDList, sPath
    'free the block of memory
    CoTaskMemFree lpIDList
    iNull = InStr(sPath, vbNullChar)
    If iNull Then
       sPath = Left$(sPath, iNull - 1)
    End If
Else
sPath = ""
End If
If Len(sPath) <> 0 Then sPath = sPath & IIf(Right(sPath, 1) = "\", "", "\")
OpenFolder = sPath
End Function

'
' DoEvents replacement, well not really a replacement but triggers DoEvents only for
'certain events
'
Public Sub NewDoEvents()
Dim m_lQueueUsed As QueueMessagesUsed
m_lQueueUsed = Standard + Messages
If GetQueueStatus(m_lQueueUsed) <> 0 Then DoEvents

End Sub

         

