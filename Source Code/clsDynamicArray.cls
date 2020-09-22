VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDynamicArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Records() As Long
Private RecordMax As Long
Private RecordCnt As Long

Private Sub Class_Initialize()
RecordCnt = -1
RecordMax = 8192
ReDim Records(0 To RecordMax - 1)
End Sub

Public Sub Add(Value As Long)
RecordCnt = RecordCnt + 1
If RecordCnt = RecordMax Then
 RecordMax = RecordMax + 8192
 ReDim Preserve Records(0 To RecordMax - 1)
End If
Records(RecordCnt) = Value
End Sub

Public Function GetValue(ListIndex As Long) As Long
If ListIndex > RecordCnt Then
 GetValue = -1
 Exit Function
End If
GetValue = Records(ListIndex)
End Function

Public Function GetCount() As Long
GetCount = RecordCnt
End Function
