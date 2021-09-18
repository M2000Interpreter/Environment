VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "safeforms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' do something
Public mylist As New LongHash

Private Sub Class_Terminate()
Dim i As Long, j As Long, start As Long
j = Forms.count
While j > 0
For i = start To Forms.count - 1
If mylist.ExistKey(Forms(i).hWnd) Then
If mylist.Value = "" Then
Forms(i).icon = LoadPicture("")
Unload Forms(i): start = i: Exit For
End If
End If
Next i
j = j - 1
Wend
End Sub
