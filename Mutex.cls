VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mutex"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Mymutexs As New FastCollection
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const MUTEX_ALL_ACCESS = &H1F0001
Public LastError As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexW" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As Long) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Function create(Name$) As Long
If Mymutexs.ExistKey(Name$) Then exist: Exit Function
Dim myMutex As Variant
myMutex = CVar(CreateMutex(0, 1, StrPtr(Name$)))
LastError = Err.LastDllError
If LastError = 0 Then Mymutexs.AddKey2 Name$, myMutex: Exit Function
If LastError = ERROR_ALREADY_EXISTS Then CloseHandle myMutex

MyEr "Mutex object exist", "Υπάρχει ήδη το αντικείμενο"
create = LastError
End Function

Private Sub notexist()
MyEr "named mutex not exist", "το αντικείμενο δεν υπάρχει"
End Sub

Private Sub exist()
MyEr "named mutex exist", "το αντικείμενο υπάρχει"
End Sub
Sub Destroy(Name$)
If Not Mymutexs.ExistKey(Name$) Then notexist: Exit Sub
ReleaseMutex CLng(Mymutexs.Value)
CloseHandle CLng(Mymutexs.Value)
Mymutexs.Remove Name$
End Sub

Private Sub Class_Terminate()
Dim i As Long, myMutex As Long
If Mymutexs.count > 0 Then
For i = 0 To Mymutexs.count - 1
Mymutexs.Index = i
ReleaseMutex CLng(Mymutexs.Value)
CloseHandle CLng(Mymutexs.Value)
Next i
End If
End Sub
