VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' The TaskMaster handles queing and timing of the
' various tasks.
Dim mymes As Long
Public tickdrop As Long
Dim holdtasks As Boolean
Public Enum DisposeAction
  tmThisTask
  tmAllTasks
End Enum

Public Enum PriorityLevel
  tmLow
  tmHigh
End Enum


Public MusicTaskNum As Long
Public OnlyMusic As Boolean
Private myRunningList As Collection
Private myMasterList As Collection
Private myInterval As Long
Private myTimerHandle As Long
Private myTimerIsRunning As Boolean



' PUBLIC PROPERTIES
Public Property Get PlayMusic() As Boolean
' check if music is there

PlayMusic = MusicTaskNum > 0
Exit Sub

Dim i As Long
rest
       If myMasterList.count > 0 Then
       For i = 1 To myMasterList.count
       If myMasterList(i).id < 100 Then
 PlayMusic = True
GoTo thh11
       End If
       
       Next i
        
        End If
       If myRunningList.count > 0 Then
       For i = 1 To myRunningList.count
       If myRunningList(i).id < 100 Then
        PlayMusic = True
 GoTo thh11
     '  End
       End If
       
       Next i
        
        End If
thh11:
        RestEnd
 
End Property
Public Property Get QueueCount() As Long
  QueueCount = myRunningList.count + myMasterList.count
End Property


Public Property Get Interval() As Long
  Interval = myInterval
End Property


Public Property Let Interval(ByVal Value As Long)
  myInterval = Abs(Value)
End Property


Public Property Get Processing() As Boolean
  Processing = myTimerIsRunning
End Property

Public Function ThrowOne(thisid As Long) As Boolean
Dim i As Long
rest
       If myMasterList.count > 0 Then
       For i = 1 To myMasterList.count
       If myMasterList(i).id = thisid Then
       myMasterList.Remove i
       ThrowOne = True
GoTo thh
       End If
       
       Next i
        
        End If
       If myRunningList.count > 0 Then
       For i = 1 To myRunningList.count
       If myRunningList(i).id = thisid Then
       
       myRunningList(i).Done = True
       myRunningList.Remove i
       ThrowOne = True
 GoTo thh
     '  End
       End If
       
       Next i
        
        End If
        ' SECONDTIME
               If myMasterList.count > 0 Then
       For i = 1 To myMasterList.count
       If myMasterList(i).id = thisid Then
       myMasterList.Remove i
       ThrowOne = True
GoTo thh
       End If
       
       Next i
        
        End If
       If myRunningList.count > 0 Then
       For i = 1 To myRunningList.count
       If myRunningList(i).id = thisid Then
       myRunningList(i).Done = True
       myRunningList.Remove i
       ThrowOne = True
 GoTo thh
     '  End
       End If
       
       Next i
        
        End If
       
thh:
        RestEnd
        If Me.QueueCount <= 0 Then
        StopProcess
        End If
End Function
Private Function Remove(thisid As Long) As Boolean
Dim i As Long
rest

       If myRunningList.count > 0 Then
       For i = 1 To myRunningList.count
       If myRunningList(i).id = thisid Then
       myRunningList.Remove i
       Remove = True
       RestEnd
       Exit Function
       
       End If
       
       Next i
        
        End If
        RestEnd
End Function

' PUBLIC METHODS

Public Sub AddTask(ByRef NewTask As TaskInterface, Optional Priority As PriorityLevel)
  ' Add to list
 
  If Priority = tmHigh Then
    If myMasterList.count > 0 Then
    
    myMasterList.Add NewTask, , 1
    Else
    myMasterList.Add NewTask
    End If
  Else
        If myMasterList.count > 0 Then
    
    myMasterList.Add NewTask, , , myMasterList.count
    Else
    myMasterList.Add NewTask
    End If

  End If

    StartTimer

End Sub
Private Property Get getone(thisid As Long) As TaskInterface
Dim i As Long
rest
       If myMasterList.count > 0 Then
       For i = 1 To myMasterList.count
       If myMasterList(i).id = thisid Then
      Set getone = myMasterList(i)
GoTo thh
       End If
       
       Next i
        
        End If
       If myRunningList.count > 0 Then
       For i = 1 To myRunningList.count
       If myRunningList(i).id = thisid Then
        Set getone = myRunningList(i)
 GoTo thh
      
       End If
       
       Next i
        
        End If
thh:
        RestEnd
End Property
Public Sub Message(Mes As Long, code As Long, Optional param As Long = 0, Optional Exec$)
Dim ok As Boolean, bs As basetask, m As mStiva
Dim task As TaskInterface
Set task = getone(Mes)
ok = False
If Not (task Is Nothing) Then
        Select Case code
        Case 0 ' hold
            task.Busy = True
        Case 1
            task.Dispose tmAllTasks
            ThrowOne Mes
        Case 2 ' Restart
            task.Busy = False
        Case 3
        task.Interval = param
        Case 4
        Set bs = GetNewBaseTask(task)
        Set bs.Process = task
        If Execute(bs, Exec$, True) = 0 Then
            task.Dispose tmAllTasks
            ThrowOne Mes
        ElseIf Not task.Process Is Nothing Then
         If task.Process.StaticCollection Is Nothing Then
            Set task.Process.StaticCollection = bs.StaticCollection
         End If
         Set m = task.Process.soros
         m.MergeBottom bs.soros
        End If

       
        Case 5
        Set bs = GetNewBaseTask(task)
        Set bs.Process = task
        If Execute(bs, Exec$, False) = 0 Then
            task.Dispose tmAllTasks
            ThrowOne Mes
         ElseIf Not task.Process Is Nothing Then
         If task.Process.StaticCollection Is Nothing Then
            Set task.Process.StaticCollection = bs.StaticCollection
         End If
         Set m = task.Process.soros
         m.MergeBottom bs.soros
          
        End If
        End Select
        

        If ok Then code = -1
End If
End Sub

Private Function GetNewBaseTask(task As TaskInterface)
Dim b As New basetask, bstack As basetask
Set bstack = task.Process
Set b.Process = Nothing
Set b.Owner = task.Owner
'Set b.Parent = b
Set b.Sorosref = bstack.soros
'If Len(bstack.StaticInUse) > 0 Then
'b.StaticInUse = bstack.StaticInUse
'End If
Set b.StaticCollection = bstack.StaticCollection
b.OriginalCode = bstack.OriginalCode
b.originalname = bstack.originalname

Set GetNewBaseTask = b
End Function



Public Sub Dispose()
  On Error Resume Next
        StopTimer
        Do While myMasterList.count > 0
          myMasterList(1).Dispose tmAllTasks
          myMasterList.Remove 1
        Loop
            Do While myRunningList.count > 0
        
          myRunningList(1).Dispose tmAllTasks
          myRunningList.Remove 1
        Loop
   
End Sub

Public Sub TimerTick()
    tickTimer = tickTimer + 1
    If tickdrop > 0 Then Exit Sub
    Dim this As TaskInterface, many As Long, dosomething As Boolean
    If myMasterList.count = 0 Then Exit Sub
    many = myMasterList.count
    Dim manybusy As Long
    
again:
    Do
        Set this = myMasterList(1)
        myMasterList.Remove 1
        If OnlyMusic Then
            If this.id > 100 Then
                AddTask this
                Exit Sub
            End If
        ElseIf holdtasks Then
            AddTask this, this.Priority
            Exit Sub
        End If
        myRunningList.Add this
        many = many - 1
    Loop While this.id < 0 And myMasterList.count <> 0
    If this.id < 0 Then Exit Sub
    dosomething = False
    If Not this.Busy And Not this.Done Then
      dosomething = this.Tick
    Else
        manybusy = manybusy + 1
    End If
    If Remove(this.id) Then
      If this.Done Then
          this.Dispose tmThisTask
      Else
          If Not NOEXECUTION Then
              AddTask this, this.Priority
              If many > 0 And Not dosomething Then GoTo again
          Else
              this.Dispose tmThisTask
          End If
    End If
  Else
    this.Dispose tmThisTask
  End If

    If myTimerIsRunning = False Or QueueCount = 0 Then
        StopTimer
    ElseIf QueueCount = 1 And taskmainonly Then
        Sleep 50
    ElseIf Not dosomething Then
        Sleep 5
    End If
    If LastErNum <> 0 Then Dispose
    If manybusy > 0 Then
        If myMasterList.count = manybusy Then
           If Not myTimerIsRunning Then Dispose
        End If
    End If
End Sub
Public Sub TimerTickOne()
If tickdrop > 0 Then Exit Sub 'tickdrop = myMasterList.Count: Exit Sub
tickdrop = 1000
Dim this As TaskInterface
  If myMasterList.count = 0 Then tickdrop = 0: Exit Sub
  
  Do
  
  Set this = myMasterList(1)
  myMasterList.Remove 1

  
 If OnlyMusic Then
 If this.id > 100 Then
 AddTask this
 tickdrop = 0
 Exit Sub
 End If
 ElseIf holdtasks Then
 AddTask this, this.Priority
 tickdrop = 0
 Exit Sub
 End If
  myRunningList.Add this
  Loop While this.id < 0 And myMasterList.count <> 0
  If this.id < 0 Then tickdrop = 0: Exit Sub
  
  If Not this.Busy And Not this.Done Then

  this.Tick
  End If
  If Remove(this.id) Then
  If this.Done Then

    this.Dispose tmThisTask
  Else
If Not NOEXECUTION Then
AddTask this, this.Priority
Else
this.Dispose tmThisTask
End If
    
  End If
  Else
    If this.Done Then

  Else

''  AddTask this, this.Priority
    
  End If
    this.Dispose tmThisTask
  End If

  If myTimerIsRunning = False Or QueueCount = 0 Then
    StopTimer

       ElseIf QueueCount = 1 And taskmainonly Then
    Sleep 50
  End If

  If LastErNum <> 0 Then Dispose
End Sub
Public Sub TimerTickNow()
tickTimer = tickTimer + 1
Dim this As TaskInterface
Dim manybusy As Long
If myTimerHandle = 0 Then Exit Sub
  If myMasterList.count = 0 Then Exit Sub
  
  Do
  
  Set this = myMasterList(1)
  myMasterList.Remove 1
 If OnlyMusic Then
 If this.id > 100 Then
 AddTask this
 Exit Sub
 End If
 ElseIf holdtasks Then
 AddTask this, this.Priority
 Exit Sub
 End If
  myRunningList.Add this
  Loop While this.id < 0 And myMasterList.count <> 0
  If this.id < 0 Then Exit Sub
  
  If Not this.Busy And Not this.Done Then

  this.Tick
  Else
  manybusy = manybusy + 1
  End If
  If Remove(this.id) Then
  If this.Done Then

    this.Dispose tmThisTask
  Else
If Not NOEXECUTION Then
AddTask this, this.Priority
Else
this.Dispose tmThisTask
End If
    
  End If
  Else
    If this.Done Then
    this.Dispose tmThisTask
  Else

    AddTask this, this.Priority
    
  End If
    this.Dispose tmThisTask
  End If

  If myTimerIsRunning = False Or QueueCount = 0 Then
    StopTimer

       ElseIf QueueCount = 1 And taskmainonly Then
    Sleep 50
  End If
  If LastErNum <> 0 Then Dispose
  If manybusy > 0 Then
        If myMasterList.count = manybusy Then
          If Not myTimerIsRunning Then Dispose
        End If
    End If
  tickdrop = 0
End Sub

Public Sub StopProcess()
  myTimerIsRunning = False
  StopTimer

End Sub


Public Sub StartProcess()
  StartTimer
End Sub


' PRIVATE ROUTINES

Private Sub Class_Initialize()
tickTimer = 0
  Set myMasterList = New Collection
  Set myRunningList = New Collection

  myInterval = 5

End Sub


Private Sub Class_Terminate()
tickTimer = 0
  If QueueCount > 0 Then
    Dispose
  End If
  Set myMasterList = Nothing
  Set myRunningList = Nothing
End Sub


Private Sub StartTimer()
  If myTimerHandle = 0 Then
  If QueueCount > 0 Then
 ''MasterTimer = 0
  
    myTimerHandle = SetTimer(0&, 0&, myInterval, AddressOf TaskMasterTick)
    myTimerIsRunning = (myTimerHandle <> 0)
  End If
  End If
End Sub


Private Sub StopTimer()
  If myTimerHandle <> 0 Then
  '' MasterTimer = 0
    KillTimer 0&, myTimerHandle
    myTimerHandle = 0
    myTimerIsRunning = False
  End If
End Sub

Public Sub rest()
'
''holdtasks = True
tickdrop = tickdrop + 1
End Sub
Public Sub RestEnd()
tickdrop = tickdrop - 1
If tickdrop <= 0 Then tickdrop = 0: holdtasks = False
''If tickdrop > 0 Then If QueueCount = 0 Then tickdrop = 0:: holdtasks = False
End Sub
Public Sub rest1()
'

tickdrop = tickdrop + 1
End Sub
Public Sub RestEnd1()
tickdrop = 0
End Sub
Public Sub CheckThreadsForThisObject(obj As Object)
Dim this As TaskInterface, i As Long
If myRunningList.count + myMasterList.count = 0 Then Exit Sub
StopTimer
For i = 0 To myRunningList.count - 1
Set this = myRunningList(i + 1)
If this.Owner Is obj Then this.Dispose tmAllTasks
Next i
For i = 0 To myMasterList.count - 1
Set this = myMasterList(i + 1)
If this.Owner Is obj Then this.Dispose tmAllTasks
Next i
StartTimer
End Sub

